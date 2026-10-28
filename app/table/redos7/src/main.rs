//! ENS-Table для РЕД ОС 7 («Муром», glibc 2.28).
//!
//! Настольная версия (Tauri) там не запускается: её окну нужен WebKitGTK 4.1.
//! Здесь окно рисует установленный Chromium (или Яндекс.Браузер) в режиме
//! приложения — без адресной строки и вкладок, со своим профилем, — а эта
//! программа даёт всё остальное, как в настоящей версии: сохранение на место
//! (атомарно), восстановление после сбоя, «Открыть с помощью», окна выбора файлов.
//!
//! Работает только внутри компьютера: сервер слушает 127.0.0.1, каждый запрос
//! к файлам требует секретный ключ, выданный окну при запуске.

#[path = "../../src-tauri/src/files.rs"]
mod files;
#[path = "../../src-tauri/src/recovery.rs"]
#[allow(dead_code)]
mod recovery;

use std::fs;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::process::{Child, Command, Stdio};
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::sync::Arc;
use std::time::{Duration, SystemTime, UNIX_EPOCH};

use serde_json::{json, Value};
use tiny_http::{Header, Method, Response, Server};

const BROWSERS: &[&str] = &["chromium-browser", "chromium", "yandex-browser-corporate", "yandex-browser", "google-chrome"];

fn now_ms() -> u64 {
    SystemTime::now().duration_since(UNIX_EPOCH).map_or(0, |d| d.as_millis() as u64)
}

fn home() -> PathBuf {
    std::env::var_os("HOME").map(PathBuf::from).unwrap_or_else(|| PathBuf::from("/tmp"))
}

fn data_dir() -> PathBuf {
    std::env::var_os("XDG_DATA_HOME")
        .map(PathBuf::from)
        .unwrap_or_else(|| home().join(".local/share"))
        .join("ens-table")
}

/// Где лежит интерфейс: рядом с программой (архив) или в /usr/share (RPM).
fn ui_dir() -> PathBuf {
    if let Some(p) = std::env::var_os("ENS_UI_DIR") {
        return PathBuf::from(p);
    }
    let exe_dir = std::env::current_exe().ok().and_then(|p| p.parent().map(Path::to_path_buf)).unwrap_or_default();
    for c in [exe_dir.join("ui"), exe_dir.join("../share/ens-table-r7/ui"), PathBuf::from("/usr/share/ens-table-r7/ui")] {
        if c.join("index.html").exists() {
            return c;
        }
    }
    exe_dir.join("ui")
}

fn random_token() -> String {
    let mut buf = [0u8; 16];
    let ok = fs::File::open("/dev/urandom").and_then(|mut f| f.read_exact(&mut buf)).is_ok();
    if !ok {
        // Нет /dev/urandom: случайные ключи хеш-таблиц стандартной библиотеки.
        use std::hash::{BuildHasher, Hasher};
        for chunk in buf.chunks_mut(8) {
            let mut h = std::collections::hash_map::RandomState::new().build_hasher();
            h.write_u64(now_ms() ^ u64::from(std::process::id()));
            chunk.copy_from_slice(&h.finish().to_le_bytes()[..chunk.len()]);
        }
    }
    buf.iter().map(|b| format!("{b:02x}")).collect()
}

fn find_browser() -> Option<String> {
    if let Ok(b) = std::env::var("ENS_BROWSER") {
        return Some(b);
    }
    let path = std::env::var_os("PATH")?;
    for name in BROWSERS {
        for dir in std::env::split_paths(&path) {
            if dir.join(name).is_file() {
                return Some(name.to_string());
            }
        }
    }
    // Яндекс.Браузер для организаций часто лежит вне PATH.
    ["/opt/yandex/browser-corporate/yandex-browser", "/opt/yandex/browser/yandex-browser"]
        .iter()
        .find(|p| Path::new(p).exists())
        .map(|p| p.to_string())
}

/// Окно ENS-Table: браузер в режиме приложения со своим профилем.
fn open_app_window(browser: &str, url: &str) -> std::io::Result<Child> {
    let profile = data_dir().join("browser");
    Command::new(browser)
        .arg(format!("--app={url}"))
        .arg(format!("--user-data-dir={}", profile.display()))
        .args([
            "--class=ENS-Table",
            "--name=ENS-Table",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-sync",
            "--disable-extensions",
            "--disable-features=TranslateUI",
            "--password-store=basic",
            "--window-size=1280,820",
        ])
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn()
}

fn message(text: &str) {
    eprintln!("ENS-Table: {text}");
    let _ = Command::new("zenity").args(["--error", "--no-wrap", "--title=ENS-Table", &format!("--text={text}")]).status();
}

/// Окна выбора файлов: zenity (MATE/GNOME), иначе kdialog (KDE).
fn file_dialog(kind: &str, title: &str, name: &str) -> Result<Option<String>, String> {
    let filters = ["--file-filter=Книги Excel | *.xlsx *.xlsm *.XLSX", "--file-filter=Все файлы | *"];
    let mut z = Command::new("zenity");
    z.current_dir(home()).arg("--file-selection").arg(format!("--title={title}"));
    match kind {
        "open" => { z.args(filters); }
        "save" => { z.args(["--save", "--confirm-overwrite"]).arg(format!("--filename={name}")).arg(filters[0]); }
        _ => { z.arg("--directory"); }
    }
    let out = match z.output() {
        Ok(o) => o,
        Err(_) => {
            let mut k = Command::new("kdialog");
            k.current_dir(home()).arg(format!("--title={title}"));
            match kind {
                "open" => { k.args(["--getopenfilename", ".", "*.xlsx *.xlsm|Книги Excel"]); }
                "save" => { k.args(["--getsavefilename"]).arg(name).arg("*.xlsx|Книга Excel"); }
                _ => { k.args(["--getexistingdirectory", "."]); }
            }
            k.output().map_err(|_| "Не найдено окно выбора файлов: установите пакет zenity".to_string())?
        }
    };
    let path = String::from_utf8_lossy(&out.stdout).trim().to_string();
    Ok((out.status.success() && !path.is_empty()).then_some(path))
}

struct App {
    token: String,
    port: u16,
    browser: String,
    recovery: recovery::Recovery,
    last_request: AtomicU64,
}

impl App {
    fn app_url(&self, query: &str) -> String {
        let sep = if query.is_empty() { "" } else { "&" };
        format!("http://127.0.0.1:{}/index.html?{query}{sep}t={}", self.port, self.token)
    }
}

type Reply = Result<Response<std::io::Cursor<Vec<u8>>>, String>;

fn json_reply(v: Value) -> Reply {
    let mut r = Response::from_data(serde_json::to_vec(&v).unwrap_or_default());
    r.add_header(Header::from_bytes("Content-Type", "application/json").unwrap());
    Ok(r)
}

fn bytes_reply(b: Vec<u8>) -> Reply {
    let mut r = Response::from_data(b);
    r.add_header(Header::from_bytes("Content-Type", "application/octet-stream").unwrap());
    Ok(r)
}

fn arg(v: &Value, name: &str) -> Result<String, String> {
    v.get(name).and_then(Value::as_str).map(str::to_string).ok_or_else(|| format!("нет параметра {name}"))
}

/// Команды — те же, что у настольной версии (см. src/platform.ts).
fn api(app: &App, cmd: &str, headers: &[(String, String)], body: Vec<u8>) -> Reply {
    let header = |n: &str| headers.iter().find(|(k, _)| k.eq_ignore_ascii_case(n)).map(|(_, v)| v.clone());
    let j = || -> Value { serde_json::from_slice(&body).unwrap_or(Value::Null) };
    let rec = &app.recovery;
    match cmd {
        "user_name" => json_reply(json!(files::user_name())),
        "inspect_file" => json_reply(json!(files::inspect_file(arg(&j(), "path")?)?)),
        "read_file" => {
            let p = PathBuf::from(arg(&j(), "path")?);
            bytes_reply(fs::read(&p).map_err(|e| format!("Не удалось открыть «{}»: {e}", p.display()))?)
        }
        "save_file" => {
            let path = files::percent_decode(&header("path").ok_or("нет пути")?)?;
            let expected = header("expected-modified").and_then(|v| v.parse().ok()).unwrap_or(0);
            let force = header("force").as_deref() == Some("1");
            json_reply(json!(files::save_atomic(Path::new(&path), &body, expected, force)?))
        }
        "recovery_heartbeat" => json_reply(json!(rec.heartbeat()?)),
        "recovery_set_root" => json_reply(json!(rec.set_root(arg(&j(), "path")?)?)),
        "recovery_root" => json_reply(json!(rec.root_text())),
        "recovery_begin" => {
            let v = j();
            let meta = serde_json::from_value(v.get("meta").cloned().unwrap_or_default()).map_err(|e| e.to_string())?;
            json_reply(json!(rec.begin(arg(&v, "id")?, meta)?))
        }
        "recovery_snapshot" => json_reply(json!(rec.snapshot(header("id").ok_or("нет id")?, body)?)),
        "recovery_snapshot_xlsx" => json_reply(json!(rec.snapshot_xlsx(header("id").ok_or("нет id")?, body)?)),
        "recovery_load_snapshot_xlsx" => { let v = j(); bytes_reply(rec.load_snapshot_xlsx(arg(&v, "session")?, arg(&v, "id")?)?) }
        "recovery_append" => json_reply(json!(rec.append(header("id").ok_or("нет id")?, body)?)),
        "recovery_forget" => json_reply(json!(rec.forget(arg(&j(), "id")?)?)),
        "recovery_end_session" => json_reply(json!(rec.end_session()?)),
        "recovery_list" => json_reply(json!(rec.list()?)),
        "recovery_load_snapshot" => { let v = j(); bytes_reply(rec.load_snapshot(arg(&v, "session")?, arg(&v, "id")?)?) }
        "recovery_load_journal" => { let v = j(); bytes_reply(rec.load_journal(arg(&v, "session")?, arg(&v, "id")?)?) }
        "recovery_discard" => { let v = j(); json_reply(json!(rec.discard(arg(&v, "session")?, arg(&v, "id")?)?)) }
        "pick_open" => json_reply(json!(file_dialog("open", &arg(&j(), "title").unwrap_or_default(), "")?)),
        "pick_save" => { let v = j(); json_reply(json!(file_dialog("save", &arg(&v, "title").unwrap_or_default(), &arg(&v, "name").unwrap_or_default())?)) }
        "pick_folder" => json_reply(json!(file_dialog("folder", &arg(&j(), "title").unwrap_or_default(), "")?)),
        "open_window" => {
            open_app_window(&app.browser, &app.app_url(&arg(&j(), "query").unwrap_or_default())).map_err(|e| e.to_string())?;
            json_reply(json!(null))
        }
        "window_closed" => {
            // Окно закрыли без несохранённых изменений — временные файлы не нужны.
            let v = j();
            if v.get("clean").and_then(Value::as_bool) == Some(true) {
                let _ = rec.forget(arg(&v, "id")?);
            }
            json_reply(json!(null))
        }
        "ping" => json_reply(json!("ens-table")),
        _ => Err(format!("неизвестная команда {cmd}")),
    }
}

fn content_type(path: &Path) -> &'static str {
    match path.extension().and_then(|e| e.to_str()).unwrap_or("") {
        "html" => "text/html; charset=utf-8",
        "js" | "mjs" => "text/javascript; charset=utf-8",
        "css" => "text/css; charset=utf-8",
        "wasm" => "application/wasm",
        "svg" => "image/svg+xml",
        "png" => "image/png",
        "ico" => "image/x-icon",
        "woff2" => "font/woff2",
        "woff" => "font/woff",
        "json" => "application/json",
        _ => "application/octet-stream",
    }
}

fn serve(app: Arc<App>, server: Server, browser_done: Arc<AtomicBool>) {
    let ui = ui_dir();
    let host_ok = format!("127.0.0.1:{}", app.port);
    loop {
        // Окно закрыто и минуту нет запросов — работа окончена.
        if browser_done.load(Ordering::SeqCst) && now_ms() - app.last_request.load(Ordering::SeqCst) > 60_000 {
            break;
        }
        let Ok(Some(mut req)) = server.recv_timeout(Duration::from_secs(2)) else { continue };
        app.last_request.store(now_ms(), Ordering::SeqCst);
        let headers: Vec<(String, String)> =
            req.headers().iter().map(|h| (h.field.as_str().to_string(), h.value.as_str().to_string())).collect();
        let get = |n: &str| headers.iter().find(|(k, _)| k.eq_ignore_ascii_case(n)).map(|(_, v)| v.as_str());
        // Защита от «перепривязки DNS»: только наш адрес.
        if get("Host") != Some(host_ok.as_str()) {
            let _ = req.respond(Response::from_string("forbidden").with_status_code(403));
            continue;
        }
        let url = req.url().to_string();
        let (path, query) = url.split_once('?').unwrap_or((&url, ""));
        let reply: Reply = if let Some(cmd) = path.strip_prefix("/api/") {
            let token_q = query.split('&').find_map(|kv| kv.strip_prefix("t=")).unwrap_or("");
            if get("x-ens-token") != Some(app.token.as_str()) && token_q != app.token {
                Err("нет доступа".to_string())
            } else if *req.method() != Method::Post {
                Err("ожидался POST".to_string())
            } else {
                let mut body = Vec::new();
                let _ = req.as_reader().read_to_end(&mut body);
                api(&app, cmd, &headers, body)
            }
        } else {
            let rel = if path == "/" { "index.html" } else { path.trim_start_matches('/') };
            if rel.split('/').any(|p| p == ".." || p.is_empty()) {
                Err("нет файла".to_string())
            } else {
                let file = ui.join(rel);
                match fs::read(&file) {
                    Ok(data) => {
                        let mut r = Response::from_data(data);
                        r.add_header(Header::from_bytes("Content-Type", content_type(&file)).unwrap());
                        r.add_header(Header::from_bytes("Cache-Control", "no-store").unwrap());
                        Ok(r)
                    }
                    Err(_) => Err("нет файла".to_string()),
                }
            }
        };
        let _ = match reply {
            Ok(r) => req.respond(r),
            Err(e) => req.respond(Response::from_string(e).with_status_code(400)),
        };
    }
}

/// Уже запущенный экземпляр: окно для нового файла открывает он.
fn existing_instance() -> Option<(u16, String)> {
    let v: Value = serde_json::from_slice(&fs::read(data_dir().join("instance.json")).ok()?).ok()?;
    let pid = v.get("pid")?.as_u64()?;
    if !Path::new(&format!("/proc/{pid}")).exists() {
        return None;
    }
    Some((v.get("port")?.as_u64()? as u16, v.get("token")?.as_str()?.to_string()))
}

fn main() {
    let file = std::env::args().skip(1).map(PathBuf::from).find(|p| p.is_file());
    let open_query = file
        .as_ref()
        .and_then(|p| fs::canonicalize(p).ok())
        .map(|p| format!("open={}", encode(&p.display().to_string())))
        .unwrap_or_default();
    let Some(browser) = find_browser() else {
        message("Не найден браузер Chromium или Яндекс.Браузер — ENS-Table открывает окно с его помощью. Попросите администратора установить пакет chromium.");
        std::process::exit(2);
    };
    if let Some((port, token)) = existing_instance() {
        let sep = if open_query.is_empty() { "" } else { "&" };
        let url = format!("http://127.0.0.1:{port}/index.html?{open_query}{sep}t={token}");
        if open_app_window(&browser, &url).is_ok() {
            return;
        }
    }
    let server = match Server::http("127.0.0.1:0") {
        Ok(s) => s,
        Err(e) => {
            message(&format!("Не удалось запустить ENS-Table: {e}"));
            std::process::exit(1);
        }
    };
    let port = server.server_addr().to_ip().map(|a| a.port()).unwrap_or(0);
    let _ = fs::create_dir_all(data_dir());
    let app = Arc::new(App {
        token: random_token(),
        port,
        browser: browser.clone(),
        recovery: recovery::Recovery::new(data_dir().join("recovery")),
        last_request: AtomicU64::new(now_ms()),
    });
    let instance = json!({ "pid": std::process::id(), "port": port, "token": app.token });
    let _ = fs::write(data_dir().join("instance.json"), instance.to_string());

    let browser_done = Arc::new(AtomicBool::new(false));
    match open_app_window(&browser, &app.app_url(&open_query)) {
        Ok(mut child) => {
            let done = browser_done.clone();
            std::thread::spawn(move || {
                let _ = child.wait();
                done.store(true, Ordering::SeqCst);
            });
        }
        Err(e) => {
            message(&format!("Не удалось открыть окно ({browser}): {e}"));
            std::process::exit(1);
        }
    }
    serve(app.clone(), server, browser_done);
    let _ = app.recovery.end_session();
    let _ = fs::remove_file(data_dir().join("instance.json"));
}

/// Кодирование пути для адреса (кириллица, пробелы).
fn encode(s: &str) -> String {
    s.bytes()
        .map(|b| match b {
            b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'-' | b'_' | b'.' | b'~' | b'/' => (b as char).to_string(),
            _ => format!("%{b:02X}"),
        })
        .collect()
}
