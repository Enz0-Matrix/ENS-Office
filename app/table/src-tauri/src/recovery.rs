//! Восстановление после сбоя (docs/03-table-ux-requirements.md, §7).
//!
//! Для каждой открытой книги в папке восстановления лежат:
//!   meta.json      — откуда книга, когда открыта, когда снимок;
//!   snapshot.bin   — вся книга (раз в N минут и после простоя);
//!   snapshot.prev.bin — предыдущий снимок;
//!   journal.bin    — правки после снимка: [длина u32][байты] …
//! Восстановление = снимок + все правки журнала: теряется не больше секунды.
//!
//! Сеанс (запуск программы) держит блокировку файла alive.lock, пока жив.
//! Если процесс убит (сбой, «Снять задачу»), система сразу снимает блокировку:
//! сеанс без блокировки и без отметки о нормальном завершении — упавший, его
//! книги предлагаются к восстановлению при следующем запуске. Книги второго
//! работающего экземпляра программы не трогаем. Для старых сеансов без
//! alive.lock (и папок, где блокировки не работают) остаётся «пульс» раз в 30 с:
//! без пульса дольше двух минут — упавший.

use std::fs::{self, File, OpenOptions};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use std::time::{SystemTime, UNIX_EPOCH};

use serde::{Deserialize, Serialize};
#[cfg(feature = "tauri-shell")]
use tauri::ipc::{Request, Response};

pub struct Recovery {
    pub root: Mutex<PathBuf>,
    pub session: String,
    /// Открытый и заблокированный alive.lock этого сеанса.
    alive: Mutex<Option<File>>,
}

fn now_ms() -> u64 {
    SystemTime::now().duration_since(UNIX_EPOCH).map_or(0, |d| d.as_millis() as u64)
}

#[derive(Serialize, Deserialize, Clone)]
pub struct SessionInfo {
    pid: u32,
    started_ms: u64,
    heartbeat_ms: u64,
    clean_exit: bool,
}

#[derive(Serialize, Deserialize, Clone, Default)]
pub struct DocMeta {
    /// Путь к исходному файлу ("" — новая, ни разу не сохранённая книга).
    pub source_path: String,
    pub title: String,
    /// Время изменения исходного файла на момент открытия.
    pub source_modified_ms: u64,
    pub snapshot_ms: u64,
    pub journal_entries: u64,
    /// Формат снимка и журнала (версия программы): снимки другой версии
    /// восстанавливаются из запасной копии в xlsx (snapshot.xlsx).
    #[serde(default)]
    pub format: String,
    /// Когда записана запасная копия в xlsx (0 — нет).
    #[serde(default)]
    pub xlsx_ms: u64,
}

#[derive(Serialize)]
pub struct RecoverableDoc {
    session: String,
    id: String,
    meta: DocMeta,
    /// Исходный файл существует и его можно открыть для сравнения.
    source_exists: bool,
    /// Исходный файл менялся после того, как книга была открыта.
    source_changed: bool,
    crashed_ms: u64,
}

impl Recovery {
    pub fn new(root: PathBuf) -> Self {
        let session = format!("{}-{}", now_ms(), std::process::id());
        let r = Recovery { root: Mutex::new(root), session, alive: Mutex::new(None) };
        let _ = r.write_session(false);
        r.hold_alive_lock();
        r
    }

    fn root(&self) -> PathBuf {
        self.root.lock().map(|r| r.clone()).unwrap_or_default()
    }

    fn session_dir(&self, session: &str) -> PathBuf {
        self.root().join("sessions").join(session)
    }

    fn doc_dir(&self, session: &str, id: &str) -> Result<PathBuf, String> {
        if id.is_empty() || !id.chars().all(|c| c.is_ascii_alphanumeric() || c == '-') {
            return Err("недопустимый идентификатор книги".into());
        }
        Ok(self.session_dir(session).join(id))
    }

    /// Заблокировать alive.lock в папке сеанса (заново — после смены папки).
    fn hold_alive_lock(&self) {
        let path = self.session_dir(&self.session).join("alive.lock");
        let file = OpenOptions::new().create(true).truncate(false).write(true).open(path).ok()
            .filter(|f| f.try_lock().is_ok());
        if let Ok(mut alive) = self.alive.lock() {
            *alive = file;
        }
    }

    fn write_session(&self, clean_exit: bool) -> std::io::Result<()> {
        let dir = self.session_dir(&self.session);
        fs::create_dir_all(&dir)?;
        let info_path = dir.join("session.json");
        let started = fs::read(&info_path)
            .ok()
            .and_then(|b| serde_json::from_slice::<SessionInfo>(&b).ok())
            .map_or_else(now_ms, |s| s.started_ms);
        let info = SessionInfo { pid: std::process::id(), started_ms: started, heartbeat_ms: now_ms(), clean_exit };
        write_synced(&info_path, &serde_json::to_vec(&info).unwrap_or_default())
    }
}

fn write_synced(path: &Path, data: &[u8]) -> std::io::Result<()> {
    let tmp = path.with_extension("tmp");
    {
        let mut f = OpenOptions::new().write(true).create(true).truncate(true).open(&tmp)?;
        f.write_all(data)?;
        f.sync_all()?;
    }
    fs::rename(tmp, path)
}

#[cfg(feature = "tauri-shell")]
fn raw(request: &Request<'_>) -> Result<Vec<u8>, String> {
    crate::files::bytes_of(request.body()).ok_or_else(|| "ожидались байты".into())
}

#[cfg(feature = "tauri-shell")]
fn header(request: &Request<'_>, name: &str) -> Result<String, String> {
    request
        .headers()
        .get(name)
        .and_then(|v| v.to_str().ok())
        .map(str::to_string)
        .ok_or_else(|| format!("нет заголовка {name}"))
}

fn io_err(e: std::io::Error) -> String {
    format!("Папка восстановления недоступна: {e}")
}

// --- Действия (общие для оболочки Tauri и для версии РЕД ОС 7) ---

impl Recovery {
    pub fn heartbeat(&self) -> Result<(), String> {
        let state = self;
        state.write_session(false).map_err(io_err)
    }

    pub fn set_root(&self, path: String) -> Result<(), String> {
        let state = self;
        let p = PathBuf::from(path);
        fs::create_dir_all(&p).map_err(io_err)?;
        *state.root.lock().map_err(|e| e.to_string())? = p;
        state.write_session(false).map_err(io_err)?;
        state.hold_alive_lock();
        Ok(())
    }

    pub fn root_text(&self) -> String {
        let state = self;
        state.root().display().to_string()
    }

    pub fn begin(&self, id: String, meta: DocMeta) -> Result<(), String> {
        let state = self;
        let dir = state.doc_dir(&state.session, &id)?;
        fs::create_dir_all(&dir).map_err(io_err)?;
        write_synced(&dir.join("meta.json"), &serde_json::to_vec(&meta).unwrap_or_default()).map_err(io_err)
    }

    pub fn snapshot(&self, id: String, data: Vec<u8>) -> Result<(), String> {
        let state = self;
        let dir = state.doc_dir(&state.session, &id)?;
        fs::create_dir_all(&dir).map_err(io_err)?;
        let snap = dir.join("snapshot.bin");
        if snap.exists() {
            let _ = fs::copy(&snap, dir.join("snapshot.prev.bin"));
        }
        write_synced(&snap, &data).map_err(io_err)?;
        let _ = fs::remove_file(dir.join("journal.bin"));
        let meta_path = dir.join("meta.json");
        let mut meta: DocMeta = fs::read(&meta_path).ok().and_then(|b| serde_json::from_slice(&b).ok()).unwrap_or_default();
        meta.snapshot_ms = now_ms();
        meta.journal_entries = 0;
        write_synced(&meta_path, &serde_json::to_vec(&meta).unwrap_or_default()).map_err(io_err)
    }

    /// Запасная копия в xlsx: её откроет любая версия программы.
    pub fn snapshot_xlsx(&self, id: String, data: Vec<u8>) -> Result<(), String> {
        let dir = self.doc_dir(&self.session, &id)?;
        fs::create_dir_all(&dir).map_err(io_err)?;
        write_synced(&dir.join("snapshot.xlsx"), &data).map_err(io_err)?;
        let meta_path = dir.join("meta.json");
        let mut meta: DocMeta = fs::read(&meta_path).ok().and_then(|b| serde_json::from_slice(&b).ok()).unwrap_or_default();
        meta.xlsx_ms = now_ms();
        write_synced(&meta_path, &serde_json::to_vec(&meta).unwrap_or_default()).map_err(io_err)
    }

    pub fn append(&self, id: String, data: Vec<u8>) -> Result<(), String> {
        let state = self;
        if data.is_empty() {
            return Ok(());
        }
        let dir = state.doc_dir(&state.session, &id)?;
        let mut f = OpenOptions::new().create(true).append(true).open(dir.join("journal.bin")).map_err(io_err)?;
        f.write_all(&(data.len() as u32).to_le_bytes()).map_err(io_err)?;
        f.write_all(&data).map_err(io_err)?;
        f.sync_data().map_err(io_err)
    }

    pub fn forget(&self, id: String) -> Result<(), String> {
        let state = self;
        let dir = state.doc_dir(&state.session, &id)?;
        if dir.exists() {
            fs::remove_dir_all(dir).map_err(io_err)?;
        }
        Ok(())
    }

    pub fn end_session(&self) -> Result<(), String> {
        let state = self;
        let dir = state.session_dir(&state.session);
        let has_docs = fs::read_dir(&dir).map(|d| d.filter_map(|e| e.ok()).any(|e| e.path().is_dir())).unwrap_or(false);
        if let Ok(mut alive) = state.alive.lock() {
            *alive = None;
        }
        if has_docs {
            // Остались книги, по которым пользователь не решил: помечаем, но храним.
            state.write_session(true).map_err(io_err)
        } else {
            fs::remove_dir_all(dir).map_err(io_err)
        }
    }

    pub fn list(&self) -> Result<Vec<RecoverableDoc>, String> {
        let state = self;
        let sessions = state.root().join("sessions");
        let mut out = Vec::new();
        let Ok(entries) = fs::read_dir(&sessions) else { return Ok(out) };
        for s in entries.filter_map(|e| e.ok()) {
            let session = s.file_name().to_string_lossy().to_string();
            if session == state.session {
                continue;
            }
            let info: Option<SessionInfo> = fs::read(s.path().join("session.json")).ok().and_then(|b| serde_json::from_slice(&b).ok());
            let Some(info) = info else { continue };
            if !info.clean_exit && session_alive(&s.path(), info.heartbeat_ms) {
                continue; // другой работающий экземпляр программы
            }
            let Ok(docs) = fs::read_dir(s.path()) else { continue };
            for d in docs.filter_map(|e| e.ok()).filter(|e| e.path().is_dir()) {
                let meta: DocMeta = fs::read(d.path().join("meta.json")).ok().and_then(|b| serde_json::from_slice(&b).ok()).unwrap_or_default();
                if !d.path().join("snapshot.bin").exists() {
                    continue;
                }
                let source = PathBuf::from(&meta.source_path);
                let source_exists = !meta.source_path.is_empty() && source.exists();
                let source_now = fs::metadata(&source).and_then(|m| m.modified()).ok()
                    .and_then(|t| t.duration_since(UNIX_EPOCH).ok()).map_or(0, |d| d.as_millis() as u64);
                out.push(RecoverableDoc {
                    session: session.clone(),
                    id: d.file_name().to_string_lossy().to_string(),
                    source_changed: source_exists && source_now.abs_diff(meta.source_modified_ms) > 2000,
                    meta,
                    source_exists,
                    crashed_ms: info.heartbeat_ms,
                });
            }
        }
        out.sort_by_key(|d| std::cmp::Reverse(d.crashed_ms));
        Ok(out)
    }

    pub fn load_snapshot(&self, session: String, id: String) -> Result<Vec<u8>, String> {
        let state = self;
        let dir = state.doc_dir(&session, &id)?;
        fs::read(dir.join("snapshot.bin")).map_err(io_err)
    }

    pub fn load_snapshot_xlsx(&self, session: String, id: String) -> Result<Vec<u8>, String> {
        let dir = self.doc_dir(&session, &id)?;
        fs::read(dir.join("snapshot.xlsx")).map_err(io_err)
    }

    pub fn load_journal(&self, session: String, id: String) -> Result<Vec<u8>, String> {
        let state = self;
        let dir = state.doc_dir(&session, &id)?;
        let data = fs::read(dir.join("journal.bin")).unwrap_or_default();
        let mut end = 0usize;
        while end + 4 <= data.len() {
            let len = u32::from_le_bytes([data[end], data[end + 1], data[end + 2], data[end + 3]]) as usize;
            if end + 4 + len > data.len() {
                break;
            }
            end += 4 + len;
        }
        Ok(data[..end].to_vec())
    }

    pub fn discard(&self, session: String, id: String) -> Result<(), String> {
        let state = self;
        let dir = state.doc_dir(&session, &id)?;
        if dir.exists() {
            fs::remove_dir_all(&dir).map_err(io_err)?;
        }
        let session_dir = state.session_dir(&session);
        let empty = fs::read_dir(&session_dir).map(|d| !d.filter_map(|e| e.ok()).any(|e| e.path().is_dir())).unwrap_or(true);
        if empty {
            let _ = fs::remove_dir_all(session_dir);
        }
        Ok(())
    }
}

/// Работает ли ещё программа, которая ведёт этот сеанс.
fn session_alive(dir: &Path, heartbeat_ms: u64) -> bool {
    match OpenOptions::new().write(true).open(dir.join("alive.lock")) {
        // Заблокировать удалось — владельца нет в живых.
        Ok(f) => f.try_lock().is_err(),
        Err(_) => now_ms().saturating_sub(heartbeat_ms) < 120_000,
    }
}

// --- Команды оболочки Tauri ---

#[cfg(feature = "tauri-shell")]
#[tauri::command]
pub fn recovery_heartbeat(state: tauri::State<'_, Recovery>) -> Result<(), String> {
    state.heartbeat()
}

#[cfg(feature = "tauri-shell")]
#[tauri::command]
pub fn recovery_set_root(state: tauri::State<'_, Recovery>, path: String) -> Result<(), String> {
    state.set_root(path)
}

#[cfg(feature = "tauri-shell")]
#[tauri::command]
pub fn recovery_root(state: tauri::State<'_, Recovery>) -> String {
    state.root_text()
}

#[cfg(feature = "tauri-shell")]
#[tauri::command]
pub fn recovery_begin(state: tauri::State<'_, Recovery>, id: String, meta: DocMeta) -> Result<(), String> {
    state.begin(id, meta)
}

#[cfg(feature = "tauri-shell")]
#[tauri::command]
pub fn recovery_snapshot(state: tauri::State<'_, Recovery>, request: Request<'_>) -> Result<(), String> {
    let id = header(&request, "id")?;
    let data = raw(&request)?;
    state.snapshot(id, data)
}

#[cfg(feature = "tauri-shell")]
#[tauri::command]
pub fn recovery_snapshot_xlsx(state: tauri::State<'_, Recovery>, request: Request<'_>) -> Result<(), String> {
    let id = header(&request, "id")?;
    let data = raw(&request)?;
    state.snapshot_xlsx(id, data)
}

#[cfg(feature = "tauri-shell")]
#[tauri::command]
pub fn recovery_load_snapshot_xlsx(state: tauri::State<'_, Recovery>, session: String, id: String) -> Result<Response, String> {
    state.load_snapshot_xlsx(session, id).map(Response::new)
}

#[cfg(feature = "tauri-shell")]
#[tauri::command]
pub fn recovery_append(state: tauri::State<'_, Recovery>, request: Request<'_>) -> Result<(), String> {
    let id = header(&request, "id")?;
    let data = raw(&request)?;
    state.append(id, data)
}

#[cfg(feature = "tauri-shell")]
#[tauri::command]
pub fn recovery_forget(state: tauri::State<'_, Recovery>, id: String) -> Result<(), String> {
    state.forget(id)
}

#[cfg(feature = "tauri-shell")]
#[tauri::command]
pub fn recovery_end_session(state: tauri::State<'_, Recovery>) -> Result<(), String> {
    state.end_session()
}

#[cfg(feature = "tauri-shell")]
#[tauri::command]
pub fn recovery_list(state: tauri::State<'_, Recovery>) -> Result<Vec<RecoverableDoc>, String> {
    state.list()
}

#[cfg(feature = "tauri-shell")]
#[tauri::command]
pub fn recovery_load_snapshot(state: tauri::State<'_, Recovery>, session: String, id: String) -> Result<Response, String> {
    state.load_snapshot(session, id).map(Response::new)
}

#[cfg(feature = "tauri-shell")]
#[tauri::command]
pub fn recovery_load_journal(state: tauri::State<'_, Recovery>, session: String, id: String) -> Result<Response, String> {
    state.load_journal(session, id).map(Response::new)
}

#[cfg(feature = "tauri-shell")]
#[tauri::command]
pub fn recovery_discard(state: tauri::State<'_, Recovery>, session: String, id: String) -> Result<(), String> {
    state.discard(session, id)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn crashed_session_is_listed_and_journal_tail_is_cut() {
        let root = std::env::temp_dir().join(format!("ens-recovery-{}", now_ms()));
        // «Упавший» сеанс: пульс давно, без отметки о завершении.
        let crashed = root.join("sessions").join("1-1");
        let doc = crashed.join("doc-1");
        fs::create_dir_all(&doc).unwrap();
        fs::write(crashed.join("session.json"), serde_json::to_vec(&SessionInfo { pid: 1, started_ms: 1, heartbeat_ms: 1, clean_exit: false }).unwrap()).unwrap();
        fs::write(doc.join("meta.json"), serde_json::to_vec(&DocMeta { title: "Книга1".into(), ..Default::default() }).unwrap()).unwrap();
        fs::write(doc.join("snapshot.bin"), b"snapshot").unwrap();
        let mut journal = Vec::new();
        journal.extend_from_slice(&3u32.to_le_bytes());
        journal.extend_from_slice(b"abc");
        journal.extend_from_slice(&10u32.to_le_bytes()); // оборванная запись
        journal.extend_from_slice(b"de");
        fs::write(doc.join("journal.bin"), &journal).unwrap();

        // Сеанс, убитый только что: пульс свежий, но alive.lock никто не держит.
        let killed = root.join("sessions").join("2-2");
        fs::create_dir_all(killed.join("doc-2")).unwrap();
        fs::write(killed.join("session.json"), serde_json::to_vec(&SessionInfo { pid: 2, started_ms: 1, heartbeat_ms: now_ms(), clean_exit: false }).unwrap()).unwrap();
        fs::write(killed.join("alive.lock"), b"").unwrap();
        assert!(!session_alive(&killed, now_ms()));
        // Живой сеанс держит блокировку.
        let live = Recovery::new(root.join("live"));
        assert!(session_alive(&live.session_dir(&live.session), 0));

        let r = Recovery::new(root.clone());
        let session = r.session.clone();
        assert_ne!(session, "1-1");
        let crashed_dir = r.doc_dir("1-1", "doc-1").unwrap();
        assert!(crashed_dir.join("snapshot.bin").exists());
        // Журнал без оборванного хвоста.
        let data = fs::read(crashed_dir.join("journal.bin")).unwrap();
        let mut end = 0usize;
        while end + 4 <= data.len() {
            let len = u32::from_le_bytes([data[end], data[end + 1], data[end + 2], data[end + 3]]) as usize;
            if end + 4 + len > data.len() { break; }
            end += 4 + len;
        }
        assert_eq!(end, 7);
        assert!(r.doc_dir("1-1", "../x").is_err());
        drop((r, live));
        fs::remove_dir_all(root).unwrap();
    }
}
