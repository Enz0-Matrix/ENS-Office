//! Работа с файлами (ADR-005).
//!
//! * Тип файла определяется по содержимому, а не по расширению: xlsx, который
//!   прислали под чужим именем, всё равно откроется как таблица (в Р7 такой
//!   файл иногда открывается текстовым редактором).
//! * Сохранение атомарное: временный файл рядом → запись на диск (fsync) →
//!   чтение обратно и сверка → замена исходного одним переименованием.
//!   Сбой на любом шаге оставляет исходный файл нетронутым.
//! * Ошибки объясняются по-человечески: что случилось и что делать.

use std::fs::{self, File, OpenOptions};
use std::io::{Read, Write};
use std::path::{Path, PathBuf};
use std::time::UNIX_EPOCH;

use serde::Serialize;
#[cfg(feature = "tauri-shell")]
use tauri::ipc::{InvokeBody, Request, Response};

#[derive(Serialize)]
pub struct FileInfo {
    /// xlsx, xlsm, docx, pptx, ods, odt, xls-or-doc, csv, unknown
    kind: String,
    size: u64,
    /// Время изменения, миллисекунды Unix: чтобы заметить, что файл изменил
    /// кто-то другой, пока он был открыт.
    modified_ms: u64,
    readonly: bool,
    name: String,
}

fn modified_ms(path: &Path) -> u64 {
    fs::metadata(path)
        .and_then(|m| m.modified())
        .ok()
        .and_then(|t| t.duration_since(UNIX_EPOCH).ok())
        .map_or(0, |d| d.as_millis() as u64)
}

/// Тип по содержимому.
pub fn detect_kind(bytes: &[u8]) -> String {
    if bytes.starts_with(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]) {
        return "xls-or-doc".into();
    }
    if bytes.starts_with(b"PK") {
        if let Ok(mut archive) = zip::ZipArchive::new(std::io::Cursor::new(bytes)) {
            let mut text = String::new();
            if let Ok(mut entry) = archive.by_name("[Content_Types].xml") {
                let _ = entry.read_to_string(&mut text);
                return if text.contains("macroEnabled") && text.contains("spreadsheetml") {
                    "xlsm"
                } else if text.contains("spreadsheetml") {
                    "xlsx"
                } else if text.contains("wordprocessingml") {
                    "docx"
                } else if text.contains("presentationml") {
                    "pptx"
                } else {
                    "unknown"
                }
                .into();
            }
            if let Ok(mut entry) = archive.by_name("mimetype") {
                let _ = entry.read_to_string(&mut text);
                if text.contains("opendocument.spreadsheet") {
                    return "ods".into();
                }
                if text.contains("opendocument.text") {
                    return "odt".into();
                }
            }
        }
        return "unknown".into();
    }
    let head = &bytes[..bytes.len().min(4096)];
    if std::str::from_utf8(head).is_ok() && head.iter().any(|&b| b == b'\n' || b == b';' || b == b',') {
        return "csv".into();
    }
    "unknown".into()
}

/// Понятное объяснение ошибки ввода-вывода.
fn explain(err: &std::io::Error, path: &Path, action: &str) -> String {
    let name = path.file_name().map_or_else(|| path.display().to_string(), |n| n.to_string_lossy().into());
    let reason = match err.raw_os_error() {
        // ERROR_SHARING_VIOLATION, ERROR_LOCK_VIOLATION
        Some(32) | Some(33) => "файл открыт в другой программе (Excel, Р7, LibreOffice?). Закройте его там или сохраните под другим именем".to_string(),
        // ERROR_HANDLE_DISK_FULL, ERROR_DISK_FULL
        Some(39) | Some(112) => "на диске недостаточно места".to_string(),
        // ERROR_BAD_NETPATH, ERROR_NETNAME_DELETED, ERROR_UNEXP_NET_ERR
        Some(53) | Some(64) | Some(59) => "сетевая папка недоступна (связь с сервером прервалась)".to_string(),
        _ => match err.kind() {
            std::io::ErrorKind::PermissionDenied => "нет прав на запись в эту папку или файл только для чтения".to_string(),
            std::io::ErrorKind::NotFound => "папка не найдена (диск отключён или файл перемещён)".to_string(),
            _ => err.to_string(),
        },
    };
    format!("Не удалось {action} «{name}»: {reason}. Попробуйте «Сохранить как…» в другое место — данные в программе не потеряны.")
}

#[cfg_attr(feature = "tauri-shell", tauri::command)]
pub fn inspect_file(path: String) -> Result<FileInfo, String> {
    let p = PathBuf::from(&path);
    let meta = fs::metadata(&p).map_err(|e| explain(&e, &p, "открыть"))?;
    let mut head = vec![0u8; 0];
    File::open(&p)
        .and_then(|mut f| f.read_to_end(&mut head))
        .map_err(|e| explain(&e, &p, "открыть"))?;
    Ok(FileInfo {
        kind: detect_kind(&head),
        size: meta.len(),
        modified_ms: modified_ms(&p),
        readonly: meta.permissions().readonly(),
        name: p.file_name().map(|n| n.to_string_lossy().into()).unwrap_or_default(),
    })
}

#[cfg(feature = "tauri-shell")]
#[tauri::command]
pub fn read_file(path: String) -> Result<Response, String> {
    let p = PathBuf::from(&path);
    fs::read(&p).map(Response::new).map_err(|e| explain(&e, &p, "открыть"))
}

#[cfg(feature = "tauri-shell")]
fn header(request: &Request<'_>, name: &str) -> Result<String, String> {
    let raw = request
        .headers()
        .get(name)
        .ok_or_else(|| format!("нет заголовка {name}"))?
        .to_str()
        .map_err(|e| e.to_string())?;
    // Путь передаётся закодированным (кириллица в заголовках недопустима).
    percent_decode(raw)
}

pub fn percent_decode(s: &str) -> Result<String, String> {
    let bytes = s.as_bytes();
    let mut out = Vec::with_capacity(bytes.len());
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'%' && i + 2 < bytes.len() {
            let hex = std::str::from_utf8(&bytes[i + 1..i + 3]).map_err(|e| e.to_string())?;
            out.push(u8::from_str_radix(hex, 16).map_err(|e| e.to_string())?);
            i += 3;
        } else {
            out.push(bytes[i]);
            i += 1;
        }
    }
    String::from_utf8(out).map_err(|e| e.to_string())
}

#[cfg(feature = "tauri-shell")]
fn body(request: &Request<'_>) -> Result<Vec<u8>, String> {
    match request.body() {
        InvokeBody::Raw(bytes) => Ok(bytes.clone()),
        InvokeBody::Json(_) => Err("ожидались байты файла".into()),
    }
}

/// Атомарное сохранение. Заголовки: `path` (закодирован), `expected-modified`
/// — время изменения файла на момент открытия (0 — не проверять).
/// Возвращает новое время изменения.
#[cfg(feature = "tauri-shell")]
#[tauri::command]
pub fn save_file(request: Request<'_>) -> Result<u64, String> {
    let path = PathBuf::from(header(&request, "path")?);
    let expected: u64 = header(&request, "expected-modified").ok().and_then(|v| v.parse().ok()).unwrap_or(0);
    let force = header(&request, "force").map(|v| v == "1").unwrap_or(false);
    let data = body(&request)?;
    save_atomic(&path, &data, expected, force)
}

pub fn save_atomic(path: &Path, data: &[u8], expected_modified: u64, force: bool) -> Result<u64, String> {
    // 1. Файл изменили снаружи, пока он был открыт: не затираем молча.
    if !force && expected_modified != 0 && path.exists() {
        let now = modified_ms(path);
        if now != 0 && now.abs_diff(expected_modified) > 2000 {
            return Err("CHANGED_ON_DISK".into());
        }
    }
    let dir = path.parent().filter(|d| !d.as_os_str().is_empty()).unwrap_or(Path::new("."));
    let name = path.file_name().map(|n| n.to_string_lossy().to_string()).unwrap_or_else(|| "книга".into());
    let tmp = dir.join(format!(".~{}.{}.ens-tmp", name, std::process::id()));

    // 2. Пишем во временный файл рядом с исходным и сбрасываем на диск.
    let write = || -> std::io::Result<()> {
        let mut f = OpenOptions::new().write(true).create(true).truncate(true).open(&tmp)?;
        f.write_all(data)?;
        f.sync_all()
    };
    if let Err(e) = write() {
        let _ = fs::remove_file(&tmp);
        return Err(explain(&e, path, "сохранить"));
    }
    // 3. Читаем обратно и сверяем: защита от сбоев сети и диска.
    match fs::read(&tmp) {
        Ok(back) if back == data => {}
        Ok(_) => {
            let _ = fs::remove_file(&tmp);
            return Err("Не удалось сохранить: записанный файл не совпал с данными (сбой диска или сети). Исходный файл не изменён — сохраните под другим именем.".into());
        }
        Err(e) => {
            let _ = fs::remove_file(&tmp);
            return Err(explain(&e, path, "проверить сохранённый"));
        }
    }
    // 4. Замена одним переименованием (на Windows — с заменой существующего).
    if let Err(e) = fs::rename(&tmp, path) {
        let _ = fs::remove_file(&tmp);
        return Err(explain(&e, path, "сохранить"));
    }
    Ok(modified_ms(path))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn atomic_save_keeps_original_on_failure_and_replaces_on_success() {
        let dir = std::env::temp_dir().join(format!("ens-save-{}", std::process::id()));
        fs::create_dir_all(&dir).unwrap();
        let file = dir.join("Книга.xlsx");
        fs::write(&file, b"old").unwrap();
        let t = modified_ms(&file);
        save_atomic(&file, b"new contents", t, false).unwrap();
        assert_eq!(fs::read(&file).unwrap(), b"new contents");
        // Файл «изменили снаружи»: без force сохранение отказывается.
        assert_eq!(save_atomic(&file, b"x", t.saturating_sub(10_000), false).unwrap_err(), "CHANGED_ON_DISK");
        assert_eq!(fs::read(&file).unwrap(), b"new contents");
        // Временных файлов не осталось.
        let leftovers: Vec<_> = fs::read_dir(&dir).unwrap().filter_map(|e| e.ok())
            .filter(|e| e.file_name().to_string_lossy().ends_with(".ens-tmp")).collect();
        assert!(leftovers.is_empty());
        fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn detects_kind_by_content() {
        assert_eq!(detect_kind(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0]), "xls-or-doc");
        assert_eq!(detect_kind(b"a;b\n1;2\n"), "csv");
        assert_eq!(detect_kind(b"PK\x03\x04garbage"), "unknown");
        assert_eq!(percent_decode("C%3A%5C%D0%9A.xlsx").unwrap(), "C:\\К.xlsx");
    }
}
