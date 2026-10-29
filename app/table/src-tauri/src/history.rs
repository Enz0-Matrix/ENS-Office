//! Копии на диске (ENS-Tekst; для ENS-Table пока не используется).
//!
//! * `backup` — предыдущее содержимое файла перед тем, как его перезаписали (последние 10 на файл, не старше 60 дней);
//! * `discarded` — документ, закрытый кнопкой «Не сохранять» (последние 20, не старше 14 дней): ошибочное
//!   «Не сохранять» — частая причина потери работы, и это единственный способ её вернуть.
//! Лежат в профиле пользователя рядом с папкой восстановления; наружу не уходят.

use std::fs;
use std::path::{Path, PathBuf};
use std::time::{SystemTime, UNIX_EPOCH};

use serde::Serialize;

pub struct History {
    root: PathBuf,
}

#[derive(Serialize)]
pub struct Entry {
    /// backup | discarded
    kind: String,
    /// Имя файла, как его передала программа («Договор.docx»).
    name: String,
    /// Полный путь к копии: читается обычной командой чтения файла.
    path: String,
    ms: u64,
    size: u64,
}

const KINDS: [&str; 2] = ["backup", "discarded"];
const DAY_MS: u64 = 86_400_000;

fn now_ms() -> u64 {
    SystemTime::now().duration_since(UNIX_EPOCH).map_or(0, |d| d.as_millis() as u64)
}

fn safe(name: &str) -> String {
    let s: String = name
        .chars()
        .map(|c| if c.is_alphanumeric() || matches!(c, ' ' | '-' | '_' | '.' | '(' | ')') { c } else { '_' })
        .collect();
    let s: String = s.trim().chars().take(80).collect();
    if s.is_empty() { "документ.docx".into() } else { s }
}

/// «1730000000000__Договор.docx» → (время, имя).
fn parse(file: &str) -> Option<(u64, String)> {
    let (ms, name) = file.split_once("__")?;
    Some((ms.parse().ok()?, name.to_string()))
}

impl History {
    pub fn new(root: PathBuf) -> Self {
        History { root }
    }

    pub fn store(&self, kind: &str, name: &str, data: &[u8]) -> Result<String, String> {
        if !KINDS.contains(&kind) {
            return Err("неизвестный вид копии".into());
        }
        let dir = self.root.join(kind);
        fs::create_dir_all(&dir).map_err(|e| format!("Не удалось создать папку копий: {e}"))?;
        let name = safe(name);
        // Время строго растёт: два сохранения в одну миллисекунду не затирают друг друга и не путают порядок.
        let newest = Self::files(&dir).first().map_or(0, |f| f.0);
        let ms = now_ms().max(newest + 1);
        let file = dir.join(format!("{ms}__{name}"));
        let tmp = dir.join(format!(".{ms}.tmp"));
        fs::write(&tmp, data).map_err(|e| format!("Не удалось записать копию: {e}"))?;
        fs::rename(&tmp, &file).map_err(|e| {
            let _ = fs::remove_file(&tmp);
            format!("Не удалось записать копию: {e}")
        })?;
        self.prune(kind);
        Ok(file.display().to_string())
    }

    fn files(dir: &Path) -> Vec<(u64, String, PathBuf)> {
        let mut v: Vec<_> = fs::read_dir(dir)
            .into_iter()
            .flatten()
            .flatten()
            .filter_map(|e| {
                let f = e.file_name().to_string_lossy().to_string();
                let (ms, name) = parse(&f)?;
                Some((ms, name, e.path()))
            })
            .collect();
        v.sort_by(|a, b| b.0.cmp(&a.0));
        v
    }

    fn prune(&self, kind: &str) {
        let now = now_ms();
        let dir = self.root.join(kind);
        let all = Self::files(&dir);
        let (per_name, max_age, total) = if kind == "backup" { (10, 60 * DAY_MS, usize::MAX) } else { (usize::MAX, 14 * DAY_MS, 20) };
        let mut seen: std::collections::HashMap<String, usize> = std::collections::HashMap::new();
        for (i, (ms, name, path)) in all.iter().enumerate() {
            let n = seen.entry(name.clone()).or_insert(0);
            *n += 1;
            if *n > per_name || i >= total || now.saturating_sub(*ms) > max_age {
                let _ = fs::remove_file(path);
            }
        }
    }

    pub fn list(&self) -> Result<Vec<Entry>, String> {
        let mut out = Vec::new();
        for kind in KINDS {
            for (ms, name, path) in Self::files(&self.root.join(kind)) {
                let size = fs::metadata(&path).map(|m| m.len()).unwrap_or(0);
                out.push(Entry { kind: kind.into(), name, path: path.display().to_string(), ms, size });
            }
        }
        out.sort_by(|a, b| b.ms.cmp(&a.ms));
        Ok(out)
    }
}

#[cfg(feature = "tauri-shell")]
#[tauri::command]
pub fn history_store(state: tauri::State<'_, History>, request: tauri::ipc::Request<'_>) -> Result<String, String> {
    let kind = super::files::header(&request, "kind")?;
    let name = super::files::header(&request, "name")?;
    let data = super::files::body(&request)?;
    state.store(&kind, &name, &data)
}

#[cfg(feature = "tauri-shell")]
#[tauri::command]
pub fn history_list(state: tauri::State<'_, History>) -> Result<Vec<Entry>, String> {
    state.list()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn dir(name: &str) -> PathBuf {
        let d = std::env::temp_dir().join(format!("ens-history-{name}-{}", now_ms()));
        let _ = fs::remove_dir_all(&d);
        d
    }

    #[test]
    fn stores_lists_and_prunes() {
        let h = History::new(dir("a"));
        for i in 0..14u8 {
            h.store("backup", "Договор.docx", &[i]).unwrap();
        }
        h.store("backup", "Другой.docx", b"x").unwrap();
        h.store("discarded", "Черновик.docx", b"y").unwrap();
        let list = h.list().unwrap();
        assert_eq!(list.iter().filter(|e| e.name == "Договор.docx").count(), 10, "10 копий на файл");
        assert_eq!(list.iter().filter(|e| e.kind == "discarded").count(), 1);
        // Новейшая копия — последняя записанная.
        let newest = list.iter().filter(|e| e.name == "Договор.docx").max_by_key(|e| e.ms).unwrap();
        assert_eq!(fs::read(&newest.path).unwrap(), vec![13u8]);
    }

    #[test]
    fn rejects_unknown_kind_and_unsafe_names() {
        let h = History::new(dir("b"));
        assert!(h.store("../x", "a.docx", b"1").is_err());
        let p = h.store("backup", "../../etc/passwd", b"1").unwrap();
        assert!(p.starts_with(&h.root.display().to_string()), "имя не должно выводить за папку");
    }
}
