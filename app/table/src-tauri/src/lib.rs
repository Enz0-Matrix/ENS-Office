//! ENS-Table: оболочка Tauri. Движок таблиц работает в окне (WebAssembly);
//! здесь — то, чего нельзя сделать в браузере: надёжная работа с файлами
//! (files.rs) и восстановление после сбоя (recovery.rs).

mod files;
mod recovery;

use std::sync::Mutex;

use tauri::Manager;

/// Файл из командной строки («Открыть с помощью», двойной щелчок по .xlsx).
/// Отдаётся интерфейсу один раз — первому окну.
struct StartupFile(Mutex<Option<String>>);

#[tauri::command]
fn startup_file(state: tauri::State<'_, StartupFile>) -> Option<String> {
    state.0.lock().ok().and_then(|mut f| f.take())
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_clipboard_manager::init())
        .setup(|app| {
            // Папка восстановления по умолчанию — в профиле пользователя.
            let root = app
                .path()
                .app_local_data_dir()
                .map(|p| p.join("recovery"))
                .unwrap_or_else(|_| std::env::temp_dir().join("ens-office-recovery"));
            app.manage(recovery::Recovery::new(root));
            let file = std::env::args_os().skip(1).map(std::path::PathBuf::from).find(|p| p.is_file());
            app.manage(StartupFile(Mutex::new(file.map(|p| p.display().to_string()))));
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            startup_file,
            files::inspect_file,
            files::read_file,
            files::save_file,
            recovery::recovery_heartbeat,
            recovery::recovery_set_root,
            recovery::recovery_root,
            recovery::recovery_begin,
            recovery::recovery_snapshot,
            recovery::recovery_append,
            recovery::recovery_forget,
            recovery::recovery_end_session,
            recovery::recovery_list,
            recovery::recovery_load_snapshot,
            recovery::recovery_load_journal,
            recovery::recovery_discard,
        ])
        .run(tauri::generate_context!())
        .expect("не удалось запустить ENS-Table");
}
