//! ENS-Table: оболочка Tauri. Движок таблиц работает в окне (WebAssembly);
//! здесь — то, чего нельзя сделать в браузере: надёжная работа с файлами
//! (files.rs) и восстановление после сбоя (recovery.rs).

mod files;
mod recovery;

use tauri::Manager;

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .setup(|app| {
            // Папка восстановления по умолчанию — в профиле пользователя.
            let root = app
                .path()
                .app_local_data_dir()
                .map(|p| p.join("recovery"))
                .unwrap_or_else(|_| std::env::temp_dir().join("ens-office-recovery"));
            app.manage(recovery::Recovery::new(root));
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
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
