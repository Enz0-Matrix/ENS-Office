//! ENS-Table: оболочка Tauri. Движок таблиц работает в окне (WebAssembly);
//! здесь — то, чего нельзя сделать в браузере: надёжная работа с файлами
//! (files.rs) и восстановление после сбоя (recovery.rs).

mod files;
mod history;
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

/// Скруглённые углы окна без системного заголовка (Windows 11). В Windows 10
/// и других системах вызов просто ничего не меняет.
fn round_corners<R: tauri::Runtime>(window: &tauri::WebviewWindow<R>) {
    #[cfg(windows)]
    {
        #[link(name = "dwmapi")]
        extern "system" {
            fn DwmSetWindowAttribute(hwnd: isize, attribute: u32, value: *const core::ffi::c_void, size: u32) -> i32;
        }
        const DWMWA_WINDOW_CORNER_PREFERENCE: u32 = 33;
        const DWMWCP_ROUND: u32 = 2;
        if let Ok(hwnd) = window.hwnd() {
            let pref = DWMWCP_ROUND;
            // SAFETY: hwnd — окно этого процесса, значение — u32 нужного размера.
            unsafe {
                DwmSetWindowAttribute(hwnd.0 as isize, DWMWA_WINDOW_CORNER_PREFERENCE, &pref as *const u32 as *const _, 4);
            }
        }
    }
    #[cfg(not(windows))]
    let _ = window;
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
            // Копии на диске (предыдущие версии файлов, закрытое без сохранения) — рядом с папкой восстановления.
            app.manage(history::History::new(root.with_file_name("history")));
            app.manage(recovery::Recovery::new(root));
            let file = std::env::args_os().skip(1).map(std::path::PathBuf::from).find(|p| p.is_file());
            app.manage(StartupFile(Mutex::new(file.map(|p| p.display().to_string()))));
            if let Some(w) = app.get_webview_window("main") {
                round_corners(&w);
            }
            Ok(())
        })
        .on_page_load(|webview, _| {
            // Окна дополнительных книг тоже со скруглёнными углами.
            if let Some(w) = webview.app_handle().get_webview_window(webview.label()) {
                round_corners(&w);
            }
        })
        .invoke_handler(tauri::generate_handler![
            startup_file,
            files::inspect_file,
            files::user_name,
            files::open_target,
            files::read_file,
            files::save_file,
            history::history_store,
            history::history_list,
            recovery::recovery_heartbeat,
            recovery::recovery_set_root,
            recovery::recovery_root,
            recovery::recovery_begin,
            recovery::recovery_snapshot,
            recovery::recovery_snapshot_xlsx,
            recovery::recovery_load_snapshot_xlsx,
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
