// В релизной сборке под Windows не открывать лишнее окно консоли.
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

fn main() {
    // Без графического экрана (например, вход по SSH без -X) GTK не запустится:
    // объясняем по-русски, вместо непонятной трассировки.
    #[cfg(target_os = "linux")]
    if std::env::var_os("DISPLAY").is_none() && std::env::var_os("WAYLAND_DISPLAY").is_none() {
        eprintln!(
            "ENS-Text: нет графического экрана.
             Запустите программу на рабочем столе или подключитесь к компьютеру
             с пробросом окон: ssh -X пользователь@адрес (на сервере нужны
             X11Forwarding yes в /etc/ssh/sshd_config и пакет xorg-x11-xauth)."
        );
        std::process::exit(2);
    }
    ens_text_lib::run()
}
