// В релизной сборке под Windows не открывать лишнее окно консоли.
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

fn main() {
    ens_table_lib::run()
}
