//! ENS-Tekst: оболочка Tauri. Это та же оболочка, что у ENS-Table (app/table/src-tauri/src/lib.rs):
//! окно, надёжная работа с файлами (files.rs, ADR-005) и восстановление после сбоя (recovery.rs).
//! Различаются интерфейс (app/text) и tauri.conf.json: название, значок, типы файлов.

#[path = "../../../table/src-tauri/src/lib.rs"]
mod shell;

pub use shell::run;
