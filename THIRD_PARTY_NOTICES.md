# Сторонние компоненты ENS-Office

ENS-Office использует свободные компоненты с разрешительными лицензиями. Их
лицензии допускают коммерческое использование, изменение и распространение в
составе ENS-Office при сохранении уведомлений об авторских правах.

| Компонент | Для чего | Лицензия | Авторы |
|---|---|---|---|
| IronCalc (форк в `core/ironcalc`, с доработками ENS-Office) | движок таблиц и чтение/запись xlsx | MIT или Apache-2.0 | Nicolás Hatcher и участники проекта IronCalc |
| Tauri 2, плагины dialog и clipboard-manager | окно программы, файлы, буфер обмена | MIT или Apache-2.0 | Tauri Programme within The Commons Conservancy |
| wasm-bindgen | связь движка с интерфейсом | MIT или Apache-2.0 | The Rust and WebAssembly Working Group |
| Библиотеки Rust (serde, zip и др., см. `Cargo.lock`) | служебные функции | MIT и/или Apache-2.0 | их авторы |
| Шрифты Carlito, Caladea, Arimo, Tinos, Cousine (пакеты @fontsource) | замена шрифтов Office с теми же размерами символов | SIL Open Font License 1.1 | tyPoland (Carlito), Huerta Tipográfica (Caladea), Steve Matteson для Google (Arimo, Tinos, Cousine) |
| Fluent UI System Icons | значки ленты | MIT | Microsoft Corporation |
| Vite, TypeScript | только сборка, в программу не входят | MIT / Apache-2.0 | их авторы |
| WebKitGTK (Linux), WebView2 (Windows) | отображение окна; системные компоненты, в установщик ENS-Office не входят (кроме AppImage) | LGPL-2.1 (WebKitGTK) / условия Microsoft (WebView2) | их авторы |

Полные тексты лицензий: `core/ironcalc/LICENSE-MIT`, `core/ironcalc/LICENSE-Apache-2.0`,
а также в каталогах соответствующих пакетов (`node_modules/*/LICENSE`, `~/.cargo/registry`).
