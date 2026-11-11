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

## Компоненты ENS-Text (добавлено 12.11.2026)

| Компонент | Версия | Для чего | Лицензия | Авторы |
|---|---|---|---|---|
| prosemirror-model | 1.25.4 | модель документа | MIT | Marijn Haverbeke и участники |
| prosemirror-view | 1.41.3 | отображение и ввод | MIT | Marijn Haverbeke и участники |
| prosemirror-state | 1.4.4 | состояние редактора | MIT | Marijn Haverbeke и участники |
| prosemirror-transform | 1.10.5 | преобразования документа | MIT | Marijn Haverbeke и участники |
| prosemirror-commands | 1.7.1 | команды редактирования | MIT | Marijn Haverbeke и участники |
| prosemirror-keymap | 1.2.3 | горячие клавиши | MIT | Marijn Haverbeke и участники |
| prosemirror-history | 1.5.0 | отмена и повтор | MIT | Marijn Haverbeke и участники |
| prosemirror-schema-list | 1.5.1 | списки | MIT | Marijn Haverbeke и участники |
| prosemirror-tables | 1.8.5 | таблицы в тексте | MIT | Marijn Haverbeke, Atlassian и участники |
| prosemirror-dropcursor | 1.8.2 | курсор перетаскивания | MIT | Marijn Haverbeke и участники |
| prosemirror-gapcursor | 1.4.0 | курсор в пустых местах | MIT | Marijn Haverbeke и участники |
| fflate | 0.8.3 | упаковка ZIP для docx | MIT | Arjun Barrett |
| @fontsource/roboto-condensed | 5.3.0 | инженерный шрифт (аналог Bahnschrift) | OFL-1.1 (шрифт), MIT (упаковка) | Google Inc.; Fontsource |

## Инструменты сборки (в распространяемые копии не входят)

| Компонент | Версия | Лицензия |
|---|---|---|
| Vite | 8.3.2 | MIT |
| TypeScript | 7.0.2 | Apache-2.0 |
| rolldown | 1.2.x | MIT |
| lightningcss | 1.33.0 | MPL-2.0 |
| postcss | 8.5.x | MIT |
| picocolors | 1.1.1 | ISC |
| source-map-js | 1.2.2 | BSD-3-Clause |
| nanoid, picomatch, fdir, tinyglobby | — | MIT |
| Playwright (только тесты) | 1.63.0 | Apache-2.0 |
| LibreOffice Writer (только тесты в CI) | — | MPL-2.0 |

MPL-2.0 у lightningcss относится к инструменту сборки CSS и не
накладывает обязательств на распространяемый продукт: в сборку
попадает результат обработки (CSS), а не сам компонент.

## Тексты лицензий в распространяемых копиях

Полные тексты прилагаются в каталоге `licenses/` каждой сборки:
`LICENSE` (AGPL-3.0), `NOTICE`, `THIRD_PARTY_NOTICES.md`,
`OFL-1.1.txt`, `MIT-fluent-icons.txt`, `LICENSE-MIT` и
`LICENSE-Apache-2.0` (IronCalc), а также лицензии npm- и cargo-пакетов.

## Системные компоненты

| Компонент | Лицензия | Как используется |
|---|---|---|
| WebKitGTK (Linux) | LGPL-2.1+ | системная библиотека, динамическое связывание; в установщик не входит (кроме AppImage) |
| WebView2 (Windows) | условия Microsoft | системный компонент; для организаций собирается offline-вариант |
| Chromium (РЕД ОС 7) | BSD-3-Clause и др. | установлен в системе, используется в режиме приложения; в пакет не входит |
