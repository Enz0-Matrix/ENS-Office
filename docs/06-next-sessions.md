# Задания для новых сессий

Как начать: открыть новую сессию Claude Code на репозитории `Enz0-Matrix/ENS-Office`, ветка `app/feedback-1`,
и написать: «Прочитай docs/06-next-sessions.md и выполни задание 1» (или 2).

Общее для обоих заданий:
- Работать в ветке `app/feedback-1`; PR не создавать, пока не попросят. Пишем по-русски, отчёты — короткие.
- Контекст — `docs/04-backlog.md`. Форк IronCalc лежит в `core/ironcalc`, менять минимально.
- Rust: `cd core/ironcalc && cargo test --workspace --exclude ironcalc_nodejs` (все тесты должны проходить);
  на каждую возможность — тест, в том числе круговой тест xlsx (`xlsx/src/export/test/test_ens_ru_export.rs`).
  Отмена — через `Diff` (образец: `SetPrintTitles` в `base/src/user_model/ens_tools.rs`, `history.rs`, `undo_redo.rs`, `scoped_undo.rs`).
- Интерфейс: `cd app/table && npm ci && npm run engine` (нужны Rust, `rustup target add wasm32-unknown-unknown`,
  `cargo install wasm-bindgen-cli --version 0.2.126`), затем `npx vite` и проверка в Chromium (Playwright);
  типы: `npx tsc --noEmit`.
- Сборка: workflow `ci.yml` запускается вручную на ветке; артефакты `ens-table-linux`, `ens-table-windows`.
- Версию (`src/version.ts`, `package.json`, `src-tauri/Cargo.toml`, `src-tauri/tauri.conf.json`, `redos7/Cargo.toml`
  и lock-файлы) менять только по просьбе выпустить версию.
- Результат — коммиты в `app/feedback-1` и отметки в `docs/04-backlog.md`.

## Задание 1. Комментарии и разрывы страниц (пункты 22 и 23)

1. **Комментарии к ячейкам.** В движке уже есть `Worksheet.comments` и чтение из xlsx
   (`xlsx/src/import/worksheets.rs`, `load_comments`), но при сохранении (`xlsx/src/export`) комментарии теряются.
   Нужно: API в `UserModel` (получить, задать, удалить; с отменой), экспорт в xlsx (`xl/commentsN.xml` и `vmlDrawing`,
   чтобы Excel показал примечания), привязки wasm (`core/ironcalc/bindings/wasm/src/lib.rs`), интерфейс в `app/table/src`:
   красный уголок в ячейке (canvas, `grid.ts`), всплывающее окно по наведению, создание, правка и удаление через
   контекстное меню и ленту, команда в `main.ts`.
2. **Разрывы страниц.** Вставить и убрать разрыв строки и столбца; в xlsx `rowBreaks` и `colBreaks`;
   учитывать в печати (`app/table/src/print.ts`, функции `layout` и `paginate`).

## Задание 2. Защита, изображения, диаграммы, rich text (пункты 19, 24, 25)

Это крупные независимые функции: делать отдельными коммитами, перед каждой кратко согласовать объём первой версии.

1. **Rich text** (форматирование части текста в ячейке, `<r>` в sharedStrings xlsx): хранение в движке, импорт и
   экспорт xlsx (`xlsx/src/import/shared_strings.rs` и экспорт), отрисовка на canvas (`grid.ts`, `drawText`),
   правка выделенной части в поле ввода.
2. **Защита** книги, листа, ячеек, со скрытием формул или без (`sheetProtection`, `protection` в стилях xlsx; пароль
   можно отложить): флаги в модели, запрет правок защищённых ячеек в `UserModel`, интерфейс (защитить и снять лист,
   «Формат ячеек» → вкладка «Защита»).
3. **Изображения, фигуры, надписи, диаграммы** со своими шаблонами: сначала изображения (drawing и media в xlsx),
   затем простые диаграммы (столбцы, линии, круговая).

## Задание 3. ENS-Tekst: оболочки и колонтитулы (после версии 0.1)

Контекст — ADR-009 и раздел «ENS-Tekst» в `docs/04-backlog.md`; приложение — `app/text` (проверки: `npm test`).
1. ~~Оболочки~~ — сделано (ADR-009). Осталось проверить установщики на живых Windows и «Муроме» и выпустить ENS-Tekst в Releases.
2. **Предпросмотр печати** — окном (сама печать идёт по нашей вёрстке, см. ADR-009).
3. ~~Колонтитулы и номера страниц~~ — сделано; остались чётные/нечётные, разделы, поля даты, сноски (см. бэклог).
