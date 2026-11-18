#![allow(clippy::unwrap_used)]

//! ENS-Office: «Переместить или скопировать лист» в другую книгу (ADR-012).

use crate::{expressions::types::Area, UserModel};

/// Книга-источник «Цены.xlsx»: «Цены» (оформленная ячейка, ссылки на себя и на «Склад»), «Склад».
fn source() -> UserModel<'static> {
    let mut m = UserModel::new_empty("Цены", "ru", "UTC", "ru").unwrap();
    m.rename_sheet(0, "Цены").unwrap();
    m.new_sheet().unwrap();
    m.rename_sheet(1, "Склад").unwrap();
    m.set_user_input(1, 1, 1, "1200").unwrap();
    m.set_user_input(1, 2, 1, "300").unwrap();
    m.set_user_input(0, 1, 1, "АИ-95").unwrap();
    m.set_user_input(0, 1, 2, "57,1").unwrap();
    m.set_user_input(0, 2, 2, "=B1*Склад!A1").unwrap();
    m.set_user_input(0, 3, 2, "=СУММ(Склад!A1:A2)+Цены!B1").unwrap();
    m.set_user_input(0, 4, 2, "=B1*2").unwrap();
    let area = Area { sheet: 0, row: 1, column: 1, width: 1, height: 1 };
    m.update_range_style(&area, "font.b", "true").unwrap();
    m.update_range_style(&area, "fill.fg_color", "#FFC000").unwrap();
    m.new_defined_name("Цена", Some(0), "Цены!$B$1").unwrap();
    m
}

#[test]
fn ens_sheet_moves_to_other_book_with_links() {
    let mut src = source();
    let bytes = src.ens_export_sheet(0, "Цены.xlsx", r"C:\Работа\Цены.xlsx").unwrap();
    // книга-источник не изменилась
    assert_eq!(src.get_cell_content(0, 2, 2).unwrap(), "=B1*Склад!A1");

    let mut m = UserModel::new_empty("Отчёт", "ru", "UTC", "ru").unwrap();
    m.rename_sheet(0, "Цены").unwrap(); // имя занято — новый лист получит «Цены (2)»
    m.set_user_input(0, 1, 1, "своё").unwrap();
    let index = m.ens_import_sheet(&bytes, 1).unwrap();
    assert_eq!(index, 1);
    let props = m.get_worksheets_properties();
    assert_eq!(props[1].name, "Цены (2)");
    assert_eq!(props[1].state, "visible");
    // значения и формулы; ссылки на другие листы источника — внешние, значения приложены
    assert_eq!(m.get_formatted_cell_value(1, 1, 1).unwrap(), "АИ-95");
    assert_eq!(m.get_formatted_cell_value(1, 2, 2).unwrap(), "68520");
    assert_eq!(m.get_cell_content(1, 2, 2).unwrap(), "=B1*'[Цены.xlsx]Склад'!A1");
    // явная ссылка листа на себя ведёт на новое имя
    assert_eq!(m.get_cell_content(1, 3, 2).unwrap(), "=СУММ('[Цены.xlsx]Склад'!A1:A2)+'Цены (2)'!B1");
    assert_eq!(m.get_formatted_cell_value(1, 3, 2).unwrap(), "1557,1");
    assert_eq!(m.get_formatted_cell_value(1, 4, 2).unwrap(), "114,2");
    // своя ячейка книги не тронута
    assert_eq!(m.get_formatted_cell_value(0, 1, 1).unwrap(), "своё");
    // оформление
    let style = m.get_cell_style(1, 1, 1).unwrap();
    assert!(style.font.b);
    assert_eq!(format!("{:?}", style.fill).contains("FFC000"), true);
    // путь книги-источника запомнен для «Обновить связи»
    assert!(m
        .get_model()
        .workbook
        .ens_external_books
        .iter()
        .any(|b| b.file == "Цены.xlsx" && b.target == r"C:\Работа\Цены.xlsx"));
    // имя листа перенесено
    assert!(m
        .get_defined_name_list()
        .iter()
        .any(|(n, scope, f)| n == "Цена" && *scope == Some(1) && f.contains("'Цены (2)'!$B$1")));

    // одна отмена убирает лист, повтор возвращает
    m.undo().unwrap();
    assert_eq!(m.get_worksheets_properties().iter().filter(|p| p.state == "visible").count(), 1);
    assert!(m.get_defined_name_list().iter().all(|(n, _, _)| n != "Цена"));
    m.redo().unwrap();
    assert_eq!(m.get_worksheets_properties()[1].name, "Цены (2)");
    assert_eq!(m.get_formatted_cell_value(1, 2, 2).unwrap(), "68520");
}

#[test]
fn ens_imported_sheet_replays_from_journal() {
    let mut src = source();
    let bytes = src.ens_export_sheet(0, "Цены.xlsx", "").unwrap();
    let mut m = UserModel::new_empty("Отчёт", "ru", "UTC", "ru").unwrap();
    let snapshot = m.to_bytes();
    m.ens_import_sheet(&bytes, 0).unwrap();
    let journal = m.flush_send_queue();
    // восстановление после сбоя: снимок до вставки + журнал
    let mut r = UserModel::from_bytes(&snapshot, "ru").unwrap();
    r.apply_external_diffs(&journal).unwrap();
    let props = r.get_worksheets_properties();
    assert_eq!(props[0].name, "Цены");
    assert_eq!(r.get_formatted_cell_value(0, 2, 2).unwrap(), "68520");
    assert!(r.get_cell_style(0, 1, 1).unwrap().font.b);
}
