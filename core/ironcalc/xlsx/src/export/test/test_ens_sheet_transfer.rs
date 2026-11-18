// SPDX-License-Identifier: AGPL-3.0-or-later
//! ENS-Office: лист, перенесённый из другой книги (ADR-012), — «сохранить → открыть»:
//! ссылки на книгу-источник пишутся как внешние связи Excel, значения — их кэшем.

use ironcalc_base::{expressions::types::Area, Model, UserModel};

use crate::export::save_to_xlsx;
use crate::import::load_from_xlsx;

#[test]
fn ens_moved_sheet_round_trip() {
    let mut src = UserModel::new_empty("Цены", "ru", "UTC", "ru").unwrap();
    src.rename_sheet(0, "Цены").unwrap();
    src.new_sheet().unwrap();
    src.rename_sheet(1, "Склад").unwrap();
    src.set_user_input(1, 1, 1, "1200").unwrap();
    src.set_user_input(0, 1, 2, "57,1").unwrap();
    src.set_user_input(0, 2, 2, "=B1*Склад!A1").unwrap();
    src.set_user_input(0, 3, 1, "Итог").unwrap();
    let area = Area { sheet: 0, row: 3, column: 1, width: 1, height: 1 };
    src.update_range_style(&area, "font.i", "true").unwrap();
    let pkg = src.ens_export_sheet(0, "Цены.xlsx", r"D:\Работа\Цены.xlsx").unwrap();

    let mut m = UserModel::new_empty("Отчёт", "ru", "UTC", "ru").unwrap();
    m.ens_import_sheet(&pkg, 1).unwrap();
    let out = "temp_ens_moved_sheet.xlsx";
    save_to_xlsx(m.get_model(), out).unwrap();
    let back: Model = load_from_xlsx(out, "ru", "UTC", "ru").unwrap();
    std::fs::remove_file(out).unwrap();
    let names = back.workbook.get_worksheet_names();
    assert_eq!(names[..2], ["Sheet1".to_string(), "Цены".to_string()][..]);
    assert_eq!(back.get_formatted_cell_value(1, 2, 2).unwrap(), "68520");
    assert_eq!(back.get_localized_cell_content(1, 2, 2).unwrap(), "=B1*'[Цены.xlsx]Склад'!A1");
    assert_eq!(back.get_formatted_cell_value(1, 3, 1).unwrap(), "Итог");
    assert!(back.get_style_for_cell(1, 3, 1).unwrap().font.i);
    assert!(back.workbook.ens_external_books.iter().any(|b| b.file == "Цены.xlsx"));
    assert!(back.workbook.ens_lost_features.is_empty(), "{:?}", back.workbook.ens_lost_features);
}
