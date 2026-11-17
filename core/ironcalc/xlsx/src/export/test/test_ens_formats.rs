// SPDX-License-Identifier: AGPL-3.0-or-later
//! ENS-Office: «всеядность» — `.xls`, `.ods` (сделаны LibreOffice из одной книги) и `.csv`.

use crate::export::save_to_xlsx;
use crate::import::{load_from_csv_bytes, load_from_other_bytes, load_from_xlsx};

const DIR: &str = "tests/ens/formats";

fn check_book(m: &ironcalc_base::Model, what: &str) {
    let v = |s: u32, r: i32, c: i32| m.get_formatted_cell_value(s, r, c).unwrap();
    let props = m.get_worksheets_properties();
    assert_eq!(props[0].name, "Топливо", "{what}");
    assert_eq!(props[1].name, "Второй лист", "{what}");
    assert_eq!(v(0, 2, 1), "АИ-92", "{what}");
    assert_eq!(v(0, 2, 2), "52,3", "{what}");
    assert_eq!(v(0, 2, 4), "01.11.2026", "{what}: дата");
    assert_eq!(v(0, 4, 5), "289880", "{what}: СУММ по произведениям");
    assert_eq!(v(0, 2, 6), "66,8", "{what}: ВПР");
    assert_eq!(v(0, 3, 6), "14", "{what}: разность дат");
    assert_eq!(v(0, 4, 6), "42", "{what}: ссылка на лист с пробелом");
    assert_eq!(
        m.get_localized_cell_content(0, 4, 5).unwrap(),
        "=СУММ(E2:E3)",
        "{what}: формула сохранилась"
    );
    assert_eq!(v(1, 1, 2), "текст", "{what}");
    assert!(
        !m.workbook.ens_lost_features.is_empty(),
        "{what}: предупреждение об оформлении"
    );
}

#[test]
fn test_ens_open_xls_and_ods() {
    for file in ["книга.xls", "книга.ods"] {
        let bytes = std::fs::read(format!("{DIR}/{file}")).unwrap();
        let m = load_from_other_bytes(&bytes, "книга", "ru", "UTC", "ru").unwrap();
        check_book(&m, file);
        // и сохраняется в xlsx без потерь значений
        let out = format!(
            "temp_ens_{}.xlsx",
            if file.ends_with("xls") { "xls" } else { "ods" }
        );
        save_to_xlsx(&m, &out).unwrap();
        let mut back = load_from_xlsx(&out, "ru", "UTC", "ru").unwrap();
        std::fs::remove_file(&out).unwrap();
        back.evaluate();
        assert_eq!(
            back.get_formatted_cell_value(0, 4, 5).unwrap(),
            "289880",
            "{file} → xlsx"
        );
    }
}

#[test]
fn test_ens_open_csv() {
    let m = load_from_csv_bytes(
        &std::fs::read(format!("{DIR}/utf8.csv")).unwrap(),
        "закупки",
        "ru",
        "UTC",
        "ru",
    )
    .unwrap();
    let v = |r: i32, c: i32| m.get_formatted_cell_value(0, r, c).unwrap();
    assert_eq!(v(1, 1), "Вид");
    assert_eq!(v(2, 2), "52,3");
    assert_eq!(
        m.get_cell_value_by_index(0, 2, 3).unwrap(),
        ironcalc_base::cell::CellValue::Number(1200.0)
    );
    assert_eq!(v(2, 4), "01.11.2026");
    assert_eq!(v(3, 1), "ДТ \"зимнее\"");
    let m = load_from_csv_bytes(
        &std::fs::read(format!("{DIR}/cp1251.csv")).unwrap(),
        "газ",
        "ru",
        "UTC",
        "ru",
    )
    .unwrap();
    assert_eq!(m.get_formatted_cell_value(0, 2, 1).unwrap(), "Газ");
    assert_eq!(m.get_formatted_cell_value(0, 2, 2).unwrap(), "28,1");
}

#[test]
fn test_ens_csv_round_trip() {
    let m = load_from_csv_bytes(
        &std::fs::read(format!("{DIR}/utf8.csv")).unwrap(),
        "закупки",
        "ru",
        "UTC",
        "ru",
    )
    .unwrap();
    let um = ironcalc_base::UserModel::from_model(m);
    let csv = um.sheet_to_csv(0).unwrap();
    let text = String::from_utf8(csv.clone()).unwrap();
    assert!(text.starts_with('\u{feff}'));
    // «1 200» введено с разрядами: формат с разрядами сохраняется, как в Excel
    assert!(
        text.contains("АИ-92;52,3;1") && text.contains("200;01.11.2026\r\n"),
        "{text}"
    );
    assert!(text.contains("\"ДТ \"\"зимнее\"\"\";66,8"), "{text}");
    let back = load_from_csv_bytes(&csv, "снова", "ru", "UTC", "ru").unwrap();
    assert_eq!(
        back.get_formatted_cell_value(0, 3, 1).unwrap(),
        "ДТ \"зимнее\""
    );
    assert_eq!(
        back.get_formatted_cell_value(0, 2, 4).unwrap(),
        "01.11.2026"
    );
    assert_eq!(
        back.get_cell_value_by_index(0, 2, 3).unwrap(),
        ironcalc_base::cell::CellValue::Number(1200.0)
    );
}
