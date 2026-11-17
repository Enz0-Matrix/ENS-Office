#![allow(clippy::unwrap_used)]

//! ENS-Office: число с пробелами в разрядах (обычный, неразрывный, узкий неразрывный пробел),
//! как его вставляют из Word, 1С, браузера, — это число.

use crate::cell::CellValue;
use crate::UserModel;

#[test]
fn ens_numbers_with_group_spaces() {
    let mut m = UserModel::new_empty("Книга", "ru", "UTC", "ru").unwrap();
    let cases = [
        ("1 234 567", 1234567.0),
        ("1\u{a0}234\u{a0}567", 1234567.0),
        ("1\u{202f}234,5", 1234.5),
        ("12 345,67", 12345.67),
        ("-1 234", -1234.0),
        ("  1 234  ", 1234.0),
        ("1 234 ₽", 1234.0),
        ("1\u{a0}234,00\u{a0}₽", 1234.0),
    ];
    for (i, (text, value)) in cases.iter().enumerate() {
        let row = i as i32 + 1;
        m.set_user_input(0, row, 1, text).unwrap();
        assert_eq!(m.get_model().get_cell_value_by_index(0, row, 1).unwrap(), CellValue::Number(*value), "{text:?}");
    }
    // а это не числа
    for (row, text) in [(20, "1 23 4"), (21, "12 3456"), (22, "Дом 12 345")] {
        m.set_user_input(0, row, 1, text).unwrap();
        assert!(matches!(m.get_model().get_cell_value_by_index(0, row, 1).unwrap(), CellValue::String(_)), "{text:?}");
    }
}

#[test]
fn ens_paste_numbers_with_group_spaces() {
    let mut m = UserModel::new_empty("Книга", "ru", "UTC", "ru").unwrap();
    let area = |row, column| crate::expressions::types::Area { sheet: 0, row, column, width: 1, height: 1 };
    // форматы ячеек, куда вставляют: общий, числовой с разрядами, денежный, дата
    m.update_range_style(&area(1, 2), "num_fmt", "#,##0").unwrap();
    m.update_range_style(&area(1, 3), "num_fmt", "#,##0.00\\ \"₽\"").unwrap();
    m.update_range_style(&area(1, 4), "num_fmt", "0.00").unwrap();
    let text = "1 234 567\t1\u{a0}234\u{a0}567\t1\u{2009}234,50\t1\u{2007}234\n12\u{202f}345,6\t-2 000\t 3 000 \t1 234 ₽";
    m.paste_csv_string(&crate::expressions::types::Area { sheet: 0, row: 1, column: 1, width: 4, height: 2 }, text).unwrap();
    let expect = [[1234567.0, 1234567.0, 1234.5, 1234.0], [12345.6, -2000.0, 3000.0, 1234.0]];
    for (r, row) in expect.iter().enumerate() {
        for (c, v) in row.iter().enumerate() {
            let got = m.get_model().get_cell_value_by_index(0, r as i32 + 1, c as i32 + 1).unwrap();
            assert_eq!(got, CellValue::Number(*v), "R{}C{}", r + 1, c + 1);
        }
    }
}

/// Вставка текста из другой программы, когда активная ячейка модели — в другом месте:
/// раньше вставка обрывалась ошибкой выделения, и значение было не видно до Enter.
#[test]
fn ens_paste_text_away_from_selected_cell() {
    let mut m = UserModel::new_empty("Книга", "ru", "UTC", "ru").unwrap();
    m.set_user_input(0, 1, 2, "=B5*2").unwrap();
    let area = crate::expressions::types::Area { sheet: 0, row: 5, column: 2, width: 1, height: 1 };
    m.paste_csv_string(&area, "21").unwrap();
    assert_eq!(m.get_formatted_cell_value(0, 5, 2).unwrap(), "21");
    assert_eq!(m.get_formatted_cell_value(0, 1, 2).unwrap(), "42"); // и пересчёт прошёл
    assert_eq!(m.get_selected_cell(), (0, 5, 2));
}

/// Скрыть несколько листов одним шагом и отменить его.
#[test]
fn ens_hide_many_sheets_undo() {
    let mut m = UserModel::new_empty("Книга", "ru", "UTC", "ru").unwrap();
    for _ in 0..3 {
        m.new_sheet().unwrap();
    }
    let mark = m.undo_mark();
    m.hide_sheet(1).unwrap();
    m.hide_sheet(2).unwrap();
    m.merge_undo_since(mark);
    let states = |m: &UserModel| m.get_worksheets_properties().iter().map(|p| p.state.clone()).collect::<Vec<_>>();
    assert_eq!(states(&m), ["visible", "hidden", "hidden", "visible"]);
    m.undo().unwrap();
    assert_eq!(states(&m), ["visible", "visible", "visible", "visible"]);
    m.redo().unwrap();
    assert_eq!(states(&m), ["visible", "hidden", "hidden", "visible"]);
}
