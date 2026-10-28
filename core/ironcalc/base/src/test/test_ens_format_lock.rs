#![allow(clippy::unwrap_used)]

//! ENS-Office §4 (docs/03-table-ux-requirements.md): a format chosen for a
//! cell is kept; what is typed is read according to it.

use crate::UserModel;

fn model() -> UserModel<'static> {
    UserModel::new_empty("model", "en", "UTC", "en").unwrap()
}

fn area(row: i32, column: i32) -> crate::expressions::types::Area {
    crate::expressions::types::Area { sheet: 0, row, column, width: 1, height: 1 }
}

fn fmt(m: &UserModel, row: i32, column: i32) -> String {
    m.get_cell_style(0, row, column).unwrap().num_fmt
}

#[test]
fn ens_number_format_keeps_numbers() {
    let mut m = model();
    m.update_range_style(&area(1, 1), "num_fmt", "#,##0.00").unwrap();
    for (input, expected) in [("22.08", "22.08"), ("22,08", "22.08"), ("-3,5", "-3.50"), ("1/2", "")] {
        m.set_user_input(0, 1, 1, input).unwrap();
        assert_eq!(fmt(&m, 1, 1), "#,##0.00", "format changed by {input:?}");
        if !expected.is_empty() {
            assert_eq!(m.get_formatted_cell_value(0, 1, 1).unwrap(), expected, "{input:?}");
        }
    }
    // A percent or currency typed in gives its value, the format stays.
    m.set_user_input(0, 1, 1, "12%").unwrap();
    assert_eq!(fmt(&m, 1, 1), "#,##0.00");
    assert_eq!(m.get_formatted_cell_value(0, 1, 1).unwrap(), "0.12");
    m.set_user_input(0, 1, 1, "$5").unwrap();
    assert_eq!(fmt(&m, 1, 1), "#,##0.00");
    assert_eq!(m.get_formatted_cell_value(0, 1, 1).unwrap(), "5.00");
}

#[test]
fn ens_text_format_keeps_text() {
    let mut m = model();
    m.update_range_style(&area(1, 1), "num_fmt", "@").unwrap();
    m.set_user_input(0, 1, 1, "00123").unwrap();
    assert_eq!(m.get_formatted_cell_value(0, 1, 1).unwrap(), "00123");
    m.set_user_input(0, 1, 1, "=1+1").unwrap();
    assert_eq!(m.get_formatted_cell_value(0, 1, 1).unwrap(), "=1+1");
    assert_eq!(fmt(&m, 1, 1), "@");
}

#[test]
fn ens_general_cells_still_recognize() {
    let mut m = model();
    m.set_user_input(0, 1, 1, "12%").unwrap();
    assert_eq!(fmt(&m, 1, 1), "#,##0%");
    m.set_user_input(0, 2, 1, "1/2/2026").unwrap();
    assert!(fmt(&m, 2, 1).contains('d') || fmt(&m, 2, 1).contains('m'), "{}", fmt(&m, 2, 1));
    m.set_user_input(0, 3, 1, "007").unwrap();
    assert_eq!(m.get_formatted_cell_value(0, 3, 1).unwrap(), "7");
}

#[test]
fn ens_date_recognition_can_be_switched_off() {
    let mut m = model();
    m.set_recognize_dates(false);
    m.set_user_input(0, 1, 1, "1/2/2026").unwrap();
    assert_eq!(fmt(&m, 1, 1), "general");
    assert_eq!(m.get_formatted_cell_value(0, 1, 1).unwrap(), "1/2/2026");
    // Other recognition (percent) still works.
    m.set_user_input(0, 2, 1, "12%").unwrap();
    assert_eq!(fmt(&m, 2, 1), "#,##0%");
}

#[test]
fn ens_date_cells_take_dates() {
    let mut m = model();
    m.update_range_style(&area(1, 1), "num_fmt", "dd.mm.yyyy").unwrap();
    m.set_user_input(0, 1, 1, "1/2/2026").unwrap();
    assert_eq!(fmt(&m, 1, 1), "dd.mm.yyyy");
    assert_eq!(m.get_formatted_cell_value(0, 1, 1).unwrap(), "02.01.2026");
}
