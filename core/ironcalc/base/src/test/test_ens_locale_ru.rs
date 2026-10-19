#![allow(clippy::unwrap_used)]

//! ENS-Office: the Russian locale (decimal comma, ";" between arguments,
//! dd.mm.yyyy dates).

use crate::UserModel;

fn model() -> UserModel<'static> {
    UserModel::new_empty("model", "ru", "UTC", "en").unwrap()
}

fn text(m: &UserModel, row: i32, column: i32) -> String {
    m.get_formatted_cell_value(0, row, column).unwrap()
}

#[test]
fn ens_ru_decimal_comma() {
    let mut m = model();
    m.set_user_input(0, 1, 1, "1,5").unwrap();
    m.set_user_input(0, 2, 1, "2,25").unwrap();
    m.set_user_input(0, 3, 1, "=SUM(A1:A2;0,25)").unwrap();
    m.set_user_input(0, 4, 1, "=ROUND(A3/3;2)").unwrap();
    assert_eq!(text(&m, 1, 1), "1,5");
    assert_eq!(text(&m, 3, 1), "4");
    assert_eq!(text(&m, 4, 1), "1,33");
    assert_eq!(m.get_cell_content(0, 3, 1).unwrap(), "=SUM(A1:A2;0,25)");
}

#[test]
fn ens_ru_dates_and_numbers() {
    let mut m = model();
    m.set_user_input(0, 1, 1, "22.08.2026").unwrap();
    assert_eq!(text(&m, 1, 1), "22.08.2026");
    m.set_user_input(0, 2, 1, "=A1+1").unwrap();
    m.update_range_style(
        &crate::expressions::types::Area { sheet: 0, row: 2, column: 1, width: 1, height: 1 },
        "num_fmt",
        "dd.mm.yyyy",
    )
    .unwrap();
    assert_eq!(text(&m, 2, 1), "23.08.2026");
    // §4: in a number cell "22.08" is the number 22,08 even in this locale.
    m.update_range_style(
        &crate::expressions::types::Area { sheet: 0, row: 3, column: 1, width: 1, height: 1 },
        "num_fmt",
        "#,##0.00",
    )
    .unwrap();
    m.set_user_input(0, 3, 1, "22.08").unwrap();
    assert_eq!(text(&m, 3, 1), "22,08");
    m.set_user_input(0, 4, 1, "1234567,891").unwrap();
    m.update_range_style(
        &crate::expressions::types::Area { sheet: 0, row: 4, column: 1, width: 1, height: 1 },
        "num_fmt",
        "#,##0.00",
    )
    .unwrap();
    assert_eq!(text(&m, 4, 1), "1\u{a0}234\u{a0}567,89");
}
