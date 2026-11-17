// SPDX-License-Identifier: AGPL-3.0-or-later
//! ENS-Office: книга в системе дат 1904 (Excel для Mac). Эталон — LibreOffice,
//! который поддерживает обе системы: 17.11.2026, 17.12.2026, 2026, 31.12.1999 18:00, 30.

use crate::export::save_to_xlsx;
use crate::import::load_from_xlsx;

fn check(m: &ironcalc_base::Model) {
    let v = |r| m.get_formatted_cell_value(0, r, 1).unwrap();
    assert_eq!(v(1), "17.11.2026");
    assert_eq!(v(2), "12:30");
    assert_eq!(v(3), "45000"); // число без формата даты не трогаем
    assert_eq!(v(4), "17.12.2026");
    assert_eq!(v(5), "2026");
    assert_eq!(v(6), "31.12.1999 18:00");
    assert_eq!(v(7), "30");
}

#[test]
fn test_ens_date1904_book() {
    let mut m = load_from_xlsx("tests/ens/date1904.xlsx", "ru", "UTC", "ru").unwrap();
    m.evaluate();
    check(&m);
    assert!(m
        .workbook
        .ens_lost_features
        .iter()
        .any(|f| f.starts_with("система дат 1904")));
    // Сохраняется в обычной системе 1900: даты те же, предупреждения больше нет
    let out = "temp_ens_date1904.xlsx";
    save_to_xlsx(&m, out).unwrap();
    let mut back = load_from_xlsx(out, "ru", "UTC", "ru").unwrap();
    std::fs::remove_file(out).unwrap();
    back.evaluate();
    check(&back);
    assert!(
        back.workbook.ens_lost_features.is_empty(),
        "{:?}",
        back.workbook.ens_lost_features
    );
}
