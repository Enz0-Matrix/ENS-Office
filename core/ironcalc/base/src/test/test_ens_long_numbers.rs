#![allow(clippy::unwrap_used)]

//! ENS-Office: длинные числа. Excel хранит 15 значащих цифр: лицевые счета, номера карт
//! и т. п. длиннее 15 цифр теряли бы цифры молча — такие записи остаются текстом.
//! Строка формул показывает число без экспоненты, как Excel.

use crate::cell::CellValue;
use crate::UserModel;

#[test]
fn ens_long_numbers() {
    let mut m = UserModel::new_empty("Книга", "ru", "UTC", "ru").unwrap();
    let text = |s: &str| CellValue::String(s.to_string());
    let cases: [(&str, CellValue, &str); 8] = [
        ("40817810099910004312", text("40817810099910004312"), "'40817810099910004312"),
        ("4081 7810 0999 1000 4312", text("4081 7810 0999 1000 4312"), "'4081 7810 0999 1000 4312"),
        ("1234567890123456", text("1234567890123456"), "'1234567890123456"),
        ("123456789012345", CellValue::Number(123456789012345.0), "123456789012345"),
        ("1010101010101", CellValue::Number(1010101010101.0), "1010101010101"),
        ("0012345", CellValue::Number(12345.0), "12345"),
        ("1e17", CellValue::Number(1e17), "100000000000000000"),
        ("-1,5e18", CellValue::Number(-1.5e18), "-1500000000000000000"),
    ];
    for (r, (input, value, content)) in cases.iter().enumerate() {
        let row = r as i32 + 1;
        m.set_user_input(0, row, 1, input).unwrap();
        assert_eq!(&m.get_model().get_cell_value_by_index(0, row, 1).unwrap(), value, "{input}");
        assert_eq!(m.get_cell_content(0, row, 1).unwrap(), *content, "{input}");
    }
    // формула с огромным результатом: в строке формул — формула, значение без потерь сверх 15 цифр
    m.set_user_input(0, 20, 1, "=654785313541516000+500").unwrap();
    assert_eq!(m.get_model().get_cell_value_by_index(0, 20, 1).unwrap(), CellValue::Number(654785313541516500.0));
}
