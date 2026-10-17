#![allow(clippy::unwrap_used)]

//! ENS-Office: совместимость с Excel на случаях, найденных в прототипе №1
//! (docs/research/proto-01-grid-bench.md). Ожидаемые значения — из Excel.

use crate::test::util::new_empty_model;

#[test]
fn ens_round_half_cases() {
    let mut model = new_empty_model();
    model._set("A1", "=ROUND(2.675,2)");
    model._set("A2", "=ROUND(1.005,2)");
    model._set("A3", "=ROUND(0.285,2)");
    model._set("A4", "=ROUND(5.015,2)");
    model._set("A5", "=ROUND(-2.5,0)");
    model._set("A6", "=ROUND(-0.285,2)");
    model._set("A7", "=MROUND(2.5,1)");
    model._set("A8", "=ROUND(1234.5678*0.1,2)");
    model._set("A9", "=ROUND(4.35*100/100,1)");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"2.68");
    assert_eq!(model._get_text("A2"), *"1.01");
    assert_eq!(model._get_text("A3"), *"0.29");
    assert_eq!(model._get_text("A4"), *"5.02");
    assert_eq!(model._get_text("A5"), *"-3");
    assert_eq!(model._get_text("A6"), *"-0.29");
    assert_eq!(model._get_text("A7"), *"3");
    assert_eq!(model._get_text("A8"), *"123.46");
    assert_eq!(model._get_text("A9"), *"4.4");
}

#[test]
fn ens_general_format_15_significant_digits() {
    // Excel в формате «Общий» показывает не больше 15 значащих цифр
    // (IronCalc 0.8.4 показывал 6366933.449999999).
    let mut model = new_empty_model();
    model._set("A1", "=6366933.45+0.0000000001");
    model._set("A2", "=0.1+0.2");
    model._set("A3", "=10.1-9.93");
    model._set("A4", "=SUM(0.1,0.2,0.3,0.4,0.5,0.6,0.7)");
    model._set("A5", "=1/3");
    // Найдено замером ens-bench: сумма 50 000 округлённых значений.
    model._set("A6", "=6366933.539999999");
    model._set("A7", "=1234567.123456789+0.0000000004");
    model.evaluate();
    assert_eq!(model._get_text("A6"), *"6366933.54");
    assert_eq!(model._get_text("A7"), *"1234567.12345679");

    assert_eq!(model._get_text("A1"), *"6366933.45");
    assert_eq!(model._get_text("A2"), *"0.3");
    assert_eq!(model._get_text("A3"), *"0.17");
    assert_eq!(model._get_text("A4"), *"2.8");
    // Как Excel при стандартной ширине столбца: «Общий» обрезает по ширине ячейки.
    assert_eq!(model._get_text("A5"), *"0.333333333");
}

#[test]
fn ens_sum_of_rounded_products() {
    // Сценарий прототипа: СУММ по ОКРУГЛ(кол-во*цена*ставка/100;2).
    let mut model = new_empty_model();
    model._set("A1", "=ROUND(37*65.43*12.57/100,2)");
    model._set("A2", "=ROUND(3*0.95*10.05/100,2)");
    model._set("A3", "=SUM(A1:A2)");
    model.evaluate();
    assert_eq!(model._get_text("A1"), *"304.31");
    assert_eq!(model._get_text("A2"), *"0.29");
    assert_eq!(model._get_text("A3"), *"304.6");
}
