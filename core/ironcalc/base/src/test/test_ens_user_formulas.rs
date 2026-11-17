#![allow(clippy::unwrap_used)]

//! ENS-Office: формулы из реальных книг пользователя (Р7/Excel), которые не работали.

use crate::UserModel;

fn book() -> UserModel<'static> {
    let mut m = UserModel::new_empty("Книга", "ru", "UTC", "ru").unwrap();
    m.new_sheet().unwrap();
    m.rename_sheet(1, "01.11.2026").unwrap();
    m.set_user_input(1, 5, 15, "123,6").unwrap(); // O5
    m.set_user_input(1, 68, 17, "7,4").unwrap(); // Q68
    m.set_user_input(1, 24, 3, "77").unwrap(); // C24
    m.set_user_input(0, 2, 24, "01.11.2026").unwrap(); // X2
    m.set_user_input(0, 85, 3, "01.11.2026").unwrap(); // C85
    m.set_user_input(0, 3, 7, "15").unwrap(); // G3
    m
}

fn val(m: &UserModel, r: i32, c: i32) -> String {
    m.get_formatted_cell_value(0, r, c).unwrap()
}

#[test]
fn ens_cell_filename_russian() {
    let mut m = book();
    m.set_user_input(0, 1, 1, "=\"Отчёт на \"&СЖПРОБЕЛЫ(ПСТР(ЯЧЕЙКА(\"имяфайла\";$A$1);НАЙТИ(\"]\";ЯЧЕЙКА(\"имяфайла\";$A$1))+1;255))").unwrap();
    m.set_user_input(0, 2, 1, "=ЯЧЕЙКА(\"имяфайла\";$A$1)")
        .unwrap();
    m.set_user_input(0, 3, 1, "=ЯЧЕЙКА(\"filename\";$A$1)")
        .unwrap();
    m.set_user_input(0, 4, 1, "=ЯЧЕЙКА(\"строка\";G3)").unwrap();
    let sheet = m.get_worksheets_properties()[0].name.clone();
    assert_eq!(val(&m, 1, 1), format!("Отчёт на {sheet}"));
    assert_eq!(val(&m, 2, 1), format!("[Книга.xlsx]{sheet}"));
    assert_eq!(val(&m, 3, 1), val(&m, 2, 1));
    assert_eq!(val(&m, 4, 1), "3");
}

#[test]
fn ens_text_russian_date_codes() {
    let mut m = book();
    m.set_user_input(0, 1, 1, "=ТЕКСТ($X$2;\"ДД.ММ.ГГГГ\")")
        .unwrap();
    m.set_user_input(0, 2, 1, "=ТЕКСТ($X$2;\"dd.mm.yyyy\")")
        .unwrap();
    m.set_user_input(0, 3, 1, "=ТЕКСТ($X$2;\"дд.мм.гг\")")
        .unwrap();
    m.set_user_input(0, 4, 1, "=ТЕКСТ(1234,5;\"# ##0,00\")")
        .unwrap();
    m.set_user_input(0, 5, 1, "=ТЕКСТ(0,75;\"чч:мм\")").unwrap();
    m.set_user_input(0, 6, 1, "=ТЕКСТ(-5;\"0;[Красный]-0\")")
        .unwrap();
    assert_eq!(val(&m, 1, 1), "01.11.2026");
    assert_eq!(val(&m, 2, 1), "01.11.2026");
    assert_eq!(val(&m, 3, 1), "01.11.26");
    assert_eq!(val(&m, 4, 1), "1\u{a0}234,50");
    assert_eq!(val(&m, 5, 1), "18:00");
    assert_eq!(val(&m, 6, 1), "-5");
}

#[test]
fn ens_indirect_by_date_sheet() {
    let mut m = book();
    m.set_user_input(
        0,
        1,
        1,
        "=ОКРУГЛ(ДВССЫЛ(\"'\"&ТЕКСТ($X$2;\"ДД.ММ.ГГГГ\")&\"'!O5\");0)",
    )
    .unwrap();
    m.set_user_input(
        0,
        2,
        1,
        "=ОКРУГЛ(ДВССЫЛ(\"'\"&ТЕКСТ($X$2;\"ДД.ММ.ГГГГ\")&\"'!Q68\");0)",
    )
    .unwrap();
    // строка 87: СТРОКА()-СТРОКА($C$85)=2 → ВЫБОР даёт 24 → '01.11.2026'!C24
    m.set_user_input(0, 87, 3, "=@ДВССЫЛ(\"'\"&ТЕКСТ($C$85;\"ДД.ММ.ГГГГ\")&\"'!\"&АДРЕС(ВЫБОР(СТРОКА()-СТРОКА($C$85);5;24;45;68);СТОЛБЕЦ();4))").unwrap();
    assert_eq!(val(&m, 1, 1), "124");
    assert_eq!(val(&m, 2, 1), "7");
    assert_eq!(val(&m, 87, 3), "77");
}

#[test]
fn ens_comma_as_argument_separator_in_russian() {
    let mut m = book();
    m.set_user_input(0, 1, 1, "=СЦЕПИТЬ(G3,\" л/час\")")
        .unwrap();
    m.set_user_input(0, 2, 1, "=СУММ(G3, 1,5)").unwrap();
    m.set_user_input(0, 3, 1, "=ОКРУГЛ(2,567;1)").unwrap();
    assert_eq!(val(&m, 1, 1), "15 л/час");
    assert_eq!(
        m.get_cell_content(0, 1, 1).unwrap(),
        "=СЦЕПИТЬ(G3;\" л/час\")"
    );
    assert_eq!(val(&m, 2, 1), "16,5");
    assert_eq!(val(&m, 3, 1), "2,6");
}
