#![allow(clippy::unwrap_used)]

//! ENS-Office: ~70 most used Excel functions, typed with their Russian names,
//! give the same results as Russian Excel (decimal comma, «;» between
//! arguments). Lists: os-market.store «Топ-50 полезных формул», Udemy Excel
//! formula cheat sheet.

use crate::UserModel;

#[test]
fn ens_top50_functions_ru() {
    let mut m = UserModel::new_empty("model", "ru", "UTC", "ru").unwrap();
    let data: &[(i32, i32, &str)] = &[
        (1, 1, "10"), (2, 1, "20"), (3, 1, "30"), (4, 1, "40"), (5, 1, "50"),
        (1, 2, "яблоко"), (2, 2, "груша"), (3, 2, "яблоко"), (4, 2, "слива"), (5, 2, "груша"),
        (1, 3, " Привет   мир "), (2, 3, "иванов иван"),
        (1, 5, "3"), (2, 5, "1"), (3, 5, "2"),
        (1, 8, "1"), (1, 9, "2"), (1, 10, "3"), (2, 8, "x"), (2, 9, "y"), (2, 10, "z"),
    ];
    for (row, column, value) in data {
        m.set_user_input(0, *row, *column, value).unwrap();
    }
    let cases: &[(&str, &str)] = &[
        ("=СУММ(A1:A5)", "150"),
        ("=СРЗНАЧ(A1:A5)", "30"),
        ("=СЧЁТ(A1:B5)", "5"),
        ("=СЧЁТЗ(A1:B5)", "10"),
        ("=СЧИТАТЬПУСТОТЫ(A1:A7)", "2"),
        ("=МИН(A1:A5)", "10"),
        ("=МАКС(A1:A5)", "50"),
        ("=ОКРУГЛ(2,345;2)", "2,35"),
        ("=ОКРУГЛВВЕРХ(2,341;1)", "2,4"),
        ("=ОКРУГЛВНИЗ(2,349;1)", "2,3"),
        ("=ЕСЛИ(A1>5;\"да\";\"нет\")", "да"),
        ("=ЕСЛИОШИБКА(1/0;\"ошибка\")", "ошибка"),
        ("=И(A1>5;A2>15)", "ИСТИНА"),
        ("=ИЛИ(A1>50;A2>15)", "ИСТИНА"),
        ("=НЕ(ИСТИНА)", "ЛОЖЬ"),
        ("=СУММЕСЛИ(B1:B5;\"яблоко\";A1:A5)", "40"),
        ("=СУММЕСЛИМН(A1:A5;B1:B5;\"груша\";A1:A5;\">20\")", "50"),
        ("=СЧЁТЕСЛИ(B1:B5;\"груша\")", "2"),
        ("=СЧЁТЕСЛИМН(B1:B5;\"яблоко\";A1:A5;\">15\")", "1"),
        ("=СРЗНАЧЕСЛИ(B1:B5;\"яблоко\";A1:A5)", "20"),
        ("=МАКСЕСЛИ(A1:A5;B1:B5;\"груша\")", "50"),
        ("=МИНЕСЛИ(A1:A5;B1:B5;\"груша\")", "20"),
        ("=ВПР(30;A1:B5;2;ЛОЖЬ)", "яблоко"),
        ("=ГПР(2;H1:J2;2;ЛОЖЬ)", "y"),
        ("=ИНДЕКС(B1:B5;4)", "слива"),
        ("=ПОИСКПОЗ(40;A1:A5;0)", "4"),
        ("=ИНДЕКС(A1:A5;ПОИСКПОЗ(\"слива\";B1:B5;0))", "40"),
        ("=ПРОСМОТРX(\"слива\";B1:B5;A1:A5)", "40"),
        ("=СЦЕП(\"a\";\"b\";1)", "ab1"),
        ("=СЦЕПИТЬ(\"a\";\"b\")", "ab"),
        ("=\"a\"&\"b\"", "ab"),
        ("=ОБЪЕДИНИТЬ(\", \";ИСТИНА;\"a\";\"\";\"b\")", "a, b"),
        ("=ЛЕВСИМВ(\"Привет\";3)", "При"),
        ("=ПРАВСИМВ(\"Привет\";3)", "вет"),
        ("=ПСТР(\"Привет\";2;3)", "рив"),
        ("=ДЛСТР(\"Привет\")", "6"),
        ("=СЖПРОБЕЛЫ(C1)", "Привет мир"),
        ("=ПРОПИСН(\"абв\")", "АБВ"),
        ("=СТРОЧН(\"АБВ\")", "абв"),
        ("=ПРОПНАЧ(C2)", "Иванов Иван"),
        ("=НАЙТИ(\"и\";\"Привет\")", "3"),
        ("=ПОИСК(\"И\";\"Привет\")", "3"),
        ("=ПОДСТАВИТЬ(\"a-b-c\";\"-\";\"+\")", "a+b+c"),
        ("=ТЕКСТ(0,5;\"0%\")", "50%"),
        ("=ЗНАЧЕН(\"12,5\")", "12,5"),
        ("=ГОД(ДАТА(2024;3;15))", "2024"),
        ("=МЕСЯЦ(ДАТА(2024;3;15))", "3"),
        ("=ДЕНЬ(ДАТА(2024;3;15))", "15"),
        ("=ДЕНЬНЕД(ДАТА(2024;3;15);2)", "5"),
        ("=РАЗНДАТ(ДАТА(2024;1;1);ДАТА(2024;3;15);\"d\")", "74"),
        ("=ЧИСТРАБДНИ(ДАТА(2024;3;1);ДАТА(2024;3;31))", "21"),
        ("=ДЕНЬ(КОНМЕСЯЦА(ДАТА(2024;2;10);0))", "29"),
        ("=ДЕНЬ(ДАТАМЕС(ДАТА(2024;1;31);1))", "29"),
        // В тестах «сегодня» — 08.11.2022 (mock_time).
        ("=СЕГОДНЯ()=ДАТА(2022;11;8)", "ИСТИНА"),
        ("=ТДАТА()>=СЕГОДНЯ()", "ИСТИНА"),
        ("=ABS(-5)", "5"),
        ("=ОСТАТ(10;3)", "1"),
        ("=СТЕПЕНЬ(2;10)", "1024"),
        ("=КОРЕНЬ(16)", "4"),
        ("=ПРОИЗВЕД(2;3;4)", "24"),
        ("=СУММПРОИЗВ(A1:A2;A3:A4)", "1100"),
        ("=СЛУЧМЕЖДУ(1;1)", "1"),
        ("=РАНГ(30;A1:A5)", "3"),
        ("=МЕДИАНА(1;3;2;10)", "2,5"),
        ("=НАИБОЛЬШИЙ(A1:A5;2)", "40"),
        ("=НАИМЕНЬШИЙ(A1:A5;2)", "20"),
        ("=ЕПУСТО(A7)", "ИСТИНА"),
        ("=ЕЧИСЛО(A1)", "ИСТИНА"),
        ("=ОКРУГЛ(ПЛТ(0,1/12;12;-1000);2)", "87,92"),
        ("=ЦЕЛОЕ(-2,5)", "-3"),
        ("=ЕСЛИМН(A1>20;\"a\";ИСТИНА;\"b\")", "b"),
        ("=ПЕРЕКЛЮЧ(2;1;\"один\";2;\"два\")", "два"),
        ("=СЧЁТЗ(УНИК(B1:B5))", "3"),
        ("=ИНДЕКС(СОРТ(E1:E3);1)", "1"),
        ("=СУММ(ФИЛЬТР(A1:A5;A1:A5>25))", "120"),
        ("=СУММ(ТРАНСП(A1:A5))", "150"),
    ];
    for (i, (formula, _)) in cases.iter().enumerate() {
        m.set_user_input(0, i as i32 + 1, 26, formula).unwrap();
    }
    let mut failed = Vec::new();
    for (i, (formula, expected)) in cases.iter().enumerate() {
        let got = m.get_formatted_cell_value(0, i as i32 + 1, 26).unwrap();
        if got != *expected {
            failed.push(format!("{formula}: ожидалось «{expected}», получено «{got}»"));
        }
    }
    assert!(failed.is_empty(), "{} из {} не совпали:\n{}", failed.len(), cases.len(), failed.join("\n"));
}

/// ENS-Office: halves are rounded away from zero when a number is shown, as in
/// Excel, and TEXT takes the format in Russian notation ("# ##0,00").
#[test]
fn ens_display_rounding_and_text_format_ru() {
    let mut m = UserModel::new_empty("model", "ru", "UTC", "ru").unwrap();
    m.set_user_input(0, 1, 1, "1234,5").unwrap();
    m.set_user_input(0, 2, 1, "2,5").unwrap();
    m.set_user_input(0, 3, 1, "2,345").unwrap();
    m.set_user_input(0, 4, 1, "-0,125").unwrap();
    let area = |r| crate::expressions::types::Area { sheet: 0, row: r, column: 1, width: 1, height: 1 };
    m.update_range_style(&area(1), "num_fmt", "0").unwrap();
    m.update_range_style(&area(2), "num_fmt", "0").unwrap();
    m.update_range_style(&area(3), "num_fmt", "0.00").unwrap();
    m.update_range_style(&area(4), "num_fmt", "0.00").unwrap();
    assert_eq!(m.get_formatted_cell_value(0, 1, 1).unwrap(), "1235");
    assert_eq!(m.get_formatted_cell_value(0, 2, 1).unwrap(), "3");
    assert_eq!(m.get_formatted_cell_value(0, 3, 1).unwrap(), "2,35");
    assert_eq!(m.get_formatted_cell_value(0, 4, 1).unwrap(), "-0,13");
    m.set_user_input(0, 1, 2, "=ТЕКСТ(A1; \"0\")").unwrap();
    m.set_user_input(0, 2, 2, "=ТЕКСТ(A1; \"# ##0,00\")").unwrap();
    m.set_user_input(0, 3, 2, "=ТЕКСТ(0,256; \"0,0%\")").unwrap();
    m.set_user_input(0, 4, 2, "=ТЕКСТ(A1; \"0,00\\\" руб.\\\"\")").unwrap();
    assert_eq!(m.get_formatted_cell_value(0, 1, 2).unwrap(), "1235");
    assert_eq!(m.get_formatted_cell_value(0, 2, 2).unwrap().replace('\u{a0}', " "), "1 234,50");
    assert_eq!(m.get_formatted_cell_value(0, 3, 2).unwrap(), "25,6%");
}

#[test]
fn ens_numbers_with_spaces_are_numbers() {
    let mut m = UserModel::new_empty("model", "ru", "UTC", "ru").unwrap();
    for (row, input, shown, value) in [
        (1, "1 234 567,89", "1 234 567,89", "1234567,89"),
        (2, "-12 345", "-12 345", "-12345"),
        (3, "1\u{a0}234", "1 234", "1234"),
    ] {
        m.set_user_input(0, row, 1, input).unwrap();
        assert_eq!(m.get_cell_type(0, row, 1).unwrap(), crate::types::CellType::Number, "{input}");
        assert_eq!(m.get_cell_content(0, row, 1).unwrap(), value, "{input}");
        assert_eq!(m.get_formatted_cell_value(0, row, 1).unwrap().replace('\u{a0}', " "), shown, "{input}");
    }
    // Не число: группы не по три цифры — остаётся текстом.
    m.set_user_input(0, 4, 1, "12 34").unwrap();
    assert_eq!(m.get_cell_type(0, 4, 1).unwrap(), crate::types::CellType::Text);
}

#[test]
fn ens_russian_dates_and_locale_tags() {
    let mut m = UserModel::new_empty("model", "ru", "UTC", "ru").unwrap();
    let area = |r| crate::expressions::types::Area { sheet: 0, row: r, column: 1, width: 1, height: 1 };
    let cases = [
        ("45000,5625", "d mmmm yyyy", "15 марта 2023"),
        ("45000", "[$-419]d mmmm yyyy", "15 марта 2023"),
        ("45000", "mmmm yyyy", "Март 2023"),
        ("45000", "[$-F800]dddd, d mmmm yyyy", "среда, 15 марта 2023"),
        ("1234,5", "#,##0.00 [$₽-419]", "1 234,50 ₽"),
        ("1234,5", "#,##0 [$руб.-419]", "1 235 руб."),
        ("1234,5", "[$$-409]#,##0.00", "$1 234,50"),
    ];
    for (r, (value, code, expected)) in cases.iter().enumerate() {
        let row = r as i32 + 1;
        m.set_user_input(0, row, 1, value).unwrap();
        m.update_range_style(&area(row), "num_fmt", code).unwrap();
        assert_eq!(m.get_formatted_cell_value(0, row, 1).unwrap().replace('\u{a0}', " "), *expected, "{code}");
    }
}
