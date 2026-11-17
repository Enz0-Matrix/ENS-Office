// SPDX-License-Identifier: AGPL-3.0-or-later
//! ENS-Office: ссылки на другие книги через xlsx. Книга-источник не открыта:
//! значения берутся из кэша файла или читаются из файла-источника на диске.

use std::io::Read;

use ironcalc_base::{Model, UserModel};

use crate::export::save_to_xlsx;
use crate::import::load_from_xlsx;

const DIR: &str = "tests/ens/external";
const SOURCE: &str = "Цены и остатки.xlsx";

/// Книга-источник: «Цены» (топливо → цена за литр) и «Остатки на складе».
fn make_source(price_95: &str) -> Model<'static> {
    let mut m = Model::new_empty("Цены и остатки", "ru", "UTC", "ru").unwrap();
    m.rename_sheet_by_index(0, "Цены").unwrap();
    m.add_sheet("Остатки на складе").unwrap();
    let rows = [
        ("Топливо", "Цена, руб./л"),
        ("АИ-92", "52,30"),
        ("АИ-95", price_95),
        ("ДТ", "66,80"),
        ("Газ", "28,10"),
    ];
    for (r, (a, b)) in rows.iter().enumerate() {
        m.set_user_input(0, r as i32 + 1, 1, a.to_string()).unwrap();
        m.set_user_input(0, r as i32 + 1, 2, b.to_string()).unwrap();
    }
    let stock = [
        ("Подразделение", "Топливо", "Остаток, л"),
        ("Участок 1", "АИ-92", "1200"),
        ("Участок 1", "ДТ", "3400"),
        ("Участок 2", "АИ-95", "800"),
        ("Участок 2", "ДТ", "1500"),
        ("Участок 3", "Газ", "640"),
    ];
    for (r, (a, b, c)) in stock.iter().enumerate() {
        m.set_user_input(1, r as i32 + 1, 1, a.to_string()).unwrap();
        m.set_user_input(1, r as i32 + 1, 2, b.to_string()).unwrap();
        m.set_user_input(1, r as i32 + 1, 3, c.to_string()).unwrap();
    }
    m.set_user_input(1, 8, 2, "Итого".to_string()).unwrap();
    m.set_user_input(1, 8, 3, "=СУММ(C2:C6)".to_string())
        .unwrap();
    m.evaluate();
    m
}

/// Книга-отчёт со ссылками на источник: ячейка, сумма, ВПР, ИНДЕКС+ПОИСКПОЗ, СУММЕСЛИ, путь.
fn make_report(source: &Model) -> UserModel<'static> {
    let mut m = UserModel::new_empty("Отчёт", "ru", "UTC", "ru").unwrap();
    m.rename_sheet(0, "Обеспеченность").unwrap();
    let cells: &[(i32, i32, &str)] = &[
        (1, 1, "Обеспеченность топливом"),
        (2, 1, "Топливо"),
        (2, 2, "Цена (ВПР)"),
        (2, 3, "Цена (ИНДЕКС+ПОИСКПОЗ)"),
        (2, 4, "Остаток, л (СУММЕСЛИ)"),
        (2, 5, "Стоимость, руб."),
        (3, 1, "АИ-92"),
        (4, 1, "АИ-95"),
        (5, 1, "ДТ"),
        (6, 1, "Газ"),
        (8, 1, "Всего на складах, л"),
        (8, 4, "='[Цены и остатки.xlsx]Остатки на складе'!C8"),
        (9, 1, "Сумма остатков диапазоном"),
        (
            9,
            4,
            "=СУММ('[Цены и остатки.xlsx]Остатки на складе'!C2:C6)",
        ),
        (10, 1, "Цена АИ-92 прямой ссылкой"),
        (10, 2, "=[Цены и остатки.xlsx]Цены!B2"),
        (11, 1, "Сколько видов топлива"),
        (11, 2, "=СЧЁТЗ([Цены и остатки.xlsx]Цены!A2:A10)"),
    ];
    for (r, c, v) in cells {
        m.set_user_input(0, *r, *c, v).unwrap();
    }
    for r in 3..=6 {
        m.set_user_input(
            0,
            r,
            2,
            &format!("=ВПР(A{r};[Цены и остатки.xlsx]Цены!$A$2:$B$10;2;ЛОЖЬ)"),
        )
        .unwrap();
        m.set_user_input(0, r, 3, &format!(
            "=ИНДЕКС([Цены и остатки.xlsx]Цены!$B$2:$B$10;ПОИСКПОЗ(A{r};[Цены и остатки.xlsx]Цены!$A$2:$A$10;0))"
        )).unwrap();
        m.set_user_input(0, r, 4, &format!(
            "=СУММЕСЛИ('[Цены и остатки.xlsx]Остатки на складе'!$B$2:$B$6;A{r};'[Цены и остатки.xlsx]Остатки на складе'!$C$2:$C$6)"
        )).unwrap();
        m.set_user_input(0, r, 5, &format!("=B{r}*D{r}")).unwrap();
    }
    m.update_external_book(SOURCE, source).unwrap();
    m
}

fn values(m: &Model) -> Vec<String> {
    [
        (3, 2),
        (4, 2),
        (4, 3),
        (5, 4),
        (4, 5),
        (8, 4),
        (9, 4),
        (10, 2),
        (11, 2),
    ]
    .iter()
    .map(|(r, c)| m.get_formatted_cell_value(0, *r, *c).unwrap())
    .collect()
}

#[test]
fn test_ens_external_report_from_closed_source() {
    let source = make_source("57,10");
    let report = make_report(&source);
    let expected = [
        "52,3", "57,1", "57,1", "4900", "45680", "7540", "7540", "52,3", "4",
    ];
    assert_eq!(values(report.get_model()), expected);

    // Сохранили: ссылки в виде Excel, кэш значений, путь к книге
    let out = "temp_ens_external_report.xlsx";
    save_to_xlsx(report.get_model(), out).unwrap();
    let mut zip = zip::ZipArchive::new(std::fs::File::open(out).unwrap()).unwrap();
    let mut sheet = String::new();
    zip.by_name("xl/worksheets/sheet1.xml")
        .unwrap()
        .read_to_string(&mut sheet)
        .unwrap();
    let mut link = String::new();
    zip.by_name("xl/externalLinks/externalLink1.xml")
        .unwrap()
        .read_to_string(&mut link)
        .unwrap();
    let mut rels = String::new();
    zip.by_name("xl/externalLinks/_rels/externalLink1.xml.rels")
        .unwrap()
        .read_to_string(&mut rels)
        .unwrap();
    let names: Vec<String> = zip.file_names().map(|s| s.to_string()).collect();
    drop(zip);
    assert!(
        sheet.contains("VLOOKUP(A3,[1]Цены!$A$2:$B$10,2,FALSE)"),
        "{sheet}"
    );
    assert!(
        sheet.contains("&apos;[1]Остатки на складе&apos;!C8")
            || sheet.contains("'[1]Остатки на складе'!C8"),
        "{sheet}"
    );
    assert!(
        link.contains("<sheetName val=\"Цены\"/>") && link.contains("<v>АИ-95</v>"),
        "{link}"
    );
    assert!(
        rels.contains("Target=\"Цены и остатки.xlsx\" TargetMode=\"External\""),
        "{rels}"
    );
    assert_eq!(
        names
            .iter()
            .filter(|n| n.starts_with("xl/worksheets/sheet"))
            .count(),
        1,
        "служебные листы не пишутся: {names:?}"
    );

    // Открыли отчёт, источник закрыт: значения из кэша, формулы на месте
    let back = load_from_xlsx(out, "ru", "UTC", "ru").unwrap();
    std::fs::remove_file(out).unwrap();
    assert_eq!(values(&back), expected);
    assert!(
        back.workbook.ens_lost_features.is_empty(),
        "{:?}",
        back.workbook.ens_lost_features
    );
    let mut back = UserModel::from_model(back);
    back.get_model();
    assert_eq!(
        back.get_cell_content(0, 4, 2).unwrap(),
        "=ВПР(A4;'[Цены и остатки.xlsx]Цены'!$A$2:$B$10;2;ЛОЖЬ)"
    );

    // Цена в источнике изменилась → «Обновить связи» читает файл-источник с диска
    let src_file = "temp_ens_external_source.xlsx";
    save_to_xlsx(&make_source("60,00"), src_file).unwrap();
    let fresh = load_from_xlsx(src_file, "ru", "UTC", "ru").unwrap();
    std::fs::remove_file(src_file).unwrap();
    back.update_external_book(SOURCE, &fresh).unwrap();
    assert_eq!(back.get_formatted_cell_value(0, 4, 2).unwrap(), "60");
    assert_eq!(back.get_formatted_cell_value(0, 4, 5).unwrap(), "48000");
}

/// Файл в том виде, как внешние связи пишет Excel (собран вручную: `[1]`, кавычки, кэш, путь file:///).
#[test]
fn test_ens_external_excel_style_file() {
    let mut m = load_from_xlsx(&format!("{DIR}/excel-style.xlsx"), "ru", "UTC", "ru").unwrap();
    m.evaluate(); // файл собран без сохранённых результатов формул
    let books = &m.workbook.ens_external_books;
    assert_eq!(books.len(), 1);
    assert_eq!(books[0].file, "Цены и остатки.xlsx");
    assert_eq!(books[0].target, r"file:///D:\Работа\Цены и остатки.xlsx");
    assert_eq!(m.get_formatted_cell_value(0, 1, 2).unwrap(), "57,1"); // ВПР по кэшу
    assert_eq!(m.get_formatted_cell_value(0, 2, 2).unwrap(), "7540"); // СУММ по листу с пробелом
    assert_eq!(m.get_formatted_cell_value(0, 3, 2).unwrap(), "#Н/Д"); // нет в кэше — как в Excel
}

/// Демо-книги для ручной проверки: `ENS_KEEP=папка cargo test … external_demo -- --ignored`.
#[test]
#[ignore]
fn test_ens_external_demo_books() {
    let dir = std::env::var("ENS_KEEP").unwrap();
    let source = make_source("57,10");
    save_to_xlsx(&source, &format!("{dir}/{SOURCE}")).unwrap();
    let report = make_report(&source);
    save_to_xlsx(report.get_model(), &format!("{dir}/Отчёт.xlsx")).unwrap();
}
