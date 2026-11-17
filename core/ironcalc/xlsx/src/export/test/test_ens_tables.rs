// SPDX-License-Identifier: AGPL-3.0-or-later
//! ENS-Office: «умные» таблицы переживают круг «открыть → сохранить → открыть»
//! и сдвигаются вместе со строками и столбцами.

use std::io::Read;

use crate::export::save_to_xlsx;
use crate::import::load_from_xlsx;

fn read_part(name: &str, part: &str) -> String {
    let f = std::fs::File::open(name).unwrap();
    let mut zip = zip::ZipArchive::new(f).unwrap();
    let mut s = String::new();
    zip.by_name(part).unwrap().read_to_string(&mut s).unwrap();
    s
}

#[test]
fn test_ens_tables_round_trip() {
    let model = load_from_xlsx("tests/calc_test_no_export/tables.xlsx", "en", "UTC", "en").unwrap();
    assert!(!model.workbook.tables.is_empty());
    let name = "temp_ens_tables_round_trip.xlsx";
    save_to_xlsx(&model, name).unwrap();
    let ct = read_part(name, "[Content_Types].xml");
    assert!(ct.contains("/xl/tables/table1.xml"), "{ct}");
    let back = load_from_xlsx(name, "en", "UTC", "en").unwrap();
    std::fs::remove_file(name).unwrap();
    let strip = |m: &ironcalc_base::Model| {
        let mut t = m.workbook.tables.clone();
        for v in t.values_mut() {
            v.ens_raw_xml.clear();
        }
        t
    };
    assert_eq!(strip(&back), strip(&model));
    // То, чего модель не знает, сохраняется: условие фильтра, стиль, uid Excel
    let t4 = &back.workbook.tables["Table4"].ens_raw_xml;
    assert!(
        t4.contains("<customFilter operator=\"greaterThan\" val=\"20\"/>"),
        "{t4}"
    );
    assert!(t4.contains("TableStyleDark5"), "{t4}");
    // Формулы считают то же самое
    for (s, ws) in model.workbook.worksheets.iter().enumerate() {
        for r in ws.sheet_data.rows() {
            for c in ws.sheet_data.columns_in_row(r) {
                assert_eq!(
                    back.get_formatted_cell_value(s as u32, r, c).unwrap(),
                    model.get_formatted_cell_value(s as u32, r, c).unwrap(),
                    "лист {s}, R{r}C{c}"
                );
            }
        }
    }
}

#[test]
fn test_ens_tables_follow_rows_and_columns() {
    let mut model =
        load_from_xlsx("tests/calc_test_no_export/tables.xlsx", "en", "UTC", "en").unwrap();
    let key = "Table4".to_string();
    let t = model.workbook.tables[&key].clone();
    let sheet = model
        .workbook
        .worksheets
        .iter()
        .position(|w| w.name == t.sheet_name)
        .unwrap() as u32;
    let a = ironcalc_base::expressions::utils::parse_reference_a1(
        t.reference.split(':').next().unwrap(),
    )
    .unwrap();
    let width = t.columns.len();
    // строки выше таблицы: таблица сдвигается вниз
    model.insert_rows(sheet, 1, 2).unwrap();
    let t2 = &model.workbook.tables[&key];
    let a2 = ironcalc_base::expressions::utils::parse_reference_a1(
        t2.reference.split(':').next().unwrap(),
    )
    .unwrap();
    assert_eq!(a2.row, a.row + 2);
    // столбец внутри таблицы: у таблицы становится на столбец больше
    model.insert_columns(sheet, a.column + 1, 1).unwrap();
    assert_eq!(model.workbook.tables[&key].columns.len(), width + 1);
    let new_name = model.workbook.tables[&key].columns[1].name.clone();
    assert_eq!(
        model
            .get_formatted_cell_value(sheet, a2.row, a.column + 1)
            .unwrap(),
        new_name
    );
    model.delete_columns(sheet, a.column + 1, 1).unwrap();
    assert_eq!(model.workbook.tables[&key].columns.len(), width);
    assert_eq!(model.workbook.tables[&key].columns, t.columns);
    // переименование листа
    model.rename_sheet_by_index(sheet, "Новое имя").unwrap();
    model.evaluate();
    assert_eq!(model.workbook.tables[&key].sheet_name, "Новое имя");
    let name = "temp_ens_tables_moved.xlsx";
    save_to_xlsx(&model, name).unwrap();
    let back = load_from_xlsx(name, "en", "UTC", "en").unwrap();
    std::fs::remove_file(name).unwrap();
    assert_eq!(
        back.workbook.tables[&key].reference,
        model.workbook.tables[&key].reference
    );
    // Столбец вставили и удалили: XML собран из модели, имена столбцов те же
    let names = |t: &ironcalc_base::types::Table| {
        t.columns.iter().map(|c| c.name.clone()).collect::<Vec<_>>()
    };
    assert_eq!(names(&back.workbook.tables[&key]), names(&t));
}

#[test]
fn test_ens_tables_undo_redo() {
    use ironcalc_base::UserModel;
    let model = load_from_xlsx("tests/calc_test_no_export/tables.xlsx", "en", "UTC", "en").unwrap();
    let sheet = model
        .workbook
        .worksheets
        .iter()
        .position(|w| w.name == "filters")
        .unwrap() as u32;
    let mut um = UserModel::from_model(model);
    let before = um.get_model().workbook.tables.clone();
    let t = |um: &UserModel| um.get_model().workbook.tables.get("Table4").cloned();

    // Строки выше таблицы: сдвиг, отмена, повтор
    um.insert_rows(sheet, 1, 2).unwrap();
    let moved = t(&um).unwrap().reference;
    assert_eq!(moved, "A3:D9");
    um.undo().unwrap();
    assert_eq!(um.get_model().workbook.tables, before);
    um.redo().unwrap();
    assert_eq!(t(&um).unwrap().reference, moved);
    um.undo().unwrap();

    // Таблица удалена вместе со строками и возвращается отменой целиком
    um.delete_rows(sheet, 1, 10).unwrap();
    assert!(t(&um).is_none());
    um.undo().unwrap();
    assert_eq!(um.get_model().workbook.tables, before);
    um.redo().unwrap();
    assert!(t(&um).is_none());
    um.undo().unwrap();

    // Столбец внутри таблицы: отмена возвращает и столбцы, и исходный XML (условия фильтра)
    um.insert_columns(sheet, 2, 1).unwrap();
    assert_eq!(t(&um).unwrap().columns.len(), 5);
    um.undo().unwrap();
    assert_eq!(um.get_model().workbook.tables, before);

    // Лист удалён с таблицей и возвращён
    um.delete_sheet(sheet).unwrap();
    assert!(t(&um).is_none());
    um.undo().unwrap();
    assert_eq!(um.get_model().workbook.tables, before);
}

/// Ручная проверка: `ENS_KEEP=путь cargo test ... keep -- --ignored` сохраняет файлы.
#[test]
#[ignore]
fn test_ens_tables_keep_for_manual_check() {
    let dir = std::env::var("ENS_KEEP").unwrap();
    let mut model =
        load_from_xlsx("tests/calc_test_no_export/tables.xlsx", "en", "UTC", "en").unwrap();
    save_to_xlsx(&model, &format!("{dir}/tables_round.xlsx")).unwrap();
    let sheet = model
        .workbook
        .worksheets
        .iter()
        .position(|w| w.name == "filters")
        .unwrap() as u32;
    model.insert_rows(sheet, 1, 3).unwrap();
    model.insert_columns(sheet, 2, 1).unwrap();
    model.evaluate();
    save_to_xlsx(&model, &format!("{dir}/tables_moved.xlsx")).unwrap();
}
