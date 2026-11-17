// SPDX-License-Identifier: AGPL-3.0-or-later
//! ENS-Office: «открыть → сохранить → открыть» без потерь: макросы .xlsm, codeName,
//! цвет ярлыка; список того, что сохранить нельзя.

use std::io::Read;

use crate::export::save_to_xlsx;
use crate::import::{load_from_xlsx, lost_features_of};
use ironcalc_base::types::Color;

fn part(name: &str, path: &str) -> Vec<u8> {
    let f = std::fs::File::open(name).unwrap();
    let mut zip = zip::ZipArchive::new(f).unwrap();
    let mut v = Vec::new();
    zip.by_name(path).unwrap().read_to_end(&mut v).unwrap();
    v
}

#[test]
fn test_ens_xlsm_macros_round_trip() {
    let src = "tests/ens/macros.xlsm";
    let model = load_from_xlsx(src, "ru", "UTC", "ru").unwrap();
    assert_eq!(model.workbook.ens_vba_parts.len(), 1);
    assert_eq!(model.workbook.ens_code_name, "ЭтаКнига");
    assert_eq!(model.workbook.worksheets[0].ens_code_name, "Sheet1");
    let out = "temp_ens_macros.xlsm";
    save_to_xlsx(&model, out).unwrap();
    let vba = part(out, "xl/vbaProject.bin");
    let ct = String::from_utf8(part(out, "[Content_Types].xml")).unwrap();
    let rels = String::from_utf8(part(out, "xl/_rels/workbook.xml.rels")).unwrap();
    let wb = String::from_utf8(part(out, "xl/workbook.xml")).unwrap();
    let back = load_from_xlsx(out, "ru", "UTC", "ru").unwrap();
    std::fs::remove_file(out).unwrap();
    assert_eq!(vba, part(src, "xl/vbaProject.bin"), "макросы изменились");
    assert!(ct.contains("macroEnabled.main+xml"), "{ct}");
    assert!(
        ct.contains(
            r#"PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject""#
        ),
        "{ct}"
    );
    assert!(
        rels.contains(r#"relationships/vbaProject" Target="vbaProject.bin""#),
        "{rels}"
    );
    assert!(wb.contains(r#"<workbookPr codeName="ЭтаКнига"/>"#), "{wb}");
    assert_eq!(back.workbook.ens_vba_parts, model.workbook.ens_vba_parts);
    assert_eq!(back.workbook.ens_code_name, "ЭтаКнига");
    assert_eq!(back.workbook.worksheets[0].ens_code_name, "Sheet1");
}

#[test]
fn test_ens_tab_color_round_trip() {
    let mut model = load_from_xlsx("tests/ens/macros.xlsm", "ru", "UTC", "ru").unwrap();
    model.evaluate();
    assert_eq!(
        model.workbook.worksheets[0].color,
        Color::Rgb("#FF0000".to_string())
    );
    let out = "temp_ens_tab_color.xlsx";
    save_to_xlsx(&model, out).unwrap();
    let back = load_from_xlsx(out, "ru", "UTC", "ru").unwrap();
    std::fs::remove_file(out).unwrap();
    assert_eq!(
        back.workbook.worksheets[0].color,
        Color::Rgb("#FF0000".to_string())
    );
    assert_eq!(
        back.workbook.worksheets[1].color,
        Color::Rgb("#00B050".to_string())
    );
    // Значения на месте
    assert_eq!(back.get_formatted_cell_value(0, 2, 1).unwrap(), "10");
}

#[test]
fn test_ens_plain_xlsx_has_no_macros() {
    let model = load_from_xlsx("tests/example.xlsx", "en", "UTC", "en").unwrap();
    assert!(model.workbook.ens_vba_parts.is_empty());
    let out = "temp_ens_plain.xlsx";
    save_to_xlsx(&model, out).unwrap();
    let ct = String::from_utf8(part(out, "[Content_Types].xml")).unwrap();
    std::fs::remove_file(out).unwrap();
    assert!(!ct.contains("macroEnabled"), "{ct}");
}

#[test]
fn test_ens_lost_features_list() {
    let names: Vec<String> = [
        "xl/workbook.xml",
        "xl/calcChain.xml",
        "xl/printerSettings/printerSettings1.bin",
        "xl/pivotTables/pivotTable1.xml",
        "xl/pivotCache/pivotCacheDefinition1.xml",
        "xl/externalLinks/externalLink1.xml",
        "xl/ctrlProps/ctrlProp1.xml",
    ]
    .iter()
    .map(|s| s.to_string())
    .collect();
    let lost = lost_features_of(&names);
    assert_eq!(lost.len(), 3, "{lost:?}");
    assert!(lost[0].starts_with("сводные таблицы"));
    assert!(lost[1].starts_with("связи с другими книгами"));
    assert!(lost[2].starts_with("элементы управления"));
    // В example.xlsx есть сводная таблица и лист диаграммы: о них нужно предупредить
    let model = load_from_xlsx("tests/example.xlsx", "en", "UTC", "en").unwrap();
    assert_eq!(
        model.workbook.ens_lost_features,
        ["сводные таблицы", "листы диаграмм"]
    );
    // Книга, сохранённая ENS-Table, ничего не теряет
    let out = "temp_ens_lost_none.xlsx";
    save_to_xlsx(&model, out).unwrap();
    let back = load_from_xlsx(out, "en", "UTC", "en").unwrap();
    std::fs::remove_file(out).unwrap();
    assert!(
        back.workbook.ens_lost_features.is_empty(),
        "{:?}",
        back.workbook.ens_lost_features
    );
}

/// Ручная проверка: `ENS_KEEP=папка cargo test ... xlsm_keep -- --ignored`.
#[test]
#[ignore]
fn test_ens_xlsm_keep_for_manual_check() {
    let dir = std::env::var("ENS_KEEP").unwrap();
    let mut model = load_from_xlsx("tests/ens/macros.xlsm", "ru", "UTC", "ru").unwrap();
    model.evaluate();
    save_to_xlsx(&model, &format!("{dir}/macros_saved.xlsm")).unwrap();
}
