use std::io::Read;

use ironcalc_base::Model;

use crate::export::save_to_xlsx;
use crate::import::load_from_xlsx;

// ENS-Office: книга, созданная с нуля по-русски (локаль и язык ru), должна
// сохраняться в xlsx, который Excel открывает без «восстановления».
#[test]
fn test_ens_ru_formulas_export() {
    let mut model = Model::new_empty("Книга1", "ru", "UTC", "ru").unwrap();
    model.rename_sheet_by_index(0, "Лист1").unwrap();
    model.add_sheet("Лист2").unwrap();
    model.set_user_input(0, 1, 1, "1,5".to_string()).unwrap();
    model.set_user_input(1, 1, 1, "=СУММ(Лист1!A1:A3)".to_string()).unwrap();
    model.set_user_input(1, 2, 1, "=Лист1!A1*2,5".to_string()).unwrap();
    model.set_user_input(1, 3, 1, "=ЕСЛИ(A1>1;\"да\";\"нет\")".to_string()).unwrap();
    model.set_user_input(1, 4, 1, "=SUM(A1:A3)".to_string()).unwrap();
    model.set_user_input(1, 5, 1, "=ОКРУГЛ(A2;1)".to_string()).unwrap();
    model.set_user_input(1, 6, 1, "=B5+".to_string()).unwrap();
    model.set_user_input(1, 7, 1, "=СУММ(".to_string()).unwrap();
    model.evaluate();
    let name = "temp_ens_ru_export.xlsx";
    save_to_xlsx(&model, name).unwrap();
    let f = std::fs::File::open(name).unwrap();
    let mut zip = zip::ZipArchive::new(f).unwrap();
    let mut xml = String::new();
    zip.by_name("xl/worksheets/sheet2.xml").unwrap().read_to_string(&mut xml).unwrap();
    let mut wb = String::new();
    zip.by_name("xl/workbook.xml").unwrap().read_to_string(&mut wb).unwrap();
    let back = load_from_xlsx(name, "ru", "UTC", "ru").unwrap();
    std::fs::remove_file(name).unwrap();
    assert_eq!(back.get_formatted_cell_value(1, 1, 1).unwrap(), "1,5");
    assert!(xml.contains("<is><t>=B5+</t></is>"), "{xml}");
    assert!(!xml.contains("#ERROR!"), "{xml}");
    assert_eq!(back.get_formatted_cell_value(1, 6, 1).unwrap(), "=B5+");
    for f in ["СУММ", "ЕСЛИ", "ОКРУГЛ", "2,5"] {
        assert!(!xml.contains(f), "в xlsx попало {f}: {xml}");
    }
}

// ENS-Office: the autofilter is saved as Excel does it.
#[test]
fn test_ens_autofilter_export() {
    let mut model = ironcalc_base::UserModel::new_empty("Книга1", "ru", "UTC", "ru").unwrap();
    model.set_user_input(0, 1, 1, "Город").unwrap();
    model.set_user_input(0, 2, 1, "Москва").unwrap();
    model.set_auto_filter(0, Some([1, 1, 2, 2])).unwrap();
    let name = "temp_ens_autofilter.xlsx";
    save_to_xlsx(model.get_model(), name).unwrap();
    let f = std::fs::File::open(name).unwrap();
    let mut zip = zip::ZipArchive::new(f).unwrap();
    let mut xml = String::new();
    zip.by_name("xl/worksheets/sheet1.xml").unwrap().read_to_string(&mut xml).unwrap();
    let mut wb = String::new();
    zip.by_name("xl/workbook.xml").unwrap().read_to_string(&mut wb).unwrap();
    let back = load_from_xlsx(name, "ru", "UTC", "ru").unwrap();
    std::fs::remove_file(name).unwrap();
    assert!(xml.contains("<autoFilter ref=\"A1:B2\"/>"), "{xml}");
    assert!(wb.contains("_xlnm._FilterDatabase"), "{wb}");
    let back = ironcalc_base::UserModel::from_model(back);
    assert_eq!(back.get_auto_filter(0).unwrap(), Some([1, 1, 2, 2]));
}

// ENS-Office: page setup round trip (pageSetup, pageMargins, printOptions, fitToPage).
#[test]
fn test_ens_page_setup_export() {
    let mut model = ironcalc_base::UserModel::new_empty("Book1", "en", "UTC", "en").unwrap();
    model.set_user_input(0, 1, 1, "x").unwrap();
    let json = r#"{"landscape":true,"paper":"A3","fit":"width","marginMm":[10,6,10,6],"gridlines":true,"centerH":true,"footer":"&CСтраница &P из &N"}"#;
    model.set_page_setup(0, json).unwrap();
    model.undo().unwrap();
    assert_eq!(model.get_page_setup(0).unwrap(), "");
    model.redo().unwrap();
    model.set_print_area(0, Some([1, 1, 20, 3])).unwrap();
    model.set_print_titles(0, [1, 2, 1, 1]).unwrap();
    let name = "temp_ens_page_setup.xlsx";
    save_to_xlsx(model.get_model(), name).unwrap();
    let f = std::fs::File::open(name).unwrap();
    let mut zip = zip::ZipArchive::new(f).unwrap();
    let mut xml = String::new();
    zip.by_name("xl/worksheets/sheet1.xml").unwrap().read_to_string(&mut xml).unwrap();
    let back = load_from_xlsx(name, "en", "UTC", "en").unwrap();
    std::fs::remove_file(name).unwrap();
    assert!(xml.contains(r#"<pageSetup paperSize="8" orientation="landscape" fitToWidth="1" fitToHeight="0"/>"#), "{xml}");
    assert!(xml.contains(r#"<sheetPr><outlinePr summaryBelow="0" summaryRight="0"/><pageSetUpPr fitToPage="1"/></sheetPr>"#), "{xml}");
    let v: serde_json::Value = serde_json::from_str(&back.workbook.worksheets[0].ens_page_setup).unwrap();
    assert_eq!(v["landscape"], true);
    assert_eq!(v["paper"], "A3");
    assert_eq!(v["fit"], "width");
    assert_eq!(v["gridlines"], true);
    assert_eq!(v["centerH"], true);
    assert_eq!(v["marginMm"][0], 10.0);
    assert_eq!(v["footer"], "&CСтраница &P из &N");
    let back = ironcalc_base::UserModel::from_model(back);
    assert_eq!(back.get_print_area(0).unwrap(), Some([1, 1, 20, 3]));
    assert_eq!(back.get_print_titles(0).unwrap(), [1, 2, 1, 1]);
}

// ENS-Office: position of the outline buttons round trip (`<outlinePr>`); Excel's default is below / right.
#[test]
fn test_ens_outline_summary_export() {
    let mut model = ironcalc_base::UserModel::new_empty("Book1", "en", "UTC", "en").unwrap();
    model.set_user_input(0, 1, 1, "x").unwrap();
    assert_eq!(model.get_outline_summary(0).unwrap(), (false, false));
    model.set_outline_summary(0, true, false).unwrap();
    model.undo().unwrap();
    assert_eq!(model.get_outline_summary(0).unwrap(), (false, false));
    model.redo().unwrap();
    let name = "temp_ens_outline_summary.xlsx";
    save_to_xlsx(model.get_model(), name).unwrap();
    let back = load_from_xlsx(name, "en", "UTC", "en").unwrap();
    std::fs::remove_file(name).unwrap();
    model.set_outline_summary(0, true, true).unwrap();
    save_to_xlsx(model.get_model(), name).unwrap();
    let f = std::fs::File::open(name).unwrap();
    let mut zip = zip::ZipArchive::new(f).unwrap();
    let mut xml = String::new();
    zip.by_name("xl/worksheets/sheet1.xml").unwrap().read_to_string(&mut xml).unwrap();
    let default_back = load_from_xlsx(name, "en", "UTC", "en").unwrap();
    std::fs::remove_file(name).unwrap();
    assert_eq!(back.workbook.worksheets[0].ens_summary, (true, false));
    assert!(!xml.contains("outlinePr"), "{xml}");
    assert_eq!(default_back.workbook.worksheets[0].ens_summary, (true, true));
}

// ENS-Office: superscript / subscript round trip (`<vertAlign>`).
#[test]
fn test_ens_vert_align_export() {
    let mut model = ironcalc_base::UserModel::new_empty("Book1", "en", "UTC", "en").unwrap();
    model.set_user_input(0, 1, 1, "x2").unwrap();
    let area = |column| ironcalc_base::expressions::types::Area { sheet: 0, row: 1, column, width: 1, height: 1 };
    model.update_range_style(&area(1), "font.vert_align", "superscript").unwrap();
    model.set_user_input(0, 1, 2, "H2O").unwrap();
    model.update_range_style(&area(2), "font.vert_align", "subscript").unwrap();
    let name = "temp_ens_vert_align.xlsx";
    save_to_xlsx(model.get_model(), name).unwrap();
    let back = load_from_xlsx(name, "en", "UTC", "en").unwrap();
    std::fs::remove_file(name).unwrap();
    let back = ironcalc_base::UserModel::from_model(back);
    let v = |c| back.get_cell_style(0, 1, c).unwrap().font.vert_align;
    assert_eq!(v(1), ironcalc_base::types::VertAlign::Superscript);
    assert_eq!(v(2), ironcalc_base::types::VertAlign::Subscript);
    assert_eq!(v(3), ironcalc_base::types::VertAlign::Baseline);
}

// ENS-Office: the change log is kept in customXml and read back.
#[test]
fn test_ens_change_log_export() {
    let mut model = ironcalc_base::UserModel::new_empty("Book1", "en", "UTC", "en").unwrap();
    model.set_user_input(0, 1, 1, "x").unwrap();
    model.set_change_log(r#"[{"u":"Иванов <И.>","t":1}]"#);
    model.set_last_modified_by("Иванов");
    let name = "temp_ens_change_log.xlsx";
    save_to_xlsx(model.get_model(), name).unwrap();
    let f = std::fs::File::open(name).unwrap();
    let mut zip = zip::ZipArchive::new(f).unwrap();
    let mut rels = String::new();
    zip.by_name("xl/_rels/workbook.xml.rels").unwrap().read_to_string(&mut rels).unwrap();
    let back = load_from_xlsx(name, "en", "UTC", "en").unwrap();
    std::fs::remove_file(name).unwrap();
    assert!(rels.contains("../customXml/item1.xml"), "{rels}");
    assert_eq!(back.workbook.ens_change_log, r#"[{"u":"Иванов <И.>","t":1}]"#);
    assert_eq!(back.workbook.metadata.last_modified_by, "Иванов");
}
