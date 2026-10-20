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
