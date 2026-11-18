// SPDX-License-Identifier: AGPL-3.0-or-later
//! ENS-Office: структурированные ссылки (`Таблица1[Сумма]`) остаются ссылками на таблицу
//! после «открыть → сохранить» (ADR-013), а не превращаются в диапазоны вида `$B$2:$B$5`.

use std::io::Read;

use ironcalc_base::UserModel;

use crate::export::save_to_xlsx;
use crate::import::load_from_xlsx;

const FILE: &str = "tests/calc_test_no_export/tables.xlsx";

/// Все формулы листа `n` книги: (ячейка, текст формулы).
fn formulas(file: &str, n: usize) -> Vec<(String, String)> {
    let mut zip = zip::ZipArchive::new(std::fs::File::open(file).unwrap()).unwrap();
    let mut xml = String::new();
    zip.by_name(&format!("xl/worksheets/sheet{n}.xml")).unwrap().read_to_string(&mut xml).unwrap();
    let mut out = Vec::new();
    for part in xml.split("<c r=\"").skip(1) {
        let cell = part.split('"').next().unwrap().to_string();
        let head = part.split("</c>").next().unwrap();
        if let Some(i) = head.find("<f") {
            let rest = &head[i..];
            if let (Some(a), Some(b)) = (rest.find('>'), rest.find("</f>")) {
                if a < b {
                    let f = rest[a + 1..b].replace("&apos;", "'").replace("&quot;", "\"").replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&");
                    out.push((cell, f));
                }
            }
        }
    }
    out
}

#[test]
fn ens_structured_references_survive_round_trip() {
    let model = load_from_xlsx(FILE, "en", "UTC", "en").unwrap();
    let out = "temp_ens_structured.xlsx";
    let _ = std::fs::remove_file(out);
    save_to_xlsx(&model, out).unwrap();
    let mut checked = 0;
    for n in 1..=8 {
        let before: Vec<_> = formulas(FILE, n).into_iter().filter(|(_, f)| f.contains('[')).collect();
        let after = formulas(out, n);
        for (cell, f) in &before {
            let got = after.iter().find(|(c, _)| c == cell).map(|(_, f)| f.clone());
            assert_eq!(got.as_deref().map(|g| g.replace(' ', "")), Some(f.replace(", ", ",").replace(' ', "")), "лист {n}, {cell}");
            checked += 1;
        }
    }
    assert!(checked > 40, "{checked}");
    // и при повторном открытии значения те же
    let back = load_from_xlsx(out, "en", "UTC", "en").unwrap();
    std::fs::remove_file(out).unwrap();
    for (i, ws) in model.workbook.worksheets.iter().enumerate() {
        for (r, c, _) in ws.sheet_data.cells() {
            assert_eq!(
                back.get_formatted_cell_value(i as u32, r, c).unwrap(),
                model.get_formatted_cell_value(i as u32, r, c).unwrap(),
                "{}!{r},{c}",
                ws.name
            );
        }
    }
}

#[test]
fn ens_structured_references_russian() {
    let model = load_from_xlsx(FILE, "ru", "UTC", "ru").unwrap();
    let mut m = UserModel::from_model(model);
    // показ — как в русском Excel
    assert_eq!(m.get_cell_content(0, 2, 6).unwrap(), "=СУММ(Table1[@[Jan]:[Dec]])");
    assert_eq!(m.get_cell_content(0, 3, 10).unwrap(), "=СУММ(Table1[#Все])");
    assert_eq!(m.get_cell_content(0, 3, 12).unwrap(), "=СУММ(Table1[[#Данные];[Jan]:[Mar]])");
    assert_eq!(m.get_cell_content(4, 2, 6).unwrap(), "=Table16[@Units]*Table16[@[Unit Cost]]");
    // ввод по-русски и по-английски, с таблицей и без (своя таблица)
    let jan_feb = m.get_formatted_cell_value(0, 2, 1).unwrap().parse::<f64>().unwrap()
        + m.get_formatted_cell_value(0, 2, 2).unwrap().parse::<f64>().unwrap();
    m.set_user_input(0, 2, 8, "=СУММ(Table1[[#Эта строка];[Jan]:[Feb]])").unwrap();
    assert_eq!(m.get_formatted_cell_value(0, 2, 8).unwrap().parse::<f64>().unwrap(), jan_feb);
    m.set_user_input(0, 2, 8, "=Table1[@Jan]+Table1[@Feb]").unwrap();
    assert_eq!(m.get_formatted_cell_value(0, 2, 8).unwrap().parse::<f64>().unwrap(), jan_feb);
    m.set_user_input(0, 2, 8, "=СУММ(table1[jan])").unwrap();
    assert_eq!(m.get_cell_content(0, 2, 8).unwrap(), "=СУММ(Table1[Jan])");
    m.set_user_input(0, 2, 8, "=СУММ(Table1[#Итоги])").unwrap();
    assert_eq!(m.get_cell_content(0, 2, 8).unwrap(), "=СУММ(Table1[#Итоги])");
}

#[test]
fn ens_structured_references_follow_table() {
    let model = load_from_xlsx(FILE, "en", "UTC", "en").unwrap();
    let mut m = UserModel::from_model(model);
    // лист filters: Table4 (заголовок в строке 1, данные 2–6), A7 = SUBTOTAL(109,Table4[Dogs])
    let sheet = 3;
    let before = m.get_formatted_cell_value(sheet, 7, 1).unwrap().parse::<f64>().unwrap();
    m.insert_rows(sheet, 4, 1).unwrap();
    m.set_user_input(sheet, 4, 1, "1000").unwrap();
    assert_eq!(m.get_cell_content(sheet, 8, 1).unwrap(), "=SUBTOTAL(109,Table4[Dogs])");
    assert_eq!(m.get_formatted_cell_value(sheet, 8, 1).unwrap().parse::<f64>().unwrap(), before + 1000.0);
    m.undo().unwrap();
    m.undo().unwrap();
    assert_eq!(m.get_formatted_cell_value(sheet, 7, 1).unwrap().parse::<f64>().unwrap(), before);
}

#[test]
fn ens_header_rename_rewrites_formulas() {
    let model = load_from_xlsx(FILE, "ru", "UTC", "ru").unwrap();
    let mut m = UserModel::from_model(model);
    // лист filters: Table4 (A1:D7), заголовок в строке 1 (A1 = Dogs), A7 = SUBTOTAL(109,Table4[Dogs])
    let sheet = 3;
    let value = m.get_formatted_cell_value(sheet, 7, 1).unwrap();
    m.set_user_input(sheet, 1, 1, "Собаки").unwrap();
    assert_eq!(m.get_cell_content(sheet, 7, 1).unwrap(), "=ПРОМЕЖУТОЧНЫЕ.ИТОГИ(109;Table4[Собаки])");
    assert_eq!(m.get_formatted_cell_value(sheet, 7, 1).unwrap(), value);
    // xlsx: имя столбца и формула согласованы
    let out = "temp_ens_header_rename.xlsx";
    let _ = std::fs::remove_file(out);
    save_to_xlsx(m.get_model(), out).unwrap();
    let f = formulas(out, 4);
    let back = load_from_xlsx(out, "ru", "UTC", "ru").unwrap();
    std::fs::remove_file(out).unwrap();
    assert!(f.iter().any(|(c, f)| c == "A7" && f == "SUBTOTAL(109,Table4[Собаки])"), "{f:?}");
    assert_eq!(back.get_formatted_cell_value(sheet, 7, 1).unwrap(), value);
    // отмена возвращает и заголовок, и формулу
    m.undo().unwrap();
    assert_eq!(m.get_cell_content(sheet, 7, 1).unwrap(), "=ПРОМЕЖУТОЧНЫЕ.ИТОГИ(109;Table4[Dogs])");
    m.redo().unwrap();
    assert_eq!(m.get_cell_content(sheet, 7, 1).unwrap(), "=ПРОМЕЖУТОЧНЫЕ.ИТОГИ(109;Table4[Собаки])");
    // пустой заголовок — прежнее имя, ссылки не ломаются
    m.set_user_input(sheet, 1, 1, "").unwrap();
    assert_eq!(m.get_cell_content(sheet, 7, 1).unwrap(), "=ПРОМЕЖУТОЧНЫЕ.ИТОГИ(109;Table4[Собаки])");
}

#[test]
fn ens_header_rename_replays_from_journal() {
    let model = load_from_xlsx(FILE, "ru", "UTC", "ru").unwrap();
    let mut m = UserModel::from_model(model);
    let snapshot = m.to_bytes();
    let _ = m.flush_send_queue();
    m.set_user_input(3, 1, 1, "Собаки").unwrap();
    let journal = m.flush_send_queue();
    let mut r = UserModel::from_bytes(&snapshot, "ru").unwrap();
    r.apply_external_diffs(&journal).unwrap();
    assert_eq!(r.get_cell_content(3, 7, 1).unwrap(), "=ПРОМЕЖУТОЧНЫЕ.ИТОГИ(109;Table4[Собаки])");
    assert_eq!(r.get_formatted_cell_value(3, 7, 1).unwrap(), m.get_formatted_cell_value(3, 7, 1).unwrap());
}
