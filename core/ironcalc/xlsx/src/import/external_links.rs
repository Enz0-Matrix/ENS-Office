// SPDX-License-Identifier: AGPL-3.0-or-later
//! ENS-Office: внешние связи xlsx (`xl/externalLinks/externalLinkN.xml`): путь к книге-источнику,
//! её листы и кэш значений, по которому Excel считает формулы, пока источник закрыт.

use std::collections::HashMap;
use std::io::Read;

use ironcalc_base::expressions::utils::parse_reference_a1;
use ironcalc_base::external::{file_name_of, ExternalValue};

use super::worksheets::Relationship;
use super::xml::XmlNode;

const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Книга-источник, прочитанная из xlsx.
pub(crate) struct LoadedBook {
    pub target: String,
    pub file: String,
    pub sheets: Vec<(String, Vec<(i32, i32, ExternalValue)>)>,
}

fn read_text<R: Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    path: &str,
) -> Option<String> {
    let mut s = String::new();
    archive.by_name(path).ok()?.read_to_string(&mut s).ok()?;
    Some(s)
}

/// Путь части из связи `xl/_rels/workbook.xml.rels` (`externalLinks/externalLink1.xml`).
fn part_path(target: &str) -> String {
    match target.strip_prefix('/') {
        Some(p) => p.to_string(),
        None => format!("xl/{target}"),
    }
}

/// Книги-источники по порядку `<externalReferences>` (порядок = номера `[1]`, `[2]` в формулах).
/// Книга, которую не удалось прочитать, всё равно занимает своё место (с пустым именем),
/// чтобы номера остальных не сдвинулись.
pub(crate) fn load_external_links<R: Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    rels: &HashMap<String, Relationship>,
) -> Vec<LoadedBook> {
    let Some(text) = read_text(archive, "xl/workbook.xml") else {
        return Vec::new();
    };
    let Ok(root) = XmlNode::parse(text.as_bytes()) else {
        return Vec::new();
    };
    let Some(refs) = root
        .children()
        .find(|n| n.has_tag_name("externalReferences"))
    else {
        return Vec::new();
    };
    let ids: Vec<String> = refs
        .children()
        .filter(|n| n.has_tag_name("externalReference"))
        .filter_map(|n| n.attribute_ns(R_NS, "id").map(|s| s.to_string()))
        .collect();
    let mut books = Vec::new();
    for id in ids {
        let book = rels
            .get(&id)
            .and_then(|r| load_book(archive, &part_path(&r.target)))
            .unwrap_or(LoadedBook {
                target: String::new(),
                file: String::new(),
                sheets: Vec::new(),
            });
        books.push(book);
    }
    books
}

fn load_book<R: Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    path: &str,
) -> Option<LoadedBook> {
    let text = read_text(archive, path)?;
    let root = XmlNode::parse(text.as_bytes()).ok()?;
    let book = root.children().find(|n| n.has_tag_name("externalBook"))?;
    // путь к файлу — в связях самой части
    let (dir, name) = path.rsplit_once('/').unwrap_or(("", path));
    let rels_text = read_text(archive, &format!("{dir}/_rels/{name}.rels"))?;
    let rels_root = XmlNode::parse(rels_text.as_bytes()).ok()?;
    let rid = book.attribute_ns(R_NS, "id")?;
    let target = rels_root
        .children()
        .filter(|n| n.has_tag_name("Relationship"))
        .find(|n| n.attribute("Id") == Some(rid))?
        .attribute("Target")?
        .to_string();
    let file = file_name_of(&target).to_string();
    let names: Vec<String> = book
        .children()
        .find(|n| n.has_tag_name("sheetNames"))
        .map(|sn| {
            sn.children()
                .filter(|n| n.has_tag_name("sheetName"))
                .filter_map(|n| n.attribute("val").map(|s| s.to_string()))
                .collect()
        })
        .unwrap_or_default();
    let mut values: Vec<Vec<(i32, i32, ExternalValue)>> = vec![Vec::new(); names.len()];
    if let Some(set) = book.children().find(|n| n.has_tag_name("sheetDataSet")) {
        for data in set.children().filter(|n| n.has_tag_name("sheetData")) {
            let Some(i) = data
                .attribute("sheetId")
                .and_then(|s| s.parse::<usize>().ok())
            else {
                continue;
            };
            let Some(list) = values.get_mut(i) else {
                continue;
            };
            for row in data.children().filter(|n| n.has_tag_name("row")) {
                for cell in row.children().filter(|n| n.has_tag_name("cell")) {
                    let Some(r) = cell.attribute("r").and_then(parse_reference_a1) else {
                        continue;
                    };
                    let Some(v) = cell
                        .children()
                        .find(|n| n.has_tag_name("v"))
                        .and_then(|n| n.text())
                        .map(|s| s.to_string())
                    else {
                        continue;
                    };
                    let value = match cell.attribute("t") {
                        Some("s") | Some("str") | Some("inlineStr") => ExternalValue::Text(v),
                        Some("b") => ExternalValue::Bool(v == "1" || v == "true"),
                        Some("e") => ExternalValue::Error(v),
                        _ => match v.parse::<f64>() {
                            Ok(n) => ExternalValue::Number(n),
                            Err(_) => ExternalValue::Text(v),
                        },
                    };
                    list.push((r.row, r.column, value));
                }
            }
        }
    }
    Some(LoadedBook {
        target,
        file,
        sheets: names.into_iter().zip(values).collect(),
    })
}
