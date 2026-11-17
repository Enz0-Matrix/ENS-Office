// SPDX-License-Identifier: AGPL-3.0-or-later
//! ENS-Office: внешние связи в xlsx — `xl/externalLinks/externalLinkN.xml` с кэшем значений
//! листов книги-источника и путём к ней (`_rels/externalLinkN.xml.rels`).

use ironcalc_base::expressions::utils::number_to_column;
use ironcalc_base::external::{external_sheet_name, ExternalValue};
use ironcalc_base::types::ExternalBook;
use ironcalc_base::Model;

use super::escape::escape_xml;
use super::xml_constants::XML_DECLARATION;

/// Путь к книге для Excel: абсолютный путь Windows или Linux — как `file:///…`.
fn excel_target(target: &str) -> String {
    let t = target.trim();
    let windows_abs = t.len() > 2
        && t.as_bytes()[1] == b':'
        && (t.as_bytes()[2] == b'\\' || t.as_bytes()[2] == b'/');
    if t.starts_with("file:")
        || t.starts_with("http")
        || !(windows_abs || t.starts_with('/') || t.starts_with("\\\\"))
    {
        t.to_string()
    } else {
        format!("file:///{}", t.trim_start_matches('/'))
    }
}

pub(crate) fn get_external_link_xml(model: &Model, book: &ExternalBook) -> String {
    let sheets = model.external_sheets_of(&book.file);
    let names: String = sheets
        .iter()
        .map(|s| format!("<sheetName val=\"{}\"/>", escape_xml(s)))
        .collect();
    let mut data = String::new();
    for (i, sheet) in sheets.iter().enumerate() {
        let name = external_sheet_name(&book.file, sheet).to_lowercase();
        let Some(index) = model
            .workbook
            .worksheets
            .iter()
            .position(|w| w.name.to_lowercase() == name)
        else {
            continue;
        };
        let mut values = model.sheet_values(index as u32);
        values.sort_by_key(|(r, c, _)| (*r, *c));
        data.push_str(&format!("<sheetData sheetId=\"{i}\""));
        if values.is_empty() {
            data.push_str("/>");
            continue;
        }
        data.push('>');
        let mut current_row = 0;
        for (row, column, value) in values {
            if row != current_row {
                if current_row != 0 {
                    data.push_str("</row>");
                }
                data.push_str(&format!("<row r=\"{row}\">"));
                current_row = row;
            }
            let cell = format!("{}{row}", number_to_column(column).unwrap_or_default());
            let (t, v) = match value {
                ExternalValue::Number(n) => ("", n.to_string()),
                ExternalValue::Bool(b) => (
                    " t=\"b\"",
                    if b { "1".to_string() } else { "0".to_string() },
                ),
                ExternalValue::Text(s) => (" t=\"str\"", escape_xml(&s).to_string()),
                ExternalValue::Error(e) => (" t=\"e\"", escape_xml(&e).to_string()),
            };
            data.push_str(&format!("<cell r=\"{cell}\"{t}><v>{v}</v></cell>"));
        }
        data.push_str("</row></sheetData>");
    }
    format!(
        "{XML_DECLARATION}\n<externalLink xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" \
xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\
<externalBook r:id=\"rId1\"><sheetNames>{names}</sheetNames><sheetDataSet>{data}</sheetDataSet></externalBook></externalLink>"
    )
}

pub(crate) fn get_external_link_rels(book: &ExternalBook) -> String {
    format!(
        "{XML_DECLARATION}\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath\" \
Target=\"{}\" TargetMode=\"External\"/></Relationships>",
        escape_xml(&excel_target(&book.target))
    )
}
