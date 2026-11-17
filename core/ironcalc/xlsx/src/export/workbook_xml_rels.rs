use ironcalc_base::types::Workbook;

use super::xml_constants::{XML_DECLARATION, XML_WORKSHEET};

pub(crate) fn get_workbook_xml_rels(workbook: &Workbook) -> String {
    let mut relationships_str: Vec<String> = vec![];
    let worksheet_count = super::real_sheets(workbook).len() + 1;
    for id in 1..worksheet_count {
        relationships_str.push(format!(
            "<Relationship Id=\"rId{id}\" Type=\"{XML_WORKSHEET}\" Target=\"worksheets/sheet{id}.xml\"/>"
        ));
    }
    let mut id = worksheet_count;
    relationships_str.push(
        format!("<Relationship Id=\"rId{id}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>")
    );
    id += 1;
    relationships_str.push(
        format!("<Relationship Id=\"rId{id}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>")
    );
    id += 1;
    relationships_str.push(
        format!("<Relationship Id=\"rId{id}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata\" Target=\"metadata.xml\"/>")
    );
    id += 1;
    relationships_str.push(
        format!("<Relationship Id=\"rId{id}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>")
    );
    if !workbook.ens_change_log.is_empty() {
        id += 1;
        relationships_str.push(format!(
            "<Relationship Id=\"rId{id}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml\" Target=\"../customXml/item1.xml\"/>"
        ));
    }
    if workbook.ens_vba_parts.iter().any(|p| p.path.eq_ignore_ascii_case("xl/vbaProject.bin")) {
        id += 1;
        relationships_str.push(format!(
            "<Relationship Id=\"rId{id}\" Type=\"http://schemas.microsoft.com/office/2006/relationships/vbaProject\" Target=\"vbaProject.bin\"/>"
        ));
    }
    for n in 1..=super::written_books(workbook).len() {
        relationships_str.push(format!(
            "<Relationship Id=\"rIdExt{n}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink\" Target=\"externalLinks/externalLink{n}.xml\"/>"
        ));
    }
    format!(
        "{XML_DECLARATION}\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{}</Relationships>",
        relationships_str.join("")
    )
}
