use std::collections::HashMap;
use std::io::{BufReader, Read};

use ironcalc_base::types::RichRun;

use super::xml::XmlNode;

use crate::error::XlsxError;

/// The workbook's shared strings while its sheets are read: the list, and an
/// index from each string to its first position in it. Inline-string and
/// formula-string (`t="str"`) cells look their text up here; a linear search
/// of the list per cell made such workbooks quadratic to load.
pub(crate) struct SharedStringTable {
    strings: Vec<String>,
    index: HashMap<String, i32>,
}

impl SharedStringTable {
    pub(crate) fn new(strings: Vec<String>) -> SharedStringTable {
        Self::new_excluding(strings, &HashMap::new())
    }

    /// As [`Self::new`], but the rich strings are not found by their text (they belong to one cell).
    pub(crate) fn new_excluding(strings: Vec<String>, rich: &HashMap<i32, Vec<RichRun>>) -> SharedStringTable {
        let mut index = HashMap::with_capacity(strings.len());
        for (i, s) in strings.iter().enumerate() {
            if rich.contains_key(&(i as i32)) {
                continue;
            }
            // The first occurrence wins, as a search of the list would find it.
            index.entry(s.clone()).or_insert(i as i32);
        }
        SharedStringTable { strings, index }
    }

    /// The position of `s` in the list, appending it if it is not there yet.
    pub(crate) fn index_of(&mut self, s: &str) -> i32 {
        if let Some(i) = self.index.get(s) {
            return *i;
        }
        let i = self.strings.len() as i32;
        self.strings.push(s.to_string());
        self.index.insert(s.to_string(), i);
        i
    }

    pub(crate) fn into_strings(self) -> Vec<String> {
        self.strings
    }
}

/// Reads the list of shared strings in an Excel workbook
/// Note than in IronCalc we lose _internal_ styling of a string
/// See Section 18.4
pub(crate) fn read_shared_strings<R: Read + std::io::Seek>(
    archive: &mut zip::read::ZipArchive<R>,
) -> Result<Vec<String>, XlsxError> {
    match archive.by_name("xl/sharedStrings.xml") {
        Ok(file) => read_shared_strings_from_xml(XmlNode::parse(BufReader::new(file))?),
        Err(_e) => Ok(Vec::new()),
    }
}

/// ENS-Office: the formatted runs (`<r>`) of the shared strings, by the index of the string.
pub(crate) fn read_rich_runs<R: Read + std::io::Seek>(
    archive: &mut zip::read::ZipArchive<R>,
) -> HashMap<i32, Vec<RichRun>> {
    let mut out = HashMap::new();
    let Ok(file) = archive.by_name("xl/sharedStrings.xml") else { return out };
    let Ok(doc) = XmlNode::parse(BufReader::new(file)) else { return out };
    for (i, si) in doc.descendants().filter(|n| n.has_tag_name("si")).enumerate() {
        let mut runs = Vec::new();
        for r in si.children().filter(|n| n.has_tag_name("r")) {
            let text = r
                .children()
                .filter(|n| n.has_tag_name("t"))
                .map(|n| decode_xlsx_escapes(n.text().unwrap_or("")))
                .collect::<String>();
            let mut run = RichRun { text, ..Default::default() };
            if let Some(pr) = r.children().find(|n| n.has_tag_name("rPr")) {
                let on = |n: &XmlNode| n.attribute("val").is_none_or(|v| v != "0" && v != "false");
                for p in pr.children() {
                    match p.tag_name() {
                        "b" => run.b = on(p),
                        "i" => run.i = on(p),
                        "strike" => run.strike = on(p),
                        "u" => run.u = p.attribute("val").is_none_or(|v| v != "none"),
                        "sz" => run.sz = p.attribute("val").and_then(|v| v.parse().ok()),
                        "rFont" => run.name = p.attribute("val").map(|v| v.to_string()),
                        "color" => {
                            run.color = p.attribute("rgb").filter(|v| v.len() == 8).map(|v| format!("#{}", &v[2..]));
                        }
                        "vertAlign" => run.vert = match p.attribute("val") { Some("superscript") => 1, Some("subscript") => 2, _ => 0 },
                        _ => {}
                    }
                }
            }
            runs.push(run);
        }
        if !runs.is_empty() {
            out.insert(i as i32, runs);
        }
    }
    out
}

#[cfg(test)]
fn read_shared_strings_from_string(text: &str) -> Result<Vec<String>, XlsxError> {
    read_shared_strings_from_xml(XmlNode::parse(text.as_bytes())?)
}

fn read_shared_strings_from_xml(doc: XmlNode) -> Result<Vec<String>, XlsxError> {
    let mut shared_strings = Vec::new();
    let nodes: Vec<&XmlNode> = doc.descendants().filter(|n| n.has_tag_name("si")).collect();
    for node in nodes {
        let text = node
            .descendants()
            .filter(|n| n.has_tag_name("t"))
            .map(|n| decode_xlsx_escapes(n.text().unwrap_or("")))
            .collect::<Vec<String>>()
            .join("");
        shared_strings.push(text);
    }
    Ok(shared_strings)
}

/// Decodes Excel's `_xXXXX_` escape sequences for characters that are invalid in XML 1.0.
/// For example, `_x0001_` → U+0001 (SOH).
pub(crate) fn decode_xlsx_escapes(s: &str) -> String {
    if !s.contains("_x") {
        return s.to_string();
    }

    let bytes = s.as_bytes();
    let len = bytes.len();
    let mut result = String::with_capacity(len);
    let mut i = 0;

    while i < len {
        // Look for _xHHHH_ pattern (7 bytes: '_', 'x', 4 hex digits, '_')
        if i + 6 < len && bytes[i] == b'_' && bytes[i + 1] == b'x' && bytes[i + 6] == b'_' {
            let hex = &s[i + 2..i + 6];

            if hex.chars().all(|c| c.is_ascii_hexdigit()) {
                if let Ok(code) = u32::from_str_radix(hex, 16) {
                    if let Some(c) = char::from_u32(code) {
                        result.push(c);
                        i += 7;
                        continue;
                    }
                }
            }
        }

        if let Some(c) = s[i..].chars().next() {
            result.push(c);
            i += c.len_utf8();
        } else {
            break;
        }
    }

    result
}

#[cfg(test)]
mod tests {
    #![allow(clippy::unwrap_used)]
    use super::*;

    #[test]
    fn test_shared_strings() {
        let xml_string = r#"
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
    <si>
        <t>A string</t>
    </si>
    <si>
        <t>A second String</t>
    </si>
    <si>
        <r>
            <t>Hello</t>
        </r>
            <r>
            <rPr>
                <b/>
                <sz val="11"/>
                <color rgb="FFFF0000"/>
                <rFont val="Inter"/>
                <family val="2"/>
                <scheme val="minor"/>
            </rPr>
            <t xml:space="preserve"> World</t>
        </r>
    </si>
</sst>"#;
        let shared_strings = read_shared_strings_from_string(xml_string.trim()).unwrap();
        assert_eq!(
            shared_strings,
            [
                "A string".to_string(),
                "A second String".to_string(),
                "Hello World".to_string()
            ]
        );
    }

    #[test]
    fn test_decode_xlsx_escapes_control_chars() {
        assert_eq!(decode_xlsx_escapes("_x0001_"), "\x01");
        assert_eq!(decode_xlsx_escapes("_x001F_"), "\x1F");
        assert_eq!(decode_xlsx_escapes("_x000B_"), "\x0B");
        assert_eq!(decode_xlsx_escapes("hello_x0001_world"), "hello\x01world");
    }

    #[test]
    fn test_decode_xlsx_escapes_literal_underscore() {
        // _x005F_ decodes to underscore; the rest is literal
        assert_eq!(decode_xlsx_escapes("_x005F_x0001_"), "_x0001_");
        assert_eq!(decode_xlsx_escapes("_x005F_x005F_"), "_x005F_");
    }

    #[test]
    fn test_decode_xlsx_escapes_non_patterns() {
        // Strings that look similar but are not valid _xHHHH_ patterns
        assert_eq!(decode_xlsx_escapes("_xGGGG_"), "_xGGGG_"); // non-hex digits
        assert_eq!(decode_xlsx_escapes("_x001"), "_x001"); // too short
        assert_eq!(decode_xlsx_escapes("_x00001_"), "_x00001_"); // 5 hex digits
        assert_eq!(decode_xlsx_escapes("plain text"), "plain text");
        assert_eq!(decode_xlsx_escapes("_hello"), "_hello");
    }

    #[test]
    fn test_decode_xlsx_escapes_nul() {
        assert_eq!(decode_xlsx_escapes("_x0000_"), "\x00");
    }
}

#[cfg(test)]
mod table_tests {
    use super::SharedStringTable;

    #[test]
    fn finds_the_first_occurrence_and_appends_new_strings() {
        let mut table =
            SharedStringTable::new(vec!["a".to_string(), "b".to_string(), "a".to_string()]);
        assert_eq!(table.index_of("a"), 0);
        assert_eq!(table.index_of("b"), 1);
        assert_eq!(table.index_of("c"), 3);
        assert_eq!(table.index_of("c"), 3);
        assert_eq!(table.into_strings(), ["a", "b", "a", "c"]);
    }
}
