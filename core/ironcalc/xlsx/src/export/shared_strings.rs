use ironcalc_base::types::{RichRun, Workbook};

use super::{escape::escape_xml, xml_constants::XML_DECLARATION};

pub(crate) fn get_shared_strings_xml(model: &Workbook) -> String {
    let mut shared_strings: Vec<String> = vec![];
    let count = &model.shared_strings.len();
    let unique_count = &model.shared_strings.len();
    for (index, shared_string) in model.shared_strings.iter().enumerate() {
        if let Some(runs) = model.ens_rich.get(&(index as i32)) {
            shared_strings.push(format!("<si>{}</si>", runs.iter().map(rich_run_xml).collect::<String>()));
            continue;
        }
        shared_strings.push(format!("<si><t>{}</t></si>", escape_xml(shared_string)));
    }
    format!("{}\n\
      <sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{count}\" uniqueCount=\"{unique_count}\">\
        {}\
      </sst>", XML_DECLARATION, shared_strings.join(""))
}

/// ENS-Office: one run of a rich text, `<r><rPr>…</rPr><t>…</t></r>`.
fn rich_run_xml(run: &RichRun) -> String {
    let mut pr = String::new();
    if let Some(n) = &run.name {
        pr.push_str(&format!("<rFont val=\"{}\"/>", escape_xml(n)));
    }
    if run.b {
        pr.push_str("<b/>");
    }
    if run.i {
        pr.push_str("<i/>");
    }
    if run.strike {
        pr.push_str("<strike/>");
    }
    if let Some(c) = run.color.as_deref().and_then(|c| c.strip_prefix('#')).filter(|c| c.len() == 6) {
        pr.push_str(&format!("<color rgb=\"FF{}\"/>", c.to_uppercase()));
    }
    if let Some(sz) = run.sz {
        pr.push_str(&format!("<sz val=\"{sz}\"/>"));
    }
    if run.u {
        pr.push_str("<u/>");
    }
    match run.vert {
        1 => pr.push_str("<vertAlign val=\"superscript\"/>"),
        2 => pr.push_str("<vertAlign val=\"subscript\"/>"),
        _ => {}
    }
    let pr = if pr.is_empty() { String::new() } else { format!("<rPr>{pr}</rPr>") };
    format!("<r>{pr}<t xml:space=\"preserve\">{}</t></r>", escape_xml(&run.text))
}
