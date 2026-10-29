#![allow(clippy::unwrap_used)]
#![allow(clippy::panic)]

//! # A note on shared formulas
//! Although both Excel and IronCalc uses shared formulas they are used in a slightly different way that cannot be mapped 1-1
//! In IronCalc _all_ formulas are shared and there is a list of shared formulas much like there is a list of shared strings.
//! In Excel the situation in more nuanced. A shared formula is shared amongst a rage of cells.
//! The top left cell would be the "mother" cell that would contain the shared formula:
//! ```xml
//! <c r="F4" t="str">
//!    <f t="shared" ref="F4:F8" si="42">A4+C4</f>
//!    <v>123</v>
//! </c>
//! ```
//! Cells in the range F4:F8 will then link to that formula like so:
//! ```xml
//! <c r="F6">
//!   <f t="shared" si="42"/>
//!   <v>1</v>
//! </c>
//! ```
//! Formula in F6 would then be 'A6+C6'
use std::collections::HashMap;

use itertools::Itertools;

use ironcalc_base::{
    expressions::{
        parser::{
            static_analysis::StaticResult,
            stringify::{to_excel_array_formula_string, to_excel_string},
            Node,
        },
        types::CellReferenceRC,
        utils::{number_to_column, parse_reference_a1},
    },
    types::{ArrayKind, Cell, FormulaValue, Link, SpillValue, Worksheet},
};

use crate::export::conditional_formatting::get_conditional_formatting_xml;

use super::{escape::escape_xml, xml_constants::XML_DECLARATION};

fn get_range_str(row: i32, column: i32, width: i32, height: i32) -> Option<String> {
    let column1 = number_to_column(column)?;
    let column2 = number_to_column(column + width - 1)?;
    let row2 = row + height - 1;
    Some(format!("{}{}:{}{}", column1, row, column2, row2))
}

fn get_cell_style_attribute(s: i32) -> String {
    if s == 0 {
        "".to_string()
    } else {
        format!(" s=\"{s}\"")
    }
}

/// The links of the worksheet sorted by (row, column) so that both the
/// `<hyperlinks>` section and the sheet rels assign relationship ids in the
/// same, deterministic order.
fn get_sorted_links(worksheet: &Worksheet) -> Vec<((i32, i32), &Link)> {
    worksheet
        .links
        .iter()
        .map(|(k, v)| (*k, v))
        .sorted_by_key(|x| x.0)
        .collect()
}

/// An external link whose target is `scheme://host/path#location` is exported with the
/// fragment in the `location` attribute of the `<hyperlink>` element, matching what
/// Excel does. Returns (target, location).
fn split_external_target(target: &str) -> (&str, Option<&str>) {
    match target.split_once('#') {
        Some((t, location)) => (t, Some(location)),
        None => (target, None),
    }
}

/// The `<hyperlinks>` section of the worksheet or an empty string if there are none.
/// External links reference a relationship (`rId1`, `rId2`...) in the sheet rels part,
/// see [`get_worksheet_xml_rels`].
fn get_hyperlinks_section(worksheet: &Worksheet) -> String {
    let links = get_sorted_links(worksheet);
    if links.is_empty() {
        return "".to_string();
    }
    let mut hyperlinks_str: Vec<String> = vec![];
    let mut rel_id = 0;
    for ((row, column), link) in links {
        let column_name = match number_to_column(column) {
            Some(name) => name,
            None => continue,
        };
        let cell_name = format!("{column_name}{row}");
        match link {
            Link::External { target, tooltip } => {
                rel_id += 1;
                let location_attribute = match split_external_target(target).1 {
                    Some(location) => format!(" location=\"{}\"", escape_xml(location)),
                    None => "".to_string(),
                };
                let tooltip_attribute = get_tooltip_attribute(tooltip);
                hyperlinks_str.push(format!(
                    "<hyperlink ref=\"{cell_name}\" r:id=\"rId{rel_id}\"{location_attribute}{tooltip_attribute}/>"
                ));
            }
            Link::Internal { location, tooltip } => {
                let location = escape_xml(location);
                let tooltip_attribute = get_tooltip_attribute(tooltip);
                hyperlinks_str.push(format!(
                    "<hyperlink ref=\"{cell_name}\" location=\"{location}\"{tooltip_attribute}/>"
                ));
            }
        }
    }
    format!("<hyperlinks>{}</hyperlinks>", hyperlinks_str.join(""))
}

fn get_tooltip_attribute(tooltip: &Option<String>) -> String {
    match tooltip {
        Some(tooltip) => format!(" tooltip=\"{}\"", escape_xml(tooltip)),
        None => "".to_string(),
    }
}

/// Number of relationships taken by external hyperlinks (they use `rId1`..).
fn external_link_count(worksheet: &Worksheet) -> usize {
    get_sorted_links(worksheet)
        .iter()
        .filter(|(_, l)| matches!(l, Link::External { .. }))
        .count()
}

/// ENS-Office: `<rowBreaks>` / `<colBreaks>` of the sheet's manual page breaks.
fn page_breaks_xml(worksheet: &Worksheet) -> String {
    let section = |tag: &str, list: &[i32], max: i32| {
        if list.is_empty() {
            return String::new();
        }
        let brks: String = list
            .iter()
            .map(|id| format!("<brk id=\"{id}\" max=\"{max}\" man=\"1\"/>"))
            .collect();
        format!("<{tag} count=\"{}\" manualBreakCount=\"{}\">{brks}</{tag}>", list.len(), list.len())
    };
    format!(
        "{}{}",
        section("rowBreaks", &worksheet.ens_row_breaks, 16383),
        section("colBreaks", &worksheet.ens_col_breaks, 1048575)
    )
}

/// ENS-Office: `xl/commentsN.xml` of the sheet's comments (notes).
pub(crate) fn get_comments_xml(worksheet: &Worksheet) -> String {
    let mut authors: Vec<&str> = vec![];
    let mut list = String::new();
    for c in &worksheet.comments {
        let a = match authors.iter().position(|x| *x == c.author_name) {
            Some(i) => i,
            None => {
                authors.push(&c.author_name);
                authors.len() - 1
            }
        };
        list.push_str(&format!(
            "<comment ref=\"{}\" authorId=\"{a}\"><text><t xml:space=\"preserve\">{}</t></text></comment>",
            escape_xml(&c.cell_ref),
            escape_xml(&c.text)
        ));
    }
    let authors: String = authors.iter().map(|a| format!("<author>{}</author>", escape_xml(a))).collect();
    format!(
        "{XML_DECLARATION}\n<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><authors>{authors}</authors><commentList>{list}</commentList></comments>"
    )
}

/// ENS-Office: `xl/drawings/vmlDrawingN.vml`, the hidden note boxes Excel needs to show comments.
pub(crate) fn get_comments_vml(worksheet: &Worksheet) -> String {
    let mut shapes = String::new();
    for (i, c) in worksheet.comments.iter().enumerate() {
        let Some(r) = parse_reference_a1(&c.cell_ref) else { continue };
        let (row, col) = (r.row - 1, r.column - 1);
        shapes.push_str(&format!(
            "<v:shape id=\"_x0000_s{}\" type=\"#_x0000_t202\" style=\"position:absolute;margin-left:80pt;margin-top:5pt;width:140pt;height:60pt;z-index:{};visibility:hidden\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\">\
<v:fill color2=\"#ffffe1\"/><v:shadow on=\"t\" color=\"black\" obscured=\"t\"/><v:path o:connecttype=\"none\"/>\
<v:textbox style=\"mso-direction-alt:auto\"><div style=\"text-align:left\"></div></v:textbox>\
<x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>{}, 15, {}, 2, {}, 15, {}, 16</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>{row}</x:Row><x:Column>{col}</x:Column></x:ClientData></v:shape>",
            1025 + i,
            i + 1,
            col + 1,
            row.max(1) - 1,
            col + 3,
            row + 3
        ));
    }
    format!(
        "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\
<o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/></o:shapelayout>\
<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/><v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype>{shapes}</xml>"
    )
}

/// The rels part of the worksheet (`xl/worksheets/_rels/sheetN.xml.rels`) holding one
/// relationship per external link, or `None` if the worksheet has no external links.
pub(crate) fn get_worksheet_xml_rels(worksheet: &Worksheet, comments_id: Option<usize>, drawing_id: Option<usize>) -> Option<String> {
    let mut relationships_str: Vec<String> = vec![];
    let mut rel_id = 0;
    for (_, link) in get_sorted_links(worksheet) {
        if let Link::External { target, .. } = link {
            rel_id += 1;
            let target = escape_xml(split_external_target(target).0);
            relationships_str.push(format!(
                "<Relationship Id=\"rId{rel_id}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"{target}\" TargetMode=\"External\"/>"
            ));
        }
    }
    if let Some(id) = comments_id {
        relationships_str.push(format!(
            "<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" Target=\"../comments{id}.xml\"/>\
<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\" Target=\"../drawings/vmlDrawing{id}.vml\"/>",
            rel_id + 1,
            rel_id + 2
        ));
    }
    if let Some(id) = drawing_id {
        let n = rel_id + if comments_id.is_some() { 2 } else { 0 } + 1;
        relationships_str.push(format!(
            "<Relationship Id=\"rId{n}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing{id}.xml\"/>"
        ));
    }
    if relationships_str.is_empty() {
        return None;
    }
    Some(format!(
        "{XML_DECLARATION}\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{}</Relationships>",
        relationships_str.join("")
    ))
}

fn get_formula_attribute(
    sheet_name: String,
    row: i32,
    column: i32,
    parsed_formula: &Node,
) -> String {
    let cell_ref = CellReferenceRC {
        sheet: sheet_name,
        row,
        column,
    };
    let formula = &to_excel_string(parsed_formula, &cell_ref);
    escape_xml(formula).to_string()
}

/// The formula of an array, CSE or dynamic. It keeps every implicit
/// intersection: nothing puts them back when the formula is read again.
fn get_array_formula_attribute(
    sheet_name: String,
    row: i32,
    column: i32,
    parsed_formula: &Node,
) -> String {
    let cell_ref = CellReferenceRC {
        sheet: sheet_name,
        row,
        column,
    };
    let formula = &to_excel_array_formula_string(parsed_formula, &cell_ref);
    escape_xml(formula).to_string()
}

/// ENS-Office: an error value Excel understands.
fn excel_error_name(name: &str) -> &str {
    match name {
        "#NULL!" | "#DIV/0!" | "#VALUE!" | "#REF!" | "#NAME?" | "#NUM!" | "#N/A" | "#SPILL!"
        | "#CALC!" => name,
        "#N/IMPL" | "#N/IMPL!" => "#NAME?",
        _ => "#VALUE!",
    }
}

pub(crate) fn get_worksheet_xml(
    auto_filter: &str,
    worksheet: &Worksheet,
    parsed_formulas: &[(std::sync::Arc<Node>, StaticResult)],
    dimension: &str,
    is_sheet_selected: bool,
) -> String {
    let mut sheet_data_str: Vec<String> = vec![];
    let mut cols_str: Vec<String> = vec![];
    let mut merged_cells_str: Vec<String> = vec![];

    for col in &worksheet.cols {
        // <col min="4" max="4" width="12" customWidth="1"/>
        let min = col.min;
        let max = col.max;
        let width = col.width;
        let custom_width = i32::from(col.custom_width);
        let column_style = match col.style {
            Some(s) => format!(" style=\"{s}\""),
            None => "".to_string(),
        };
        let hidden = if col.hidden { " hidden=\"1\"" } else { "" };
        // ENS-Office: grouping (outline) of columns.
        let outline = if col.outline_level > 0 { format!(" outlineLevel=\"{}\"", col.outline_level) } else { String::new() };
        let collapsed = if col.collapsed { " collapsed=\"1\"" } else { "" };
        cols_str.push(format!(
            "<col min=\"{min}\" max=\"{max}\" width=\"{width}\" customWidth=\"{custom_width}\"{column_style}{hidden}{outline}{collapsed}/>"
        ));
    }

    // this is a bit of an overkill. A dictionary of the row styles by row_index
    let mut row_style_dict = HashMap::new();
    for row in &worksheet.rows {
        // {
        //     "height": 13,
        //     "r": 7,
        //     "custom_format": false,
        //     "custom_height": true,
        //     "s": 0
        //     "hidden": false,
        //   },
        row_style_dict.insert(row.r, row.clone());
    }

    // A row with no cells is still written when it has a height or is hidden.
    let mut row_indices = worksheet.sheet_data.rows();
    row_indices.extend(row_style_dict.keys());
    row_indices.sort_unstable();
    row_indices.dedup();
    for row_index in row_indices {
        let mut row_data_str: Vec<String> = vec![];
        for (column_index, cell) in worksheet.sheet_data.cells_in_row(row_index) {
            let column_name = number_to_column(column_index).unwrap();
            let cell_name = format!("{column_name}{row_index}");
            match &*cell {
                Cell::EmptyCell { s } => {
                    // they only hold the style
                    let style = get_cell_style_attribute(*s);
                    row_data_str.push(format!("<c r=\"{cell_name}\"{style}/>"));
                }
                Cell::BooleanCell { v, s }
                | Cell::SpillCell {
                    v: SpillValue::Boolean(v),
                    s,
                    ..
                } => {
                    // <c r="A8" t="b" s="1">
                    //     <v>1</v>
                    // </c>
                    let b = i32::from(*v);
                    let style = get_cell_style_attribute(*s);
                    row_data_str.push(format!(
                        "<c r=\"{cell_name}\" t=\"b\"{style}><v>{b}</v></c>"
                    ));
                }
                Cell::NumberCell { v, s }
                | Cell::SpillCell {
                    v: SpillValue::Number(v),
                    s,
                    ..
                } => {
                    // Normally the type number is left out. Example:
                    // <c r="C6" s="1">
                    //     <v>3</v>
                    // </c>
                    let style = get_cell_style_attribute(*s);
                    row_data_str.push(format!("<c r=\"{cell_name}\"{style}><v>{v}</v></c>"));
                }
                Cell::ErrorCell { ei, s }
                | Cell::SpillCell {
                    v: SpillValue::Error(ei),
                    s,
                    ..
                } => {
                    let style = get_cell_style_attribute(*s);
                    row_data_str.push(format!(
                        "<c r=\"{cell_name}\" t=\"e\"{style}><v>{ei}</v></c>"
                    ));
                }
                Cell::SharedString { si, s } => {
                    // Example:
                    // <c r="A1" s="1" t="s">
                    //    <v>5</v>
                    // </c>
                    // Cell on A1 contains a string (t="s") of style="1". The string is the 6th in the list of shared strings
                    let style = get_cell_style_attribute(*s);
                    row_data_str.push(format!(
                        "<c r=\"{cell_name}\" t=\"s\"{style}><v>{si}</v></c>"
                    ));
                }
                Cell::CellFormula {
                    v: FormulaValue::Unevaluated,
                    ..
                }
                | Cell::ArrayFormula {
                    v: FormulaValue::Unevaluated,
                    ..
                } => {
                    // TODO: We should NOT panic here.
                    // Is there like an "import error" we can set in Excel?
                    panic!("Model needs to be evaluated before saving!");
                }
                Cell::CellFormula {
                    f,
                    v: FormulaValue::Boolean(v),
                    s,
                } => {
                    // <c r="A4" t="b" s="3">
                    //   <f>ISTEXT(A5)</f>
                    //   <v>1</v>
                    // </c>
                    let style = get_cell_style_attribute(*s);
                    let formula = get_formula_attribute(
                        worksheet.get_name(),
                        row_index,
                        column_index,
                        &parsed_formulas[*f as usize].0,
                    );
                    let b = i32::from(*v);
                    row_data_str.push(format!(
                        "<c r=\"{cell_name}\" t=\"b\"{style}><f>{formula}</f><v>{b}</v></c>"
                    ));
                }
                Cell::CellFormula {
                    f,
                    v: FormulaValue::Number(v),
                    s,
                } => {
                    // Note again type is skipped
                    // <c r="C4" s="3">
                    //   <f>A5+C3</f>
                    //   <v>123</v>
                    // </c>
                    let formula = get_formula_attribute(
                        worksheet.get_name(),
                        row_index,
                        column_index,
                        &parsed_formulas[*f as usize].0,
                    );
                    let style = get_cell_style_attribute(*s);
                    row_data_str.push(format!(
                        "<c r=\"{cell_name}\"{style}><f>{formula}</f><v>{v}</v></c>"
                    ));
                }
                Cell::CellFormula {
                    f,
                    v: FormulaValue::Text(v),
                    s,
                } => {
                    // <c r="C6" t="str" s="5">
                    //   <f>CONCATENATE(A1, A2)</f>
                    //   <v>Hello world!</v>
                    // </c>
                    let formula = get_formula_attribute(
                        worksheet.get_name(),
                        row_index,
                        column_index,
                        &parsed_formulas[*f as usize].0,
                    );
                    let style = get_cell_style_attribute(*s);
                    let escaped_v = escape_xml(v);
                    row_data_str.push(format!(
                        "<c r=\"{cell_name}\" t=\"str\"{style}><f>{formula}</f><v>{escaped_v}</v></c>"
                    ));
                }
                Cell::CellFormula {
                    f,
                    v: FormulaValue::Error { ei, .. },
                    s,
                } => {
                    // <c r="C6" t="e" s="4">
                    //   <f>A1/A3<f/>
                    //   <v>#DIV/0!</v>
                    // </c>
                    let node = &parsed_formulas[*f as usize].0;
                    let style = get_cell_style_attribute(*s);
                    // ENS-Office: a formula that does not parse (say "=B5+") is not
                    // valid in Excel, which "repairs" the whole file and drops it.
                    // Keep the text the user typed as an inline string instead.
                    if let Node::ParseErrorKind { formula, .. } = node.as_ref() {
                        let text = escape_xml(&format!("={formula}")).to_string();
                        row_data_str.push(format!(
                            "<c r=\"{cell_name}\" t=\"inlineStr\"{style}><is><t>{text}</t></is></c>"
                        ));
                        continue;
                    }
                    let formula = get_formula_attribute(
                        worksheet.get_name(),
                        row_index,
                        column_index,
                        node,
                    );
                    // ENS-Office: #ERROR!, #N/IMPL and #CIRC! are IronCalc's own
                    // errors; Excel only accepts its standard ones.
                    let ei_name = ei.to_string();
                    let ei = excel_error_name(&ei_name);
                    row_data_str.push(format!(
                        "<c r=\"{cell_name}\" t=\"e\"{style}><f>{formula}</f><v>{ei}</v></c>"
                    ));
                }
                Cell::SpillCell {
                    v: SpillValue::Text(v),
                    s,
                    ..
                } => {
                    // inline string
                    // <c r="C6" t="str" s="5">
                    //   <v>Hello world!</v>
                    // </c>
                    let style = get_cell_style_attribute(*s);
                    let escaped_v = escape_xml(v);
                    row_data_str.push(format!(
                        "<c r=\"{cell_name}\" t=\"str\"{style}><v>{escaped_v}</v></c>"
                    ));
                }
                // The difference between dynamic and array formulas is that dynamic formulas have cm="1"
                Cell::ArrayFormula {
                    f,
                    v: FormulaValue::Boolean(v),
                    s,
                    r,
                    kind,
                } => {
                    // Dynamic: <c r="D1" s="3" t="b" cm="1">
                    //            <f t="array" ref="D1:D5">A1:A5</f>
                    //            <v>TRUE</v>
                    //          </c>
                    // CSE:     <c r="D1" s="3" t="b">
                    //            <f t="array" ref="D1:D5">A1:A5</f>
                    //            <v>TRUE</v>
                    //          </c>
                    let node = match parsed_formulas.get(*f as usize) {
                        Some(node) => node,
                        None => continue,
                    };
                    let formula = get_array_formula_attribute(
                        worksheet.get_name(),
                        row_index,
                        column_index,
                        &node.0,
                    );
                    let style = get_cell_style_attribute(*s);
                    let range = match get_range_str(row_index, column_index, r.0, r.1) {
                        Some(range) => range,
                        None => continue,
                    };
                    let b = i32::from(*v);
                    let cm = if matches!(kind, ArrayKind::Dynamic) {
                        r#" cm="1""#
                    } else {
                        ""
                    };
                    row_data_str.push(format!(
                        r#"<c r="{cell_name}"{style} t="b"{cm}><f t="array" ref="{range}">{formula}</f><v>{b}</v></c>"#
                    ));
                }
                Cell::ArrayFormula {
                    f,
                    v: FormulaValue::Number(v),
                    s,
                    r,
                    kind,
                } => {
                    // Dynamic: <c r="D1" s="3" cm="1">
                    //            <f t="array" ref="D1:D5">A1:A5</f>
                    //            <v>123</v>
                    //          </c>
                    // CSE:     <c r="D1" s="3">
                    //            <f t="array" ref="D1:D5">A1:A5</f>
                    //            <v>123</v>
                    //          </c>
                    let node = match parsed_formulas.get(*f as usize) {
                        Some(node) => node,
                        None => continue,
                    };
                    let formula = get_array_formula_attribute(
                        worksheet.get_name(),
                        row_index,
                        column_index,
                        &node.0,
                    );
                    let style = get_cell_style_attribute(*s);
                    let range = match get_range_str(row_index, column_index, r.0, r.1) {
                        Some(range) => range,
                        None => continue,
                    };
                    let cm = if matches!(kind, ArrayKind::Dynamic) {
                        r#" cm="1""#
                    } else {
                        ""
                    };
                    row_data_str.push(format!(
                        r#"<c r="{cell_name}"{style}{cm}><f t="array" ref="{range}">{formula}</f><v>{v}</v></c>"#
                    ));
                }
                Cell::ArrayFormula {
                    f,
                    v: FormulaValue::Text(v),
                    s,
                    r,
                    kind,
                } => {
                    // Dynamic: <c r="D1" s="3" t="str" cm="1">
                    //            <f t="array" ref="D1:D5">A1:A5</f>
                    //            <v>Hello world!</v>
                    //          </c>
                    // CSE:     <c r="D1" s="3" t="str">
                    //            <f t="array" ref="D1:D5">A1:A5</f>
                    //            <v>Hello world!</v>
                    //          </c>
                    let node = match parsed_formulas.get(*f as usize) {
                        Some(node) => node,
                        None => continue,
                    };
                    let formula = get_array_formula_attribute(
                        worksheet.get_name(),
                        row_index,
                        column_index,
                        &node.0,
                    );
                    let style = get_cell_style_attribute(*s);
                    let range = match get_range_str(row_index, column_index, r.0, r.1) {
                        Some(range) => range,
                        None => continue,
                    };
                    let escaped_v = escape_xml(v);
                    let cm = if matches!(kind, ArrayKind::Dynamic) {
                        r#" cm="1""#
                    } else {
                        ""
                    };
                    row_data_str.push(format!(
                        r#"<c r="{cell_name}"{style} t="str"{cm}><f t="array" ref="{range}">{formula}</f><v>{escaped_v}</v></c>"#
                    ));
                }
                Cell::ArrayFormula {
                    f,
                    v: FormulaValue::Error { ei, .. },
                    s,
                    r,
                    kind,
                } => {
                    // Dynamic: <c r="D1" s="3" t="e" cm="1">
                    //            <f t="array" ref="D1:D5">A1:A5</f>
                    //            <v>#N/A!</v>
                    //          </c>
                    // CSE:     <c r="D1" s="3" t="e">
                    //            <f t="array" ref="D1:D5">A1:A5</f>
                    //            <v>#N/A!</v>
                    //          </c>
                    let node = match parsed_formulas.get(*f as usize) {
                        Some(node) => node,
                        None => continue,
                    };
                    let formula = get_array_formula_attribute(
                        worksheet.get_name(),
                        row_index,
                        column_index,
                        &node.0,
                    );
                    let style = get_cell_style_attribute(*s);
                    let range = match get_range_str(row_index, column_index, r.0, r.1) {
                        Some(range) => range,
                        None => continue,
                    };
                    let cm = if matches!(kind, ArrayKind::Dynamic) {
                        r#" cm="1""#
                    } else {
                        ""
                    };
                    row_data_str.push(format!(
                        r#"<c r="{cell_name}"{style} t="e"{cm}><f t="array" ref="{range}">{formula}</f><v>{ei}</v></c>"#
                    ));
                }
            }
        }
        let row_style_str = match row_style_dict.get(&row_index) {
            Some(row_style) => {
                let hidden_str = if row_style.hidden {
                    r#" hidden="1""#
                } else {
                    ""
                };
                // ENS-Office: grouping (outline) of rows.
                let outline_str = if row_style.outline_level > 0 {
                    format!(r#" outlineLevel="{}""#, row_style.outline_level)
                } else {
                    String::new()
                };
                let collapsed_str = if row_style.collapsed { r#" collapsed="1""# } else { "" };
                format!(
                    r#" s="{}" ht="{}" customHeight="{}" customFormat="{}"{}{}{}"#,
                    row_style.s,
                    row_style.height,
                    i32::from(row_style.custom_height),
                    i32::from(row_style.custom_format),
                    hidden_str,
                    outline_str,
                    collapsed_str,
                )
            }
            None => "".to_string(),
        };
        sheet_data_str.push(format!(
            "<row r=\"{row_index}\"{row_style_str}>{}</row>",
            row_data_str.join("")
        ))
    }
    let sheet_data = sheet_data_str.join("");

    for merged_cell in &worksheet.merged_cells {
        let first_column = number_to_column(merged_cell.column).unwrap_or("A".to_string());
        let last_column = number_to_column(merged_cell.last_column()).unwrap_or("A".to_string());
        merged_cells_str.push(format!(
            "<mergeCell ref=\"{first_column}{}:{last_column}{}\"/>",
            merged_cell.row,
            merged_cell.last_row()
        ))
    }
    let merged_cells_count = merged_cells_str.len();

    let cols = cols_str.join("");
    let cols = if cols.is_empty() {
        "".to_string()
    } else {
        format!("<cols>{cols}</cols>")
    };

    let tab_selected = if is_sheet_selected {
        " tabSelected=\"1\""
    } else {
        ""
    };

    let show_grid_lines = if !worksheet.show_grid_lines {
        " showGridLines=\"0\""
    } else {
        ""
    };

    let mut active_cell = "A1".to_string();
    let mut sqref = "A1".to_string();

    let views = &worksheet.views;
    if let Some(view) = views.get(&0) {
        let range = view.range;
        let row = view.row;
        let column = view.column;
        let column_name = number_to_column(column).unwrap_or("A".to_string());
        active_cell = format!("{column_name}{row}");

        let column_start = number_to_column(range[1]).unwrap_or("A".to_string());
        let column_end = number_to_column(range[3]).unwrap_or("A".to_string());
        if range[0] == range[2] && range[1] == range[3] {
            sqref = format!("{column_start}{}", range[0]);
        } else {
            sqref = format!("{}{}:{}{}", column_start, range[0], column_end, range[2]);
        }
    }

    let merge_cells_section = if merged_cells_count > 0 {
        format!(
            "<mergeCells count=\"{}\">{}</mergeCells>",
            merged_cells_count,
            merged_cells_str.join("")
        )
    } else {
        "".to_string()
    };

    let frozen_rows = worksheet.frozen_rows;
    let frozen_columns = worksheet.frozen_columns;

    let pane = if frozen_rows > 0 && frozen_columns > 0 {
        // There are both frozen rows and columns. There are four panes.
        // The first column is the first column after the last frozen column.
        let first_column = number_to_column(frozen_columns + 1).unwrap_or("A".to_string());
        // This is the top left cell of the bottom right pane.
        let top_left_cell = format!("{}{}", first_column, frozen_rows + 1);
        // The meaning of the next two is irrelevant for IronCalc.
        let top_right_active_cell = format!("{first_column}1");
        let bottom_left_active_cell = format!("A{}", frozen_rows + 1);
        // The bottom right active cell is the "true" selected cell and it does not need to reside on this pane.
        format!(
            "<pane xSplit=\"{frozen_columns}\" ySplit=\"{frozen_rows}\" topLeftCell=\"{top_left_cell}\" activePane=\"bottomRight\" state=\"frozen\"/>\
             <selection pane=\"topRight\" activeCell=\"{top_right_active_cell}\" sqref=\"{top_right_active_cell}\"/>\
             <selection pane=\"bottomLeft\" activeCell=\"{bottom_left_active_cell}\" sqref=\"{bottom_left_active_cell}\"/>\
             <selection pane=\"bottomRight\" activeCell=\"{active_cell}\" sqref=\"{sqref}\"/>",
        )
    } else if frozen_rows > 0 {
        // Only frozen rows
        let top_left_cell = format!("A{}", frozen_rows + 1);
        format!(
            "<pane ySplit=\"{frozen_rows}\" topLeftCell=\"{top_left_cell}\" activePane=\"bottomLeft\" state=\"frozen\"/>\
            <selection pane=\"bottomLeft\" activeCell=\"{active_cell}\" sqref=\"{sqref}\"/>",
        )
    } else if frozen_columns > 0 {
        let top_left_cell = format!(
            "{}1",
            number_to_column(frozen_columns + 1).unwrap_or("A".to_string())
        );
        format!(
            "<pane xSplit=\"{frozen_columns}\" topLeftCell=\"{top_left_cell}\" activePane=\"topRight\" state=\"frozen\"/>\
             <selection pane=\"topRight\" activeCell=\"{active_cell}\" sqref=\"{sqref}\"/>"
        )
    } else {
        // No frozen rows or columns
        format!(r#"<selection activeCell="{active_cell}" sqref="{sqref}"/>"#)
    };

    let (cf_sections, cf_ext_lst) =
        get_conditional_formatting_xml(&worksheet.conditional_formatting);

    let hyperlinks_section = get_hyperlinks_section(worksheet);
    // ENS-Office: the deepest outline levels, so Excel draws the group bar.
    let level_row = worksheet.rows.iter().map(|r| r.outline_level).max().unwrap_or(0);
    let level_col = worksheet.cols.iter().map(|c| c.outline_level).max().unwrap_or(0);
    let sheet_format = if level_row > 0 || level_col > 0 {
        format!(r#"<sheetFormatPr defaultRowHeight="15" outlineLevelRow="{level_row}" outlineLevelCol="{level_col}"/>"#)
    } else {
        String::new()
    };
    let sheet_protection = if worksheet.ens_protected {
        format!("<sheetProtection sheet=\"1\" objects=\"1\" scenarios=\"1\"{}/>", worksheet.ens_protect_attrs)
    } else {
        String::new()
    };
    let page_breaks = page_breaks_xml(worksheet);
    let drawing = if worksheet.ens_images.is_empty() && worksheet.ens_charts.is_empty() && worksheet.ens_shapes.is_empty() {
        String::new()
    } else {
        let n = external_link_count(worksheet) + if worksheet.comments.is_empty() { 0 } else { 2 } + 1;
        format!("<drawing r:id=\"rId{n}\"/>")
    };
    let legacy_drawing = if worksheet.comments.is_empty() {
        String::new()
    } else {
        format!("<legacyDrawing r:id=\"rId{}\"/>", external_link_count(worksheet) + 2)
    };
    let auto_filter_section = auto_filter;
    let (sheet_pr, page_setup) = page_setup_xml(&worksheet.ens_page_setup);
    // Excel's default is below / right; anything else is written to `<outlinePr>`.
    let sheet_pr = if worksheet.ens_summary == (true, true) {
        sheet_pr
    } else {
        let outline = format!(
            r#"<outlinePr summaryBelow="{}" summaryRight="{}"/>"#,
            i32::from(worksheet.ens_summary.0),
            i32::from(worksheet.ens_summary.1)
        );
        match sheet_pr.strip_prefix("<sheetPr>") {
            Some(rest) => format!("<sheetPr>{outline}{rest}"),
            None => format!("<sheetPr>{outline}</sheetPr>"),
        }
    };

    format!(
        "{XML_DECLARATION}\
<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\
  {sheet_pr}\
  <dimension ref=\"{dimension}\"/>\
  <sheetViews>\
    <sheetView workbookViewId=\"0\"{show_grid_lines}{tab_selected}>\
        {pane}\
    </sheetView>\
  </sheetViews>\
  {sheet_format}\
  {cols}\
  <sheetData>\
  {sheet_data}\
  </sheetData>\
  {sheet_protection}\
  {auto_filter_section}\
  {merge_cells_section}\
  {cf_sections}\
  {hyperlinks_section}\
  {page_setup}\
  {page_breaks}\
  {drawing}\
  {legacy_drawing}\
  {cf_ext_lst}\
</worksheet>"
    )
}


/// ENS-Office: the sheet's page setup (JSON, see `Worksheet::ens_page_setup`) as Excel's
/// `<sheetPr>` (fit to page) and `<printOptions>`, `<pageMargins>`, `<pageSetup>`.
fn page_setup_xml(json: &str) -> (String, String) {
    let Ok(v) = serde_json::from_str::<serde_json::Value>(json) else {
        return (String::new(), String::new());
    };
    let flag = |k: &str| v.get(k).and_then(|x| x.as_bool()).unwrap_or(false);
    let fit = v.get("fit").and_then(|x| x.as_str()).unwrap_or("none");
    let sheet_pr = if fit == "none" {
        String::new()
    } else {
        r#"<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>"#.to_string()
    };
    let mut out = String::new();
    if flag("gridlines") || flag("centerH") {
        out.push_str("<printOptions");
        if flag("centerH") {
            out.push_str(r#" horizontalCentered="1""#);
        }
        if flag("gridlines") {
            out.push_str(r#" gridLines="1""#);
        }
        out.push_str("/>");
    }
    let m: Vec<f64> = v
        .get("marginMm")
        .and_then(|x| x.as_array())
        .map(|a| a.iter().filter_map(|x| x.as_f64()).collect())
        .unwrap_or_default();
    let (top, right, bottom, left) = if m.len() == 4 { (m[0], m[1], m[2], m[3]) } else { (19.05, 17.78, 19.05, 17.78) };
    let inch = |mm: f64| (mm / 25.4 * 1000.0).round() / 1000.0;
    out.push_str(&format!(
        r#"<pageMargins left="{}" right="{}" top="{}" bottom="{}" header="0.3" footer="0.3"/>"#,
        inch(left),
        inch(right),
        inch(top),
        inch(bottom)
    ));
    let paper = match v.get("paper").and_then(|x| x.as_str()) {
        Some("Letter") => 1,
        Some("A3") => 8,
        Some("A5") => 11,
        _ => 9,
    };
    let orientation = if flag("landscape") { "landscape" } else { "portrait" };
    let fit_attrs = match fit {
        "width" => r#" fitToWidth="1" fitToHeight="0""#,
        "page" => r#" fitToWidth="1" fitToHeight="1""#,
        _ => "",
    };
    out.push_str(&format!(r#"<pageSetup paperSize="{paper}" orientation="{orientation}"{fit_attrs}/>"#));
    let text = |k: &str| v.get(k).and_then(|x| x.as_str()).unwrap_or("").to_string();
    let (header, footer) = (text("header"), text("footer"));
    if !header.is_empty() || !footer.is_empty() {
        out.push_str("<headerFooter>");
        if !header.is_empty() {
            out.push_str(&format!("<oddHeader>{}</oddHeader>", super::escape::escape_xml(&header)));
        }
        if !footer.is_empty() {
            out.push_str(&format!("<oddFooter>{}</oddFooter>", super::escape::escape_xml(&footer)));
        }
        out.push_str("</headerFooter>");
    }
    (sheet_pr, out)
}
