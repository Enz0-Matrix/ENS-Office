// SPDX-License-Identifier: AGPL-3.0-or-later
//! ENS-Office: «умные» таблицы (ListObject) в xlsx: `xl/tables/tableN.xml`,
//! связь из листа и `<tableParts>`. Раньше таблицы читались, но при сохранении пропадали.

use ironcalc_base::types::{Table, Workbook};
use ironcalc_base::Model;

use super::escape::escape_xml;

/// Таблицы листа в устойчивом порядке (по имени).
pub(crate) fn sheet_tables<'a>(workbook: &'a Workbook, sheet_name: &str) -> Vec<&'a Table> {
    let mut v: Vec<&Table> = workbook
        .tables
        .values()
        .filter(|t| t.sheet_name == sheet_name && is_valid(t))
        .collect();
    v.sort_by(|a, b| a.name.cmp(&b.name));
    v
}

/// Таблицу можно записать: диапазон читается, ширина совпадает с числом столбцов.
/// Повреждённую не пишем (Excel «восстанавливал» бы книгу); о ней предупреждает интерфейс.
pub fn is_valid(t: &Table) -> bool {
    match bounds(&t.reference) {
        Some((r1, c1, r2, c2)) => {
            !t.columns.is_empty()
                && r2 >= r1
                && c2 >= c1
                && (c2 - c1 + 1) as usize == t.columns.len()
        }
        None => false,
    }
}

fn bounds(reference: &str) -> Option<(i32, i32, i32, i32)> {
    let mut parts = reference.split(':');
    let a = ironcalc_base::expressions::utils::parse_reference_a1(parts.next()?)?;
    let b = match parts.next() {
        Some(p) => ironcalc_base::expressions::utils::parse_reference_a1(p)?,
        None => a.clone(),
    };
    Some((a.row, a.column, b.row, b.column))
}

/// Имена столбцов: Excel требует, чтобы они совпадали с текстом ячеек заголовка
/// и не повторялись. Правило одно с движком (он по нему же переписывает формулы).
fn column_names(model: &Model, sheet: u32, table: &Table) -> Vec<String> {
    model.ens_header_names(sheet, table)
}

fn opt_attr(name: &str, v: Option<u32>) -> String {
    v.map(|x| format!(" {name}=\"{x}\"")).unwrap_or_default()
}

/// Заменяет (или добавляет) атрибут в открывающем теге, который начинается с `tag_start`.
fn set_attr(xml: &mut String, tag_start: usize, attr: &str, value: &str) {
    let Some(rel_end) = xml[tag_start..].find('>') else {
        return;
    };
    let tag_end = tag_start + rel_end;
    let needle = format!(" {attr}=\"");
    if let Some(p) = xml[tag_start..tag_end].find(&needle) {
        let v_start = tag_start + p + needle.len();
        if let Some(q) = xml[v_start..].find('"') {
            xml.replace_range(v_start..v_start + q, value);
        }
    } else {
        let insert_at = if xml[..tag_end].ends_with('/') {
            tag_end - 1
        } else {
            tag_end
        };
        xml.insert_str(insert_at, &format!(" {attr}=\"{value}\""));
    }
}

/// Начало открывающего тега `<name` (а не `<nameXxx`) начиная с `from`.
fn find_tag(xml: &str, name: &str, from: usize) -> Option<usize> {
    let pat = format!("<{name}");
    let mut at = from;
    while let Some(p) = xml[at..].find(&pat) {
        let start = at + p;
        let next = xml[start + pat.len()..].chars().next();
        if matches!(next, Some(' ' | '>' | '/' | '\n' | '\r' | '\t')) {
            return Some(start);
        }
        at = start + pat.len();
    }
    None
}

/// Исходный XML таблицы с новым номером, диапазоном и именами столбцов.
fn patch_raw_xml(
    table: &Table,
    id: usize,
    names: &[String],
    auto_filter_ref: &str,
) -> Option<String> {
    let mut xml = table.ens_raw_xml.clone();
    let root = find_tag(&xml, "table", 0)?;
    set_attr(&mut xml, root, "id", &id.to_string());
    set_attr(&mut xml, root, "ref", &escape_xml(&table.reference));
    if let Some(af) = find_tag(&xml, "autoFilter", root + 1) {
        let old_ref = {
            let end = af + xml[af..].find('>')?;
            let tag = &xml[af..end];
            tag.find(" ref=\"").and_then(|p| {
                tag[p + 6..]
                    .find('"')
                    .map(|q| tag[p + 6..p + 6 + q].to_string())
            })
        };
        if old_ref.as_deref() != Some(auto_filter_ref) {
            set_attr(&mut xml, af, "ref", auto_filter_ref);
            // Сортировка хранит свой диапазон: после сдвига она устарела, Excel её пересчитает.
            if let Some(ss) = find_tag(&xml, "sortState", af) {
                let self_closing = xml[ss..]
                    .find('>')
                    .map(|e| xml[ss..ss + e].ends_with('/'))?;
                let end = if self_closing {
                    ss + xml[ss..].find('>')? + 1
                } else {
                    ss + xml[ss..].find("</sortState>")? + "</sortState>".len()
                };
                xml.replace_range(ss..end, "");
            }
        }
    }
    let mut at = root + 1;
    for name in names {
        let tc = find_tag(&xml, "tableColumn", at)?;
        set_attr(&mut xml, tc, "name", &escape_xml(name));
        at = tc + 1;
    }
    if find_tag(&xml, "tableColumn", at).is_some() {
        return None; // столбцов в XML больше, чем в модели
    }
    Some(xml)
}

/// `None` — таблица повреждена (нет диапазона или столбцов): не пишем её, чтобы Excel
/// не «восстанавливал» книгу.
pub(crate) fn get_table_xml(model: &Model, sheet: u32, table: &Table, id: usize) -> Option<String> {
    let (r1, c1, r2, c2) = bounds(&table.reference)?;
    if table.columns.is_empty() || (c2 - c1 + 1) as usize != table.columns.len() || r2 < r1 {
        return None;
    }
    let names = column_names(model, sheet, table);
    let auto_filter_ref = {
        let a = ironcalc_base::expressions::utils::number_to_column(c1)?;
        let b = ironcalc_base::expressions::utils::number_to_column(c2)?;
        format!(
            "{a}{r1}:{b}{}",
            (r2 - table.totals_row_count as i32).max(r1)
        )
    };
    if !table.ens_raw_xml.is_empty() {
        if let Some(xml) = patch_raw_xml(table, id, &names, &auto_filter_ref) {
            return Some(xml);
        }
    }
    let mut cols = String::new();
    for (i, col) in table.columns.iter().enumerate() {
        let mut attrs = format!(
            " id=\"{}\" name=\"{}\"",
            col.id.max(1),
            escape_xml(&names[i])
        );
        if let Some(f) = col
            .totals_row_function
            .as_deref()
            .filter(|f| *f != "custom" && *f != "none")
        {
            attrs.push_str(&format!(" totalsRowFunction=\"{}\"", escape_xml(f)));
        }
        if let Some(l) = &col.totals_row_label {
            attrs.push_str(&format!(" totalsRowLabel=\"{}\"", escape_xml(l)));
        }
        attrs.push_str(&opt_attr("headerRowDxfId", col.header_row_dxf_id));
        attrs.push_str(&opt_attr("dataDxfId", col.data_dxf_id));
        attrs.push_str(&opt_attr("totalsRowDxfId", col.totals_row_dxf_id));
        cols.push_str(&format!("<tableColumn{attrs}/>"));
    }
    let mut head = format!(
        "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"{id}\" name=\"{}\" displayName=\"{}\" ref=\"{}\"",
        escape_xml(&table.name),
        escape_xml(&table.display_name),
        escape_xml(&table.reference)
    );
    if table.header_row_count == 0 {
        head.push_str(" headerRowCount=\"0\"");
    }
    if table.totals_row_count > 0 {
        head.push_str(&format!(" totalsRowCount=\"{}\"", table.totals_row_count));
    } else {
        head.push_str(" totalsRowShown=\"0\"");
    }
    head.push_str(&opt_attr("headerRowDxfId", table.header_row_dxf_id));
    head.push_str(&opt_attr("dataDxfId", table.data_dxf_id));
    head.push_str(&opt_attr("totalsRowDxfId", table.totals_row_dxf_id));
    head.push('>');
    // Кнопки фильтра в заголовке, как у новой таблицы Excel.
    let auto_filter = if table.header_row_count > 0 {
        format!("<autoFilter ref=\"{auto_filter_ref}\"/>")
    } else {
        String::new()
    };
    let s = &table.style_info;
    let style_name = s
        .name
        .as_ref()
        .map(|n| format!(" name=\"{}\"", escape_xml(n)))
        .unwrap_or_default();
    let style = format!(
        "<tableStyleInfo{style_name} showFirstColumn=\"{}\" showLastColumn=\"{}\" showRowStripes=\"{}\" showColumnStripes=\"{}\"/>",
        u8::from(s.show_first_column),
        u8::from(s.show_last_column),
        u8::from(s.show_row_stripes),
        u8::from(s.show_column_stripes)
    );
    Some(format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{head}{auto_filter}<tableColumns count=\"{}\">{cols}</tableColumns>{style}</table>",
        table.columns.len()
    ))
}

/// `<tableParts>` листа: связи `rId{first_rel}`, `rId{first_rel+1}`, …
pub(crate) fn table_parts_xml(count: usize, first_rel: usize) -> String {
    if count == 0 {
        return String::new();
    }
    let parts: String = (0..count)
        .map(|i| format!("<tablePart r:id=\"rId{}\"/>", first_rel + i))
        .collect();
    format!("<tableParts count=\"{count}\">{parts}</tableParts>")
}

/// Вставляет `<tableParts>` в XML листа: после рисунков, перед `<extLst>` (порядок схемы).
pub(crate) fn insert_table_parts(sheet_xml: String, parts: &str) -> String {
    if parts.is_empty() {
        return sheet_xml;
    }
    let after_data = sheet_xml.find("</sheetData>").unwrap_or(0);
    let pos = sheet_xml[after_data..]
        .find("<extLst")
        .map(|p| after_data + p)
        .or_else(|| sheet_xml.rfind("</worksheet>"))
        .unwrap_or(sheet_xml.len());
    let mut out = sheet_xml;
    out.insert_str(pos, parts);
    out
}
