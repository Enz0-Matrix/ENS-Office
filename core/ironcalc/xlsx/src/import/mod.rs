mod conditional_formatting;
mod metadata;
pub(crate) mod shared_strings;
mod drawings;
mod external_links;
mod sheet_data;
mod styles;
mod tables;
mod theme;
mod util;
mod workbook;
mod worksheets;
mod xml;

use ironcalc_base::expressions::token::get_error_by_english_name;
use ironcalc_base::external::ExternalValue;

use std::{
    collections::HashMap,
    fs,
    io::{BufReader, Cursor, Read},
};

use xml::XmlNode;

use ironcalc_base::{
    expressions::{
        parser::{new_parser_english, stringify::to_english_string},
        types::CellReferenceRC,
    },
    types::{ExternalBook, Metadata, RawPart, Workbook, WorkbookSettings, WorkbookView},
    Model,
};

use crate::error::XlsxError;

use shared_strings::{read_rich_runs, read_shared_strings, SharedStringTable};

use metadata::load_metadata;
use styles::load_styles;
use util::get_attribute;
use workbook::load_workbook;
use worksheets::{load_sheets, Relationship};

fn load_relationships<R: Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
) -> Result<HashMap<String, Relationship>, XlsxError> {
    let file = archive.by_name("xl/_rels/workbook.xml.rels")?;
    let doc = XmlNode::parse(BufReader::new(file))?;
    let nodes: Vec<&XmlNode> = doc
        .descendants()
        .filter(|n| n.has_tag_name("Relationship"))
        .collect();
    let mut rels = HashMap::new();
    for node in nodes {
        rels.insert(
            get_attribute(node, "Id")?.to_string(),
            Relationship {
                rel_type: get_attribute(node, "Type")?.to_string(),
                target: get_attribute(node, "Target")?.to_string(),
            },
        );
    }
    Ok(rels)
}

fn resolve_theme_path(rels: &HashMap<String, Relationship>) -> Option<String> {
    let target = rels
        .values()
        .find(|r| r.rel_type.ends_with("/theme"))?
        .target
        .clone();
    Some(if let Some(absolute) = target.strip_prefix('/') {
        absolute.to_string()
    } else {
        format!("xl/{target}")
    })
}

// FIXME: This is a bit of a HACK. We basically re-parse all the defined names assuming the context is `A1`
// and there is no tables or defined names.
fn reparse_formula_hack(formula: &str, worksheets: &[String]) -> Result<String, XlsxError> {
    let defined_names = Vec::new();
    let tables = HashMap::new();
    let mut parser = new_parser_english(worksheets.to_owned(), defined_names, tables);
    let cell_reference = CellReferenceRC {
        sheet: worksheets[0].clone(),
        column: 1,
        row: 1,
    };
    let t = parser.parse(formula, &cell_reference);

    Ok(to_english_string(&t, &cell_reference))
}

fn load_xlsx_from_reader<R: Read + std::io::Seek>(
    name: String,
    reader: R,
    locale: &str,
    tz: &str,
) -> Result<Workbook, XlsxError> {
    let mut archive = zip::ZipArchive::new(reader)?;

    let strings = read_shared_strings(&mut archive)?;
    let ens_rich = read_rich_runs(&mut archive);
    let mut shared_strings = SharedStringTable::new_excluding(strings, &ens_rich);
    let mut workbook = load_workbook(&mut archive)?;
    let rels = load_relationships(&mut archive)?;
    let theme_path = resolve_theme_path(&rels);
    let theme = theme::load(&mut archive, theme_path.as_deref());
    let mut tables = HashMap::new();
    // Styles must be loaded before the worksheets: conditional-formatting rules
    // stored in x14 `extLst` extensions carry inline `<x14:dxf>` formats that we
    // append to `styles.dxfs`, referencing them back by index from the rule.
    let mut styles = load_styles(&mut archive, &theme)?;
    // ENS-Office: книги-источники внешних ссылок и их листы (служебные, после настоящих).
    let external = external_links::load_external_links(&mut archive, &rels);
    let book_files: Vec<String> = external.iter().map(|b| b.file.clone()).collect();
    let external_sheet_names: Vec<String> = external
        .iter()
        .filter(|b| !b.file.is_empty())
        .flat_map(|b| b.sheets.iter().map(|(name, _)| ironcalc_base::external::external_sheet_name(&b.file, name)))
        .collect();
    let (mut worksheets, selected_sheet) = load_sheets(
        &mut archive,
        &rels,
        &workbook,
        &mut tables,
        &mut shared_strings,
        &theme,
        &mut styles.dxfs,
        &book_files,
        &external_sheet_names,
    )?;
    let mut next_sheet_id = worksheets.iter().map(|w| w.sheet_id).max().unwrap_or(0) + 1;
    for book in external.iter().filter(|b| !b.file.is_empty()) {
        for (name, values) in &book.sheets {
            let mut ws = ironcalc_base::external::new_external_worksheet(&book.file, name, next_sheet_id);
            next_sheet_id += 1;
            for (row, column, value) in values {
                let _ = match value {
                    ExternalValue::Number(n) => ws.set_cell_with_number(*row, *column, *n, 0),
                    ExternalValue::Bool(b) => ws.set_cell_with_boolean(*row, *column, *b, 0),
                    ExternalValue::Text(t) => ws.set_cell_with_string(*row, *column, shared_strings.index_of(t), 0),
                    ExternalValue::Error(e) => ws.set_cell_with_error(
                        *row,
                        *column,
                        get_error_by_english_name(e).unwrap_or(ironcalc_base::expressions::token::Error::NA),
                        0,
                    ),
                };
            }
            worksheets.push(ws);
        }
    }
    // reparse formulas in defined names, since they may refer to sheets and tables that have been loaded
    let worksheet_names = worksheets
        .iter()
        .map(|s| s.name.clone())
        .collect::<Vec<_>>();
    for dn in &mut workbook.defined_names {
        // ENS-Office: Excel's own names (`_xlnm.Print_Titles` is a list of ranges) stay as written.
        if dn.name.starts_with("_xlnm.") {
            continue;
        }
        let formula = ironcalc_base::external::to_internal_refs(&dn.formula, &book_files);
        dn.formula = reparse_formula_hack(&formula, &worksheet_names)?;
    }
    let metadata = match load_metadata(&mut archive) {
        Ok(metadata) => metadata,
        Err(_) => {
            // In case there is no metadata, add some
            Metadata {
                application: "Unknown application".to_string(),
                app_version: "".to_string(),
                creator: "".to_string(),
                last_modified_by: "".to_string(),
                created: "".to_string(),
                last_modified: "".to_string(),
            }
        }
    };
    let ens_change_log = load_change_log(&mut archive);
    let ens_vba_parts = load_vba_parts(&mut archive);
    let part_names: Vec<String> = archive.file_names().map(|n| n.to_string()).collect();
    let mut ens_lost_features = lost_features_of(&part_names);
    // ENS-Office: связи с другими книгами теперь переносятся, если все они прочитаны.
    if !external.is_empty() && external.iter().all(|b| !b.file.is_empty()) {
        ens_lost_features.retain(|f| !f.starts_with("связи с другими книгами"));
    }
    let ens_external_books: Vec<ExternalBook> = external
        .iter()
        .map(|b| ExternalBook { target: b.target.clone(), file: b.file.clone() })
        .collect();
    let ens_code_name = load_workbook_code_name(&mut archive);
    // ENS-Office: система дат 1904 (Excel для Mac) → обычная система 1900.
    if load_date1904(&mut archive) {
        convert_date1904(&mut worksheets, &styles);
        ens_lost_features.push(
            "система дат 1904 (Excel для Mac): даты пересчитаны в обычную систему, \
             проверьте числа без формата даты, которые формулы используют как даты"
                .to_string(),
        );
    }
    let (ens_structure_locked, ens_structure_attrs) = load_structure_locked(&mut archive);
    let mut views = HashMap::new();
    views.insert(
        0,
        WorkbookView {
            sheet: selected_sheet,
            window_width: 800,
            window_height: 600,
        },
    );
    Ok(Workbook {
        shared_strings: shared_strings.into_strings(),
        defined_names: workbook.defined_names,
        worksheets,
        styles,
        name,
        settings: WorkbookSettings {
            tz: tz.to_string(),
            locale: locale.to_string(),
        },
        metadata,
        tables,
        views,
        theme,
        ens_change_log,
        ens_structure_locked,
        ens_structure_attrs,
        ens_rich,
        ens_vba_parts,
        ens_lost_features,
        ens_code_name,
        ens_external_books,
    })
}

/// ENS-Office: `<workbookPr codeName>`.
fn load_workbook_code_name<R: Read + std::io::Seek>(archive: &mut zip::read::ZipArchive<R>) -> String {
    let mut text = String::new();
    let read = archive.by_name("xl/workbook.xml").map(|mut f| f.read_to_string(&mut text));
    if !matches!(read, Ok(Ok(_))) {
        return String::new();
    }
    let Ok(root) = XmlNode::parse(text.as_bytes()) else { return String::new() };
    let name = root
        .children()
        .find(|n| n.has_tag_name("workbookPr"))
        .and_then(|n| n.attribute("codeName"))
        .unwrap_or_default()
        .to_string();
    name
}

/// ENS-Office: макросы (VBA) и всё, что к ним относится, переносятся байт в байт.
fn load_vba_parts<R: Read + std::io::Seek>(archive: &mut zip::read::ZipArchive<R>) -> Vec<RawPart> {
    let names: Vec<String> = archive
        .file_names()
        .filter(|n| {
            let l = n.to_ascii_lowercase();
            l.starts_with("xl/vbaproject") || l.starts_with("xl/_rels/vbaproject")
        })
        .map(|n| n.to_string())
        .collect();
    let mut parts = Vec::new();
    for path in names {
        let mut data = Vec::new();
        if let Ok(mut f) = archive.by_name(&path) {
            if f.read_to_end(&mut data).is_ok() {
                parts.push(RawPart { path, data });
            }
        }
    }
    parts
}

/// ENS-Office: части пакета, которые ENS-Table пока не сохраняет, — по-русски, без повторов.
/// Безвредное (calcChain, настройки принтера, оформление диаграмм) не перечисляется.
pub fn lost_features_of(names: &[String]) -> Vec<String> {
    const KNOWN: &[(&str, &str)] = &[
        ("xl/pivottables/", "сводные таблицы"),
        ("xl/pivotcache/", "сводные таблицы"),
        ("xl/externallinks/", "связи с другими книгами (сохранённые значения из них)"),
        ("xl/slicers/", "срезы"),
        ("xl/slicercaches/", "срезы"),
        ("xl/timelines/", "временные шкалы"),
        ("xl/timelinecaches/", "временные шкалы"),
        ("xl/connections.xml", "подключения к внешним данным"),
        ("xl/querytables/", "запросы к внешним данным"),
        ("xl/ctrlprops/", "элементы управления (флажки, списки, кнопки)"),
        ("xl/activex/", "элементы ActiveX"),
        ("xl/embeddings/", "внедрённые объекты (OLE)"),
        ("xl/threadedcomments/", "цепочки комментариев (останутся как примечания)"),
        ("xl/richdata/", "изображения в ячейках"),
        ("xl/model/", "модель данных Power Pivot"),
        ("xl/webextensions/", "надстройки Office"),
        ("xl/dialogsheets/", "листы диалогов Excel 5"),
        ("xl/macrosheets/", "листы макросов Excel 4"),
        ("xl/chartsheets/", "листы диаграмм"),
        ("docprops/custom.xml", "дополнительные свойства документа"),
    ];
    // Порядок — как в KNOWN (по важности), а не как части лежат в архиве.
    let lower: Vec<String> = names.iter().map(|n| n.to_ascii_lowercase()).collect();
    let mut out: Vec<String> = Vec::new();
    for (prefix, what) in KNOWN {
        if lower.iter().any(|l| l.starts_with(prefix)) && !out.iter().any(|x| x == what) {
            out.push(what.to_string());
        }
    }
    out
}

/// ENS-Office: `<workbookProtection lockStructure>` of `xl/workbook.xml`: locked or not, and the password
/// hash as attributes named as in `<sheetProtection>`.
fn load_structure_locked<R: Read + std::io::Seek>(archive: &mut zip::read::ZipArchive<R>) -> (bool, String) {
    let mut text = String::new();
    let read = archive.by_name("xl/workbook.xml").map(|mut f| f.read_to_string(&mut text));
    if !matches!(read, Ok(Ok(_))) {
        return (false, String::new());
    }
    let Ok(root) = XmlNode::parse(text.as_bytes()) else { return (false, String::new()) };
    let Some(n) = root.children().find(|n| n.has_tag_name("workbookProtection")) else { return (false, String::new()) };
    let locked = n.attribute("lockStructure").is_some_and(|v| v == "1" || v == "true");
    let attrs = [
        ("workbookPassword", "password"),
        ("workbookAlgorithmName", "algorithmName"),
        ("workbookHashValue", "hashValue"),
        ("workbookSaltValue", "saltValue"),
        ("workbookSpinCount", "spinCount"),
    ]
    .iter()
    .filter_map(|(from, to)| n.attribute(from).map(|v| format!(" {to}=\"{}\"", v.replace('&', "&amp;").replace('"', "&quot;").replace('<', "&lt;"))))
    .collect::<String>();
    (locked, attrs)
}

/// ENS-Office: the change log from `customXml/item*.xml` (root `<ensChangeLog>`).
fn load_change_log<R: Read + std::io::Seek>(archive: &mut zip::read::ZipArchive<R>) -> String {
    let names: Vec<String> = archive
        .file_names()
        .filter(|n| n.starts_with("customXml/item") && !n.contains("Props") && n.ends_with(".xml"))
        .map(|n| n.to_string())
        .collect();
    for name in names {
        let mut text = String::new();
        if let Ok(mut f) = archive.by_name(&name) {
            if f.read_to_string(&mut text).is_err() {
                continue;
            }
        }
        let Ok(root) = XmlNode::parse(text.as_bytes()) else {
            continue;
        };
        if root.has_tag_name("ensChangeLog") {
            return root.text().unwrap_or_default().to_string();
        }
    }
    String::new()
}

// Imports a file from disk into an internal representation
fn load_from_excel(file_name: &str, locale: &str, tz: &str) -> Result<Workbook, XlsxError> {
    let file_path = std::path::Path::new(file_name);
    let file = fs::File::open(file_path)?;
    let reader = BufReader::new(file);
    let name = file_path
        .file_stem()
        .ok_or_else(|| XlsxError::IO("Could not extract workbook name".to_string()))?
        .to_string_lossy()
        .to_string();
    load_xlsx_from_reader(name, reader, locale, tz)
}

/// Loads a [Workbook] from the bytes of an xlsx file.
/// This is useful, for instance, when bytes are transferred over the network
pub fn load_from_xlsx_bytes(
    bytes: &[u8],
    name: &str,
    locale: &str,
    tz: &str,
) -> Result<Workbook, XlsxError> {
    let cursor = Cursor::new(bytes);
    let reader = BufReader::new(cursor);
    load_xlsx_from_reader(name.to_string(), reader, locale, tz)
}

/// Loads a [Model] from an xlsx file
pub fn load_from_xlsx<'a>(
    file_name: &str,
    locale: &str,
    tz: &str,
    language: &'a str,
) -> Result<Model<'a>, XlsxError> {
    let workbook = load_from_excel(file_name, locale, tz)?;
    Model::from_workbook(workbook, language).map_err(XlsxError::Workbook)
}

/// Loads a [Model] from an `ic` file (a file in the IronCalc internal representation)
pub fn load_from_icalc<'a>(file_name: &str, language_id: &'a str) -> Result<Model<'a>, XlsxError> {
    let contents = fs::read(file_name)
        .map_err(|e| XlsxError::IO(format!("Could not extract workbook name: {e}")))?;
    let workbook: Workbook = bitcode::decode(&contents)
        .map_err(|e| XlsxError::IO(format!("Failed to decode file: {e}")))?;
    Model::from_workbook(workbook, language_id).map_err(XlsxError::Workbook)
}

/// ENS-Office: `<workbookPr date1904>` — книга в системе дат 1904 (Excel для Mac).
fn load_date1904<R: Read + std::io::Seek>(archive: &mut zip::read::ZipArchive<R>) -> bool {
    let mut text = String::new();
    let read = archive.by_name("xl/workbook.xml").map(|mut f| f.read_to_string(&mut text));
    if !matches!(read, Ok(Ok(_))) {
        return false;
    }
    let Ok(root) = XmlNode::parse(text.as_bytes()) else { return false };
    let on = root
        .children()
        .find(|n| n.has_tag_name("workbookPr"))
        .and_then(|n| n.attribute("date1904"))
        .is_some_and(|v| v == "1" || v.eq_ignore_ascii_case("true"));
    on
}

/// ENS-Office: книга в системе дат 1904 переводится в систему 1900: числа с форматом даты
/// сдвигаются на 1462 дня, в ячейках остаются те же даты, функции дат работают как обычно.
/// Числа без формата даты не трогаются. Возвращает, сколько ячеек сдвинуто.
fn convert_date1904(worksheets: &mut [ironcalc_base::types::Worksheet], styles: &ironcalc_base::types::Styles) -> usize {
    use ironcalc_base::types::Cell;
    use ironcalc_base::formatter::dates::{shows_calendar_date, DATE1904_OFFSET};
    let is_date: Vec<bool> = styles
        .cell_xfs
        .iter()
        .map(|xf| shows_calendar_date(&ironcalc_base::number_format::get_num_fmt(xf.num_fmt_id, &styles.num_fmts)))
        .collect();
    let mut count = 0;
    for ws in worksheets.iter_mut() {
        if ironcalc_base::external::is_external_sheet_name(&ws.name) {
            continue;
        }
        let cells: Vec<(i32, i32)> = ws
            .sheet_data
            .cells()
            .filter(|(_, _, c)| matches!(c.as_ref(), Cell::NumberCell { s, .. } if is_date.get(*s as usize).copied().unwrap_or(false)))
            .map(|(r, c, _)| (r, c))
            .collect();
        for (r, c) in cells {
            if let Some(Cell::NumberCell { v, .. }) = ws.sheet_data.cell_mut(r, c) {
                *v += DATE1904_OFFSET;
                count += 1;
            }
        }
    }
    count
}
