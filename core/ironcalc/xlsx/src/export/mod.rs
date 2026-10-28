#![allow(clippy::unwrap_used)]

mod _rels;
mod conditional_formatting;
mod doc_props;
mod drawings;
mod dxfs_styles;
mod escape;
mod shared_strings;
mod styles;
mod styles_util;
mod theme;
mod workbook;
mod workbook_xml_rels;
mod worksheets;
mod xml_constants;

use std::io::BufWriter;
use std::{
    fs,
    io::{Seek, Write},
};

use ironcalc_base::expressions::utils::number_to_column;
use ironcalc_base::types::Workbook;
use ironcalc_base::{get_milliseconds_since_epoch, Model};

use self::xml_constants::XML_DECLARATION;

use crate::error::XlsxError;

#[cfg(test)]
mod test;

fn get_content_types_xml(workbook: &Workbook) -> String {
    // A list of all files in the zip
    let mut content = vec![
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#.to_string(),
        r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#.to_string(),
        r#"<Default Extension="xml" ContentType="application/xml"/>"#.to_string(),
        r#"<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>"#.to_string(),
    ];
    for worksheet in 0..workbook.worksheets.len() {
        let sheet = format!(
            r#"<Override PartName="/xl/worksheets/sheet{}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
            worksheet + 1
        );
        content.push(sheet);
    }
    if workbook.worksheets.iter().any(|w| !w.comments.is_empty()) {
        content.push(r#"<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>"#.to_string());
    }
    for ext in ["png", "jpeg", "gif"] {
        if workbook.worksheets.iter().any(|w| w.ens_images.iter().any(|i| i.ext == ext)) {
            content.push(format!(r#"<Default Extension="{ext}" ContentType="image/{ext}"/>"#));
        }
    }
    for (i, w) in workbook.worksheets.iter().enumerate() {
        if !w.ens_images.is_empty() {
            content.push(format!(
                r#"<Override PartName="/xl/drawings/drawing{}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>"#,
                i + 1
            ));
        }
        if !w.comments.is_empty() {
            content.push(format!(
                r#"<Override PartName="/xl/comments{}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>"#,
                i + 1
            ));
        }
    }
    content.push(
        r#"<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>"#.to_string(),
    );
    content.extend([
        r#"<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"#.to_string(),
        r#"<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>"#.to_string(),
        r#"<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>"#.to_string(),
        r#"<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>"#.to_string(),
        r#"<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>"#.to_string(),
    ]);
    if !workbook.ens_change_log.is_empty() {
        content.push(
            r#"<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>"#.to_string(),
        );
    }
    content.push(r#"</Types>"#.to_string());
    format!("{XML_DECLARATION}\n{}", content.join(""))
}

/// Exports a model to an xlsx file
pub fn save_to_xlsx(model: &Model, file_name: &str) -> Result<(), XlsxError> {
    let file_path = std::path::Path::new(&file_name);
    if file_path.exists() {
        return Err(XlsxError::IO(format!("file {file_name} already exists")));
    }
    let file = fs::File::create(file_path).unwrap();
    let writer = BufWriter::new(file);
    save_xlsx_to_writer(model, writer)?;

    Ok(())
}

pub fn save_xlsx_to_writer<W: Write + Seek>(model: &Model, writer: W) -> Result<W, XlsxError> {
    let workbook = &model.workbook;
    let selected_sheet = match workbook.views.get(&0) {
        Some(view) => view.sheet,
        _ => 0,
    };
    let mut zip = zip::ZipWriter::new(writer);

    let options = zip::write::FileOptions::default();

    // root folder
    zip.start_file("[Content_Types].xml", options)?;
    zip.write_all(get_content_types_xml(workbook).as_bytes())?;

    zip.add_directory("docProps", options)?;
    zip.start_file("docProps/app.xml", options)?;
    zip.write_all(doc_props::get_app_xml(workbook).as_bytes())?;
    zip.start_file("docProps/core.xml", options)?;
    let milliseconds = get_milliseconds_since_epoch();
    zip.write_all(doc_props::get_core_xml(workbook, milliseconds)?.as_bytes())?;

    // Package-relationship item
    zip.add_directory("_rels", options)?;
    zip.start_file("_rels/.rels", options)?;
    zip.write_all(_rels::get_dot_rels(workbook).as_bytes())?;

    zip.add_directory("xl", options)?;
    zip.start_file("xl/sharedStrings.xml", options)?;
    zip.write_all(shared_strings::get_shared_strings_xml(workbook).as_bytes())?;
    zip.start_file("xl/styles.xml", options)?;
    zip.write_all(styles::get_styles_xml(workbook).as_bytes())?;
    zip.start_file("xl/workbook.xml", options)?;
    zip.write_all(workbook::get_workbook_xml(workbook, selected_sheet).as_bytes())?;
    zip.start_file("xl/metadata.xml", options)?;
    let metadata_xml = include_str!("metadata.xml");
    zip.write_all(metadata_xml.as_bytes())?;

    // ENS-Office: the change log as a custom XML part (Excel keeps it when it saves the file).
    if !workbook.ens_change_log.is_empty() {
        zip.add_directory("customXml", options)?;
        zip.start_file("customXml/item1.xml", options)?;
        zip.write_all(
            format!(
                "{XML_DECLARATION}\n<ensChangeLog xmlns=\"urn:ens-office:change-log\">{}</ensChangeLog>",
                escape::escape_xml(&workbook.ens_change_log)
            )
            .as_bytes(),
        )?;
        zip.start_file("customXml/itemProps1.xml", options)?;
        zip.write_all(
            format!(
                "{XML_DECLARATION}\n<ds:datastoreItem ds:itemID=\"{{6E1F3C2A-5B7D-4E8A-9C0B-1D2E3F4A5B6C}}\" xmlns:ds=\"http://schemas.openxmlformats.org/officeDocument/2006/customXml\"><ds:schemaRefs/></ds:datastoreItem>"
            )
            .as_bytes(),
        )?;
        zip.add_directory("customXml/_rels", options)?;
        zip.start_file("customXml/_rels/item1.xml.rels", options)?;
        zip.write_all(
            format!(
                "{XML_DECLARATION}\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps\" Target=\"itemProps1.xml\"/></Relationships>"
            )
            .as_bytes(),
        )?;
    }

    zip.add_directory("xl/theme", options)?;
    zip.start_file("xl/theme/theme1.xml", options)?;
    zip.write_all(theme::get_theme_xml(&workbook.theme).as_bytes())?;

    zip.add_directory("xl/_rels", options)?;
    zip.start_file("xl/_rels/workbook.xml.rels", options)?;
    zip.write_all(workbook_xml_rels::get_workbook_xml_rels(workbook).as_bytes())?;

    zip.add_directory("xl/worksheets", options)?;
    // sheet rels parts (one per sheet with external hyperlinks)
    let mut has_worksheet_rels = false;
    for (sheet_index, worksheet) in workbook.worksheets.iter().enumerate() {
        let comments_id = (!worksheet.comments.is_empty()).then_some(sheet_index + 1);
        let drawing_id = (!worksheet.ens_images.is_empty()).then_some(sheet_index + 1);
        if let Some(rels_xml) = worksheets::get_worksheet_xml_rels(worksheet, comments_id, drawing_id) {
            if !has_worksheet_rels {
                zip.add_directory("xl/worksheets/_rels", options)?;
                has_worksheet_rels = true;
            }
            zip.start_file(
                format!("xl/worksheets/_rels/sheet{}.xml.rels", sheet_index + 1),
                options,
            )?;
            zip.write_all(rels_xml.as_bytes())?;
        }
    }
    // ENS-Office: comments (notes) and the VML boxes Excel shows them in.
    let mut has_drawings_dir = false;
    for (sheet_index, worksheet) in workbook.worksheets.iter().enumerate() {
        if worksheet.comments.is_empty() {
            continue;
        }
        let id = sheet_index + 1;
        zip.start_file(format!("xl/comments{id}.xml"), options)?;
        zip.write_all(worksheets::get_comments_xml(worksheet).as_bytes())?;
        if !has_drawings_dir {
            zip.add_directory("xl/drawings", options)?;
            has_drawings_dir = true;
        }
        zip.start_file(format!("xl/drawings/vmlDrawing{id}.vml"), options)?;
        zip.write_all(worksheets::get_comments_vml(worksheet).as_bytes())?;
    }
    // ENS-Office: pictures.
    let mut media_count = 1;
    let mut has_media_dir = false;
    for (sheet_index, worksheet) in workbook.worksheets.iter().enumerate() {
        let images = &worksheet.ens_images;
        if images.is_empty() {
            continue;
        }
        let id = sheet_index + 1;
        if !has_drawings_dir {
            zip.add_directory("xl/drawings", options)?;
            has_drawings_dir = true;
        }
        if !has_media_dir {
            zip.add_directory("xl/media", options)?;
            zip.add_directory("xl/drawings/_rels", options)?;
            has_media_dir = true;
        }
        zip.start_file(format!("xl/drawings/drawing{id}.xml"), options)?;
        zip.write_all(drawings::get_drawing_xml(images).as_bytes())?;
        zip.start_file(format!("xl/drawings/_rels/drawing{id}.xml.rels"), options)?;
        zip.write_all(drawings::get_drawing_rels(images, media_count).as_bytes())?;
        for (i, img) in images.iter().enumerate() {
            zip.start_file(format!("xl/media/{}", drawings::media_name(media_count, i, img)), options)?;
            zip.write_all(&img.data)?;
        }
        media_count += images.len();
    }
    for (sheet_index, worksheet) in workbook.worksheets.iter().enumerate() {
        let id = sheet_index + 1;
        zip.start_file(format!("xl/worksheets/sheet{id}.xml"), options)?;
        let dimension = model
            .workbook
            .worksheet(sheet_index as u32)
            .unwrap()
            .dimension();
        let column_min_str = number_to_column(dimension.min_column).unwrap();
        let column_max_str = number_to_column(dimension.max_column).unwrap();
        let min_row = dimension.min_row;
        let max_row = dimension.max_row;
        let sheet_dimension_str = &format!("{column_min_str}{min_row}:{column_max_str}{max_row}");
        let is_sheet_selected = selected_sheet as usize == sheet_index;
        // ENS-Office: the autofilter range lives in the sheet's _xlnm._FilterDatabase name.
        let auto_filter = workbook
            .defined_names
            .iter()
            .find(|d| d.sheet_id == Some(worksheet.sheet_id) && d.name.eq_ignore_ascii_case(ironcalc_base::FILTER_DATABASE))
            .and_then(|d| ironcalc_base::parse_filter_range(&d.formula))
            .map(|(r1, c1, r2, c2)| {
                let range = format!(
                    "{}{r1}:{}{r2}",
                    number_to_column(c1).unwrap_or_default(),
                    number_to_column(c2).unwrap_or_default()
                );
                let columns = filter_columns_xml(&worksheet.ens_filter_criteria);
                if columns.is_empty() {
                    format!("<autoFilter ref=\"{range}\"/>")
                } else {
                    format!("<autoFilter ref=\"{range}\">{columns}</autoFilter>")
                }
            })
            .unwrap_or_default();
        zip.write_all(
            worksheets::get_worksheet_xml(
                &auto_filter,
                worksheet,
                &model.parsed_formulas[sheet_index],
                sheet_dimension_str,
                is_sheet_selected,
            )
            .as_bytes(),
        )?;
    }

    let writer = zip.finish()?;
    Ok(writer)
}

/// Exports a model to an icalc file
pub fn save_to_icalc(model: &Model, file_name: &str) -> Result<(), XlsxError> {
    let file_path = std::path::Path::new(&file_name);
    if file_path.exists() {
        return Err(XlsxError::IO(format!("file {file_name} already exists")));
    }
    let s = bitcode::encode(&model.workbook);
    let mut file = fs::File::create(file_path)?;
    file.write_all(&s)?;

    Ok(())
}


/// ENS-Office: the first filter's choices as Excel's `<filterColumn>` elements.
fn filter_columns_xml(criteria: &str) -> String {
    let Ok(serde_json::Value::Array(filters)) = serde_json::from_str::<serde_json::Value>(criteria) else {
        return String::new();
    };
    let Some(serde_json::Value::Array(columns)) = filters.first() else {
        return String::new();
    };
    let mut out = String::new();
    for c in columns {
        let (Some(col), Some(serde_json::Value::Array(values))) = (c.get(0).and_then(|v| v.as_i64()), c.get(1)) else {
            continue;
        };
        // Отбор по цвету («fill:…», «font:…») Excel в таком виде не знает: не пишем.
        if values.iter().filter_map(|v| v.as_str()).any(|v| v.starts_with("fill:") || v.starts_with("font:")) {
            continue;
        }
        let blank = values.iter().any(|v| v.as_str() == Some("(пустые)"));
        let items: String = values
            .iter()
            .filter_map(|v| v.as_str())
            .filter(|v| *v != "(пустые)")
            .map(|v| format!("<filter val=\"{}\"/>", escape::escape_xml(v)))
            .collect();
        let blank_attr = if blank { " blank=\"1\"" } else { "" };
        out.push_str(&format!("<filterColumn colId=\"{col}\"><filters{blank_attr}>{items}</filters></filterColumn>"));
    }
    out
}
