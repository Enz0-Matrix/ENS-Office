//! ENS-Office: pictures of a sheet (`<drawing>` → `xl/drawings/drawingN.xml` → `xl/media/*`).

use std::io::{BufReader, Read};

use ironcalc_base::types::{Chart, Image};

use super::xml::XmlNode;

const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const EMU_PER_PX: i64 = 9525;

/// `dir` + relative `target` (`../media/a.png`) as a path inside the archive.
fn resolve(dir: &str, target: &str) -> String {
    if let Some(t) = target.strip_prefix('/') {
        return t.to_string();
    }
    let mut parts: Vec<&str> = dir.split('/').filter(|p| !p.is_empty()).collect();
    for p in target.split('/') {
        match p {
            ".." => {
                parts.pop();
            }
            "." | "" => {}
            _ => parts.push(p),
        }
    }
    parts.join("/")
}

fn dir_of(path: &str) -> &str {
    path.rsplit_once('/').map_or("", |(d, _)| d)
}

fn rels_of<R: Read + std::io::Seek>(archive: &mut zip::read::ZipArchive<R>, path: &str) -> Vec<(String, String, String)> {
    let (dir, file) = path.rsplit_once('/').unwrap_or(("", path));
    let Ok(f) = archive.by_name(&format!("{dir}/_rels/{file}.rels")) else { return vec![] };
    let Ok(doc) = XmlNode::parse(BufReader::new(f)) else { return vec![] };
    doc.children()
        .filter(|n| n.has_tag_name("Relationship"))
        .filter_map(|n| Some((n.attribute("Id")?.to_string(), n.attribute("Type")?.to_string(), n.attribute("Target")?.to_string())))
        .collect()
}

fn number(n: Option<&XmlNode>, child: &str) -> i64 {
    n.and_then(|n| n.children().find(|c| c.has_tag_name(child)))
        .and_then(|c| c.text().and_then(|t| t.trim().parse().ok()))
        .unwrap_or(0)
}

/// The pictures of the sheet stored at `sheet_path`; anything unreadable is skipped.
pub(crate) fn load_images<R: Read + std::io::Seek>(
    archive: &mut zip::read::ZipArchive<R>,
    sheet_path: &str,
    sheet_name: &str,
) -> (Vec<Image>, Vec<Chart>) {
    let mut out = Vec::new();
    let mut charts = Vec::new();
    let sheet_dir = dir_of(sheet_path).to_string();
    for (_, t, target) in rels_of(archive, sheet_path) {
        if !t.ends_with("/drawing") {
            continue;
        }
        let drawing_path = resolve(&sheet_dir, &target);
        let drawing_rels = rels_of(archive, &drawing_path);
        let mut text = String::new();
        if !matches!(archive.by_name(&drawing_path).map(|mut f| f.read_to_string(&mut text)), Ok(Ok(_))) {
            continue;
        }
        let Ok(doc) = XmlNode::parse(text.as_bytes()) else { continue };
        for anchor in doc.children() {
            if !matches!(anchor.tag_name(), "oneCellAnchor" | "twoCellAnchor" | "absoluteAnchor") {
                continue;
            }
            if let Some(frame) = anchor.children().find(|n| n.has_tag_name("graphicFrame")) {
                if let Some(c) = load_chart(archive, &drawing_rels, &drawing_path, anchor, frame, sheet_name) {
                    charts.push(c);
                }
                continue;
            }
            let Some(pic) = anchor.descendants().find(|n| n.has_tag_name("pic")) else { continue };
            let Some(embed) = pic.descendants().find(|n| n.has_tag_name("blip")).and_then(|b| b.attribute_ns(R_NS, "embed")) else { continue };
            let Some((_, _, media)) = drawing_rels.iter().find(|(id, _, _)| id == embed) else { continue };
            let media_path = resolve(dir_of(&drawing_path), media);
            let ext = match media_path.rsplit('.').next().map(str::to_lowercase).as_deref() {
                Some("png") => "png",
                Some("jpg") | Some("jpeg") => "jpeg",
                Some("gif") => "gif",
                _ => continue,
            };
            let mut data = Vec::new();
            if !matches!(archive.by_name(&media_path).map(|mut f| f.read_to_end(&mut data)), Ok(Ok(_))) {
                continue;
            }
            let from = anchor.children().find(|n| n.has_tag_name("from"));
            let ext_node = anchor
                .children()
                .find(|n| n.has_tag_name("ext"))
                .or_else(|| pic.descendants().find(|n| n.has_tag_name("xfrm")).and_then(|x| x.children().find(|n| n.has_tag_name("ext"))));
            let size = |a: &str| ext_node.and_then(|n| n.attribute(a)).and_then(|v| v.parse::<i64>().ok()).map_or(0, |v| (v / EMU_PER_PX) as i32);
            let (mut width, mut height) = (size("cx"), size("cy"));
            if width <= 0 || height <= 0 {
                (width, height) = (160, 120);
            }
            out.push(Image {
                name: pic.descendants().find(|n| n.has_tag_name("cNvPr")).and_then(|n| n.attribute("name")).unwrap_or("").to_string(),
                row: number(from, "row") as i32 + 1,
                column: number(from, "col") as i32 + 1,
                dx: (number(from, "colOff") / EMU_PER_PX) as i32,
                dy: (number(from, "rowOff") / EMU_PER_PX) as i32,
                width,
                height,
                ext: ext.to_string(),
                data,
            });
        }
    }
    (out, charts)
}

/// "'My sheet'!$A$2:$A$5" → (sheet name, column1, row1, column2, row2).
fn parse_ref(f: &str) -> Option<(String, i32, i32, i32, i32)> {
    let (sheet, cells) = f.rsplit_once('!')?;
    let sheet = sheet.trim_matches('\'').replace("''", "'");
    let cells = cells.replace('$', "");
    let (a, b) = cells.split_once(':').unwrap_or((&cells, &cells));
    let a = ironcalc_base::expressions::utils::parse_reference_a1(a)?;
    let b = ironcalc_base::expressions::utils::parse_reference_a1(b)?;
    Some((sheet, a.column, a.row, b.column, b.row))
}

fn load_chart<R: Read + std::io::Seek>(
    archive: &mut zip::read::ZipArchive<R>,
    drawing_rels: &[(String, String, String)],
    drawing_path: &str,
    anchor: &XmlNode,
    frame: &XmlNode,
    sheet_name: &str,
) -> Option<Chart> {
    let id = frame.descendants().find(|n| n.has_tag_name("chart"))?.attribute_ns(R_NS, "id")?;
    let (_, _, target) = drawing_rels.iter().find(|(i, _, _)| i == id)?;
    let path = resolve(dir_of(drawing_path), target);
    let mut text = String::new();
    archive.by_name(&path).ok()?.read_to_string(&mut text).ok()?;
    let doc = XmlNode::parse(text.as_bytes()).ok()?;
    let plot = doc.descendants().find(|n| n.has_tag_name("plotArea"))?;
    let (kind, chart) = plot.children().find_map(|n| match n.tag_name() {
        "barChart" => Some((0u8, n)),
        "lineChart" => Some((1, n)),
        "pieChart" => Some((2, n)),
        _ => None,
    })?;
    let title: String = doc
        .children()
        .flat_map(|c| c.children())
        .find(|n| n.has_tag_name("title"))
        .map(|t| t.descendants().filter(|n| n.has_tag_name("t")).filter_map(|n| n.text()).collect())
        .unwrap_or_default();
    let formula = |ser: &XmlNode, tag: &str| {
        ser.children()
            .find(|n| n.has_tag_name(tag))
            .and_then(|n| n.descendants().find(|n| n.has_tag_name("f")))
            .and_then(|n| n.text().map(|t| t.to_string()))
            .and_then(|f| parse_ref(&f))
    };
    let sers: Vec<&XmlNode> = chart.children().filter(|n| n.has_tag_name("ser")).collect();
    let first = sers.first()?;
    let last = sers.last()?;
    let vals_first = formula(first, "val")?;
    let vals_last = formula(last, "val")?;
    if vals_first.0 != sheet_name {
        return None;
    }
    let cat = formula(first, "cat");
    let c1 = cat.as_ref().map_or(vals_first.1 - 1, |c| c.1).max(1);
    let r1 = (vals_first.2 - 1).max(1);
    let (c2, r2) = (vals_last.3, vals_first.4);
    let from = anchor.children().find(|n| n.has_tag_name("from"));
    let ext = anchor.children().find(|n| n.has_tag_name("ext"));
    let size = |a: &str| ext.and_then(|n| n.attribute(a)).and_then(|v| v.parse::<i64>().ok()).map_or(0, |v| (v / EMU_PER_PX) as i32);
    let (mut width, mut height) = (size("cx"), size("cy"));
    if width <= 0 || height <= 0 {
        (width, height) = (480, 288);
    }
    let rgb = |n: &XmlNode| {
        n.descendants()
            .find(|c| c.has_tag_name("srgbClr"))
            .and_then(|c| c.attribute("val").map(|v| format!("#{}", v.to_uppercase())))
    };
    let colors: Vec<String> = if kind == 2 {
        let mut by_idx: Vec<(usize, String)> = first
            .children()
            .filter(|n| n.has_tag_name("dPt"))
            .filter_map(|d| Some((d.children().find(|n| n.has_tag_name("idx"))?.attribute("val")?.parse().ok()?, rgb(d)?)))
            .collect();
        by_idx.sort();
        // A palette is kept only when every slice has its colour.
        if by_idx.iter().enumerate().all(|(i, (k, _))| i == *k) { by_idx.into_iter().map(|(_, c)| c).collect() } else { vec![] }
    } else {
        let cs: Vec<Option<String>> = sers.iter().map(|x| x.children().find(|n| n.has_tag_name("spPr")).and_then(&rgb)).collect();
        if cs.iter().all(Option::is_some) { cs.into_iter().flatten().collect() } else { vec![] }
    };
    let legend = doc
        .descendants()
        .find(|n| n.has_tag_name("legend"))
        .map_or(2, |l| match l.children().find(|n| n.has_tag_name("legendPos")).and_then(|p| p.attribute("val")) {
            Some("r") | Some("l") | Some("tr") => 1,
            _ => 0,
        });
    let labels = chart.descendants().filter(|n| n.has_tag_name("dLbls")).any(|d| {
        d.children().any(|n| matches!(n.tag_name(), "showVal" | "showPercent") && n.attribute("val").is_some_and(|v| v == "1" || v == "true"))
    });
    Some(Chart {
        colors,
        legend,
        labels,
        kind,
        title,
        range: format!(
            "{}{r1}:{}{r2}",
            ironcalc_base::expressions::utils::number_to_column(c1)?,
            ironcalc_base::expressions::utils::number_to_column(c2)?
        ),
        row: number(from, "row") as i32 + 1,
        column: number(from, "col") as i32 + 1,
        dx: (number(from, "colOff") / EMU_PER_PX) as i32,
        dy: (number(from, "rowOff") / EMU_PER_PX) as i32,
        width,
        height,
    })
}
