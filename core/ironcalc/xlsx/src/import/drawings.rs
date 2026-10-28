//! ENS-Office: pictures of a sheet (`<drawing>` → `xl/drawings/drawingN.xml` → `xl/media/*`).

use std::io::{BufReader, Read};

use ironcalc_base::types::Image;

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
pub(crate) fn load_images<R: Read + std::io::Seek>(archive: &mut zip::read::ZipArchive<R>, sheet_path: &str) -> Vec<Image> {
    let mut out = Vec::new();
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
    out
}
