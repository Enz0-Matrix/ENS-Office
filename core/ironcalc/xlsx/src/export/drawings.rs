//! ENS-Office: pictures over a sheet (`xl/drawings/drawingN.xml`, `xl/media/imageK.*`).

use ironcalc_base::types::Image;

use super::{escape::escape_xml, xml_constants::XML_DECLARATION};

const EMU_PER_PX: i64 = 9525;

/// `xl/drawings/drawingN.xml`: one `<xdr:oneCellAnchor>` with a picture per image.
pub(crate) fn get_drawing_xml(images: &[Image]) -> String {
    let mut anchors = String::new();
    for (i, img) in images.iter().enumerate() {
        let (cx, cy) = (i64::from(img.width) * EMU_PER_PX, i64::from(img.height) * EMU_PER_PX);
        anchors.push_str(&format!(
            "<xdr:oneCellAnchor><xdr:from><xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff></xdr:from>\
<xdr:ext cx=\"{cx}\" cy=\"{cy}\"/>\
<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"{}\" name=\"{}\"/><xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr></xdr:nvPicPr>\
<xdr:blipFill><a:blip r:embed=\"rId{}\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>\
<xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>\
<xdr:clientData/></xdr:oneCellAnchor>",
            img.column - 1,
            i64::from(img.dx) * EMU_PER_PX,
            img.row - 1,
            i64::from(img.dy) * EMU_PER_PX,
            i + 2,
            escape_xml(&img.name),
            i + 1
        ));
    }
    format!(
        "{XML_DECLARATION}\n<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">{anchors}</xdr:wsDr>"
    )
}

/// The media file of an image (`first` — number of the sheet's first image in the book).
pub(crate) fn media_name(first: usize, i: usize, img: &Image) -> String {
    format!("image{}.{}", first + i, img.ext)
}

/// `xl/drawings/_rels/drawingN.xml.rels`.
pub(crate) fn get_drawing_rels(images: &[Image], first: usize) -> String {
    let rels: String = images
        .iter()
        .enumerate()
        .map(|(i, img)| {
            format!(
                "<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/{}\"/>",
                i + 1,
                media_name(first, i, img)
            )
        })
        .collect();
    format!("{XML_DECLARATION}\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{rels}</Relationships>")
}
