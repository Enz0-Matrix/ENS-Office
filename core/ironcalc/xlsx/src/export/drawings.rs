//! ENS-Office: pictures over a sheet (`xl/drawings/drawingN.xml`, `xl/media/imageK.*`).

use ironcalc_base::{
    cell::CellValue,
    expressions::utils::{number_to_column, quote_name},
    parse_chart_range,
    types::{Chart, Image},
    Model,
};

use super::{escape::escape_xml, xml_constants::XML_DECLARATION};

const EMU_PER_PX: i64 = 9525;

/// `xl/drawings/drawingN.xml`: one `<xdr:oneCellAnchor>` with a picture per image.
pub(crate) fn get_drawing_xml(images: &[Image], charts: &[Chart]) -> String {
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
    for (j, ch) in charts.iter().enumerate() {
        let (cx, cy) = (i64::from(ch.width) * EMU_PER_PX, i64::from(ch.height) * EMU_PER_PX);
        let rel = images.len() + j + 1;
        anchors.push_str(&format!(
            "<xdr:oneCellAnchor><xdr:from><xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff></xdr:from>\
<xdr:ext cx=\"{cx}\" cy=\"{cy}\"/>\
<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"{}\" name=\"Диаграмма {}\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>\
<xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></xdr:xfrm>\
<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" r:id=\"rId{rel}\"/></a:graphicData></a:graphic>\
</xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>",
            ch.column - 1,
            i64::from(ch.dx) * EMU_PER_PX,
            ch.row - 1,
            i64::from(ch.dy) * EMU_PER_PX,
            images.len() + j + 2,
            j + 1
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
pub(crate) fn get_drawing_rels(images: &[Image], first: usize, charts: &[Chart], first_chart: usize) -> String {
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
        .collect::<String>()
        + &(0..charts.len())
            .map(|j| {
                format!(
                    "<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\" Target=\"../charts/chart{}.xml\"/>",
                    images.len() + j + 1,
                    first_chart + j
                )
            })
            .collect::<String>();
    format!("{XML_DECLARATION}\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{rels}</Relationships>")
}

fn col(c: i32) -> String {
    number_to_column(c).unwrap_or_default()
}

fn range_ref(sheet: &str, c1: i32, r1: i32, c2: i32, r2: i32) -> String {
    let (a, b) = (format!("${}${r1}", col(c1)), format!("${}${r2}", col(c2)));
    let quoted = quote_name(sheet);
    let q = escape_xml(&quoted);
    if (c1, r1) == (c2, r2) {
        format!("{q}!{a}")
    } else {
        format!("{q}!{a}:{b}")
    }
}

/// `xl/charts/chartN.xml`: the chart with the values of its cells as a cache.
pub(crate) fn get_chart_xml(model: &Model, sheet: u32, chart: &Chart) -> String {
    let Some((r1, c1, r2, c2)) = parse_chart_range(&chart.range) else { return String::new() };
    let name = model.workbook.worksheet(sheet).map(|w| w.name.clone()).unwrap_or_default();
    let text = |r: i32, c: i32| model.get_formatted_cell_value(sheet, r, c).unwrap_or_default();
    let num = |r: i32, c: i32| match model.get_cell_value_by_index(sheet, r, c) {
        Ok(CellValue::Number(n)) => Some(n),
        _ => None,
    };
    let cats: String = format!(
        "<c:cat><c:strRef><c:f>{}</c:f><c:strCache><c:ptCount val=\"{}\"/>{}</c:strCache></c:strRef></c:cat>",
        range_ref(&name, c1, r1 + 1, c1, r2),
        r2 - r1,
        (r1 + 1..=r2).map(|r| format!("<c:pt idx=\"{}\"><c:v>{}</c:v></c:pt>", r - r1 - 1, escape_xml(&text(r, c1)))).collect::<String>()
    );
    let last_col = if chart.kind == 2 { c1 + 1 } else { c2 };
    let color = |i: usize| chart.colors.get(i).and_then(|c| c.strip_prefix('#')).filter(|c| c.len() == 6).map(|c| c.to_uppercase());
    let fill = |i: usize| color(i).map(|c| format!("<a:solidFill><a:srgbClr val=\"{c}\"/></a:solidFill>"));
    let mut series = String::new();
    for (i, c) in (c1 + 1..=last_col).enumerate() {
        let vals: String = (r1 + 1..=r2)
            .filter_map(|r| num(r, c).map(|n| format!("<c:pt idx=\"{}\"><c:v>{n}</c:v></c:pt>", r - r1 - 1)))
            .collect();
        let sp_pr = match (chart.kind, fill(i)) {
            (0, Some(f)) => format!("<c:spPr>{f}</c:spPr>"),
            (1, Some(f)) => format!("<c:spPr><a:ln w=\"28575\" cap=\"rnd\">{f}<a:round/></a:ln></c:spPr>"),
            _ => String::new(),
        };
        let d_pt: String = if chart.kind == 2 {
            (0..r2 - r1)
                .filter_map(|k| fill(k as usize).map(|f| format!("<c:dPt><c:idx val=\"{k}\"/><c:bubble3D val=\"0\"/><c:spPr>{f}</c:spPr></c:dPt>")))
                .collect()
        } else {
            String::new()
        };
        let d_lbls = if chart.labels {
            format!(
                "<c:dLbls><c:showLegendKey val=\"0\"/><c:showVal val=\"{}\"/><c:showCatName val=\"0\"/><c:showSerName val=\"0\"/><c:showPercent val=\"{}\"/><c:showBubbleSize val=\"0\"/></c:dLbls>",
                u8::from(chart.kind != 2),
                u8::from(chart.kind == 2)
            )
        } else {
            String::new()
        };
        series.push_str(&format!(
            "<c:ser><c:idx val=\"{i}\"/><c:order val=\"{i}\"/><c:tx><c:strRef><c:f>{}</c:f><c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>{}</c:v></c:pt></c:strCache></c:strRef></c:tx>{sp_pr}{}{d_pt}{d_lbls}{cats}\
<c:val><c:numRef><c:f>{}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"{}\"/>{vals}</c:numCache></c:numRef></c:val>{}</c:ser>",
            range_ref(&name, c, r1, c, r1),
            escape_xml(&text(r1, c)),
            if chart.kind == 1 { "<c:marker><c:symbol val=\"none\"/></c:marker>" } else { "" },
            range_ref(&name, c, r1 + 1, c, r2),
            r2 - r1,
            if chart.kind == 1 { "<c:smooth val=\"0\"/>" } else { "" }
        ));
    }
    let axes = "<c:catAx><c:axId val=\"111\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:numFmt formatCode=\"General\" sourceLinked=\"0\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"222\"/><c:crosses val=\"autoZero\"/><c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/></c:catAx>\
<c:valAx><c:axId val=\"222\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/><c:majorGridlines/><c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"111\"/><c:crosses val=\"autoZero\"/></c:valAx>";
    let plot = match chart.kind {
        0 => format!("<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/><c:varyColors val=\"0\"/>{series}<c:gapWidth val=\"150\"/><c:axId val=\"111\"/><c:axId val=\"222\"/></c:barChart>{axes}"),
        1 => format!("<c:lineChart><c:grouping val=\"standard\"/><c:varyColors val=\"0\"/>{series}<c:marker val=\"1\"/><c:axId val=\"111\"/><c:axId val=\"222\"/></c:lineChart>{axes}"),
        _ => format!("<c:pieChart><c:varyColors val=\"1\"/>{series}<c:firstSliceAng val=\"0\"/></c:pieChart>"),
    };
    let title = if chart.title.is_empty() {
        "<c:autoTitleDeleted val=\"1\"/>".to_string()
    } else {
        format!("<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>{}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title><c:autoTitleDeleted val=\"0\"/>", escape_xml(&chart.title))
    };
    let legend = match chart.legend {
        2 => String::new(),
        1 => "<c:legend><c:legendPos val=\"r\"/></c:legend>".to_string(),
        _ => "<c:legend><c:legendPos val=\"b\"/></c:legend>".to_string(),
    };
    format!(
        "{XML_DECLARATION}\n<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\
<c:chart>{title}<c:plotArea><c:layout/>{plot}</c:plotArea>{legend}<c:plotVisOnly val=\"1\"/></c:chart></c:chartSpace>"
    )
}
