use std::{collections::HashMap, io::Read};

use ironcalc_base::types::{
    Alignment, Border, BorderItem, BorderStyle, CellStyleXfs, CellStyles, CellXfs, Color, Dxf,
    DxfFont, Fill, Font, FontScheme, HorizontalAlignment, NumFmt, Protection, Styles, Theme, VerticalAlignment,
};
use std::io::BufReader;

use super::xml::XmlNode;

use crate::error::XlsxError;

use super::util::{get_attribute, get_bool, get_bool_false, get_color_indexed, get_number};

/// Reads the workbook's `<colors><indexedColors>` palette override (OOXML §18.8.27) from the
/// styleSheet root as `#RRGGBB` strings positioned by index. Returns `None` when the file
/// supplies no override (the common case), so callers keep the legacy default indexed palette.
/// A malformed entry becomes an empty string, which the colour resolver treats as "fall back".

/// ENS-Office: a boolean font property (`<b/>`, `<b val="1"/>`, `<b val="0"/>`,
/// `<u val="none"/>`). ECMA-376: no `val` means on.
fn flag_on(val: Option<&str>) -> bool {
    !matches!(val, Some("0" | "false" | "none"))
}

/// ENS-Office: font size, also written as "8.000000" or "10.5".
fn font_size(val: Option<&str>) -> i32 {
    val.and_then(|v| v.trim().parse::<f64>().ok())
        .filter(|v| v.is_finite() && *v > 0.0)
        .map_or(11, |v| v.round() as i32)
}

fn parse_indexed_colors(style_sheet: &XmlNode) -> Option<Vec<String>> {
    let colors = style_sheet.children().find(|n| n.has_tag_name("colors"))?;
    let indexed = colors
        .children()
        .find(|n| n.has_tag_name("indexedColors"))?;
    let palette: Vec<String> = indexed
        .children()
        .filter(|n| n.has_tag_name("rgbColor"))
        .map(|n| match n.attribute("rgb") {
            // ARGB (drop the alpha byte) or bare RGB; `is_ascii` guards the byte slice.
            Some(raw) if raw.len() == 8 && raw.is_ascii() => {
                format!("#{}", raw[2..].to_ascii_uppercase())
            }
            Some(raw) if raw.len() == 6 && raw.is_ascii() => {
                format!("#{}", raw.to_ascii_uppercase())
            }
            _ => String::new(),
        })
        .collect();
    if palette.is_empty() {
        None
    } else {
        Some(palette)
    }
}

fn get_border(
    node: &XmlNode,
    name: &str,
    theme: &Theme,
    indexed: Option<&[String]>,
) -> Result<Option<BorderItem>, XlsxError> {
    let style;
    let color;
    let border_nodes = node
        .children()
        .filter(|n| n.has_tag_name(name))
        .collect::<Vec<&XmlNode>>();
    if border_nodes.len() == 1 {
        let border = border_nodes[0];
        style = match border.attribute("style") {
            Some("thin") => BorderStyle::Thin,
            Some("medium") => BorderStyle::Medium,
            Some("thick") => BorderStyle::Thick,
            Some("double") => BorderStyle::Double,
            Some("slantdashdot") => BorderStyle::SlantDashDot,
            Some("mediumdashed") => BorderStyle::MediumDashed,
            Some("mediumdashdot") => BorderStyle::MediumDashDot,
            Some("mediumdashdotdot") => BorderStyle::MediumDashDotDot,
            // TODO: Should we fail in this case or set the border to None?
            Some(_) => BorderStyle::Thin,
            None => {
                return Ok(None);
            }
        };

        let color_node = border
            .children()
            .filter(|n| n.has_tag_name("color"))
            .collect::<Vec<&XmlNode>>();
        if color_node.len() == 1 {
            color = get_color_indexed(color_node[0], theme, indexed)?;
        } else {
            color = Color::None;
        }
    } else {
        return Ok(None);
    }
    Ok(Some(BorderItem { style, color }))
}

pub(super) fn load_styles<R: Read + std::io::Seek>(
    archive: &mut zip::read::ZipArchive<R>,
    theme: &Theme,
) -> Result<Styles, XlsxError> {
    let file = archive.by_name("xl/styles.xml")?;
    let doc = XmlNode::parse(BufReader::new(file))?;
    let style_sheet = &doc;

    // The workbook's `<colors><indexedColors>` override (if any), applied to every `indexed=`
    // fill/font/border colour below so the file's palette wins over the legacy default one.
    let indexed_colors = parse_indexed_colors(style_sheet);
    let indexed = indexed_colors.as_deref();

    let mut num_fmts = Vec::new();
    let num_fmts_nodes = style_sheet
        .children()
        .filter(|n| n.has_tag_name("numFmts"))
        .collect::<Vec<&XmlNode>>();
    if num_fmts_nodes.len() == 1 {
        for num_fmt in num_fmts_nodes[0].children() {
            let num_fmt_id = get_number(num_fmt, "numFmtId");
            let format_code = num_fmt.attribute("formatCode").unwrap_or("").to_string();
            num_fmts.push(NumFmt {
                num_fmt_id,
                format_code,
            });
        }
    }

    let mut fonts = Vec::new();
    let font_nodes = style_sheet
        .children()
        .filter(|n| n.has_tag_name("fonts"))
        .collect::<Vec<&XmlNode>>()[0];
    for font in font_nodes.children() {
        let mut sz = 11;
        let mut name = "Inter".to_string();
        // NOTE: In Excel you can have simple underline or double underline
        // In IronCalc convert double underline to simple
        // This in excel is u with a value of "double"
        let mut u = false;
        let mut b = false;
        let mut i = false;
        let mut strike = false;
        // No <color> child is semantically equivalent to <color auto="1"/> in OOXML —
        // both mean "use the automatic/default color", which we represent as None.
        // Collapsing this to Some("#000000") would make it indistinguishable from an
        // explicit <color rgb="FF000000"/>.
        let mut color: Color = Color::None;
        let mut family = 2;
        let mut scheme = FontScheme::default();
        let mut vert_align = ironcalc_base::types::VertAlign::default();
        for feature in font.children() {
            match feature.tag_name() {
                "sz" => {
                    // ENS-Office: Р7/LibreOffice write "8.000000"
                    sz = font_size(feature.attribute("val"));
                }
                "color" => {
                    color = get_color_indexed(feature, theme, indexed)?;
                }
                // ENS-Office: <b val="0"/>, <strike val="0"/>, <u val="none"/> mean OFF
                // (Р7 and LibreOffice write them for every font).
                "u" => {
                    u = flag_on(feature.attribute("val"));
                }
                "b" => {
                    b = flag_on(feature.attribute("val"));
                }
                "i" => {
                    i = flag_on(feature.attribute("val"));
                }
                "strike" => {
                    strike = flag_on(feature.attribute("val"));
                }
                "vertAlign" => {
                    vert_align = match feature.attribute("val") {
                        Some("superscript") => ironcalc_base::types::VertAlign::Superscript,
                        Some("subscript") => ironcalc_base::types::VertAlign::Subscript,
                        _ => ironcalc_base::types::VertAlign::Baseline,
                    };
                }
                // ECMA-376 §18.8.29 (`name`): the `val` attribute is the font's typeface.
                "name" => {
                    name = feature.attribute("val").unwrap_or("Inter").to_string();
                }
                "family" => {
                    family = feature
                        .attribute("val")
                        .unwrap_or("2")
                        .parse::<i32>()
                        .unwrap_or(2);
                }
                "scheme" => {
                    scheme = match feature.attribute("val") {
                        None => FontScheme::default(),
                        Some("minor") => FontScheme::Minor,
                        Some("major") => FontScheme::Major,
                        Some("none") => FontScheme::None,
                        // TODO: Should we fail?
                        Some(_) => FontScheme::default(),
                    }
                }
                "charset" => {}
                _ => {
                    println!("Unexpected feature {feature:?}");
                }
            }
        }
        fonts.push(Font {
            strike,
            u,
            b,
            i,
            sz,
            color,
            name,
            family,
            scheme,
            vert_align,
        });
    }

    let mut fills = Vec::new();
    let fill_nodes = style_sheet
        .children()
        .filter(|n| n.has_tag_name("fills"))
        .collect::<Vec<&XmlNode>>()[0];
    for fill in fill_nodes.children() {
        let pattern_fill = fill
            .children()
            .filter(|n| n.has_tag_name("patternFill"))
            .collect::<Vec<&XmlNode>>();
        if pattern_fill.len() != 1 {
            // safety belt
            // Some fills do not have a patternFill, but they have gradientFill
            fills.push(Fill::default());
            continue;
        }
        let pattern_fill = pattern_fill[0];

        let mut fg_color = Color::None;
        let mut bg_color = Color::None;
        for feature in pattern_fill.children() {
            match feature.tag_name() {
                "fgColor" => {
                    fg_color = get_color_indexed(feature, theme, indexed)?;
                }
                "bgColor" => {
                    bg_color = get_color_indexed(feature, theme, indexed)?;
                }
                _ => {
                    println!("Unexpected pattern");
                    dbg!(feature);
                }
            }
        }
        // Prefer fgColor (solid fill convention); fall back to bgColor
        fills.push(Fill {
            color: if fg_color.is_some() {
                fg_color
            } else {
                bg_color
            },
        })
    }

    let mut borders = Vec::new();
    let border_nodes = style_sheet
        .children()
        .filter(|n| n.has_tag_name("borders"))
        .collect::<Vec<&XmlNode>>()[0];
    for border in border_nodes.children() {
        let diagonal_up = get_bool_false(border, "diagonal_up");
        let diagonal_down = get_bool_false(border, "diagonal_down");
        let left = get_border(border, "left", theme, indexed)?;
        let right = get_border(border, "right", theme, indexed)?;
        let top = get_border(border, "top", theme, indexed)?;
        let bottom = get_border(border, "bottom", theme, indexed)?;
        let diagonal = get_border(border, "diagonal", theme, indexed)?;
        borders.push(Border {
            diagonal_up,
            diagonal_down,
            left,
            right,
            top,
            bottom,
            diagonal,
        });
    }

    let mut cell_style_xfs = Vec::new();
    let cell_style_xfs_nodes = style_sheet
        .children()
        .filter(|n| n.has_tag_name("cellStyleXfs"))
        .collect::<Vec<&XmlNode>>()[0];
    for xfs in cell_style_xfs_nodes.children() {
        let num_fmt_id = get_number(xfs, "numFmtId");
        let font_id = get_number(xfs, "fontId");
        let fill_id = get_number(xfs, "fillId");
        let border_id = get_number(xfs, "borderId");
        let apply_number_format = get_bool(xfs, "applyNumberFormat");
        let apply_border = get_bool(xfs, "applyBorder");
        let apply_alignment = get_bool(xfs, "applyAlignment");
        let apply_protection = get_bool(xfs, "applyProtection");
        let apply_font = get_bool(xfs, "applyFont");
        let apply_fill = get_bool(xfs, "applyFill");

        cell_style_xfs.push(CellStyleXfs {
            num_fmt_id,
            font_id,
            fill_id,
            border_id,
            apply_number_format,
            apply_border,
            apply_alignment,
            apply_protection,
            apply_font,
            apply_fill,
        });
    }

    let mut cell_styles = Vec::new();
    let mut style_names = HashMap::new();
    let cell_style_nodes = style_sheet
        .children()
        .filter(|n| n.has_tag_name("cellStyles"))
        .collect::<Vec<&XmlNode>>()[0];
    for cell_style in cell_style_nodes.children() {
        let name = get_attribute(cell_style, "name")?.to_string();
        let xf_id = get_number(cell_style, "xfId");
        let builtin_id = get_number(cell_style, "builtinId");
        // NB: A builtin style could be hidden (this is removed in the UI)
        // <cellStyle name="Linked Cell" xfId="8" builtinId="24" hidden="1"/>
        // NB: A builtin style could be modified
        // <cellStyle name="Good" xfId="4" builtinId="26" customBuiltin="1"/>
        // let hidden = get_bool(cell_style, "hidden");
        // let custom_builtin = get_bool(cell_style, "customBuiltin");
        style_names.insert(xf_id, name.clone());
        cell_styles.push(CellStyles {
            name,
            xf_id,
            builtin_id,
        })
    }

    let mut cell_xfs = Vec::new();
    let cell_xfs_nodes = style_sheet
        .children()
        .filter(|n| n.has_tag_name("cellXfs"))
        .collect::<Vec<&XmlNode>>()[0];
    for xfs in cell_xfs_nodes.children() {
        // `xfId` is optional on a cellXfs <xf> (it references cellStyleXfs;
        // many Excel/LibreOffice files omit it). Default to 0 when absent.
        let xf_id = xfs
            .attribute("xfId")
            .map(|s| s.parse::<i32>())
            .transpose()?
            .unwrap_or(0);
        let num_fmt_id = get_number(xfs, "numFmtId");
        let font_id = get_number(xfs, "fontId");
        let fill_id = get_number(xfs, "fillId");
        let border_id = get_number(xfs, "borderId");
        let apply_number_format = get_bool_false(xfs, "applyNumberFormat");
        let apply_border = get_bool_false(xfs, "applyBorder");
        let apply_alignment = get_bool_false(xfs, "applyAlignment");
        let apply_protection = get_bool_false(xfs, "applyProtection");
        let apply_font = get_bool_false(xfs, "applyFont");
        let apply_fill = get_bool_false(xfs, "applyFill");
        let quote_prefix = get_bool_false(xfs, "quotePrefix");

        // TODO: Pivot Tables
        // let pivotButton = get_bool(xfs, "pivotButton");

        let alignment_nodes = xfs
            .children()
            .filter(|n| n.has_tag_name("alignment"))
            .collect::<Vec<&XmlNode>>();
        let alignment = if alignment_nodes.len() == 1 {
            let alignment_node = alignment_nodes[0];
            let wrap_text = get_bool_false(alignment_node, "wrapText");
            let text_rotation = alignment_node
                .attribute("textRotation")
                .and_then(|v| v.parse::<i32>().ok())
                .unwrap_or(0);

            let horizontal = match alignment_node.attribute("horizontal") {
                Some("center") => HorizontalAlignment::Center,
                Some("centerContinuous") => HorizontalAlignment::CenterContinuous,
                Some("distributed") => HorizontalAlignment::Distributed,
                Some("fill") => HorizontalAlignment::Fill,
                Some("general") => HorizontalAlignment::General,
                Some("justify") => HorizontalAlignment::Justify,
                Some("left") => HorizontalAlignment::Left,
                Some("right") => HorizontalAlignment::Right,
                // TODO: Should we fail in this case or set the alignment to default?
                Some(_) => HorizontalAlignment::default(),
                None => HorizontalAlignment::default(),
            };

            let vertical = match alignment_node.attribute("vertical") {
                Some("bottom") => VerticalAlignment::Bottom,
                Some("center") => VerticalAlignment::Center,
                Some("distributed") => VerticalAlignment::Distributed,
                Some("justify") => VerticalAlignment::Justify,
                Some("top") => VerticalAlignment::Top,
                // TODO: Should we fail in this case or set the alignment to default?
                Some(_) => VerticalAlignment::default(),
                None => VerticalAlignment::default(),
            };

            Some(Alignment {
                horizontal,
                vertical,
                wrap_text,
                text_rotation,
            })
        } else {
            None
        };

        // ENS-Office: `<protection locked hidden>`; absent — locked, not hidden.
        let protection = xfs.children().find(|n| n.has_tag_name("protection")).and_then(|n| {
            let flag = |a: &str, d: bool| n.attribute(a).map_or(d, |v| v == "1" || v == "true");
            let p = Protection { locked: flag("locked", true), hidden: flag("hidden", false) };
            (p.locked != true || p.hidden).then_some(p)
        });
        cell_xfs.push(CellXfs {
            xf_id,
            num_fmt_id,
            font_id,
            fill_id,
            border_id,
            apply_number_format,
            apply_border,
            apply_alignment,
            apply_protection,
            apply_font,
            apply_fill,
            quote_prefix,
            alignment,
            protection,
        });
    }

    let dxfs = load_dxfs(style_sheet, theme, indexed)?;

    Ok(Styles {
        num_fmts,
        fonts,
        fills,
        borders,
        cell_style_xfs,
        cell_xfs,
        cell_styles,
        dxfs,
    })
}

fn load_dxfs(
    style_sheet: &XmlNode,
    theme: &Theme,
    indexed: Option<&[String]>,
) -> Result<Vec<Dxf>, XlsxError> {
    let mut dxfs = Vec::new();
    let dxfs_nodes = style_sheet
        .children()
        .filter(|n| n.has_tag_name("dxfs"))
        .collect::<Vec<&XmlNode>>();
    if dxfs_nodes.is_empty() {
        return Ok(dxfs);
    }
    for dxf_node in dxfs_nodes[0].children() {
        dxfs.push(parse_dxf(dxf_node, theme, indexed)?);
    }
    Ok(dxfs)
}

/// Parses a single `<dxf>` (differential format) node into a [`Dxf`].
/// Matches children by local tag name, so it also works on namespaced
/// `<x14:dxf>` nodes found in conditional-formatting `extLst` extensions.
pub(super) fn parse_dxf(
    dxf_node: &XmlNode,
    theme: &Theme,
    indexed: Option<&[String]>,
) -> Result<Dxf, XlsxError> {
    let mut font = None;
    let mut fill = None;
    let mut border = None;
    let mut num_fmt = None;
    let mut alignment = None;

    for child in dxf_node.children() {
        match child.tag_name() {
            "font" => {
                let mut f = DxfFont::default();
                for feat in child.children() {
                    match feat.tag_name() {
                        "color" => {
                            f.color = get_color_indexed(feat, theme, indexed)?;
                        }
                        "b" => {
                            f.b = Some(flag_on(feat.attribute("val")));
                        }
                        "i" => {
                            f.i = Some(flag_on(feat.attribute("val")));
                        }
                        "u" => {
                            f.u = Some(flag_on(feat.attribute("val")));
                        }
                        "strike" => {
                            f.strike = Some(flag_on(feat.attribute("val")));
                        }
                        "sz" => {
                            f.sz = Some(font_size(feat.attribute("val")));
                        }
                        _ => {}
                    }
                }
                font = Some(f);
            }
            "fill" => {
                let pattern_fill_nodes = child
                    .children()
                    .filter(|n| n.has_tag_name("patternFill"))
                    .collect::<Vec<&XmlNode>>();
                if pattern_fill_nodes.len() == 1 {
                    let pf = pattern_fill_nodes[0];
                    let mut fg_color = Color::None;
                    let mut bg_color = Color::None;
                    for feat in pf.children() {
                        match feat.tag_name() {
                            "fgColor" => fg_color = get_color_indexed(feat, theme, indexed)?,
                            "bgColor" => bg_color = get_color_indexed(feat, theme, indexed)?,
                            _ => {}
                        }
                    }
                    // Prefer fgColor (solid fill convention); fall back to bgColor
                    fill = Some(Fill {
                        color: if fg_color.is_some() {
                            fg_color
                        } else {
                            bg_color
                        },
                    });
                }
            }
            "border" => {
                let left = get_border(child, "left", theme, indexed)?;
                let right = get_border(child, "right", theme, indexed)?;
                let top = get_border(child, "top", theme, indexed)?;
                let bottom = get_border(child, "bottom", theme, indexed)?;
                let diagonal = get_border(child, "diagonal", theme, indexed)?;
                border = Some(Border {
                    diagonal_up: false,
                    diagonal_down: false,
                    left,
                    right,
                    top,
                    bottom,
                    diagonal,
                });
            }
            "numFmt" => {
                let num_fmt_id = get_number(child, "numFmtId");
                let format_code = child.attribute("formatCode").unwrap_or("").to_string();
                num_fmt = Some(NumFmt {
                    num_fmt_id,
                    format_code,
                });
            }
            "alignment" => {
                let wrap_text = matches!(child.attribute("wrapText"), Some("1"));
                let horizontal = match child.attribute("horizontal") {
                    Some("center") => HorizontalAlignment::Center,
                    Some("left") => HorizontalAlignment::Left,
                    Some("right") => HorizontalAlignment::Right,
                    Some("justify") => HorizontalAlignment::Justify,
                    _ => HorizontalAlignment::default(),
                };
                let vertical = match child.attribute("vertical") {
                    Some("bottom") => VerticalAlignment::Bottom,
                    Some("center") => VerticalAlignment::Center,
                    Some("top") => VerticalAlignment::Top,
                    _ => VerticalAlignment::default(),
                };
                alignment = Some(Alignment {
                    horizontal,
                    vertical,
                    wrap_text,
                    text_rotation: 0,
                });
            }
            _ => {}
        }
    }

    Ok(Dxf {
        font,
        fill,
        border,
        num_fmt,
        alignment,
    })
}

#[cfg(test)]
mod indexed_color_tests {
    #![allow(clippy::unwrap_used)]
    use super::*;
    use std::io::{Cursor, Write};

    // A structurally complete styles.xml (no namespace, matching the other import tests) whose
    // fill 2 uses `indexed="13"` and font 1 uses `indexed="14"`. `{colors}` is filled in per case.
    fn styles_xml(colors: &str) -> String {
        format!(
            concat!(
                "<styleSheet>",
                "<fonts count=\"2\"><font><sz val=\"11\"/></font>",
                "<font><color indexed=\"14\"/><sz val=\"11\"/></font></fonts>",
                "<fills count=\"3\"><fill><patternFill patternType=\"none\"/></fill>",
                "<fill><patternFill patternType=\"gray125\"/></fill>",
                "<fill><patternFill patternType=\"solid\"><fgColor indexed=\"13\"/></patternFill></fill></fills>",
                "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>",
                "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>",
                "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>",
                "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>",
                "{}",
                "</styleSheet>",
            ),
            colors
        )
    }

    // The workbook's `<colors><indexedColors>` override: index 13 -> #FFD931, 14 -> #FE634D
    // (mirrors a Numbers export). Standard entries fill 0..12.
    const OVERRIDE: &str = concat!(
        "<colors><indexedColors>",
        "<rgbColor rgb=\"FF000000\"/><rgbColor rgb=\"FFFFFFFF\"/><rgbColor rgb=\"FFFF0000\"/>",
        "<rgbColor rgb=\"FF00FF00\"/><rgbColor rgb=\"FF0000FF\"/><rgbColor rgb=\"FFFFFF00\"/>",
        "<rgbColor rgb=\"FFFF00FF\"/><rgbColor rgb=\"FF00FFFF\"/><rgbColor rgb=\"FF000000\"/>",
        "<rgbColor rgb=\"FFBDC0BF\"/><rgbColor rgb=\"FFA5A5A5\"/><rgbColor rgb=\"FF3F3F3F\"/>",
        "<rgbColor rgb=\"FFDBDBDB\"/><rgbColor rgb=\"FFFFD931\"/><rgbColor rgb=\"FFFE634D\"/>",
        "</indexedColors></colors>",
    );

    fn archive_of(styles: &str) -> zip::ZipArchive<Cursor<Vec<u8>>> {
        let mut buf = Vec::new();
        {
            let mut zw = zip::ZipWriter::new(Cursor::new(&mut buf));
            zw.start_file("xl/styles.xml", zip::write::FileOptions::default())
                .unwrap();
            zw.write_all(styles.as_bytes()).unwrap();
            zw.finish().unwrap();
        }
        zip::ZipArchive::new(Cursor::new(buf)).unwrap()
    }

    // With an `<indexedColors>` override, `indexed=` fills/fonts resolve to the FILE's palette.
    #[test]
    fn indexed_override_wins_over_default_palette() {
        let mut archive = archive_of(&styles_xml(OVERRIDE));
        let styles = load_styles(&mut archive, &Theme::default()).unwrap();
        // fill 2: fgColor indexed=13 -> the file's #FFD931 (default 13 is #FFFF00).
        assert_eq!(styles.fills[2].color, Color::Rgb("#FFD931".to_string()));
        // font 1: color indexed=14 -> the file's #FE634D (default 14 is #FF00FF).
        assert_eq!(styles.fonts[1].color, Color::Rgb("#FE634D".to_string()));
    }

    // Without an override, `indexed=` resolves against the legacy default palette (unchanged).
    #[test]
    fn no_override_keeps_default_palette() {
        let mut archive = archive_of(&styles_xml(""));
        let styles = load_styles(&mut archive, &Theme::default()).unwrap();
        assert_eq!(styles.fills[2].color, Color::Rgb("#FFFF00".to_string()));
        assert_eq!(styles.fonts[1].color, Color::Rgb("#FF00FF".to_string()));
    }

    // `parse_indexed_colors` reads the palette positionally and returns None when absent.
    #[test]
    fn parse_indexed_colors_reads_override_or_none() {
        let with = styles_xml(OVERRIDE);
        let doc = XmlNode::parse_str(&with).unwrap();
        let palette = parse_indexed_colors(&doc).unwrap();
        assert_eq!(palette.len(), 15);
        assert_eq!(palette[13], "#FFD931");
        assert_eq!(palette[14], "#FE634D");

        let without = styles_xml("");
        let doc = XmlNode::parse_str(&without).unwrap();
        assert!(parse_indexed_colors(&doc).is_none());
    }

    // `get_color_indexed` precedence + guards, independent of the file path.
    #[test]
    fn get_color_indexed_precedence_and_guards() {
        let theme = Theme::default();
        let palette = vec!["#000000".to_string(), "#FFFFFF".to_string()];
        let resolve = |xml: &str| {
            let doc = XmlNode::parse_str(xml).unwrap();
            get_color_indexed(&doc, &theme, Some(&palette)).unwrap()
        };
        // In-range index -> the override slot.
        assert_eq!(
            resolve(r#"<c indexed="1"/>"#),
            Color::Rgb("#FFFFFF".to_string())
        );
        // rgb= takes precedence over any palette.
        assert_eq!(
            resolve(r#"<c rgb="FF112233"/>"#),
            Color::Rgb("#112233".to_string())
        );
        // System index 64 (transparent) -> None, never a palette lookup.
        assert_eq!(resolve(r#"<c indexed="64"/>"#), Color::None);
        // Out-of-range index -> default palette fallback (index 5 default = #FFFF00).
        assert_eq!(
            resolve(r#"<c indexed="5"/>"#),
            Color::Rgb("#FFFF00".to_string())
        );
        // Malformed negative index -> default fallback, not a panic (would underflow
        // `index as usize` in `get_indexed_color` without the guard).
        assert_eq!(
            resolve(r#"<c indexed="-1"/>"#),
            Color::Rgb("#000000".to_string())
        );
        // theme= is unaffected by the indexed override.
        assert_eq!(
            resolve(r#"<c theme="3" tint="0.0"/>"#),
            Color::Theme(3, 0.0)
        );
    }
}

#[cfg(test)]
mod ens_tests {
    use super::{flag_on, font_size};

    #[test]
    fn ens_font_flags_and_sizes_from_r7() {
        assert!(flag_on(None));
        assert!(flag_on(Some("1")));
        assert!(flag_on(Some("true")));
        assert!(!flag_on(Some("0")));
        assert!(!flag_on(Some("false")));
        assert!(!flag_on(Some("none")));
        assert!(flag_on(Some("double")));
        assert_eq!(font_size(Some("8.000000")), 8);
        assert_eq!(font_size(Some("9")), 9);
        assert_eq!(font_size(Some("10.5")), 11);
        assert_eq!(font_size(None), 11);
    }

    /// Проверка на настоящем файле: ENS_CHECK_XLSX=путь cargo test -- --ignored
    #[test]
    #[ignore]
    fn ens_check_real_file_has_no_strike() {
        let Ok(path) = std::env::var("ENS_CHECK_XLSX") else { return };
        let model = crate::import::load_from_xlsx(&path, "ru", "UTC", "ru").unwrap();
        for (i, ws) in model.workbook.worksheets.iter().enumerate() {
            for (row, column, _) in ws.sheet_data.cells() {
                let style = model.get_style_for_cell(i as u32, row, column).unwrap();
                assert!(!style.font.strike, "зачёркнуто: лист {i}, {row}:{column}");
            }
        }
        let st = model.get_style_for_cell(0, 1, 1).unwrap();
        println!("A1: шрифт {} {}", st.font.name, st.font.sz);
    }
}
