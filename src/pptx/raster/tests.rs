//! Slides assembled part by part and painted at 72 dpi, where one slide point is one pixel:
//! a slide of 100 × 50 points is a 100 × 50 image, and an EMU position divides by 12,700.

use std::io::{Cursor, Write};

use crate::pptx::PptxParser;
use crate::raster::SlideRasterOptions;

const NS: &str = r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;
const REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// EMU for `pt` points.
fn emu(pt: u32) -> u32 {
    pt * 12_700
}

fn rels(items: &[(&str, &str, &str)]) -> String {
    let body: String = items
        .iter()
        .map(|(id, kind, target)| {
            format!(r#"<Relationship Id="{id}" Type="{REL}/{kind}" Target="{target}"/>"#)
        })
        .collect();
    format!(r#"<?xml version="1.0"?><Relationships xmlns="{REL_NS}">{body}</Relationships>"#)
}

/// A presentation of one 100 × 50 pt slide whose shape tree is `shapes`, with a layout, a
/// master (background `master_bg`, which may be empty) and a theme (`accent1` 4472C4).
fn deck(shapes: &str, master_bg: &str) -> Vec<u8> {
    let parts: Vec<(&str, String)> = vec![
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/></Types>"#.to_string(),
        ),
        (
            "_rels/.rels",
            rels(&[("rId1", "officeDocument", "ppt/presentation.xml")]),
        ),
        (
            "ppt/presentation.xml",
            format!(
                r#"<?xml version="1.0"?><p:presentation {NS}><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="{}" cy="{}"/></p:presentation>"#,
                emu(100),
                emu(50)
            ),
        ),
        (
            "ppt/_rels/presentation.xml.rels",
            rels(&[("rId2", "slide", "slides/slide1.xml")]),
        ),
        (
            "ppt/slides/slide1.xml",
            format!(
                r#"<?xml version="1.0"?><p:sld {NS}><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>{shapes}</p:spTree></p:cSld></p:sld>"#
            ),
        ),
        (
            "ppt/slides/_rels/slide1.xml.rels",
            rels(&[("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml")]),
        ),
        (
            "ppt/slideLayouts/slideLayout1.xml",
            format!(r#"<?xml version="1.0"?><p:sldLayout {NS}><p:cSld><p:spTree/></p:cSld></p:sldLayout>"#),
        ),
        (
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
            rels(&[("rId1", "slideMaster", "../slideMasters/slideMaster1.xml")]),
        ),
        (
            "ppt/slideMasters/slideMaster1.xml",
            format!(
                r#"<?xml version="1.0"?><p:sldMaster {NS}><p:cSld>{master_bg}<p:spTree/></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:sldMaster>"#
            ),
        ),
        (
            "ppt/slideMasters/_rels/slideMaster1.xml.rels",
            rels(&[("rId1", "theme", "../theme/theme1.xml")]),
        ),
        (
            "ppt/theme/theme1.xml",
            format!(
                r#"<?xml version="1.0"?><a:theme {NS} name="T"><a:themeElements><a:clrScheme name="C"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fmtScheme name="F"><a:fillStyleLst/><a:lnStyleLst><a:ln w="6350"/><a:ln w="38100"/><a:ln w="19050"/></a:lnStyleLst></a:fmtScheme></a:themeElements></a:theme>"#
            ),
        ),
    ];
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for (name, body) in parts {
        zip.start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
        zip.write_all(body.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

/// A preset shape at `(x, y)` points, `w × h` points, with `props` inside its `spPr` after
/// the geometry, and `extra` after the `spPr` (a style, a text body).
fn shape(
    prst: &str,
    (x, y, w, h): (u32, u32, u32, u32),
    xfrm_attrs: &str,
    props: &str,
    extra: &str,
) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="S"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm{xfrm_attrs}><a:off x="{}" y="{}"/><a:ext cx="{}" cy="{}"/></a:xfrm><a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>{props}</p:spPr>{extra}</p:sp>"#,
        emu(x),
        emu(y),
        emu(w),
        emu(h)
    )
}

fn solid(hex: &str) -> String {
    format!(r#"<a:solidFill><a:srgbClr val="{hex}"/></a:solidFill><a:ln><a:noFill/></a:ln>"#)
}

fn render(pptx: Vec<u8>) -> crate::raster::RasteredSlide {
    PptxParser::from_bytes(pptx)
        .unwrap()
        .render_slide(0, &SlideRasterOptions { dpi: 72.0 })
        .unwrap()
}

fn pixel(slide: &crate::raster::RasteredSlide, x: u32, y: u32) -> [u8; 3] {
    let i = ((y * slide.width + x) * 4) as usize;
    [slide.rgba[i], slide.rgba[i + 1], slide.rgba[i + 2]]
}

const WHITE: [u8; 3] = [255, 255, 255];
const RED: [u8; 3] = [255, 0, 0];
const ACCENT1: [u8; 3] = [0x44, 0x72, 0xC4];

#[test]
fn a_filled_rectangle_covers_its_place_on_a_white_slide() {
    let slide = render(deck(
        &shape("rect", (0, 0, 50, 50), "", &solid("FF0000"), ""),
        "",
    ));
    assert_eq!((slide.width, slide.height), (100, 50));
    assert_eq!(pixel(&slide, 25, 25), RED);
    assert_eq!(pixel(&slide, 75, 25), WHITE);
    assert!(slide.gaps.is_empty(), "{:?}", slide.gaps);
    assert!(slide.to_png().starts_with(b"\x89PNG"));
}

#[test]
fn scheme_colors_resolve_through_the_color_map_and_theme() {
    let fill = r#"<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>"#;
    let slide = render(deck(&shape("rect", (0, 0, 100, 50), "", fill, ""), ""));
    assert_eq!(pixel(&slide, 50, 25), ACCENT1);
    // bg1 maps to lt1 (white): a white rectangle on a red master background.
    let bg =
        r#"<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:bgPr></p:bg>"#;
    let fill = r#"<a:solidFill><a:schemeClr val="bg1"/></a:solidFill>"#;
    let slide = render(deck(&shape("rect", (0, 0, 50, 50), "", fill, ""), bg));
    assert_eq!(pixel(&slide, 25, 25), WHITE);
    assert_eq!(pixel(&slide, 75, 25), RED);
}

#[test]
fn a_color_modifier_changes_the_color() {
    // accent1 darkened to half its luminance.
    let fill = r#"<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="50000"/></a:schemeClr></a:solidFill>"#;
    let slide = render(deck(&shape("rect", (0, 0, 100, 50), "", fill, ""), ""));
    let [r, g, b] = pixel(&slide, 50, 25);
    assert!(r < 0x44 && g < 0x72 && b < 0xC4 && b > g, "{:?}", (r, g, b));
}

/// A right arrow's head reaches the shape's right edge at mid-height only.
#[test]
fn a_preset_shape_is_drawn_from_its_geometry() {
    let slide = render(deck(
        &shape("rightArrow", (0, 0, 100, 50), "", &solid("FF0000"), ""),
        "",
    ));
    assert_eq!(pixel(&slide, 97, 25), RED, "the tip");
    assert_eq!(pixel(&slide, 97, 3), WHITE, "beside the tip");
    assert_eq!(pixel(&slide, 10, 25), RED, "the shaft");
    assert_eq!(pixel(&slide, 10, 3), WHITE, "above the shaft");
}

#[test]
fn rotation_turns_a_shape_about_its_centre() {
    // A 60 × 10 bar centred at (50, 25), turned 90°: now 10 wide and 60 tall (clipped).
    let slide = render(deck(
        &shape(
            "rect",
            (20, 20, 60, 10),
            r#" rot="5400000""#,
            &solid("FF0000"),
            "",
        ),
        "",
    ));
    assert_eq!(pixel(&slide, 50, 5), RED);
    assert_eq!(pixel(&slide, 25, 25), WHITE);
}

#[test]
fn a_group_maps_its_children_into_its_place() {
    // The child space 0..1000 × 0..1000 (EMU) maps onto the group's 50 × 50 pt at (50, 0).
    let group = format!(
        r#"<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="G"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="{}" y="0"/><a:ext cx="{}" cy="{}"/><a:chOff x="0" y="0"/><a:chExt cx="1000" cy="1000"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="4" name="C"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="500" cy="1000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{}</p:spPr></p:sp></p:grpSp>"#,
        emu(50),
        emu(50),
        emu(50),
        solid("FF0000")
    );
    let slide = render(deck(&group, ""));
    assert_eq!(
        pixel(&slide, 60, 25),
        RED,
        "the child's left half of the group"
    );
    assert_eq!(pixel(&slide, 90, 25), WHITE);
    assert_eq!(pixel(&slide, 25, 25), WHITE);
}

#[test]
fn a_shape_without_its_own_fill_takes_its_style_color_and_theme_line_width() {
    let style = r#"<p:style><a:lnRef idx="2"><a:srgbClr val="00FF00"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>"#;
    let slide = render(deck(&shape("rect", (10, 10, 80, 30), "", "", style), ""));
    assert_eq!(pixel(&slide, 50, 25), ACCENT1);
    // lnRef idx 2 is the theme's 3 pt line: green across x 10 ± 1.5.
    assert_eq!(pixel(&slide, 10, 25), [0, 255, 0]);
}

#[test]
fn what_is_not_painted_is_counted() {
    let text = r#"<p:txBody><a:bodyPr/><a:p><a:r><a:t>Label</a:t></a:r><a:r><a:t>two</a:t></a:r></a:p></p:txBody>"#;
    let pic = r#"<p:pic><p:nvPicPr><p:cNvPr id="5" name="P"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId9"/></p:blipFill><p:spPr/></p:pic>"#;
    let chart = r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="Ch"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/></a:graphic></p:graphicFrame>"#;
    let custom = r#"<p:sp><p:nvSpPr><p:cNvPr id="7" name="X"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:custGeom/><a:solidFill><a:srgbClr val="000000"/></a:solidFill></p:spPr></p:sp>"#;
    let gradient = r#"<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst></a:gradFill>"#;
    let shapes = format!(
        "{}{pic}{chart}{custom}{}",
        shape("rect", (0, 0, 10, 10), "", &solid("FF0000"), text),
        shape("rect", (60, 0, 40, 50), "", gradient, "")
    );
    let slide = render(deck(&shapes, ""));
    let g = slide.gaps;
    assert_eq!(
        (
            g.text_runs,
            g.images,
            g.charts,
            g.shapes,
            g.approximated_fills
        ),
        (2, 1, 1, 1, 1)
    );
    assert_eq!(
        pixel(&slide, 80, 25),
        RED,
        "a gradient stands in as its first stop"
    );
}

#[test]
fn a_slide_index_out_of_range_is_an_error() {
    let parser = PptxParser::from_bytes(deck("", "")).unwrap();
    assert!(parser
        .render_slide(1, &SlideRasterOptions::default())
        .is_err());
}
