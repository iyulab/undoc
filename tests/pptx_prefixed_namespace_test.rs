//! A presentation written with other namespace prefixes reads like the conventional one.
//!
//! The PPTX readers match local names, so they were prefix-independent before the
//! prefix-normalising read; this pins that, together with the relationship files, for a
//! deck whose presentation, slide and relationship parts all use unconventional prefixes.

use std::io::{Cursor, Write};

use undoc::parse_bytes;
use undoc::render::{to_markdown, RenderOptions};
use zip::write::SimpleFileOptions;

const P_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn presentation() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="{P_NS}" xmlns:r="{R_NS}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>"#
    )
}

fn slide() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="{A_NS}" xmlns:p="{P_NS}" xmlns:r="{R_NS}"><p:cSld><p:spTree>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>Quarterly review</a:t></a:r></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>Revenue grew</a:t></a:r></a:p></p:txBody></p:sp>
</p:spTree></p:cSld></p:sld>"#
    )
}

fn respell(xml: &str) -> String {
    xml.replace("<p:", "<ns0:")
        .replace("</p:", "</ns0:")
        .replace("xmlns:p=", "xmlns:ns0=")
        .replace("<a:", "<ns1:")
        .replace("</a:", "</ns1:")
        .replace("xmlns:a=", "xmlns:ns1=")
        .replace(" r:id=", " ns2:id=")
        .replace("xmlns:r=", "xmlns:ns2=")
}

fn rels(prefix: &str, items: &[(&str, &str, &str)]) -> String {
    let (root, tag, decl) = if prefix.is_empty() {
        (
            "Relationships".to_string(),
            "Relationship".to_string(),
            r#"xmlns="http://schemas.openxmlformats.org/package/2006/relationships""#.to_string(),
        )
    } else {
        (
            format!("{prefix}:Relationships"),
            format!("{prefix}:Relationship"),
            format!(
                r#"xmlns:{prefix}="http://schemas.openxmlformats.org/package/2006/relationships""#
            ),
        )
    };
    let body: String = items
        .iter()
        .map(|(id, kind, target)| format!(r#"<{tag} Id="{id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}" Target="{target}"/>"#))
        .collect();
    format!(r#"<?xml version="1.0" encoding="UTF-8"?><{root} {decl}>{body}</{root}>"#)
}

fn build(prefixed: bool) -> Vec<u8> {
    let part = |xml: String| if prefixed { respell(&xml) } else { xml };
    let rp = if prefixed { "pr" } else { "" };
    let files: Vec<(&str, String)> = vec![
        ("[Content_Types].xml", r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>"#.to_string()),
        ("_rels/.rels", rels(rp, &[("rId1", "officeDocument", "ppt/presentation.xml")])),
        ("ppt/_rels/presentation.xml.rels", rels(rp, &[("rId1", "slide", "slides/slide1.xml")])),
        ("ppt/slides/_rels/slide1.xml.rels", rels(rp, &[])),
        ("ppt/presentation.xml", part(presentation())),
        ("ppt/slides/slide1.xml", part(slide())),
    ];
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = SimpleFileOptions::default();
    for (name, body) in files {
        zip.start_file(name, opts).unwrap();
        zip.write_all(body.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn markdown(prefixed: bool) -> String {
    let doc = parse_bytes(&build(prefixed)).expect("deck parses");
    to_markdown(&doc, &RenderOptions::default()).expect("renders")
}

#[test]
fn the_control_deck_reads_its_text() {
    let md = markdown(false);
    assert!(
        md.contains("Quarterly review") && md.contains("Revenue grew"),
        "{md}"
    );
}

#[test]
fn a_deck_with_unconventional_prefixes_reads_like_the_conventional_one() {
    assert_eq!(markdown(true), markdown(false));
}
