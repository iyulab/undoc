//! A line break inside a paragraph separates the words around it in Markdown.
//!
//! Word's `<w:br/>` (Shift+Enter), its page break and PowerPoint's `<a:br/>` end a line; the
//! words on either side are two words. With `preserve_line_breaks` off the Markdown writer
//! used to write nothing for the break, so `Zero<w:br/>Charlie` came out as `ZeroCharlie`
//! while the plain-text writer kept a newline. The break is now at least the space a reflowed
//! paragraph has there, and still a hard break (two spaces and a newline) when asked for.

use std::io::{Cursor, Write};

use undoc::parse_bytes;
use undoc::render::{to_markdown, RenderOptions};
use zip::write::SimpleFileOptions;

const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const P_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn zip_of(files: &[(&str, String)]) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = SimpleFileOptions::default();
    for (name, data) in files {
        zip.start_file(*name, opts).unwrap();
        zip.write_all(data.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn rels(kind: &str, target: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}" Target="{target}"/></Relationships>"#
    )
}

/// A .docx of only the three parts a reader needs, holding `body`.
fn docx(body: &str) -> Vec<u8> {
    zip_of(&[
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"#
                .to_string(),
        ),
        ("_rels/.rels", rels("officeDocument", "word/document.xml")),
        ("word/_rels/document.xml.rels", rels("styles", "styles.xml")),
        (
            "word/styles.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="{W_NS}"><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style></w:styles>"#
            ),
        ),
        (
            "word/document.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="{W_NS}"><w:body>{body}</w:body></w:document>"#
            ),
        ),
    ])
}

/// A one-slide .pptx whose body shape holds `paragraphs`.
fn pptx(paragraphs: &str) -> Vec<u8> {
    zip_of(&[
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#
                .to_string(),
        ),
        ("_rels/.rels", rels("officeDocument", "ppt/presentation.xml")),
        (
            "ppt/_rels/presentation.xml.rels",
            rels("slide", "slides/slide1.xml"),
        ),
        (
            "ppt/presentation.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:p="{P_NS}" xmlns:r="{R_NS}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>"#
            ),
        ),
        (
            "ppt/slides/slide1.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:a="{A_NS}" xmlns:p="{P_NS}" xmlns:r="{R_NS}"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:txBody><a:bodyPr/>{paragraphs}</p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#
            ),
        ),
    ])
}

fn markdown(bytes: &[u8], options: &RenderOptions) -> String {
    to_markdown(&parse_bytes(bytes).unwrap(), options).unwrap()
}

fn word(paragraph_body: &str, options: &RenderOptions) -> String {
    markdown(&docx(&format!("<w:p>{paragraph_body}</w:p>")), options)
}

/// The placements of a break the report measured: inside a run, in a run of its own, at the
/// end of a run and at the start of the next.
const PLACEMENTS: [(&str, &str); 4] = [
    (
        "inside one run",
        "<w:r><w:t>Zero</w:t><w:br/><w:t>Charlie</w:t></w:r>",
    ),
    (
        "in its own run",
        "<w:r><w:t>Zero</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>Charlie</w:t></w:r>",
    ),
    (
        "at the end of a run",
        "<w:r><w:t>Zero</w:t><w:br/></w:r><w:r><w:t>Charlie</w:t></w:r>",
    ),
    (
        "at the start of a run",
        "<w:r><w:t>Zero</w:t></w:r><w:r><w:br/><w:t>Charlie</w:t></w:r>",
    ),
];

#[test]
fn a_word_line_break_keeps_the_words_on_either_side_apart() {
    for (placement, body) in PLACEMENTS {
        let md = word(body, &RenderOptions::default());
        assert_eq!(md.trim(), "Zero Charlie", "break {placement}: {md:?}");
    }
}

#[test]
fn a_word_page_break_keeps_the_words_on_either_side_apart() {
    let md = word(
        r#"<w:r><w:t>Zero</w:t><w:br w:type="page"/><w:t>Charlie</w:t></w:r>"#,
        &RenderOptions::default(),
    );
    assert_eq!(md.trim(), "Zero Charlie", "{md:?}");
}

#[test]
fn a_line_break_is_a_hard_break_when_preserved() {
    let options = RenderOptions::new().with_preserve_breaks(true);
    for (placement, body) in PLACEMENTS {
        let md = word(body, &options);
        assert_eq!(md.trim(), "Zero  \nCharlie", "break {placement}: {md:?}");
    }
}

#[test]
fn a_break_next_to_a_space_adds_no_second_one() {
    // A space on one side already parts the words; the break adds none.
    for body in [
        r#"<w:r><w:t xml:space="preserve">Zero </w:t><w:br/><w:t>Charlie</w:t></w:r>"#,
        r#"<w:r><w:t>Zero</w:t><w:br/><w:t xml:space="preserve"> Charlie</w:t></w:r>"#,
    ] {
        let md = word(body, &RenderOptions::default());
        assert_eq!(md.trim(), "Zero Charlie", "{md:?}");
    }
    // Spaces on both sides are the document's own, and are kept as written.
    let md = word(
        r#"<w:r><w:t xml:space="preserve">Zero </w:t><w:br/><w:t xml:space="preserve"> Charlie</w:t></w:r>"#,
        &RenderOptions::default(),
    );
    assert_eq!(md.trim(), "Zero  Charlie", "{md:?}");
}

#[test]
fn a_break_that_ends_the_paragraph_leaves_nothing_behind() {
    let md = word(
        "<w:r><w:t>Zero</w:t><w:br/></w:r>",
        &RenderOptions::default(),
    );
    assert_eq!(md.trim_end_matches('\n'), "Zero", "{md:?}");
}

#[test]
fn emphasis_ends_before_the_break_and_the_words_stay_apart() {
    let md = word(
        "<w:r><w:rPr><w:b/></w:rPr><w:t>Zero.</w:t><w:br/></w:r><w:r><w:t>Charlie</w:t></w:r>",
        &RenderOptions::default(),
    );
    assert_eq!(md.trim(), "**Zero.** Charlie", "{md:?}");
}

#[test]
fn a_break_in_a_table_cell_keeps_the_words_apart() {
    let table = r#"<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Address</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>Street</w:t><w:br/><w:t>City</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#;
    let md = markdown(&docx(table), &RenderOptions::default());
    assert!(md.contains("| Street City |"), "{md:?}");
    let md = markdown(
        &docx(table),
        &RenderOptions::new().with_preserve_breaks(true),
    );
    assert!(md.contains("| Street<br>City |"), "{md:?}");
}

#[test]
fn a_powerpoint_line_break_keeps_the_words_on_either_side_apart() {
    let slide = pptx("<a:p><a:r><a:t>Zero</a:t></a:r><a:br/><a:r><a:t>Bravo</a:t></a:r></a:p>");
    let md = markdown(&slide, &RenderOptions::default());
    assert!(md.contains("Zero Bravo"), "{md:?}");
    assert!(!md.contains("ZeroBravo"), "{md:?}");
    let md = markdown(&slide, &RenderOptions::new().with_preserve_breaks(true));
    assert!(md.contains("Zero  \nBravo"), "{md:?}");
}

#[test]
fn an_emitted_page_break_in_a_paragraph_is_a_thematic_break() {
    let options = RenderOptions::new().with_emit_page_breaks(true);
    let md = word(
        r#"<w:r><w:t>Zero</w:t><w:br w:type="page"/><w:t>Charlie</w:t></w:r>"#,
        &options,
    );
    // The thematic break is structure: it was written as `\---` when the whole paragraph was
    // escaped as text.
    assert_eq!(md.trim(), "Zero\n\n---\n\nCharlie", "{md:?}");
}

#[test]
fn the_text_after_a_page_break_opens_a_block_and_is_escaped_as_one() {
    let options = RenderOptions::new().with_emit_page_breaks(true);
    let md = word(
        r#"<w:r><w:t>Zero</w:t><w:br w:type="page"/><w:t xml:space="preserve">1. Charlie</w:t></w:r>"#,
        &options,
    );
    assert_eq!(md.trim(), "Zero\n\n---\n\n1\\. Charlie", "{md:?}");
}

#[test]
fn a_page_break_in_a_table_cell_is_a_space_when_emitted() {
    let options = RenderOptions::new().with_emit_page_breaks(true);
    let table = r#"<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Notes</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>Zero</w:t><w:br w:type="page"/><w:t>Charlie</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#;
    let md = markdown(&docx(table), &options);
    assert!(md.contains("| Zero Charlie |"), "{md:?}");
}

#[test]
fn a_break_in_a_heading_leaves_the_heading_on_one_line() {
    let body = r#"<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Zero</w:t><w:br/><w:t>Charlie</w:t></w:r></w:p>"#;
    for options in [
        RenderOptions::default(),
        RenderOptions::new().with_preserve_breaks(true),
    ] {
        let md = markdown(&docx(body), &options);
        assert_eq!(md.trim(), "# Zero Charlie", "{md:?}");
    }
}
