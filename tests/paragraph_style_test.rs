//! A paragraph's style (`w:pStyle`) carries run formatting too: a Quote style in italic makes
//! its text italic. Word resolves a run's formatting as the paragraph style's run properties,
//! then the run's character style, then the run's own properties — each overriding the last.
//! Checked in body text and in a table cell — the two places the reader parses runs.

use std::io::{Cursor, Write};

use undoc::parse_bytes;
use undoc::render::{to_markdown, RenderOptions};
use zip::write::SimpleFileOptions;

const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"#;

const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;

const DOC_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#;

const STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:rPr><w:i/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="LeadIn"><w:name w:val="Lead In"/><w:basedOn w:val="Quote"/><w:rPr><w:b/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>
<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>
</w:styles>"#;

fn para(style: &str, runs: &str) -> String {
    format!(r#"<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr>{runs}</w:p>"#)
}

fn run(text: &str, rpr: &str) -> String {
    format!(r#"<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>"#)
}

fn markdown(body: &str) -> String {
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}</w:body></w:document>"#
    );
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = SimpleFileOptions::default();
    for (name, data) in [
        ("[Content_Types].xml", CONTENT_TYPES),
        ("_rels/.rels", ROOT_RELS),
        ("word/_rels/document.xml.rels", DOC_RELS),
        ("word/styles.xml", STYLES),
        ("word/document.xml", document.as_str()),
    ] {
        zip.start_file(name, opts).unwrap();
        zip.write_all(data.as_bytes()).unwrap();
    }
    let bytes = zip.finish().unwrap().into_inner();
    to_markdown(&parse_bytes(&bytes).unwrap(), &RenderOptions::default()).unwrap()
}

#[test]
fn a_paragraph_style_formats_its_runs() {
    let md = markdown(&para("Quote", &run("quoted words", "")));
    assert!(md.contains("*quoted words*"), "{md}");
}

#[test]
fn a_runs_own_property_overrides_its_paragraph_style() {
    let runs = run("slanted", "") + &run(" upright", r#"<w:i w:val="0"/>"#);
    let md = markdown(&para("Quote", &runs));
    assert!(md.contains("*slanted* upright"), "{md}");
}

#[test]
fn a_paragraph_style_inherits_through_based_on_and_a_character_style_adds_to_it() {
    // LeadIn is bold and, through Quote, italic.
    let md = markdown(&para("LeadIn", &run("both", "")));
    assert!(md.contains("***both***"), "{md}");
    // Normal has no run formatting: a character style alone decides.
    let md = markdown(&para(
        "Normal",
        &run("strong", r#"<w:rStyle w:val="Strong"/>"#),
    ));
    assert!(md.contains("**strong**"), "{md}");
}

#[test]
fn a_cell_paragraph_style_formats_its_runs_too() {
    let cell = |inner: &str| format!("<w:tc>{inner}</w:tc>");
    let table = format!(
        "<w:tbl><w:tr>{}{}</w:tr><w:tr>{}{}</w:tr></w:tbl>",
        cell(&para("Normal", &run("Name", ""))),
        cell(&para("Normal", &run("Note", ""))),
        cell(&para("Quote", &run("quoted", ""))),
        cell(&para("Normal", &run("plain", ""))),
    );
    let md = markdown(&table);
    assert!(md.contains("| *quoted* | plain |"), "{md}");
}

#[test]
fn a_bold_heading_style_does_not_bold_its_heading_text() {
    let md = markdown(&para("Heading1", &run("Introduction", "")));
    assert!(md.contains("# Introduction"), "{md}");
    assert!(!md.contains("**"), "{md}");
}
