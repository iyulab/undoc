//! A WordprocessingML element is identified by its namespace, not by the prefix a writer
//! happened to bind to it.
//!
//! Python's `ElementTree` writes `ns0:` / `ns1:` unless a namespace is registered, and some
//! generators bind the main namespace as the default. The documents are valid, and other
//! readers open them; a reader that compares the written name (`w:p`) with the element it
//! wants sees no paragraph, no run and no table, and returns an empty document without an
//! error.
//!
//! Each test builds one document in the conventional spelling (the control) and again with
//! other prefixes, and requires identical Markdown. The document covers the parts that match
//! names: body text, character styles, a heading style from `styles.xml`, a list from
//! `numbering.xml`, a table, and a hyperlink through a relationship id.

use std::io::{Cursor, Write};

use undoc::parse_bytes;
use undoc::render::{to_markdown, RenderOptions};
use zip::write::SimpleFileOptions;

const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>"#;

const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;

const DOC_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/spec" TargetMode="External"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>"#;

/// Parts in the conventional spelling: `w:` for the main namespace, `r:` for relationships.
fn styles() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="{W_NS}">
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0000FF"/><w:u w:val="single"/></w:rPr></w:style>
</w:styles>"#
    )
}

fn numbering() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<w:numbering xmlns:w="{W_NS}">
<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="-"/></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>"#
    )
}

fn document() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"><w:body>
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Overview</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">Plain, </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r><w:r><w:t xml:space="preserve"> and a </w:t></w:r><w:hyperlink r:id="rId2"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>link</w:t></w:r></w:hyperlink><w:r><w:t>.</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>first item</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>second item</w:t></w:r></w:p>
<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
<w:tr><w:tc><w:p><w:r><w:t>key</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>value</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>1</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>
</w:body></w:document>"#
    )
}

/// Respell a part: the main namespace as `w_prefix` and the relationships namespace as
/// `r_prefix`. An empty `w_prefix` makes the main namespace the default for elements while
/// attributes keep a bound `w` prefix (an unprefixed attribute has no namespace).
fn respell(xml: &str, w_prefix: &str, r_prefix: &str) -> String {
    let mut out = xml.to_string();
    if w_prefix.is_empty() {
        out = out.replace("<w:", "<").replace("</w:", "</");
        // the root keeps `xmlns:w` for the attributes and gains the default declaration
        out = out.replacen(
            &format!(r#"xmlns:w="{W_NS}""#),
            &format!(r#"xmlns="{W_NS}" xmlns:w="{W_NS}""#),
            1,
        );
    } else {
        out = out
            .replace("<w:", &format!("<{w_prefix}:"))
            .replace("</w:", &format!("</{w_prefix}:"))
            .replace(" w:", &format!(" {w_prefix}:"))
            .replace("xmlns:w=", &format!("xmlns:{w_prefix}="));
    }
    out.replace(" r:id=", &format!(" {r_prefix}:id="))
        .replace("xmlns:r=", &format!("xmlns:{r_prefix}="))
}

fn build(w_prefix: Option<&str>, r_prefix: &str) -> Vec<u8> {
    let part = |xml: String| match w_prefix {
        Some(prefix) => respell(&xml, prefix, r_prefix),
        None => xml,
    };
    let document = part(document());
    let styles = part(styles());
    let numbering = part(numbering());

    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = SimpleFileOptions::default();
    for (name, data) in [
        ("[Content_Types].xml", CONTENT_TYPES.to_string()),
        ("_rels/.rels", ROOT_RELS.to_string()),
        ("word/_rels/document.xml.rels", DOC_RELS.to_string()),
        ("word/styles.xml", styles),
        ("word/numbering.xml", numbering),
        ("word/document.xml", document),
    ] {
        zip.start_file(name, opts).unwrap();
        zip.write_all(data.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn markdown(w_prefix: Option<&str>, r_prefix: &str) -> String {
    let doc = parse_bytes(&build(w_prefix, r_prefix)).expect("document parses");
    to_markdown(&doc, &RenderOptions::default()).expect("renders")
}

#[test]
fn the_control_document_reads_all_its_parts() {
    let md = markdown(None, "r");
    for expected in [
        "# Overview",
        "**bold**",
        "[link](https://example.com/spec)",
        "first item",
        "second item",
        "| key | value |",
    ] {
        assert!(
            md.contains(expected),
            "control is missing {expected:?}:\n{md}"
        );
    }
}

#[test]
fn an_elementtree_style_prefix_reads_like_the_conventional_one() {
    let control = markdown(None, "r");
    let prefixed = markdown(Some("ns0"), "ns1");
    assert!(
        !prefixed.trim().is_empty(),
        "prefixed document read as empty"
    );
    assert_eq!(prefixed, control);
}

#[test]
fn the_main_namespace_as_the_default_namespace_reads_like_the_conventional_one() {
    assert_eq!(markdown(Some(""), "r"), markdown(None, "r"));
}

#[test]
fn only_the_relationship_prefix_differing_still_resolves_the_link() {
    // `w:` stays, the hyperlink's id is `rel:id`.
    let md = markdown(Some("w"), "rel");
    assert!(md.contains("[link](https://example.com/spec)"), "{md}");
    assert_eq!(md, markdown(None, "r"));
}
