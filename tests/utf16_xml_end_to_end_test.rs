//! A UTF-16 encoded package must survive the whole path, not just the decoder.
//!
//! `container.rs` unit-tests `decode_xml_bytes` on byte literals, which proves the
//! decoder converts. It cannot show that a package whose parts are all UTF-16 parses
//! into a document with its text intact — and the tests that would have shown it read a
//! fixture from `test-files/`, which is gitignored, so they have never run.
//!
//! XML 1.0 requires UTF-16 entities to carry a BOM, so the BOM-bearing cases are the
//! conformant ones. The BOM-less case is pinned too: it is the one where a wrong answer
//! looks like an empty document rather than an error.

use std::io::{Cursor, Write};

use undoc::{parse_bytes, Block};
use zip::write::SimpleFileOptions;

const DOCUMENT_XML: &str = r#"<?xml version="1.0" encoding="UTF-16" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:r><w:t>Encoded in sixteen bits</w:t></w:r></w:p>
<w:p><w:r><w:t>한글도 살아남아야 한다</w:t></w:r></w:p></w:body></w:document>"#;

const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-16"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;

const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-16"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;

/// How a part is written. A byte-order mark is present unless the variant says otherwise.
#[derive(Clone, Copy)]
enum Encoding {
    Le,
    Be,
    /// No BOM — non-conformant per XML 1.0, but the shape a wrong answer hides in.
    LeUnmarked,
}

fn encode(text: &str, encoding: Encoding) -> Vec<u8> {
    let units: Vec<u16> = text.encode_utf16().collect();
    let mut out = Vec::with_capacity(units.len() * 2 + 2);
    match encoding {
        Encoding::Le => {
            out.extend_from_slice(&[0xFF, 0xFE]);
            units
                .iter()
                .for_each(|u| out.extend_from_slice(&u.to_le_bytes()));
        }
        Encoding::Be => {
            out.extend_from_slice(&[0xFE, 0xFF]);
            units
                .iter()
                .for_each(|u| out.extend_from_slice(&u.to_be_bytes()));
        }
        Encoding::LeUnmarked => {
            units
                .iter()
                .for_each(|u| out.extend_from_slice(&u.to_le_bytes()));
        }
    }
    out
}

/// The same document with nothing outside ASCII. Encoded as UTF-16 its every other byte
/// is NUL — and NUL is valid UTF-8, so this is the variant that can decode "successfully"
/// into nonsense while the non-ASCII one above fails loudly and gets a second chance.
const ASCII_DOCUMENT_XML: &str = r#"<?xml version="1.0" encoding="UTF-16" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:r><w:t>Encoded in sixteen bits</w:t></w:r></w:p></w:body></w:document>"#;

/// Every part of the package is encoded — a real UTF-16 document is not mixed.
fn docx(encoding: Encoding) -> Vec<u8> {
    docx_with_document(encoding, DOCUMENT_XML)
}

fn docx_with_document(encoding: Encoding, document_xml: &str) -> Vec<u8> {
    let parts: [(&str, &str); 3] = [
        ("[Content_Types].xml", CONTENT_TYPES),
        ("_rels/.rels", ROOT_RELS),
        ("word/document.xml", document_xml),
    ];

    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default();
    for (name, body) in parts {
        zip.start_file(name, options).expect("start zip entry");
        zip.write_all(&encode(body, encoding))
            .expect("write zip entry");
    }
    zip.finish().expect("finish zip").into_inner()
}

/// The decoder sits in the shared container, below the per-format parsers, so a
/// spreadsheet exercises the same path a document does. Pinning one second format keeps
/// that from being an assumption.
const SHEET_XML: &str = r#"<?xml version="1.0" encoding="UTF-16"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Sixteen</t></is></c>
<c r="B1" t="inlineStr"><is><t>비트</t></is></c></row></sheetData></worksheet>"#;

const WORKBOOK_XML: &str = r#"<?xml version="1.0" encoding="UTF-16"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#;

const XLSX_CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-16"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>"#;

const XLSX_ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-16"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

const XLSX_WORKBOOK_RELS: &str = r#"<?xml version="1.0" encoding="UTF-16"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

fn xlsx(encoding: Encoding) -> Vec<u8> {
    let parts: [(&str, &str); 5] = [
        ("[Content_Types].xml", XLSX_CONTENT_TYPES),
        ("_rels/.rels", XLSX_ROOT_RELS),
        ("xl/workbook.xml", WORKBOOK_XML),
        ("xl/_rels/workbook.xml.rels", XLSX_WORKBOOK_RELS),
        ("xl/worksheets/sheet1.xml", SHEET_XML),
    ];

    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default();
    for (name, body) in parts {
        zip.start_file(name, options).expect("start zip entry");
        zip.write_all(&encode(body, encoding))
            .expect("write zip entry");
    }
    zip.finish().expect("finish zip").into_inner()
}

fn paragraph_texts(doc: &undoc::Document) -> Vec<String> {
    doc.sections
        .iter()
        .flat_map(|s| &s.content)
        .filter_map(|b| match b {
            Block::Paragraph(p) => Some(p.runs.iter().map(|r| r.text.as_str()).collect::<String>()),
            _ => None,
        })
        .collect()
}

#[test]
fn utf16_le_with_bom_parses_with_its_text_intact() {
    let doc = parse_bytes(&docx(Encoding::Le)).expect("a UTF-16 LE package must parse");
    assert_eq!(
        paragraph_texts(&doc),
        vec!["Encoded in sixteen bits", "한글도 살아남아야 한다"],
    );
}

#[test]
fn utf16_be_with_bom_parses_with_its_text_intact() {
    let doc = parse_bytes(&docx(Encoding::Be)).expect("a UTF-16 BE package must parse");
    assert_eq!(
        paragraph_texts(&doc),
        vec!["Encoded in sixteen bits", "한글도 살아남아야 한다"],
    );
}

/// Non-ASCII text makes the UTF-8 attempt fail outright, which is what hands the bytes
/// to the UTF-16 fallback. This case therefore says nothing about ASCII-only content —
/// see the next test for that.
#[test]
fn utf16_le_without_bom_recovers_non_ascii_text() {
    match parse_bytes(&docx(Encoding::LeUnmarked)) {
        Ok(doc) => assert_eq!(
            paragraph_texts(&doc),
            vec!["Encoded in sixteen bits", "한글도 살아남아야 한다"],
            "a BOM-less UTF-16 package parsed, so its text must be there"
        ),
        Err(e) => panic!("a BOM-less UTF-16 package failed to parse: {e}"),
    }
}

/// The discriminating case: ASCII-only UTF-16 is a valid UTF-8 byte sequence, so nothing
/// forces a second look at it. A document must not come back empty and successful.
#[test]
fn utf16_le_without_bom_does_not_silently_lose_ascii_text() {
    match parse_bytes(&docx_with_document(
        Encoding::LeUnmarked,
        ASCII_DOCUMENT_XML,
    )) {
        Ok(doc) => assert_eq!(
            paragraph_texts(&doc),
            vec!["Encoded in sixteen bits"],
            "the package parsed, so its text must be there rather than silently dropped"
        ),
        Err(e) => panic!("a BOM-less UTF-16 package failed to parse: {e}"),
    }
}

/// A spreadsheet takes the same decoder, and its cells must arrive with their text.
#[test]
fn a_utf16_spreadsheet_reaches_markdown_with_its_cells() {
    let doc = parse_bytes(&xlsx(Encoding::Le)).expect("a UTF-16 workbook must parse");
    let md = undoc::render::to_markdown(&doc, &undoc::render::RenderOptions::default())
        .expect("render must succeed");

    assert!(
        md.contains("Sixteen") && md.contains("비트"),
        "both cells must survive the decode, got {md:?}"
    );
}
