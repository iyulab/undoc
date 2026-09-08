//! A package whose XML carries bytes that are not valid UTF-8 must fail, and say so.
//!
//! `container.rs` unit-tests `decode_xml_bytes` on a byte literal, which proves the
//! decoder rejects. It cannot show what a *package* does — and that is the part worth
//! pinning, because the failure mode a wrong answer hides in is not an error but an
//! empty document that parsed "successfully".
//!
//! These tests state the policy this crate has always had, at the only layer that can
//! observe it: invalid UTF-8 in an XML part is an `ErrorKind::Encoding` failure of the
//! whole part, not a fragment silently replaced or dropped. The policy lives upstream of
//! the XML reader — the reader is handed a `&str`, so by the time it runs the bytes are
//! valid by construction. Anything that moves UTF-8 handling into or out of the reader
//! must leave these assertions true.

use std::io::{Cursor, Write};

use undoc::{parse_bytes, Block, ErrorKind};
use zip::write::SimpleFileOptions;

const CONTENT_TYPES: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;

const ROOT_RELS: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;

const DOC_HEAD: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:r><w:t>"#;

const DOC_TAIL: &[u8] = br#"</w:t></w:r></w:p></w:body></w:document>"#;

/// Where the byte that is not valid UTF-8 sits. The XML reader reaches these three
/// positions by different paths, so a change that only covers one of them is visible.
#[derive(Clone, Copy, Debug)]
enum BadByteIn {
    TextContent,
    AttributeValue,
    ElementName,
}

/// `0x80` is a UTF-8 continuation byte with nothing in front of it — invalid in any
/// position, and not part of any BOM or UTF-16 pattern the decoder probes for first.
const BAD: u8 = 0x80;

fn document_xml(where_: BadByteIn) -> Vec<u8> {
    let mut xml = Vec::new();
    match where_ {
        BadByteIn::TextContent => {
            xml.extend_from_slice(DOC_HEAD);
            xml.extend_from_slice(b"before");
            xml.push(BAD);
            xml.extend_from_slice(b"after");
            xml.extend_from_slice(DOC_TAIL);
        }
        BadByteIn::AttributeValue => {
            xml.extend_from_slice(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:pPr><w:pStyle w:val=""#,
            );
            xml.push(BAD);
            xml.extend_from_slice(
                br#""/></w:pPr><w:r><w:t>text</w:t></w:r></w:p></w:body></w:document>"#,
            );
        }
        BadByteIn::ElementName => {
            xml.extend_from_slice(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:r><w:"#,
            );
            xml.push(BAD);
            xml.extend_from_slice(br#"t>text</w:t></w:r></w:p></w:body></w:document>"#);
        }
    }
    xml
}

/// The same shape with every byte valid — including text outside ASCII, so the control
/// exercises multi-byte decoding rather than merely avoiding it.
fn valid_document_xml() -> Vec<u8> {
    let mut xml = Vec::new();
    xml.extend_from_slice(DOC_HEAD);
    xml.extend_from_slice("유효한 다바이트 본문".as_bytes());
    xml.extend_from_slice(DOC_TAIL);
    xml
}

fn docx(document_xml: &[u8]) -> Vec<u8> {
    let parts: [(&str, &[u8]); 3] = [
        ("[Content_Types].xml", CONTENT_TYPES),
        ("_rels/.rels", ROOT_RELS),
        ("word/document.xml", document_xml),
    ];

    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default();
    for (name, body) in parts {
        zip.start_file(name, options).expect("start zip entry");
        zip.write_all(body).expect("write zip entry");
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

/// The policy, stated once per position the bad byte can occupy.
///
/// The assertion that matters is not merely `is_err` — it is that the failure is
/// *classified*. `ErrorKind` is a public ABI contract, so a consumer branching on
/// `Encoding` keeps working only while this holds.
#[test]
fn invalid_utf8_in_a_part_fails_the_package_and_is_classified() {
    for where_ in [
        BadByteIn::TextContent,
        BadByteIn::AttributeValue,
        BadByteIn::ElementName,
    ] {
        let err = match parse_bytes(&docx(&document_xml(where_))) {
            Err(e) => e,
            Ok(doc) => panic!(
                "a package with an invalid UTF-8 byte in {where_:?} parsed instead of \
                 failing; it produced {:?}",
                paragraph_texts(&doc)
            ),
        };

        assert_eq!(
            err.kind(),
            ErrorKind::Encoding,
            "invalid UTF-8 in {where_:?} must be reported as an encoding failure, got: {err}"
        );
    }
}

/// The other half of the pin. Without it, "everything fails" would also pass.
#[test]
fn valid_multibyte_utf8_still_parses_with_its_text_intact() {
    let doc = parse_bytes(&docx(&valid_document_xml()))
        .expect("a package whose bytes are valid UTF-8 must parse");

    assert_eq!(paragraph_texts(&doc), vec!["유효한 다바이트 본문"]);
}

/// The decoder sits in the shared container, below the per-format parsers. Pinning a
/// second format keeps "the policy is shared" from being an assumption.
#[test]
fn invalid_utf8_fails_the_same_way_in_a_spreadsheet() {
    const XLSX_CONTENT_TYPES: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>"#;

    const XLSX_ROOT_RELS: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    const XLSX_WORKBOOK_RELS: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    const WORKBOOK_XML: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#;

    let mut sheet_xml = Vec::new();
    sheet_xml.extend_from_slice(
        br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>cell"#,
    );
    sheet_xml.push(BAD);
    sheet_xml.extend_from_slice(br#"</t></is></c></row></sheetData></worksheet>"#);

    let parts: [(&str, &[u8]); 5] = [
        ("[Content_Types].xml", XLSX_CONTENT_TYPES),
        ("_rels/.rels", XLSX_ROOT_RELS),
        ("xl/workbook.xml", WORKBOOK_XML),
        ("xl/_rels/workbook.xml.rels", XLSX_WORKBOOK_RELS),
        ("xl/worksheets/sheet1.xml", &sheet_xml),
    ];

    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default();
    for (name, body) in parts {
        zip.start_file(name, options).expect("start zip entry");
        zip.write_all(body).expect("write zip entry");
    }
    let bytes = zip.finish().expect("finish zip").into_inner();

    let err = match parse_bytes(&bytes) {
        Err(e) => e,
        Ok(doc) => panic!(
            "a spreadsheet with an invalid UTF-8 byte parsed instead of failing; it \
             produced {:?}",
            paragraph_texts(&doc)
        ),
    };

    assert_eq!(
        err.kind(),
        ErrorKind::Encoding,
        "the encoding policy must be the same for every format sharing the container, got: {err}"
    );
}
