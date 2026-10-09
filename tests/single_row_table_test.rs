//! A table with a single row keeps its content: Markdown renders it as a one-row table (its
//! row as the header, nothing below the delimiter), plain text renders its cells, and the
//! model keeps the row.

use std::io::{Cursor, Write};

use undoc::parse_bytes;
use undoc::render::{to_markdown, to_text, RenderOptions};
use zip::write::SimpleFileOptions;

const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;

const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;

fn docx(body: &str) -> Vec<u8> {
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}</w:body></w:document>"#
    );
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = SimpleFileOptions::default();
    for (name, data) in [
        ("[Content_Types].xml", CONTENT_TYPES),
        ("_rels/.rels", ROOT_RELS),
        ("word/document.xml", document.as_str()),
    ] {
        zip.start_file(name, opts).unwrap();
        zip.write_all(data.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn cell(text: &str) -> String {
    format!(r#"<w:tc><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>"#)
}

fn one_row_table() -> String {
    format!(
        r#"<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:tbl><w:tr>{}{}</w:tr></w:tbl><w:p><w:r><w:t>After</w:t></w:r></w:p>"#,
        cell("alpha"),
        cell("beta")
    )
}

#[test]
fn a_one_row_table_keeps_its_cells_in_markdown() {
    let md = to_markdown(
        &parse_bytes(&docx(&one_row_table())).unwrap(),
        &RenderOptions::default(),
    )
    .unwrap();
    assert!(md.contains("| alpha | beta |"), "{md}");
    assert!(md.contains("Before") && md.contains("After"), "{md}");
}

#[test]
fn a_one_row_table_keeps_its_cells_in_text() {
    let text = to_text(
        &parse_bytes(&docx(&one_row_table())).unwrap(),
        &RenderOptions::default(),
    )
    .unwrap();
    assert!(text.contains("alpha") && text.contains("beta"), "{text}");
}
