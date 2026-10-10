//! `undoc tables` writes each table of a document as CSV: merged cells on their grid, a
//! table nested in a cell as a table of its own, right after the one that holds it.
//!
//! The DOCX is built here rather than read from a fixture directory, so the test runs
//! everywhere instead of quietly skipping when a fixture is absent.

use std::io::{Cursor, Write};
use std::process::Command;

fn bin() -> &'static str {
    env!("CARGO_BIN_EXE_undoc")
}

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

// ┌────────┬───────────────┐
// │ Region │ Sales, total  │   «Region» merged down, «Sales, total» across two columns
// │        ├───────┬───────┤
// │        │ 2024  │ 2025  │
// ├────────┼───────┼───────┤
// │ North* │  10   │  12   │   * the cell also holds a one-cell table, «inner»
// └────────┴───────┴───────┘
const DOCUMENT_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Before the table.</w:t></w:r></w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>Region</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Sales, total</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>
        <w:tc><w:p><w:r><w:t>2024</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>2025</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc>
          <w:tbl><w:tr><w:tc><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
          <w:p><w:r><w:t>North</w:t></w:r></w:p>
        </w:tc>
        <w:tc><w:p><w:r><w:t>10</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>12</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>After the table.</w:t></w:r></w:p>
  </w:body>
</w:document>"#;

const OUTER: &str = "Region,\"Sales, total\",\r\n,2024,2025\r\nNorth,10,12\r\n";
const INNER: &str = "inner\r\n";

fn docx_with_merged_and_nested_tables() -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, body) in [
        ("[Content_Types].xml", CONTENT_TYPES),
        ("_rels/.rels", ROOT_RELS),
        ("word/document.xml", DOCUMENT_XML),
    ] {
        zip.start_file(name, options).unwrap();
        zip.write_all(body.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn input(dir: &std::path::Path) -> std::path::PathBuf {
    let path = dir.join("tables.docx");
    std::fs::write(&path, docx_with_merged_and_nested_tables()).unwrap();
    path
}

#[test]
fn tables_to_stdout_are_csv_a_blank_line_apart() {
    let tmp = tempfile::tempdir().unwrap();
    let out = Command::new(bin())
        .arg("tables")
        .arg(input(tmp.path()))
        .output()
        .unwrap();
    assert!(
        out.status.success(),
        "{}",
        String::from_utf8_lossy(&out.stderr)
    );
    assert_eq!(
        String::from_utf8(out.stdout).unwrap(),
        format!("{OUTER}\r\n{INNER}")
    );
}

#[test]
fn tables_to_a_directory_are_one_file_each_named_for_their_place() {
    let tmp = tempfile::tempdir().unwrap();
    let dir = tmp.path().join("out");
    let status = Command::new(bin())
        .arg("tables")
        .arg(input(tmp.path()))
        .arg("-o")
        .arg(&dir)
        .status()
        .unwrap();
    assert!(status.success());

    let mut names: Vec<String> = std::fs::read_dir(&dir)
        .unwrap()
        .map(|e| e.unwrap().file_name().to_string_lossy().into_owned())
        .collect();
    names.sort();
    assert_eq!(names, ["s1-t1.csv", "s1-t2.csv"]);
    assert_eq!(
        std::fs::read_to_string(dir.join("s1-t1.csv")).unwrap(),
        OUTER
    );
    assert_eq!(
        std::fs::read_to_string(dir.join("s1-t2.csv")).unwrap(),
        INNER
    );
}

#[test]
fn tsv_separates_fields_with_tabs() {
    let tmp = tempfile::tempdir().unwrap();
    let out = Command::new(bin())
        .args(["tables", "--tsv"])
        .arg(input(tmp.path()))
        .output()
        .unwrap();
    assert!(out.status.success());
    let text = String::from_utf8(out.stdout).unwrap();
    assert!(
        text.starts_with("Region\tSales, total\t\r\n\t2024\t2025\r\n"),
        "{text:?}"
    );
}
