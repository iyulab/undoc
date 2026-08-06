//! A spreadsheet's merged header must survive the whole path, not just the renderer.
//!
//! `merged_header_table_test.rs` builds tables through the model, so it proves the
//! renderer places merges correctly — but it cannot see whether the XLSX parser recorded
//! the merge in the first place. Those are the two halves of the same guarantee, and a
//! break in either produces the same user-visible symptom: labels under the wrong
//! columns, with no error.
//!
//! The workbook is assembled in memory. `test-files/` is gitignored, so a fixture-based
//! test skips silently in CI — which is how the parser's own merge test has been passing
//! without running.

use std::io::{Cursor, Write};

use undoc::render::{to_markdown, RenderOptions};
use undoc::{parse_bytes, Block, Table};
use zip::write::SimpleFileOptions;

/// Minimal XLSX: one sheet, inline strings, and whatever `<mergeCells>` is given.
fn workbook(sheet_rows: &str, merges: &str) -> Vec<u8> {
    let sheet = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>{sheet_rows}</sheetData>{merges}</worksheet>"#
    );

    let files: Vec<(&str, String)> = vec![
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"#
                .to_string(),
        ),
        (
            "_rels/.rels",
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#
                .to_string(),
        ),
        (
            "xl/workbook.xml",
            r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#
                .to_string(),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#
                .to_string(),
        ),
        ("xl/worksheets/sheet1.xml", sheet),
    ];

    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default();
    for (name, body) in files {
        zip.start_file(name, options).expect("start zip entry");
        zip.write_all(body.as_bytes()).expect("write zip entry");
    }
    zip.finish().expect("finish zip").into_inner()
}

/// `<c r="A1" t="inlineStr"><is><t>…</t></is></c>` for each value, skipping empties.
fn row(index: u32, values: &[&str]) -> String {
    let cells: String = values
        .iter()
        .enumerate()
        .filter(|(_, v)| !v.is_empty())
        .map(|(i, v)| {
            let column = (b'A' + i as u8) as char;
            format!(r#"<c r="{column}{index}" t="inlineStr"><is><t>{v}</t></is></c>"#)
        })
        .collect();
    format!(r#"<row r="{index}">{cells}</row>"#)
}

/// Row 1 holds two group labels, each merged across two columns; row 2 the real labels;
/// row 3 the data. Exactly the shape the field reports describe.
fn grouped_header_workbook() -> Vec<u8> {
    let rows = format!(
        "{}{}{}",
        row(1, &["Group A", "", "Group B", ""]),
        row(2, &["a1", "a2", "b1", "b2"]),
        row(3, &["1", "2", "3", "4"]),
    );
    let merges = r#"<mergeCells count="2">
<mergeCell ref="A1:B1"/><mergeCell ref="C1:D1"/></mergeCells>"#;
    workbook(&rows, merges)
}

fn only_table(doc: &undoc::Document) -> &Table {
    doc.sections
        .iter()
        .flat_map(|s| &s.content)
        .find_map(|b| match b {
            Block::Table(t) => Some(t),
            _ => None,
        })
        .expect("the sheet must parse into a table")
}

#[test]
fn parser_records_the_merge_as_a_span() {
    let doc = parse_bytes(&grouped_header_workbook()).expect("workbook must parse");
    let table = only_table(&doc);

    let header = table.rows.first().expect("a header row must exist");
    assert_eq!(
        header.cells.len(),
        2,
        "two merged labels stay two cells — the span is on the cell, not expanded rows"
    );
    assert!(
        header.cells.iter().all(|c| c.col_span == 2),
        "each label must carry its span, got {:?}",
        header.cells.iter().map(|c| c.col_span).collect::<Vec<_>>()
    );
}

#[test]
fn merged_labels_reach_markdown_in_their_own_columns() {
    let doc = parse_bytes(&grouped_header_workbook()).expect("workbook must parse");
    let md = to_markdown(&doc, &RenderOptions::default()).expect("render must succeed");

    let header = md
        .lines()
        .find(|l| l.contains("Group A"))
        .expect("the header row must survive to the output");
    let cells: Vec<&str> = header.trim_matches('|').split('|').map(str::trim).collect();

    assert_eq!(
        cells,
        vec!["Group A", "", "Group B", ""],
        "each label must sit at its merge's first column, got {header:?}"
    );
    assert!(
        !header.contains('#'),
        "no cell held a '#'; none may appear, got {header:?}"
    );
}

/// The row under the header is where a misplaced label shows up as wrong data: if the
/// header shifted, these values would no longer line up beneath their own group.
#[test]
fn data_rows_stay_under_their_own_columns() {
    let doc = parse_bytes(&grouped_header_workbook()).expect("workbook must parse");
    let md = to_markdown(&doc, &RenderOptions::default()).expect("render must succeed");

    let widths: Vec<usize> = md
        .lines()
        .filter(|l| l.starts_with('|'))
        .map(|l| l.matches('|').count())
        .collect();
    assert!(
        widths.windows(2).all(|w| w[0] == w[1]),
        "every row must have the same column count, got {widths:?}"
    );

    let labels = md.lines().find(|l| l.contains("a1")).expect("labels row");
    let cells: Vec<&str> = labels.trim_matches('|').split('|').map(str::trim).collect();
    assert_eq!(cells, vec!["a1", "a2", "b1", "b2"]);
}

/// The control: the same sheet without merges must be unaffected by any of this.
#[test]
fn workbook_without_merges_is_unaffected() {
    let rows = format!("{}{}", row(1, &["A", "B", "C"]), row(2, &["1", "2", "3"]),);
    let doc = parse_bytes(&workbook(&rows, "")).expect("workbook must parse");
    let md = to_markdown(&doc, &RenderOptions::default()).expect("render must succeed");

    assert!(
        md.contains("| A | B | C |"),
        "a plain sheet must render plainly, got {md:?}"
    );
    // Only the table rows: `## Sheet1` is the sheet-name heading and its `#` is real.
    let invented: Vec<&str> = md
        .lines()
        .filter(|l| l.starts_with('|') && l.contains('#'))
        .collect();
    assert!(
        invented.is_empty(),
        "no table row may contain an invented '#', got {invented:?}"
    );
}
