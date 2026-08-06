//! A table with merged header cells must keep its columns aligned.
//!
//! Spreadsheets and documents routinely put a row of merged group labels above the real
//! column headings. Markdown and the plain-text table cannot express a merge, so the
//! merged cell has to be anchored at the column it starts in and the columns it covers
//! filled with empty ones. Get that wrong and the failure is silent: the output is still
//! a well-formed table, just one where every value sits under the wrong heading.
//!
//! Tables are built through the public model rather than read from a fixture file so the
//! test states exactly which merge it is about, and runs in CI where `test-files/` is
//! absent.

use undoc::render::{to_markdown, to_text, RenderOptions, TableFallback};
use undoc::{Block, Cell, Document, Row, Section, Table};

fn spanning(text: &str, col_span: u32, row_span: u32) -> Cell {
    Cell {
        col_span,
        row_span,
        ..Cell::with_text(text)
    }
}

fn row(cells: Vec<Cell>) -> Row {
    Row {
        cells,
        is_header: false,
        height: None,
    }
}

fn document_with(table: Table) -> Document {
    let mut doc = Document::new();
    doc.sections.push(Section {
        index: 0,
        name: Some("Sheet1".to_string()),
        content: vec![Block::Table(table)],
        ..Default::default()
    });
    doc
}

fn markdown_of(table: Table) -> String {
    to_markdown(&document_with(table), &RenderOptions::default()).expect("render must succeed")
}

/// Six columns: three merged group labels of two columns each, then the real labels,
/// then a data row. This is the shape the reports describe, shrunk to what a reader
/// can check by eye.
fn grouped_header_table() -> Table {
    let mut table = Table::new();
    table.add_row(row(vec![
        spanning("Group A", 2, 1),
        spanning("Group B", 2, 1),
        spanning("Group C", 2, 1),
    ]));
    table.add_row(row(vec![
        Cell::with_text("a1"),
        Cell::with_text("a2"),
        Cell::with_text("b1"),
        Cell::with_text("b2"),
        Cell::with_text("c1"),
        Cell::with_text("c2"),
    ]));
    table.add_row(row(vec![
        Cell::with_text("1"),
        Cell::with_text("2"),
        Cell::with_text("3"),
        Cell::with_text("4"),
        Cell::with_text("5"),
        Cell::with_text("6"),
    ]));
    table
}

#[test]
fn merged_group_labels_stay_in_their_own_columns() {
    let md = markdown_of(grouped_header_table());
    let header = md
        .lines()
        .find(|l| l.contains("Group A"))
        .expect("the header row must survive");

    assert!(
        header.starts_with("| Group A |"),
        "the first group label must open the row, not be pushed right, got {header:?}"
    );
    let cells: Vec<&str> = header.trim_matches('|').split('|').map(str::trim).collect();
    assert_eq!(
        cells,
        vec!["Group A", "", "Group B", "", "Group C", ""],
        "each label must sit at its merge's first column"
    );
}

#[test]
fn no_character_is_invented_that_the_input_did_not_contain() {
    let md = markdown_of(grouped_header_table());
    let header = md.lines().find(|l| l.contains("Group A")).unwrap();

    assert!(
        !header.contains('#'),
        "no cell held a '#'; the renderer must not invent one, got {header:?}"
    );
}

#[test]
fn every_row_has_the_same_column_count() {
    let md = markdown_of(grouped_header_table());
    let widths: Vec<usize> = md
        .lines()
        .filter(|l| l.starts_with('|'))
        .map(|l| l.matches('|').count())
        .collect();

    assert!(
        widths.windows(2).all(|w| w[0] == w[1]),
        "rows disagree on column count: {widths:?}"
    );
}

/// A vertically merged cell holds its column in the rows below, so the cells of those
/// rows have to start one column further right.
#[test]
fn vertically_merged_cell_holds_its_column() {
    let mut table = Table::new();
    table.add_row(row(vec![
        spanning("Side", 1, 2),
        Cell::with_text("X"),
        Cell::with_text("Y"),
    ]));
    table.add_row(row(vec![Cell::with_text("x"), Cell::with_text("y")]));

    let md = markdown_of(table);
    let second = md
        .lines()
        .filter(|l| l.starts_with('|'))
        .nth(2)
        .expect("a second data row must exist");
    let cells: Vec<&str> = second.trim_matches('|').split('|').map(str::trim).collect();

    assert_eq!(
        cells,
        vec!["", "x", "y"],
        "row 2 must start in column 2 — column 1 is still occupied from above"
    );
}

/// The plain-text renderer had the same defect, so it gets the same guarantee.
#[test]
fn plain_text_renderer_aligns_merged_headers_too() {
    let text = to_text(
        &document_with(grouped_header_table()),
        &RenderOptions::default(),
    )
    .unwrap();
    let header = text
        .lines()
        .find(|l| l.contains("Group A"))
        .expect("the header row must survive");

    assert!(
        !header.contains('#'),
        "the text renderer must not invent a '#' either, got {header:?}"
    );
    let before_a = header.find("Group A").unwrap();
    let before_b = header.find("Group B").unwrap();
    assert!(
        before_a < before_b,
        "labels must keep their order and position, got {header:?}"
    );
}

/// AC#6 — the fix must not touch tables that never had a merge. A plain table's
/// rendering is pinned byte-for-byte so a later change to the grid cannot quietly
/// reshape ordinary output.
#[test]
fn table_without_merges_renders_exactly_as_before() {
    let mut table = Table::new();
    table.add_row(row(vec![Cell::header("A"), Cell::header("B")]));
    table.add_row(row(vec![Cell::with_text("1"), Cell::with_text("2")]));

    let md = markdown_of(table);
    assert!(
        md.contains("| A | B |\n| --- | --- |\n| 1 | 2 |"),
        "plain tables must be untouched by the merge handling, got {md:?}"
    );
}

/// The HTML fallback exists because markdown cannot express a merge, and it keys off
/// `Table::has_merged_cells()`. Laying merges out on a grid must stay a *rendering*
/// step: expand them into the model instead and that predicate goes false, silently
/// disabling this opt-in path with no other test to notice.
#[test]
fn html_fallback_still_fires_for_merged_tables() {
    let options = RenderOptions::default().with_table_fallback(TableFallback::Html);
    let md = to_markdown(&document_with(grouped_header_table()), &options).unwrap();

    assert!(
        md.contains("<table>") && md.contains("colspan"),
        "a merged table must still reach the HTML renderer, got {md:?}"
    );
}
