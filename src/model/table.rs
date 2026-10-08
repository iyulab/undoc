//! Table model structures.

use super::Paragraph;
use serde::{Deserialize, Serialize};

/// Horizontal alignment for table cells.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum CellAlignment {
    #[default]
    Left,
    Center,
    Right,
}

/// Vertical alignment for table cells.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum VerticalAlignment {
    #[default]
    Top,
    Middle,
    Bottom,
}

/// A cell in a table.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Cell {
    /// Cell content (paragraphs)
    #[serde(default)]
    pub content: Vec<Paragraph>,

    /// Nested tables within this cell
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub nested_tables: Vec<Table>,

    /// Horizontal span (colspan)
    #[serde(default = "default_span", skip_serializing_if = "is_default_span")]
    pub col_span: u32,

    /// Vertical span (rowspan)
    #[serde(default = "default_span", skip_serializing_if = "is_default_span")]
    pub row_span: u32,

    /// Horizontal alignment
    #[serde(default, skip_serializing_if = "is_default_cell_alignment")]
    pub alignment: CellAlignment,

    /// Vertical alignment
    #[serde(default, skip_serializing_if = "is_default_vertical_alignment")]
    pub vertical_alignment: VerticalAlignment,

    /// Whether this is a header cell
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub is_header: bool,

    /// Background color (hex)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub background: Option<String>,

    /// Resource id of the picture that fills the cell. A fill is not content: the cell's
    /// text stays its text and the image is only referenced from here.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub background_image: Option<String>,
}

fn default_span() -> u32 {
    1
}

fn is_default_span(n: &u32) -> bool {
    *n == 1
}

fn is_default_cell_alignment(a: &CellAlignment) -> bool {
    *a == CellAlignment::Left
}

fn is_default_vertical_alignment(a: &VerticalAlignment) -> bool {
    *a == VerticalAlignment::Top
}

impl Cell {
    /// Create a new empty cell.
    pub fn new() -> Self {
        Self {
            col_span: 1,
            row_span: 1,
            ..Default::default()
        }
    }

    /// Create a cell with text content.
    pub fn with_text(text: impl Into<String>) -> Self {
        Self {
            content: vec![Paragraph::with_text(text)],
            col_span: 1,
            row_span: 1,
            ..Default::default()
        }
    }

    /// Create a header cell with text.
    pub fn header(text: impl Into<String>) -> Self {
        Self {
            content: vec![Paragraph::with_text(text)],
            col_span: 1,
            row_span: 1,
            is_header: true,
            ..Default::default()
        }
    }

    /// Get the plain text content.
    pub fn plain_text(&self) -> String {
        self.content
            .iter()
            .map(|p| p.plain_text())
            .collect::<Vec<_>>()
            .join("\n")
    }

    /// Check if this cell is empty.
    pub fn is_empty(&self) -> bool {
        self.content.is_empty() || self.content.iter().all(|p| p.is_empty())
    }

    /// Check if this cell spans multiple columns.
    pub fn has_col_span(&self) -> bool {
        self.col_span > 1
    }

    /// Check if this cell spans multiple rows.
    pub fn has_row_span(&self) -> bool {
        self.row_span > 1
    }

    /// Check if this cell has any spans.
    pub fn has_spans(&self) -> bool {
        self.col_span > 1 || self.row_span > 1
    }
}

/// A row in a table.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Row {
    /// Cells in this row
    #[serde(default)]
    pub cells: Vec<Cell>,

    /// Whether this is a header row
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub is_header: bool,

    /// Row height in twips (optional)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub height: Option<u32>,
}

impl Row {
    /// Create a new empty row.
    pub fn new() -> Self {
        Self::default()
    }

    /// Create a header row.
    pub fn header(cells: Vec<Cell>) -> Self {
        Self {
            cells,
            is_header: true,
            height: None,
        }
    }

    /// Add a cell to this row.
    pub fn add_cell(&mut self, cell: Cell) {
        self.cells.push(cell);
    }

    /// Get the number of cells.
    pub fn len(&self) -> usize {
        self.cells.len()
    }

    /// Check if the row is empty — true when there are no cells or all cells are content-empty.
    pub fn is_empty(&self) -> bool {
        self.cells.is_empty() || self.cells.iter().all(|c| c.is_empty())
    }

    /// Get the effective column count (accounting for spans).
    pub fn effective_columns(&self) -> usize {
        self.cells.iter().map(|c| c.col_span as usize).sum()
    }
}

/// A table.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Table {
    /// Rows in this table
    #[serde(default)]
    pub rows: Vec<Row>,

    /// Column widths in twips (optional)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub column_widths: Option<Vec<u32>>,

    /// Table caption
    #[serde(skip_serializing_if = "Option::is_none")]
    pub caption: Option<String>,

    /// Table style ID
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style_id: Option<String>,
}

impl Table {
    /// Create a new empty table.
    pub fn new() -> Self {
        Self::default()
    }

    /// Add a row to this table.
    pub fn add_row(&mut self, row: Row) {
        self.rows.push(row);
    }

    /// Get the number of rows.
    pub fn row_count(&self) -> usize {
        self.rows.len()
    }

    /// Get the number of grid columns the table occupies.
    ///
    /// Derived from [`cell_columns`](Self::cell_columns), not from summing each row's
    /// spans: a vertical span from above pushes a row's cells rightward, so a row can
    /// reach further than its own spans add up to.
    pub fn column_count(&self) -> usize {
        self.cell_columns()
            .iter()
            .zip(&self.rows)
            .flat_map(|(cols, row)| {
                cols.iter()
                    .zip(&row.cells)
                    .map(|(&col, cell)| col + cell.col_span.max(1) as usize)
            })
            .max()
            .unwrap_or(0)
    }

    /// The grid column each cell starts in, one `Vec` per row, parallel to `row.cells`.
    ///
    /// A merge is recorded once, on the cell that owns it: the positions it covers —
    /// the tail of a horizontal span, and the columns a vertical span occupies in the
    /// rows below — have no `Cell` of their own. So a cell's index in its row is not its
    /// column once a vertical span from above sits to its left; this walks the spans to
    /// recover it. Every parser builds tables to that contract, and every renderer that
    /// needs a flat grid places cells with this.
    pub fn cell_columns(&self) -> Vec<Vec<usize>> {
        // For each column, how many further rows a vertical span from above still covers.
        let mut carried: Vec<usize> = Vec::new();
        let mut out = Vec::with_capacity(self.rows.len());

        for row in &self.rows {
            let mut cols = Vec::with_capacity(row.cells.len());
            // Columns this row has already passed; a span ending here frees them for the
            // next row.
            let mut col = 0usize;
            let mut seen = 0usize;

            for cell in &row.cells {
                while carried.get(col).is_some_and(|&rows| rows > 0) {
                    col += 1;
                }
                // Each covered column this row steps over is consumed once.
                for c in seen..col {
                    if let Some(rows) = carried.get_mut(c) {
                        *rows = rows.saturating_sub(1);
                    }
                }

                let col_span = cell.col_span.max(1) as usize;
                let row_span = cell.row_span.max(1) as usize;
                cols.push(col);

                if carried.len() < col + col_span {
                    carried.resize(col + col_span, 0);
                }
                for covered in &mut carried[col..col + col_span] {
                    // The span covers this many rows *below* this one.
                    *covered = row_span - 1;
                }

                col += col_span;
                seen = col;
            }

            // Columns still covered past this row's last cell belong to this row too.
            for rows in carried.iter_mut().skip(seen) {
                *rows = rows.saturating_sub(1);
            }

            out.push(cols);
        }
        out
    }

    /// Check if the table is empty.
    pub fn is_empty(&self) -> bool {
        self.rows.is_empty()
    }

    /// Check if the table has merged cells.
    pub fn has_merged_cells(&self) -> bool {
        self.rows
            .iter()
            .any(|r| r.cells.iter().any(|c| c.has_spans()))
    }

    /// Get the header rows.
    pub fn header_rows(&self) -> Vec<&Row> {
        self.rows.iter().filter(|r| r.is_header).collect()
    }

    /// Get the data rows (non-header).
    pub fn data_rows(&self) -> Vec<&Row> {
        self.rows.iter().filter(|r| !r.is_header).collect()
    }

    /// Get plain text representation.
    pub fn plain_text(&self) -> String {
        let mut text = String::new();
        for row in &self.rows {
            let cells: Vec<String> = row.cells.iter().map(|c| c.plain_text()).collect();
            text.push_str(&cells.join("\t"));
            text.push('\n');
        }
        text
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_creation() {
        let cell = Cell::with_text("Hello");
        assert_eq!(cell.plain_text(), "Hello");
        assert!(!cell.is_empty());
        assert!(!cell.has_spans());
    }

    #[test]
    fn test_cell_spans() {
        let mut cell = Cell::with_text("Merged");
        cell.col_span = 2;
        cell.row_span = 3;
        assert!(cell.has_col_span());
        assert!(cell.has_row_span());
        assert!(cell.has_spans());
    }

    fn spanned(text: &str, col_span: u32, row_span: u32) -> Cell {
        Cell {
            col_span,
            row_span,
            ..Cell::with_text(text)
        }
    }

    #[test]
    fn test_cell_columns_follow_spans() {
        // | A (2 rows) | B (2 cols)  |
        // |            | c  | d      |
        // | e          | f  | g      |
        let mut table = Table::new();
        table.add_row(Row {
            cells: vec![spanned("A", 1, 2), spanned("B", 2, 1)],
            ..Row::new()
        });
        table.add_row(Row {
            cells: vec![Cell::with_text("c"), Cell::with_text("d")],
            ..Row::new()
        });
        table.add_row(Row {
            cells: vec![
                Cell::with_text("e"),
                Cell::with_text("f"),
                Cell::with_text("g"),
            ],
            ..Row::new()
        });
        assert_eq!(
            table.cell_columns(),
            vec![vec![0, 1], vec![1, 2], vec![0, 1, 2]]
        );
        assert_eq!(table.column_count(), 3);
    }

    #[test]
    fn test_column_count_includes_columns_pushed_right_by_a_vertical_span() {
        // Row 2 sums to two columns but starts in column 1.
        let mut table = Table::new();
        table.add_row(Row {
            cells: vec![spanned("A", 1, 2), Cell::with_text("b")],
            ..Row::new()
        });
        table.add_row(Row {
            cells: vec![Cell::with_text("c"), Cell::with_text("d")],
            ..Row::new()
        });
        assert_eq!(table.column_count(), 3);
    }

    #[test]
    fn test_row_creation() {
        let mut row = Row::new();
        row.add_cell(Cell::with_text("A"));
        row.add_cell(Cell::with_text("B"));
        assert_eq!(row.len(), 2);
        assert_eq!(row.effective_columns(), 2);
    }

    #[test]
    fn test_row_with_spans() {
        let mut row = Row::new();
        let mut cell = Cell::with_text("Merged");
        cell.col_span = 2;
        row.add_cell(cell);
        row.add_cell(Cell::with_text("Single"));
        assert_eq!(row.len(), 2);
        assert_eq!(row.effective_columns(), 3);
    }

    #[test]
    fn test_table_creation() {
        let mut table = Table::new();
        let mut header = Row::new();
        header.add_cell(Cell::header("Name"));
        header.add_cell(Cell::header("Value"));
        header.is_header = true;
        table.add_row(header);

        let mut row = Row::new();
        row.add_cell(Cell::with_text("foo"));
        row.add_cell(Cell::with_text("bar"));
        table.add_row(row);

        assert_eq!(table.row_count(), 2);
        assert_eq!(table.column_count(), 2);
        assert_eq!(table.header_rows().len(), 1);
        assert_eq!(table.data_rows().len(), 1);
    }

    #[test]
    fn test_table_has_merged_cells() {
        let mut table = Table::new();
        let mut row = Row::new();
        row.add_cell(Cell::with_text("Normal"));
        table.add_row(row);
        assert!(!table.has_merged_cells());

        let mut row2 = Row::new();
        let mut merged = Cell::with_text("Merged");
        merged.col_span = 2;
        row2.add_cell(merged);
        table.add_row(row2);
        assert!(table.has_merged_cells());
    }

    #[test]
    fn test_table_plain_text() {
        let mut table = Table::new();
        let mut row = Row::new();
        row.add_cell(Cell::with_text("A1"));
        row.add_cell(Cell::with_text("B1"));
        table.add_row(row);

        let text = table.plain_text();
        assert!(text.contains("A1"));
        assert!(text.contains("B1"));
    }

    #[test]
    fn test_row_is_empty_with_content_empty_cells() {
        let mut row = Row::new();
        row.add_cell(Cell::new()); // empty cell, no text
        row.add_cell(Cell::new());
        // Row has cells but all are content-empty — should be considered empty
        assert!(
            row.is_empty(),
            "row with only empty cells should be is_empty()"
        );
    }

    #[test]
    fn test_row_is_empty_false_when_cell_has_content() {
        let mut row = Row::new();
        row.add_cell(Cell::with_text("hello"));
        assert!(!row.is_empty());
    }

    #[test]
    fn test_row_is_empty_true_with_no_cells() {
        let row = Row::new();
        assert!(row.is_empty());
    }
}
