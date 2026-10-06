//! Laying a table's cells out on a flat column grid.
//!
//! The document model records a merge as a span on the cell that owns it, which is how
//! the source formats describe it. Markdown and the plain-text table have no equivalent:
//! every row there is a flat sequence of columns. So a merged cell has to be placed at
//! the column it starts in and the columns it covers filled with empty ones.
//!
//! Without that step a row holding four merged group labels is four columns wide instead
//! of seventeen, and whatever the renderer does to reconcile that — pad one end or the
//! other — moves every value in the row out from under its heading.

use crate::model::{Cell, Table};

/// Place a table's cells on a grid: one `Vec` per row, one entry per column.
///
/// `None` is a column covered by a merge — the tail of a horizontal span, or a column a
/// vertical span occupies from a row above — and renders as an empty cell. Every row
/// comes back the same width, so callers do not pad.
///
/// Cells are placed at [`Table::cell_columns`], which is also where the grid's width
/// comes from: a vertical span pushes later cells rightward, so a row can need more
/// columns than its own spans sum to.
pub(super) fn lay_out(table: &Table) -> Vec<Vec<Option<&Cell>>> {
    let width = table.column_count();
    table
        .cell_columns()
        .into_iter()
        .zip(&table.rows)
        .map(|(cols, row)| {
            let mut slots: Vec<Option<&Cell>> = vec![None; width];
            for (col, cell) in cols.into_iter().zip(&row.cells) {
                slots[col] = Some(cell);
            }
            slots
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::Row;

    fn row(cells: Vec<Cell>) -> Row {
        Row {
            cells,
            is_header: false,
            height: None,
        }
    }

    fn cell(text: &str, col_span: u32, row_span: u32) -> Cell {
        Cell {
            col_span,
            row_span,
            ..Cell::with_text(text)
        }
    }

    fn texts(grid: &[Vec<Option<&Cell>>]) -> Vec<Vec<String>> {
        grid.iter()
            .map(|row| {
                row.iter()
                    .map(|slot| slot.map(|c| c.plain_text()).unwrap_or_default())
                    .collect()
            })
            .collect()
    }

    #[test]
    fn plain_table_is_unchanged() {
        let mut table = Table::new();
        table.add_row(row(vec![
            Cell::with_text("A"),
            Cell::with_text("B"),
            Cell::with_text("C"),
        ]));
        let grid = lay_out(&table);
        assert_eq!(texts(&grid), vec![vec!["A", "B", "C"]]);
    }

    #[test]
    fn horizontal_span_anchors_at_its_first_column() {
        let mut table = Table::new();
        // A group label over columns 1-2, then a plain cell in column 3.
        table.add_row(row(vec![cell("Group", 2, 1), Cell::with_text("C")]));
        table.add_row(row(vec![
            Cell::with_text("a"),
            Cell::with_text("b"),
            Cell::with_text("c"),
        ]));

        let grid = texts(&lay_out(&table));
        assert_eq!(
            grid[0],
            vec!["Group", "", "C"],
            "the label must sit in column 1, not be pushed right by padding"
        );
        assert_eq!(grid[1], vec!["a", "b", "c"]);
    }

    #[test]
    fn vertical_span_holds_its_column_in_later_rows() {
        let mut table = Table::new();
        // Column 1 spans both rows; row 2 supplies only the remaining columns.
        table.add_row(row(vec![
            cell("Side", 1, 2),
            Cell::with_text("X"),
            Cell::with_text("Y"),
        ]));
        table.add_row(row(vec![Cell::with_text("x"), Cell::with_text("y")]));

        let grid = texts(&lay_out(&table));
        assert_eq!(grid[0], vec!["Side", "X", "Y"]);
        assert_eq!(
            grid[1],
            vec!["", "x", "y"],
            "row 2's cells must start in column 2 — column 1 is still occupied"
        );
    }

    #[test]
    fn every_row_comes_back_the_same_width() {
        let mut table = Table::new();
        table.add_row(row(vec![cell("Wide", 3, 1)]));
        table.add_row(row(vec![Cell::with_text("only one")]));

        let grid = lay_out(&table);
        assert!(grid.iter().all(|r| r.len() == 3), "got {:?}", texts(&grid));
    }

    #[test]
    fn vertical_span_past_the_last_cell_still_widens_the_row() {
        let mut table = Table::new();
        table.add_row(row(vec![Cell::with_text("A"), cell("Tall", 1, 2)]));
        table.add_row(row(vec![Cell::with_text("b")]));

        let grid = texts(&lay_out(&table));
        assert_eq!(
            grid[1],
            vec!["b", ""],
            "the carried column belongs to row 2"
        );
    }
}
