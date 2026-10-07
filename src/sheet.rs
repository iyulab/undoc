//! Assembling a worksheet's cells into a table — shared by the spreadsheet readers.
//!
//! A worksheet arrives as rows of positioned cells plus a list of merged ranges, whatever its
//! file format. Turning that into the model's [`Table`] is the same work for every format:
//! fill the positions a row skips, record each merge once on the cell that owns it, drop the
//! positions a merge covers, trim the empty edges, and keep row spans within the rows that are
//! actually there. Each reader supplies cells; this does the rest.

use crate::model::{Cell, Row, Table};

/// One merged range, in 0-based columns and 1-based sheet rows.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct MergeRange {
    pub left: u32,
    pub top: u32,
    pub right: u32,
    pub bottom: u32,
}

/// What a sheet position is with respect to the merged ranges.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Coverage {
    /// Not inside any merge.
    Free,
    /// The top-left position of a merge: the cell that owns it, with `(col_span, row_span)`.
    Origin(u32, u32),
    /// Inside a merge but not its top-left — the model has no cell here.
    Covered,
}

/// A sheet's merged ranges, queried row by row.
///
/// Worksheets list their rows in ascending order, so only the ranges spanning the current
/// row are kept at hand; a sheet with thousands of merges does not cost a scan of all of
/// them per cell, and a range spanning whole columns costs nothing per position.
struct SheetMerges {
    /// Sorted by `top`.
    ranges: Vec<MergeRange>,
    next: usize,
    active: Vec<MergeRange>,
    row: u32,
}

impl SheetMerges {
    fn new(mut ranges: Vec<MergeRange>) -> Self {
        ranges.sort_by_key(|r| r.top);
        Self {
            ranges,
            next: 0,
            active: Vec::new(),
            row: 0,
        }
    }

    fn seek(&mut self, row: u32) {
        if row == self.row {
            return;
        }
        if row < self.row {
            // Out-of-order rows (not written by Excel, but not rejected either): start over.
            self.next = 0;
            self.active.clear();
        }
        self.row = row;
        self.active.retain(|r| r.bottom >= row);
        while let Some(r) = self.ranges.get(self.next).copied() {
            if r.top > row {
                break;
            }
            if r.bottom >= row {
                self.active.push(r);
            }
            self.next += 1;
        }
    }

    /// Coverage of `(col, row)`.
    fn at(&mut self, col: u32, row: u32) -> Coverage {
        self.seek(row);
        match self
            .active
            .iter()
            .find(|r| (r.left..=r.right).contains(&col))
        {
            None => Coverage::Free,
            Some(r) if r.left == col && r.top == row => {
                Coverage::Origin(r.right - r.left + 1, r.bottom - r.top + 1)
            }
            Some(_) => Coverage::Covered,
        }
    }
}

/// Builds a [`Table`] from a worksheet's rows, one row and one cell at a time.
pub(crate) struct SheetTableBuilder {
    table: Table,
    merges: SheetMerges,
    /// Sheet row number of each table row, parallel to `table.rows` — a merge's `row_span`
    /// counts sheet rows, and rows the worksheet omits are not table rows.
    sheet_rows: Vec<Option<u32>>,
    /// The row being read, its sheet row number, and the next grid column not yet
    /// accounted for in it.
    row: Option<Row>,
    sheet_row: Option<u32>,
    next_col: u32,
}

impl SheetTableBuilder {
    pub fn new(merges: Vec<MergeRange>) -> Self {
        Self {
            table: Table::new(),
            merges: SheetMerges::new(merges),
            sheet_rows: Vec::new(),
            row: None,
            sheet_row: None,
            next_col: 0,
        }
    }

    /// Whether the row being read is the table's first — its cells are header cells.
    pub fn in_first_row(&self) -> bool {
        self.table.rows.is_empty()
    }

    /// Begin a row. `sheet_row` is its 1-based sheet row number, when the format states it.
    pub fn start_row(&mut self, sheet_row: Option<u32>) {
        self.end_row();
        self.row = Some(Row {
            cells: Vec::new(),
            is_header: self.in_first_row(),
            height: None,
        });
        self.sheet_row = sheet_row;
        self.next_col = 0;
    }

    /// Finish the row being read, if any.
    pub fn end_row(&mut self) {
        if let Some(row) = self.row.take() {
            self.table.add_row(row);
            self.sheet_rows.push(self.sheet_row);
        }
    }

    /// Append a cell to the row being read, honouring the merged ranges.
    ///
    /// The model records a merge once, on the cell that owns it ([`Table::cell_columns`]):
    /// the owner gets `col_span`/`row_span` and the positions it covers get no cell. A
    /// worksheet may write a cell for a covered position (Excel does when it is styled), and
    /// omits positions that are neither styled nor filled. So gaps are filled up to the
    /// cell's column — skipping covered positions, giving an unwritten merge owner its spans
    /// — and a covered cell is dropped; its value is one Excel no longer shows.
    ///
    /// `at` is `(0-based column, 1-based sheet row)`; a cell without a position takes the next
    /// column. Gap reconstruction stays inside the current row, up to the highest column it
    /// references. This is not whole-sheet densification.
    pub fn place(&mut self, mut cell: Cell, at: Option<(u32, u32)>) {
        let is_header = self.in_first_row();
        let Some(row) = self.row.as_mut() else {
            return;
        };
        let Some((col, ref_row)) = at else {
            self.next_col += cell.col_span.max(1);
            row.cells.push(cell);
            return;
        };
        let sheet_row = *self.sheet_row.get_or_insert(ref_row);

        for gap in self.next_col..col {
            match self.merges.at(gap, sheet_row) {
                Coverage::Covered => {}
                Coverage::Free => row.cells.push(Cell {
                    is_header,
                    ..Cell::new()
                }),
                Coverage::Origin(col_span, row_span) => row.cells.push(Cell {
                    col_span,
                    row_span,
                    is_header,
                    ..Cell::new()
                }),
            }
        }
        self.next_col = self.next_col.max(col + 1);

        match self.merges.at(col, sheet_row) {
            Coverage::Covered => {}
            Coverage::Free => row.cells.push(cell),
            Coverage::Origin(col_span, row_span) => {
                cell.col_span = col_span;
                cell.row_span = row_span;
                row.cells.push(cell);
            }
        }
    }

    /// The table, with its empty edges trimmed and its row spans clamped.
    pub fn finish(mut self) -> Table {
        self.end_row();
        let mut table = self.table;
        let mut sheet_rows = self.sheet_rows;

        // Trim trailing rows where every cell is content-empty (no text, no value).
        // These arise from formatting-only rows, deleted-but-kept rows, and conditional
        // formatting applied to entire columns — all common causes of massive row inflation.
        while table.rows.last().is_some_and(|r| r.is_empty()) {
            table.rows.pop();
            sheet_rows.pop();
        }

        // Trim trailing columns where every row has empty cells: keep up to the rightmost
        // grid column any content reaches. Measured on the grid, not by index in the row —
        // a row under a vertical merge has fewer cells than columns.
        let columns = table.cell_columns();
        let keep = columns
            .iter()
            .zip(&table.rows)
            .flat_map(|(cols, row)| {
                cols.iter()
                    .zip(&row.cells)
                    .filter(|(_, c)| !c.is_empty())
                    .map(|(&col, c)| col + c.col_span.max(1) as usize)
            })
            .max();
        if let Some(keep) = keep {
            for (cols, row) in columns.iter().zip(&mut table.rows) {
                let kept = cols.iter().take_while(|&&col| col < keep).count();
                row.cells.truncate(kept);
            }
        }

        clamp_row_spans(&mut table, &sheet_rows);
        table
    }
}

/// Bring each merge owner's `row_span` down to the table rows it actually covers.
///
/// A merge's extent is in sheet rows, but the table holds only the rows the worksheet
/// writes — an empty row inside a merge may be omitted, and trailing empty rows are
/// trimmed. A span left counting rows that are not there would push cells of later
/// rows out from under their headings.
fn clamp_row_spans(table: &mut Table, sheet_rows: &[Option<u32>]) {
    for i in 0..table.rows.len() {
        let Some(top) = sheet_rows.get(i).copied().flatten() else {
            continue;
        };
        let below = &sheet_rows[i..];
        for cell in &mut table.rows[i].cells {
            if cell.row_span <= 1 {
                continue;
            }
            let bottom = top + cell.row_span - 1;
            let present = below
                .iter()
                .take_while(|r| r.is_none_or(|r| r <= bottom))
                .count();
            cell.row_span = present.max(1) as u32;
        }
    }
}

/// Built-in number formats that show a date or a time: 14–22 and 45–47.
pub(crate) fn is_builtin_date_format(num_fmt_id: u32) -> bool {
    (14..=22).contains(&num_fmt_id) || (45..=47).contains(&num_fmt_id)
}

/// Check if a format code string represents a date format.
pub(crate) fn is_date_format_code(format_code: &str) -> bool {
    // Date patterns: d, m, y (case insensitive, not in quotes or brackets)
    // Time patterns: h, s (case insensitive)
    // We need to exclude patterns in square brackets [Red] or quotes "text"

    let mut in_bracket = false;
    let mut in_quote = false;
    let mut prev_char = '\0';

    for c in format_code.chars() {
        match c {
            '[' if !in_quote => in_bracket = true,
            ']' if !in_quote => in_bracket = false,
            '"' => in_quote = !in_quote,
            _ if !in_bracket && !in_quote => {
                // Check for date/time patterns
                let lower = c.to_ascii_lowercase();
                match lower {
                    // 'd' for day, 'm' for month (but not 'mm:ss' which is minutes)
                    'd' => return true,
                    'y' => return true,
                    // 'h' for hour indicates time, which is often stored as fractional day
                    // But we mainly want date, so check for 'm' after 'd' or before 'd'
                    'm' => {
                        // 'm' could be month or minute
                        // If preceded by 'd' or 'y', it's likely month
                        // If preceded by 'h' or followed by 's', it's likely minute
                        // For simplicity, check surrounding context
                        let lower_prev = prev_char.to_ascii_lowercase();
                        if lower_prev == 'd' || lower_prev == 'y' {
                            return true; // Month after day/year
                        }
                        // Could also be month at start or standalone
                        // Check if format contains 'd' or 'y' anywhere
                        let lower_format = format_code.to_lowercase();
                        if lower_format.contains('d') || lower_format.contains('y') {
                            return true;
                        }
                    }
                    _ => {}
                }
            }
            _ => {}
        }
        prev_char = c;
    }

    false
}

/// Convert Excel serial date number to ISO 8601 date string.
pub(crate) fn serial_to_date(serial: f64) -> Option<String> {
    // Excel date system: days since December 30, 1899
    // (Excel incorrectly treats 1900 as a leap year for Lotus 1-2-3 compatibility)

    if serial < 0.0 {
        return None;
    }

    // Handle the "Lotus 1-2-3" bug: Excel thinks Feb 29, 1900 exists
    // Serial 60 = Feb 29, 1900 (doesn't exist)
    // Serial 61 = Mar 1, 1900
    let adjusted_serial = if serial > 60.0 { serial - 1.0 } else { serial };

    // Days since January 1, 1900 (day 1 = Jan 1, 1900)
    let days = adjusted_serial.floor() as i64;

    // Convert to date
    // January 1, 1900 is day 1
    // Using a simple calculation:
    // Base date: 1899-12-31 (so day 1 = 1900-01-01)

    // Calculate year, month, day
    let (year, month, day) = days_to_ymd(days)?;

    // Check if there's a time component
    let time_fraction = serial.fract();
    if time_fraction > 0.0001 {
        // Has time component
        let total_seconds = (time_fraction * 86400.0).round() as u32;
        let hours = total_seconds / 3600;
        let minutes = (total_seconds % 3600) / 60;
        let seconds = total_seconds % 60;
        Some(format!(
            "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}",
            year, month, day, hours, minutes, seconds
        ))
    } else {
        Some(format!("{:04}-{:02}-{:02}", year, month, day))
    }
}

/// Convert days since December 31, 1899 to (year, month, day).
fn days_to_ymd(days: i64) -> Option<(i32, u32, u32)> {
    if days < 1 {
        return None;
    }

    // Start from 1900-01-01, which is serial day 1
    let mut year = 1900;
    let mut remaining_days = days;

    // Year loop
    loop {
        let days_in_year = if is_leap_year(year) { 366 } else { 365 };
        if remaining_days <= days_in_year {
            break;
        }
        remaining_days -= days_in_year;
        year += 1;
    }

    // Month loop
    let months_days = if is_leap_year(year) {
        [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    } else {
        [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    };

    let mut month = 1u32;
    for &days_in_month in &months_days {
        if remaining_days <= days_in_month as i64 {
            break;
        }
        remaining_days -= days_in_month as i64;
        month += 1;
    }

    let day = remaining_days.max(1) as u32;

    Some((year, month, day))
}

/// Check if a year is a leap year.
fn is_leap_year(year: i32) -> bool {
    (year % 4 == 0 && year % 100 != 0) || (year % 400 == 0)
}
