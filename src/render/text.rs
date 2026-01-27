//! Plain text renderer implementation.

use crate::error::Result;
use crate::model::{Block, Document, Paragraph, Table};
use unicode_width::UnicodeWidthStr;

use super::options::RenderOptions;

/// Convert a Document to plain text.
pub fn to_text(doc: &Document, options: &RenderOptions) -> Result<String> {
    let mut output = String::new();

    // Render each section
    for (i, section) in doc.sections.iter().enumerate() {
        if i > 0 && options.paragraph_spacing {
            output.push_str("\n\n");
        }

        // Add section name if present
        if let Some(ref name) = section.name {
            output.push_str(name);
            output.push_str("\n\n");
        }

        // Render content blocks
        for block in &section.content {
            match block {
                Block::Paragraph(para) => {
                    let text = render_paragraph_text(para);
                    if !text.is_empty() || options.include_empty_paragraphs {
                        output.push_str(&text);
                        output.push('\n');
                        if options.paragraph_spacing {
                            output.push('\n');
                        }
                    }
                }
                Block::Table(table) => {
                    output.push_str(&render_table_text(table));
                    output.push_str("\n\n");
                }
                Block::PageBreak | Block::SectionBreak => {
                    output.push_str("\n---\n\n");
                }
                Block::Image { alt_text, .. } => {
                    if let Some(alt) = alt_text {
                        output.push_str(&format!("[Image: {}]\n", alt));
                    } else {
                        output.push_str("[Image]\n");
                    }
                }
            }
        }

        // Render notes if present (for PPTX)
        if let Some(ref notes) = section.notes {
            if !notes.is_empty() {
                output.push_str("\nNotes:\n");
                for note in notes {
                    let text = render_paragraph_text(note);
                    if !text.is_empty() {
                        output.push_str(&text);
                        output.push('\n');
                    }
                }
            }
        }
    }

    // Apply cleanup if configured
    let result = if let Some(ref cleanup) = options.cleanup {
        super::cleanup::clean_text(&output, cleanup)
    } else {
        output.trim().to_string()
    };

    Ok(result)
}

/// Render a paragraph to plain text.
fn render_paragraph_text(para: &Paragraph) -> String {
    let mut output = String::new();

    // Handle list items
    if let Some(ref list_info) = para.list_info {
        let indent = "  ".repeat(list_info.level as usize);
        output.push_str(&indent);
        match list_info.list_type {
            crate::model::ListType::Bullet => {
                output.push_str("• ");
            }
            crate::model::ListType::Numbered => {
                let num = list_info.number.unwrap_or(1);
                output.push_str(&format!("{}. ", num));
            }
            crate::model::ListType::None => {}
        }
    }

    // Concatenate text runs with smart spacing
    for (i, run) in para.runs.iter().enumerate() {
        // Add space between runs if needed
        if i > 0 && !run.text.is_empty() && !output.is_empty() {
            let last_char = output.chars().last();
            let first_char = run.text.chars().next();

            if let (Some(last), Some(first)) = (last_char, first_char) {
                let needs_space =
                    !last.is_whitespace() && !first.is_whitespace() && !is_no_space_before(first);
                if needs_space {
                    output.push(' ');
                }
            }
        }

        output.push_str(&run.text);
    }

    output
}

/// Check if a character should NOT have a space before it.
fn is_no_space_before(c: char) -> bool {
    matches!(
        c,
        '.' | ',' | ':' | ';' | '!' | '?' | ')' | ']' | '}' | '"' | '\'' | '…'
    )
}

/// Render a table to plain text (ASCII table).
/// Uses unicode-width to correctly align CJK characters (which take 2 columns in terminals).
fn render_table_text(table: &Table) -> String {
    if table.is_empty() {
        return String::new();
    }

    // Calculate column widths
    let col_count = table.column_count();
    if col_count == 0 {
        return String::new();
    }

    // Check if header row has fewer cells than data rows
    let header_missing = if let Some(first_row) = table.rows.first() {
        if first_row.cells.len() < col_count {
            col_count - first_row.cells.len()
        } else {
            0
        }
    } else {
        0
    };

    let mut widths: Vec<usize> = vec![0; col_count];

    // Calculate widths accounting for header placeholders
    // Use display width (unicode-width) instead of character count
    for (row_idx, row) in table.rows.iter().enumerate() {
        let offset = if row_idx == 0 { header_missing } else { 0 };
        for (i, cell) in row.cells.iter().enumerate() {
            let col_idx = i + offset;
            if col_idx < col_count {
                let text = cell.plain_text().replace('\n', " ");
                // Use display width for correct CJK alignment
                widths[col_idx] = widths[col_idx].max(text.width());
            }
        }
    }

    // Add width for header placeholders
    if header_missing > 0 {
        widths[0] = widths[0].max(1); // "#" placeholder
    }

    // Minimum width of 3 for readability
    for w in &mut widths {
        *w = (*w).max(3);
    }

    let mut output = String::new();

    // Top border
    output.push('+');
    for w in &widths {
        output.push_str(&"-".repeat(*w + 2));
        output.push('+');
    }
    output.push('\n');

    // Rows
    for (row_idx, row) in table.rows.iter().enumerate() {
        output.push('|');

        // For header row, prepend placeholder columns
        if row_idx == 0 && header_missing > 0 {
            for (j, width) in widths.iter().take(header_missing).enumerate() {
                let placeholder = if j == 0 { "#" } else { "" };
                output.push_str(&format!(" {} |", pad_to_width(placeholder, *width)));
            }
        }

        for (i, cell) in row.cells.iter().enumerate() {
            let col_idx = if row_idx == 0 { i + header_missing } else { i };
            if col_idx < col_count {
                let text = cell.plain_text().replace('\n', " ");
                output.push_str(&format!(" {} |", pad_to_width(&text, widths[col_idx])));
            }
        }

        // Pad data rows if they have fewer cells
        if row_idx > 0 {
            for width in widths.iter().take(col_count).skip(row.cells.len()) {
                output.push_str(&format!(" {} |", pad_to_width("", *width)));
            }
        }
        output.push('\n');

        // Separator after header row
        if row_idx == 0 && row.is_header {
            output.push('+');
            for w in &widths {
                output.push_str(&"=".repeat(*w + 2));
                output.push('+');
            }
            output.push('\n');
        }
    }

    // Bottom border
    output.push('+');
    for w in &widths {
        output.push_str(&"-".repeat(*w + 2));
        output.push('+');
    }

    output
}

/// Pad a string to reach a target display width.
/// Accounts for CJK characters which take 2 columns in terminals.
fn pad_to_width(s: &str, target_width: usize) -> String {
    let current_width = s.width();
    if current_width >= target_width {
        s.to_string()
    } else {
        let padding = target_width - current_width;
        format!("{}{}", s, " ".repeat(padding))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{Cell, HeadingLevel, Row, Section};

    #[test]
    fn test_basic_paragraph() {
        let para = Paragraph::with_text("Hello, World!");
        let text = render_paragraph_text(&para);
        assert_eq!(text, "Hello, World!");
    }

    #[test]
    fn test_document_to_text() {
        let mut doc = Document::new();
        let mut section = Section::new(0);
        section.add_paragraph(Paragraph::heading(HeadingLevel::H1, "Test"));
        section.add_paragraph(Paragraph::with_text("Content."));
        doc.add_section(section);

        let options = RenderOptions::default();
        let text = to_text(&doc, &options).unwrap();
        assert!(text.contains("Test"));
        assert!(text.contains("Content."));
    }

    #[test]
    fn test_table_text() {
        let mut table = Table::new();
        let mut header = Row::header(vec![Cell::header("A"), Cell::header("B")]);
        header.is_header = true;
        table.add_row(header);
        table.add_row(Row {
            cells: vec![Cell::with_text("1"), Cell::with_text("2")],
            is_header: false,
            height: None,
        });

        let text = render_table_text(&table);
        assert!(text.contains("| A "));
        assert!(text.contains("| B "));
        assert!(text.contains("| 1 "));
        assert!(text.contains("| 2 "));
    }

    #[test]
    fn test_list_items() {
        let mut para = Paragraph::with_text("Item");
        para.list_info = Some(crate::model::ListInfo {
            list_type: crate::model::ListType::Bullet,
            level: 0,
            number: None,
        });

        let text = render_paragraph_text(&para);
        assert!(text.contains("• Item"));
    }

    #[test]
    fn test_cjk_table_alignment() {
        use unicode_width::UnicodeWidthStr;

        // Test that CJK characters (which take 2 display columns) are properly aligned
        let mut table = Table::new();
        let mut header = Row::header(vec![Cell::header("이름"), Cell::header("값")]);
        header.is_header = true;
        table.add_row(header);
        table.add_row(Row {
            cells: vec![Cell::with_text("홍길동"), Cell::with_text("100")],
            is_header: false,
            height: None,
        });
        table.add_row(Row {
            cells: vec![Cell::with_text("Kim"), Cell::with_text("200")],
            is_header: false,
            height: None,
        });

        let text = render_table_text(&table);
        let lines: Vec<&str> = text.lines().collect();

        // Get border display width from first line (uses ASCII only, so bytes == display width)
        let border_width = lines[0].width();

        // All lines should have the same DISPLAY width for proper alignment
        // (byte length will differ due to UTF-8 encoding)
        for (i, line) in lines.iter().enumerate() {
            assert_eq!(
                line.width(),
                border_width,
                "Line {} has different display width: '{}' (expected {})",
                i,
                line,
                border_width
            );
        }

        // Verify the table looks correct
        assert!(text.contains("| 이름"));
        assert!(text.contains("| 홍길동"));
        assert!(text.contains("| Kim"));
    }

    #[test]
    fn test_pad_to_width() {
        // ASCII characters: 1 char = 1 width
        assert_eq!(pad_to_width("abc", 5), "abc  ");

        // CJK characters: 1 char = 2 width
        // "한글" has 2 chars but 4 display width
        assert_eq!(pad_to_width("한글", 6), "한글  ");

        // Mixed: "A한" = 1 + 2 = 3 width
        assert_eq!(pad_to_width("A한", 5), "A한  ");

        // Already at or exceeds target
        assert_eq!(pad_to_width("한글", 4), "한글");
        assert_eq!(pad_to_width("한글", 3), "한글");
    }
}
