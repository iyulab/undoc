//! Plain text renderer implementation.

use crate::error::Result;
use crate::model::{Block, Document, Paragraph, Table};

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

    // Concatenate text runs
    for run in &para.runs {
        output.push_str(&run.text);
    }

    output
}

/// Render a table to plain text (ASCII table).
fn render_table_text(table: &Table) -> String {
    if table.is_empty() {
        return String::new();
    }

    // Calculate column widths
    let col_count = table.column_count();
    if col_count == 0 {
        return String::new();
    }

    let mut widths: Vec<usize> = vec![0; col_count];

    for row in &table.rows {
        for (i, cell) in row.cells.iter().enumerate() {
            if i < col_count {
                let text = cell.plain_text().replace('\n', " ");
                widths[i] = widths[i].max(text.len());
            }
        }
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
        for (i, cell) in row.cells.iter().enumerate() {
            if i < col_count {
                let text = cell.plain_text().replace('\n', " ");
                output.push_str(&format!(" {:width$} |", text, width = widths[i]));
            }
        }
        // Pad missing cells
        for i in row.cells.len()..col_count {
            output.push_str(&format!(" {:width$} |", "", width = widths[i]));
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
}
