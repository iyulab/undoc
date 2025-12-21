//! Markdown renderer implementation.

use crate::error::Result;
use crate::model::{Block, Document, Paragraph, Table, TextRun};

use super::options::RenderOptions;

/// Convert a Document to Markdown.
pub fn to_markdown(doc: &Document, options: &RenderOptions) -> Result<String> {
    let mut output = String::new();

    // Add frontmatter if requested
    if options.include_frontmatter {
        output.push_str(&render_frontmatter(doc));
    }

    // Render each section
    for (i, section) in doc.sections.iter().enumerate() {
        // Add section name as heading if present
        if let Some(ref name) = section.name {
            if i > 0 {
                output.push_str("\n---\n\n");
            }
            output.push_str(&format!("## {}\n\n", name));
        }

        // Render content blocks
        for block in &section.content {
            match block {
                Block::Paragraph(para) => {
                    let md = render_paragraph(para, options);
                    if !md.is_empty() || options.include_empty_paragraphs {
                        output.push_str(&md);
                        if options.paragraph_spacing {
                            output.push_str("\n\n");
                        } else {
                            output.push('\n');
                        }
                    }
                }
                Block::Table(table) => {
                    output.push_str(&render_table(table, options));
                    output.push_str("\n\n");
                }
                Block::PageBreak => {
                    output.push_str("\n---\n\n");
                }
                Block::SectionBreak => {
                    output.push_str("\n---\n\n");
                }
                Block::Image {
                    resource_id,
                    alt_text,
                    ..
                } => {
                    let alt = alt_text.as_deref().unwrap_or("image");
                    let path = format!("{}{}", options.image_path_prefix, resource_id);
                    output.push_str(&format!("![{}]({})\n\n", alt, path));
                }
            }
        }

        // Render notes if present (for PPTX)
        if let Some(ref notes) = section.notes {
            if !notes.is_empty() {
                output.push_str("\n> **Notes:**\n");
                for note in notes {
                    let text = render_paragraph(note, options);
                    if !text.is_empty() {
                        output.push_str(&format!("> {}\n", text));
                    }
                }
                output.push('\n');
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

/// Render YAML frontmatter from document metadata.
fn render_frontmatter(doc: &Document) -> String {
    let mut fm = String::from("---\n");
    let meta = &doc.metadata;

    if let Some(ref title) = meta.title {
        fm.push_str(&format!("title: \"{}\"\n", escape_yaml(title)));
    }
    if let Some(ref author) = meta.author {
        fm.push_str(&format!("author: \"{}\"\n", escape_yaml(author)));
    }
    if let Some(ref created) = meta.created {
        fm.push_str(&format!("created: \"{}\"\n", created));
    }
    if let Some(ref modified) = meta.modified {
        fm.push_str(&format!("modified: \"{}\"\n", modified));
    }

    fm.push_str("---\n\n");
    fm
}

/// Escape special characters in YAML strings.
fn escape_yaml(s: &str) -> String {
    s.replace('\\', "\\\\").replace('"', "\\\"")
}

/// Render a paragraph to Markdown.
fn render_paragraph(para: &Paragraph, options: &RenderOptions) -> String {
    let mut output = String::new();

    // Merge adjacent runs with the same style to avoid issues like:
    // **시** **험** **합** -> **시험합**
    let merged_para = para.with_merged_runs();

    // Handle heading
    if merged_para.heading.is_heading() {
        let level = merged_para.heading.level().min(options.max_heading_level);
        output.push_str(&"#".repeat(level as usize));
        output.push(' ');
    }

    // Handle list items
    if let Some(ref list_info) = merged_para.list_info {
        let indent = "  ".repeat(list_info.level as usize);
        output.push_str(&indent);
        match list_info.list_type {
            crate::model::ListType::Bullet => {
                output.push(options.list_marker);
                output.push(' ');
            }
            crate::model::ListType::Numbered => {
                let num = list_info.number.unwrap_or(1);
                output.push_str(&format!("{}. ", num));
            }
            crate::model::ListType::None => {}
        }
    }

    // Render text runs with smart spacing
    for (i, run) in merged_para.runs.iter().enumerate() {
        let run_text = render_run(run, options);

        // Add space between runs if needed
        if i > 0 && !run_text.is_empty() && !output.is_empty() {
            let last_char = output.chars().last();
            let first_char = run_text.chars().next();

            // Add space if:
            // - Previous run doesn't end with space/newline
            // - Current run doesn't start with space/punctuation
            if let (Some(last), Some(first)) = (last_char, first_char) {
                let needs_space =
                    !last.is_whitespace() && !first.is_whitespace() && !is_no_space_before(first);
                if needs_space {
                    output.push(' ');
                }
            }
        }

        output.push_str(&run_text);
    }

    // Render inline images
    for image in &para.images {
        if !output.is_empty() {
            output.push('\n');
        }
        let alt = image.alt_text.as_deref().unwrap_or("image");
        let path = format!("{}{}", options.image_path_prefix, image.resource_id);
        output.push_str(&format!("![{}]({})", alt, path));
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

/// Render a text run to Markdown.
fn render_run(run: &TextRun, options: &RenderOptions) -> String {
    if run.text.is_empty() {
        return String::new();
    }

    let mut text = if options.escape_special_chars {
        escape_markdown(&run.text)
    } else {
        run.text.clone()
    };

    // Apply formatting (innermost first)
    if run.style.code {
        text = format!("`{}`", text.replace('`', "\\`"));
    }
    if run.style.strikethrough {
        text = format!("~~{}~~", text);
    }
    if run.style.bold && run.style.italic {
        text = format!("***{}***", text);
    } else if run.style.bold {
        text = format!("**{}**", text);
    } else if run.style.italic {
        text = format!("*{}*", text);
    }

    // Handle hyperlinks
    if let Some(ref url) = run.hyperlink {
        text = format!("[{}]({})", text, url);
    }

    text
}

/// Escape Markdown special characters.
///
/// Context-aware escaping - only escapes when the character could actually
/// trigger markdown formatting:
///
/// - `\` - always escape (escape character)
/// - `` ` `` - always escape (inline code)
/// - `|` - always escape (table delimiter)
/// - `*` and `_` - only escape when they could trigger emphasis:
///   - NOT escaped after `(`, `[`, or whitespace (can't start emphasis)
///   - NOT escaped before `)`, `]`, or whitespace (can't end emphasis)
///
/// Characters NOT escaped (only special in specific contexts):
/// - `()`, `[]`, `{}` - only special in link/image syntax `[text](url)`
/// - `#` - only special at start of line (headings)
/// - `+`, `-` - only special at start of line (lists) or `---` (rules)
/// - `!` - only special before `[` (images)
/// - `.` - only special in ordered lists at line start (e.g., "1.")
fn escape_markdown(s: &str) -> String {
    let mut result = String::with_capacity(s.len());
    let chars: Vec<char> = s.chars().collect();

    for (i, &c) in chars.iter().enumerate() {
        match c {
            // Always escape
            '\\' | '`' | '|' => {
                result.push('\\');
                result.push(c);
            }
            // Context-aware escaping for emphasis markers
            '*' | '_' => {
                let prev = if i > 0 { Some(chars[i - 1]) } else { None };
                let next = chars.get(i + 1).copied();

                // Don't escape if:
                // 1. After opening bracket/paren, whitespace, or start of string
                // 2. Before closing bracket/paren, whitespace, or end of string
                // 3. Before colon (common in `*NOTE:` patterns)
                //
                // In CommonMark, emphasis requires BOTH:
                // - A left-flanking `*` (followed by non-whitespace)
                // - A matching right-flanking `*` (preceded by non-whitespace)
                // If there's no matching pair, it won't render as emphasis.
                let after_opener = prev.is_none_or(|p| {
                    matches!(p, '(' | '[' | '{' | ':' | '-' | '/' | '\\') || p.is_whitespace()
                });
                let before_closer = next.is_none_or(|n| {
                    matches!(n, ')' | ']' | '}' | ':' | '-' | '/' | '\\') || n.is_whitespace()
                });

                if after_opener || before_closer {
                    // Safe to use without escaping
                    result.push(c);
                } else {
                    // Could potentially trigger emphasis, escape it
                    result.push('\\');
                    result.push(c);
                }
            }
            _ => result.push(c),
        }
    }
    result
}

/// Render a table cell's content with formatting preserved.
/// Multiple paragraphs are joined with `<br>` for inline display.
fn render_cell_content(cell: &crate::model::Cell, options: &RenderOptions) -> String {
    let mut parts = Vec::new();

    for para in &cell.content {
        // Merge adjacent runs with same style (like render_paragraph does)
        let merged_para = para.with_merged_runs();
        let mut para_text = String::new();

        for (i, run) in merged_para.runs.iter().enumerate() {
            let run_text = render_run(run, options);

            // Add smart spacing between runs (like render_paragraph does)
            if i > 0 && !run_text.is_empty() && !para_text.is_empty() {
                let last_char = para_text.chars().last();
                let first_char = run_text.chars().next();

                if let (Some(last), Some(first)) = (last_char, first_char) {
                    let needs_space = !last.is_whitespace()
                        && !first.is_whitespace()
                        && !is_no_space_before(first);
                    if needs_space {
                        para_text.push(' ');
                    }
                }
            }

            para_text.push_str(&run_text);
        }

        if !para_text.is_empty() {
            parts.push(para_text);
        }
    }

    // Render nested tables inline (flatten their content)
    for nested_table in &cell.nested_tables {
        for row in &nested_table.rows {
            for nested_cell in &row.cells {
                let nested_content = render_cell_content(nested_cell, options);
                if !nested_content.is_empty() {
                    parts.push(nested_content);
                }
            }
        }
    }

    // Join paragraphs with <br> for markdown table cells
    let text = parts.join("<br>");

    // Only replace newlines - pipes are already escaped by escape_markdown in render_run
    text.replace('\n', " ")
}

/// Render a table to Markdown.
fn render_table(table: &Table, options: &RenderOptions) -> String {
    if table.is_empty() {
        return String::new();
    }

    // Check if we need HTML fallback
    if table.has_merged_cells() && matches!(options.table_fallback, super::TableFallback::Html) {
        return render_table_html(table);
    }

    let mut output = String::new();

    // Determine column count
    let col_count = table.column_count();
    if col_count == 0 {
        return String::new();
    }

    // Render rows
    for (i, row) in table.rows.iter().enumerate() {
        output.push('|');

        // For header row, prepend placeholder columns if header has fewer cells than data
        if i == 0 && row.cells.len() < col_count {
            let missing_cols = col_count - row.cells.len();
            for j in 0..missing_cols {
                // Use "#" for first missing column (likely row number), empty for others
                let placeholder = if j == 0 { "#" } else { "" };
                output.push_str(&format!(" {} |", placeholder));
            }
        }

        for cell in &row.cells {
            let text = render_cell_content(cell, options);
            output.push_str(&format!(" {} |", text));
        }

        // Pad data rows if they have fewer cells
        if i > 0 {
            for _ in row.cells.len()..col_count {
                output.push_str(" |");
            }
        }
        output.push('\n');

        // Add separator after first row (markdown tables always need header separator)
        // In markdown, the first row is always treated as header regardless of source formatting
        if i == 0 {
            output.push('|');
            for _ in 0..col_count {
                output.push_str(" --- |");
            }
            output.push('\n');
        }
    }

    output
}

/// Render a table as HTML (for complex layouts).
fn render_table_html(table: &Table) -> String {
    let mut html = String::from("<table>\n");

    for row in &table.rows {
        html.push_str("  <tr>\n");
        for cell in &row.cells {
            let tag = if cell.is_header || row.is_header {
                "th"
            } else {
                "td"
            };
            let mut attrs = String::new();
            if cell.col_span > 1 {
                attrs.push_str(&format!(" colspan=\"{}\"", cell.col_span));
            }
            if cell.row_span > 1 {
                attrs.push_str(&format!(" rowspan=\"{}\"", cell.row_span));
            }
            let text = cell.plain_text();
            html.push_str(&format!("    <{}{}>{}</{}>\n", tag, attrs, text, tag));
        }
        html.push_str("  </tr>\n");
    }

    html.push_str("</table>");
    html
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{Cell, HeadingLevel, Row, Section, TextStyle};

    #[test]
    fn test_basic_paragraph() {
        let para = Paragraph::with_text("Hello, World!");
        let options = RenderOptions::default();
        let md = render_paragraph(&para, &options);
        // Most punctuation is NOT escaped - only special in specific contexts
        // Only `\`, `` ` ``, `*`, `_`, `|` are always escaped
        assert_eq!(md, "Hello, World!");
    }

    #[test]
    fn test_heading() {
        let para = Paragraph::heading(HeadingLevel::H2, "Title");
        let options = RenderOptions::default();
        let md = render_paragraph(&para, &options);
        assert_eq!(md, "## Title");
    }

    #[test]
    fn test_formatted_text() {
        let mut para = Paragraph::new();
        para.runs.push(TextRun::styled("bold", TextStyle::bold()));
        para.runs.push(TextRun::plain(" and "));
        para.runs
            .push(TextRun::styled("italic", TextStyle::italic()));

        let options = RenderOptions::default();
        let md = render_paragraph(&para, &options);
        assert!(md.contains("**bold**"));
        assert!(md.contains("*italic*"));
    }

    #[test]
    fn test_hyperlink() {
        let mut para = Paragraph::new();
        para.runs
            .push(TextRun::link("click here", "https://example.com"));

        let options = RenderOptions::default();
        let md = render_paragraph(&para, &options);
        assert!(md.contains("[click here](https://example.com)"));
    }

    #[test]
    fn test_simple_table() {
        let mut table = Table::new();
        let mut header = Row::header(vec![Cell::header("A"), Cell::header("B")]);
        header.is_header = true;
        table.add_row(header);
        table.add_row(Row {
            cells: vec![Cell::with_text("1"), Cell::with_text("2")],
            is_header: false,
            height: None,
        });

        let options = RenderOptions::default();
        let md = render_table(&table, &options);
        assert!(md.contains("| A | B |"));
        assert!(md.contains("| --- | --- |"));
        assert!(md.contains("| 1 | 2 |"));
    }

    #[test]
    fn test_document_to_markdown() {
        let mut doc = Document::new();
        let mut section = Section::new(0);
        section.add_paragraph(Paragraph::heading(HeadingLevel::H1, "Test Document"));
        section.add_paragraph(Paragraph::with_text("This is a test."));
        doc.add_section(section);

        let options = RenderOptions::default();
        let md = to_markdown(&doc, &options).unwrap();
        assert!(md.contains("# Test Document"));
        // Period is NOT escaped (only special in ordered list context)
        assert!(md.contains("This is a test."));
    }

    #[test]
    fn test_frontmatter() {
        let mut doc = Document::new();
        doc.metadata.title = Some("Test Title".to_string());
        doc.metadata.author = Some("Test Author".to_string());

        let options = RenderOptions::new().with_frontmatter(true);
        let md = to_markdown(&doc, &options).unwrap();
        assert!(md.starts_with("---\n"));
        assert!(md.contains("title: \"Test Title\""));
        assert!(md.contains("author: \"Test Author\""));
    }

    #[test]
    fn test_table_cell_with_bold_text() {
        let mut table = Table::new();

        // Create header row
        let header = Row::header(vec![Cell::header("Header")]);
        table.add_row(header);

        // Create data row with bold text in cell
        let mut bold_para = Paragraph::new();
        bold_para
            .runs
            .push(TextRun::styled("ClusterPlex v5.0", TextStyle::bold()));

        let cell = Cell {
            content: vec![bold_para],
            nested_tables: Vec::new(),
            col_span: 1,
            row_span: 1,
            alignment: crate::model::CellAlignment::Left,
            vertical_alignment: crate::model::VerticalAlignment::Top,
            is_header: false,
            background: None,
        };

        table.add_row(Row {
            cells: vec![cell],
            is_header: false,
            height: None,
        });

        let options = RenderOptions::default();
        let md = render_table(&table, &options);

        // Should contain bold formatting
        assert!(
            md.contains("**ClusterPlex v5.0**"),
            "Expected bold formatting, got: {}",
            md
        );
    }

    #[test]
    fn test_table_cell_with_italic_text() {
        let mut table = Table::new();

        // Create header row
        let header = Row::header(vec![Cell::header("Header")]);
        table.add_row(header);

        // Create data row with italic text in cell
        let mut italic_para = Paragraph::new();
        italic_para
            .runs
            .push(TextRun::styled("emphasis", TextStyle::italic()));

        let cell = Cell {
            content: vec![italic_para],
            nested_tables: Vec::new(),
            col_span: 1,
            row_span: 1,
            alignment: crate::model::CellAlignment::Left,
            vertical_alignment: crate::model::VerticalAlignment::Top,
            is_header: false,
            background: None,
        };

        table.add_row(Row {
            cells: vec![cell],
            is_header: false,
            height: None,
        });

        let options = RenderOptions::default();
        let md = render_table(&table, &options);

        // Should contain italic formatting
        assert!(
            md.contains("*emphasis*"),
            "Expected italic formatting, got: {}",
            md
        );
    }

    #[test]
    fn test_table_cell_with_multiple_paragraphs() {
        let mut table = Table::new();

        // Create header row
        let header = Row::header(vec![Cell::header("Steps")]);
        table.add_row(header);

        // Create data row with multiple paragraphs in cell
        let para1 = Paragraph::with_text("1. Active 서버 어댑터 Disable");
        let para2 = Paragraph::with_text("2. Standby 서버 어댑터 Enable");

        let cell = Cell {
            content: vec![para1, para2],
            nested_tables: Vec::new(),
            col_span: 1,
            row_span: 1,
            alignment: crate::model::CellAlignment::Left,
            vertical_alignment: crate::model::VerticalAlignment::Top,
            is_header: false,
            background: None,
        };

        table.add_row(Row {
            cells: vec![cell],
            is_header: false,
            height: None,
        });

        let options = RenderOptions::default();
        let md = render_table(&table, &options);

        // Should contain <br> between paragraphs
        assert!(
            md.contains("<br>"),
            "Expected <br> separator between paragraphs, got: {}",
            md
        );
        assert!(
            md.contains("1. Active"),
            "Expected first paragraph content, got: {}",
            md
        );
        assert!(
            md.contains("2. Standby"),
            "Expected second paragraph content, got: {}",
            md
        );
    }

    #[test]
    fn test_table_cell_with_mixed_formatting() {
        let mut table = Table::new();

        // Create header row
        let header = Row::header(vec![Cell::header("OS"), Cell::header("리소스 타입")]);
        table.add_row(header);

        // Create data row with bold header label and normal value
        let mut para1 = Paragraph::new();
        para1.runs.push(TextRun::styled("OS", TextStyle::bold()));

        let mut para2 = Paragraph::new();
        para2.runs.push(TextRun::plain("Linux"));

        let cell1 = Cell {
            content: vec![para1],
            nested_tables: Vec::new(),
            col_span: 1,
            row_span: 1,
            alignment: crate::model::CellAlignment::Left,
            vertical_alignment: crate::model::VerticalAlignment::Top,
            is_header: false,
            background: None,
        };

        let cell2 = Cell {
            content: vec![para2],
            nested_tables: Vec::new(),
            col_span: 1,
            row_span: 1,
            alignment: crate::model::CellAlignment::Left,
            vertical_alignment: crate::model::VerticalAlignment::Top,
            is_header: false,
            background: None,
        };

        table.add_row(Row {
            cells: vec![cell1, cell2],
            is_header: false,
            height: None,
        });

        let options = RenderOptions::default();
        let md = render_table(&table, &options);

        // Should contain both bold and plain text
        assert!(md.contains("**OS**"), "Expected bold OS, got: {}", md);
        assert!(md.contains("Linux"), "Expected Linux text, got: {}", md);
    }
}
