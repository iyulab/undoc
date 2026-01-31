//! Markdown renderer implementation.

use std::collections::HashMap;

use crate::error::Result;
use crate::model::{
    Block, CellAlignment, Document, HeadingLevel, Paragraph, RevisionType, Table, TextRun,
};

use super::heading_analyzer::{HeadingAnalyzer, HeadingDecision};
use super::options::{RenderOptions, RevisionHandling};

/// Map of resource IDs to their filenames
type ResourceMap = HashMap<String, String>;

/// Maximum character length for a heading.
/// Text longer than this is unlikely to be a semantic heading.
const MAX_HEADING_TEXT_LENGTH: usize = 80;

/// Common list/bullet markers including Korean characters.
/// Used to detect paragraphs that should not be rendered as headings.
const LIST_MARKERS: &[char] = &[
    // ASCII markers
    '-', '*', '>', // Korean/Asian markers
    '※', '○', '•', '●', '◦', '◎', '□', '■', '▪', '▫', '◇', '◆', '☐', '☑', '☒', '✓', '✗',
    'ㅇ', // Korean jamo (circle)
    'ㆍ', // Korean middle dot (U+318D)
    '·',  // Middle dot (U+00B7)
    '∙',  // Bullet operator (U+2219)
    // Arrows (commonly used as list markers in Korean documents)
    '→', '←', '↔', '⇒', '⇐', '⇔', '►', '▶', '▷', '◀', '◁', '▻',
];

/// Convert a Document to Markdown.
pub fn to_markdown(doc: &Document, options: &RenderOptions) -> Result<String> {
    // If heading analysis is enabled, use the analyzer
    if let Some(ref config) = options.heading_config {
        return to_markdown_with_analyzer(doc, options, config);
    }

    // Standard rendering without sophisticated heading analysis
    to_markdown_standard(doc, options)
}

/// Build a map from resource IDs to their suggested filenames.
fn build_resource_map(doc: &Document) -> ResourceMap {
    doc.resources
        .iter()
        .map(|(id, resource)| (id.clone(), resource.suggested_filename(id)))
        .collect()
}

/// Get image path from resource ID, resolving to actual filename if available.
fn resolve_image_path(resource_id: &str, resource_map: &ResourceMap, prefix: &str) -> String {
    let filename = resource_map
        .get(resource_id)
        .cloned()
        .unwrap_or_else(|| resource_id.to_string());
    format!("{}{}", prefix, filename)
}

/// Standard markdown conversion (without heading analyzer).
fn to_markdown_standard(doc: &Document, options: &RenderOptions) -> Result<String> {
    let mut output = String::new();
    let resource_map = build_resource_map(doc);

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
                    let md = render_paragraph(para, options, None, &resource_map);
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
                    output.push_str(&render_table(table, options, &resource_map));
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
                    let path =
                        resolve_image_path(resource_id, &resource_map, &options.image_path_prefix);
                    output.push_str(&format!("![{}]({})\n\n", alt, path));
                }
            }
        }

        // Render notes if present (for PPTX)
        if let Some(ref notes) = section.notes {
            if !notes.is_empty() {
                output.push_str("\n> **Notes:**\n");
                for note in notes {
                    let text = render_paragraph(note, options, None, &resource_map);
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

/// Convert a Document to Markdown with sophisticated heading analysis.
fn to_markdown_with_analyzer(
    doc: &Document,
    options: &RenderOptions,
    config: &super::heading_analyzer::HeadingConfig,
) -> Result<String> {
    // Run two-pass heading analysis
    let mut analyzer = HeadingAnalyzer::new(config.clone());
    let decisions = analyzer.analyze(doc);

    let mut output = String::new();
    let resource_map = build_resource_map(doc);

    // Add frontmatter if requested
    if options.include_frontmatter {
        output.push_str(&render_frontmatter(doc));
    }

    // Render each section with pre-computed heading decisions
    for (section_idx, section) in doc.sections.iter().enumerate() {
        // Add section name as heading if present
        if let Some(ref name) = section.name {
            if section_idx > 0 {
                output.push_str("\n---\n\n");
            }
            output.push_str(&format!("## {}\n\n", name));
        }

        // Get decisions for this section
        let section_decisions = decisions.get(section_idx);

        // Track paragraph index within section (only count Paragraph blocks)
        let mut para_idx = 0;

        // Render content blocks
        for block in &section.content {
            match block {
                Block::Paragraph(para) => {
                    // Get the pre-computed decision for this paragraph
                    let decision = section_decisions.and_then(|d| d.get(para_idx)).copied();
                    let md = render_paragraph(para, options, decision, &resource_map);

                    if !md.is_empty() || options.include_empty_paragraphs {
                        output.push_str(&md);
                        if options.paragraph_spacing {
                            output.push_str("\n\n");
                        } else {
                            output.push('\n');
                        }
                    }

                    para_idx += 1;
                }
                Block::Table(table) => {
                    output.push_str(&render_table(table, options, &resource_map));
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
                    let path =
                        resolve_image_path(resource_id, &resource_map, &options.image_path_prefix);
                    output.push_str(&format!("![{}]({})\n\n", alt, path));
                }
            }
        }

        // Render notes if present (for PPTX)
        if let Some(ref notes) = section.notes {
            if !notes.is_empty() {
                output.push_str("\n> **Notes:**\n");
                for note in notes {
                    // Notes don't use heading analysis
                    let text = render_paragraph(note, options, None, &resource_map);
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

    // Core metadata
    if let Some(ref title) = meta.title {
        fm.push_str(&format!("title: \"{}\"\n", escape_yaml(title)));
    }
    if let Some(ref author) = meta.author {
        fm.push_str(&format!("author: \"{}\"\n", escape_yaml(author)));
    }
    if let Some(ref subject) = meta.subject {
        fm.push_str(&format!("subject: \"{}\"\n", escape_yaml(subject)));
    }

    // Dates
    if let Some(ref created) = meta.created {
        fm.push_str(&format!("created: \"{}\"\n", created));
    }
    if let Some(ref modified) = meta.modified {
        fm.push_str(&format!("modified: \"{}\"\n", modified));
    }

    // Document statistics
    if let Some(page_count) = meta.page_count {
        // Use appropriate label based on document type (inferred from section names)
        let label = if doc
            .sections
            .first()
            .and_then(|s| s.name.as_ref())
            .is_some_and(|n| n.starts_with("Slide"))
        {
            "slides"
        } else if doc
            .sections
            .first()
            .and_then(|s| s.name.as_ref())
            .is_some_and(|n| n.starts_with("Sheet"))
        {
            "sheets"
        } else {
            "pages"
        };
        fm.push_str(&format!("{}: {}\n", label, page_count));
    }
    if let Some(word_count) = meta.word_count {
        fm.push_str(&format!("words: {}\n", word_count));
    }

    // Keywords as YAML list
    if !meta.keywords.is_empty() {
        fm.push_str("keywords:\n");
        for keyword in &meta.keywords {
            fm.push_str(&format!("  - \"{}\"\n", escape_yaml(keyword)));
        }
    }

    // Application info
    if let Some(ref app) = meta.application {
        fm.push_str(&format!("application: \"{}\"\n", escape_yaml(app)));
    }

    fm.push_str("---\n\n");
    fm
}

/// Escape special characters in YAML strings.
fn escape_yaml(s: &str) -> String {
    s.replace('\\', "\\\\").replace('"', "\\\"")
}

/// Render a paragraph to Markdown.
///
/// If `heading_decision` is provided (from HeadingAnalyzer), it takes precedence
/// over the simple heading detection logic.
fn render_paragraph(
    para: &Paragraph,
    options: &RenderOptions,
    heading_decision: Option<HeadingDecision>,
    resource_map: &ResourceMap,
) -> String {
    let mut output = String::new();

    // Merge adjacent runs with the same style to avoid issues like:
    // **시** **험** **합** -> **시험합**
    let merged_para = para.with_merged_runs();

    // Handle heading based on decision or fallback to simple logic
    let effective_heading: Option<HeadingLevel> = if let Some(decision) = heading_decision {
        // Use analyzer's decision
        match decision {
            HeadingDecision::Explicit(level) | HeadingDecision::Inferred(level) => Some(level),
            HeadingDecision::Demoted | HeadingDecision::None => None,
        }
    } else {
        // Fallback: simple heading detection (legacy behavior)
        if merged_para.heading.is_heading() {
            let plain_text = merged_para.plain_text();
            let trimmed_text = plain_text.trim();

            // Check if paragraph looks like a list item (starts with list-like markers)
            let looks_like_list_item = trimmed_text
                .chars()
                .next()
                .is_some_and(|c| LIST_MARKERS.contains(&c));

            // Check if text is too long to be a meaningful heading
            let text_too_long = trimmed_text.chars().count() > MAX_HEADING_TEXT_LENGTH;

            // Apply heading only if it's truly semantic
            if !looks_like_list_item && !text_too_long {
                let level = merged_para.heading.level().min(options.max_heading_level);
                Some(HeadingLevel::from_number(level))
            } else {
                None
            }
        } else {
            None
        }
    };

    // Apply heading formatting if determined
    if let Some(level) = effective_heading {
        let capped_level = level.level().min(options.max_heading_level);
        if capped_level > 0 {
            output.push_str(&"#".repeat(capped_level as usize));
            output.push(' ');
        }
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
        let path = resolve_image_path(&image.resource_id, resource_map, &options.image_path_prefix);
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
    // Handle tracked changes based on revision_handling option
    match (&run.revision, &options.revision_handling) {
        // AcceptAll: show inserted text, hide deleted text
        (RevisionType::Deleted, RevisionHandling::AcceptAll) => {
            // Return only break markers if present
            if run.page_break {
                return "\n\n---\n\n".to_string();
            } else if run.line_break && options.preserve_line_breaks {
                return "  \n".to_string();
            }
            return String::new();
        }
        // RejectAll: show deleted text, hide inserted text
        (RevisionType::Inserted, RevisionHandling::RejectAll) => {
            // Return only break markers if present
            if run.page_break {
                return "\n\n---\n\n".to_string();
            } else if run.line_break && options.preserve_line_breaks {
                return "  \n".to_string();
            }
            return String::new();
        }
        // ShowMarkup or normal text: continue with rendering
        _ => {}
    }

    // Handle empty runs with line/page breaks
    if run.text.is_empty() {
        if run.page_break {
            return "\n\n---\n\n".to_string();
        } else if run.line_break && options.preserve_line_breaks {
            return "  \n".to_string();
        } else {
            return String::new();
        }
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

    // Apply revision markup for ShowMarkup mode
    match (&run.revision, &options.revision_handling) {
        (RevisionType::Deleted, RevisionHandling::ShowMarkup) => {
            // Show deleted text with strikethrough
            text = format!("~~{}~~", text);
        }
        (RevisionType::Inserted, RevisionHandling::ShowMarkup) => {
            // Show inserted text with underline markers (using HTML since Markdown lacks insert markup)
            text = format!("<ins>{}</ins>", text);
        }
        _ => {}
    }

    // Append page break (horizontal rule) or line break
    if run.page_break {
        text.push_str("\n\n---\n\n");
    } else if run.line_break && options.preserve_line_breaks {
        text.push_str("  \n");
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
///
/// Note: Nested tables are NOT rendered here to avoid content duplication.
/// The nested_tables field contains tables that are already structurally
/// separate from cell.content. Rendering both would cause duplication
/// in documents where Word places the same content in both locations.
fn render_cell_content(
    cell: &crate::model::Cell,
    options: &RenderOptions,
    resource_map: &ResourceMap,
) -> String {
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

        // Render inline images from paragraph (like render_paragraph does)
        for image in &para.images {
            let alt = image.alt_text.as_deref().unwrap_or("image");
            let path =
                resolve_image_path(&image.resource_id, resource_map, &options.image_path_prefix);
            parts.push(format!("![{}]({})", alt, path));
        }
    }

    // NOTE: nested_tables are intentionally NOT rendered here.
    // They are extracted as separate Table blocks during parsing and should
    // be rendered independently to preserve structure and avoid duplication.
    // See: render_nested_tables_as_blocks() for proper nested table rendering.

    // Join paragraphs with <br> for markdown table cells
    let text = parts.join("<br>");

    // Only replace newlines - pipes are already escaped by escape_markdown in render_run
    text.replace('\n', " ")
}

/// Get column alignments from the first data row (or first row if no data rows).
/// Returns a vector of alignments for each column.
fn get_column_alignments(table: &Table, col_count: usize) -> Vec<CellAlignment> {
    // Try to get alignments from the first data row (non-header)
    // If no data rows, use the first row
    let source_row = table
        .rows
        .iter()
        .find(|r| !r.is_header)
        .or_else(|| table.rows.first());

    let mut alignments = Vec::with_capacity(col_count);

    if let Some(row) = source_row {
        for cell in &row.cells {
            // Add alignment for each column the cell spans
            for _ in 0..cell.col_span {
                alignments.push(cell.alignment);
            }
        }
    }

    // Fill remaining columns with Left alignment
    while alignments.len() < col_count {
        alignments.push(CellAlignment::Left);
    }

    alignments.truncate(col_count);
    alignments
}

/// Render a table to Markdown.
fn render_table(table: &Table, options: &RenderOptions, resource_map: &ResourceMap) -> String {
    if table.is_empty() {
        return String::new();
    }

    // Check if we need HTML fallback
    if table.has_merged_cells() && matches!(options.table_fallback, super::TableFallback::Html) {
        return render_table_html(table);
    }

    let mut output = String::new();
    let mut nested_tables: Vec<&Table> = Vec::new();

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
            let text = render_cell_content(cell, options, resource_map);
            output.push_str(&format!(" {} |", text));

            // Collect nested tables for rendering after the main table
            for nested in &cell.nested_tables {
                nested_tables.push(nested);
            }
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
            // Collect alignments from cells, filling with Left for missing columns
            let alignments = get_column_alignments(table, col_count);
            for alignment in &alignments {
                let separator = match alignment {
                    CellAlignment::Center => " :---: |",
                    CellAlignment::Right => " ---: |",
                    CellAlignment::Left => " --- |",
                };
                output.push_str(separator);
            }
            output.push('\n');
        }
    }

    // Render nested tables after the main table
    // This preserves their structure instead of flattening into cell content
    for nested in nested_tables {
        output.push('\n');
        output.push_str(&render_table(nested, options, resource_map));
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
    use crate::model::{Cell, HeadingLevel, RevisionType, Row, Section, TextStyle};

    /// Helper to create an empty resource map for tests
    fn empty_resource_map() -> ResourceMap {
        HashMap::new()
    }

    #[test]
    fn test_basic_paragraph() {
        let para = Paragraph::with_text("Hello, World!");
        let options = RenderOptions::default();
        let md = render_paragraph(&para, &options, None, &empty_resource_map());
        // Most punctuation is NOT escaped - only special in specific contexts
        // Only `\`, `` ` ``, `*`, `_`, `|` are always escaped
        assert_eq!(md, "Hello, World!");
    }

    #[test]
    fn test_heading() {
        let para = Paragraph::heading(HeadingLevel::H2, "Title");
        let options = RenderOptions::default();
        let md = render_paragraph(&para, &options, None, &empty_resource_map());
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
        let md = render_paragraph(&para, &options, None, &empty_resource_map());
        assert!(md.contains("**bold**"));
        assert!(md.contains("*italic*"));
    }

    #[test]
    fn test_hyperlink() {
        let mut para = Paragraph::new();
        para.runs
            .push(TextRun::link("click here", "https://example.com"));

        let options = RenderOptions::default();
        let md = render_paragraph(&para, &options, None, &empty_resource_map());
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
        let md = render_table(&table, &options, &empty_resource_map());
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
    fn test_korean_bullet_marker_not_heading() {
        // Paragraphs starting with Korean bullet markers should not be headings
        let para = Paragraph::heading(HeadingLevel::H2, "ㅇ항목 내용입니다");
        let options = RenderOptions::default();
        let md = render_paragraph(&para, &options, None, &empty_resource_map());

        assert!(
            !md.contains("##"),
            "Korean bullet marker should not be heading: {}",
            md
        );
        assert!(
            md.contains("ㅇ항목"),
            "Content should still be present: {}",
            md
        );
    }

    #[test]
    fn test_long_text_not_heading() {
        // Very long text (>80 chars) should not be treated as a heading
        let long_text = "이것은 매우 긴 문장입니다. 제목으로 사용하기에는 너무 길어서 본문으로 처리되어야 합니다. 일반적인 제목은 짧고 간결해야 하며, 본문과 구분되어야 합니다.";
        assert!(
            long_text.chars().count() > 80,
            "Test text should be longer than 80 chars"
        );

        let para = Paragraph::heading(HeadingLevel::H3, long_text);
        let options = RenderOptions::default();
        let md = render_paragraph(&para, &options, None, &empty_resource_map());

        assert!(
            !md.contains("###"),
            "Long text should not have heading markers: {}",
            md
        );
        assert!(
            md.contains("이것은 매우"),
            "Content should still be present: {}",
            md
        );
    }

    #[test]
    fn test_max_heading_level_capped() {
        // Heading levels beyond max (4) should be capped
        let para = Paragraph::heading(HeadingLevel::H6, "Deep Heading");
        let options = RenderOptions::default();
        let md = render_paragraph(&para, &options, None, &empty_resource_map());

        // Default max_heading_level is now 4
        assert!(
            md.contains("#### Deep Heading"),
            "Heading level 6 should be capped to 4: {}",
            md
        );
        assert!(
            !md.contains("######"),
            "Should not have 6 hash marks: {}",
            md
        );
    }

    #[test]
    fn test_arrow_marker_not_heading() {
        // Paragraphs starting with arrow markers should not be headings
        let para = Paragraph::heading(HeadingLevel::H2, "→ 다음 단계로 이동");
        let options = RenderOptions::default();
        let md = render_paragraph(&para, &options, None, &empty_resource_map());

        assert!(
            !md.contains("##"),
            "Arrow marker should not be heading: {}",
            md
        );
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
        let md = render_table(&table, &options, &empty_resource_map());

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
        let md = render_table(&table, &options, &empty_resource_map());

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
        let md = render_table(&table, &options, &empty_resource_map());

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
        let md = render_table(&table, &options, &empty_resource_map());

        // Should contain both bold and plain text
        assert!(md.contains("**OS**"), "Expected bold OS, got: {}", md);
        assert!(md.contains("Linux"), "Expected Linux text, got: {}", md);
    }

    #[test]
    fn test_line_break_rendering() {
        let mut para = Paragraph::new();
        para.runs.push(TextRun {
            text: "First line".to_string(),
            style: TextStyle::default(),
            hyperlink: None,
            line_break: true,
            page_break: false,
            revision: RevisionType::None,
        });
        para.runs.push(TextRun::plain("Second line"));

        // Without preserve_line_breaks option
        let options = RenderOptions::default();
        let md = render_paragraph(&para, &options, None, &empty_resource_map());
        assert!(
            !md.contains("  \n"),
            "Should not contain line break when preserve_line_breaks is false: {}",
            md
        );

        // With preserve_line_breaks option
        let options_with_breaks = RenderOptions::new().with_preserve_breaks(true);
        let md_with_breaks =
            render_paragraph(&para, &options_with_breaks, None, &empty_resource_map());
        assert!(
            md_with_breaks.contains("First line  \n"),
            "Should contain Markdown line break: {}",
            md_with_breaks
        );
        assert!(
            md_with_breaks.contains("Second line"),
            "Should contain second line: {}",
            md_with_breaks
        );
    }

    #[test]
    fn test_table_cell_alignment_rendering() {
        let mut table = Table::new();

        // Create header row
        let header = Row::header(vec![
            Cell::header("Left"),
            Cell::header("Center"),
            Cell::header("Right"),
        ]);
        table.add_row(header);

        // Create data row with different alignments
        let left_cell = Cell {
            content: vec![Paragraph::with_text("L")],
            nested_tables: Vec::new(),
            col_span: 1,
            row_span: 1,
            alignment: CellAlignment::Left,
            vertical_alignment: crate::model::VerticalAlignment::Top,
            is_header: false,
            background: None,
        };

        let center_cell = Cell {
            content: vec![Paragraph::with_text("C")],
            nested_tables: Vec::new(),
            col_span: 1,
            row_span: 1,
            alignment: CellAlignment::Center,
            vertical_alignment: crate::model::VerticalAlignment::Top,
            is_header: false,
            background: None,
        };

        let right_cell = Cell {
            content: vec![Paragraph::with_text("R")],
            nested_tables: Vec::new(),
            col_span: 1,
            row_span: 1,
            alignment: CellAlignment::Right,
            vertical_alignment: crate::model::VerticalAlignment::Top,
            is_header: false,
            background: None,
        };

        table.add_row(Row {
            cells: vec![left_cell, center_cell, right_cell],
            is_header: false,
            height: None,
        });

        let options = RenderOptions::default();
        let md = render_table(&table, &options, &empty_resource_map());

        // Should contain alignment markers in separator row
        assert!(
            md.contains("| --- |"),
            "Expected left alignment marker, got: {}",
            md
        );
        assert!(
            md.contains("| :---: |"),
            "Expected center alignment marker, got: {}",
            md
        );
        assert!(
            md.contains("| ---: |"),
            "Expected right alignment marker, got: {}",
            md
        );
    }
}
