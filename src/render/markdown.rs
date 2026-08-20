//! Markdown renderer implementation.

use std::collections::HashMap;

use crate::detect::FormatType;
use crate::error::Result;
use crate::model::{
    Block, Cell, CellAlignment, Document, HeadingLevel, Paragraph, RevisionType, Table, TextRun,
};

use super::heading_analyzer::{HeadingAnalyzer, HeadingDecision};
use super::options::{RenderOptions, RevisionHandling, SectionMarkerStyle};

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

/// Render a single section to Markdown.
///
/// This is the streaming counterpart to [`to_markdown`]: it renders one
/// section at a time without requiring a full `Document`. Heading analysis
/// is not available on this path (requires all sections); heading levels come
/// from `section.content` directly.
///
/// `resource_map` maps resource IDs to filenames (e.g., `"rId1" → "image1.png"`).
/// In the streaming path, supply the `image_map` from [`crate::ParseEvent::DocumentStart`].
pub fn render_section_to_string(
    section: &crate::model::Section,
    section_index: usize,
    doc_format: crate::detect::FormatType,
    options: &RenderOptions,
    resource_map: &HashMap<String, String>,
) -> String {
    let mut output = String::new();
    render_section_impl(
        section,
        section_index,
        doc_format,
        options,
        resource_map,
        &mut output,
    );
    output
}

/// Build a map from resource IDs to their suggested filenames.
fn build_resource_map(doc: &Document) -> ResourceMap {
    doc.resources
        .iter()
        .map(|(id, resource)| (id.clone(), resource.suggested_filename(id)))
        .collect()
}

/// Render header or footer paragraphs as a blockquote with italic label.
fn render_header_footer(
    label: &str,
    paragraphs: &[Paragraph],
    options: &RenderOptions,
    resource_map: &ResourceMap,
    output: &mut String,
) {
    let texts: Vec<String> = paragraphs
        .iter()
        .map(|p| render_paragraph(p, options, None, resource_map))
        .filter(|t| !t.is_empty())
        .collect();
    if !texts.is_empty() {
        output.push_str(&format!("> *{}: {}*\n\n", label, texts.join(" | ")));
    }
}

/// Get image path from resource ID, resolving to actual filename if available.
fn resolve_image_path(resource_id: &str, resource_map: &ResourceMap, prefix: &str) -> String {
    let filename = resource_map
        .get(resource_id)
        .cloned()
        .unwrap_or_else(|| resource_id.to_string());
    format_link_destination(&format!("{}{}", prefix, filename))
}

/// Format a link/image destination for `[text](destination)` syntax, wrapping it in
/// `<...>` when it contains a character CommonMark's bare-parenthesis destination form
/// forbids. A destination with a raw space is not valid CommonMark outside `<...>` at
/// all -- `pulldown-cmark` does not even produce a `Link`/`Image` event for it, so a
/// consumer sees the brackets as literal text instead of a link. A hyperlink target or
/// an `image_path_prefix`-built path can legitimately contain a space (a local file
/// path, for instance).
fn format_link_destination(url: &str) -> String {
    if url.contains(' ') || url.contains(['<', '>']) {
        // Backslash-escape, not percent-encode: a link destination is data, and
        // percent-encoding `<`/`>` would silently change the target (e.g. a real
        // file path) instead of just escaping it for Markdown syntax.
        let escaped = url
            .replace('\\', "\\\\")
            .replace('<', "\\<")
            .replace('>', "\\>");
        format!("<{}>", escaped)
    } else {
        url.to_string()
    }
}

/// Build a section boundary marker comment for PPTX slides or XLSX sheets.
///
/// Returns an empty string when markers are disabled or the format is DOCX.
fn section_marker(
    format: FormatType,
    style: SectionMarkerStyle,
    idx: usize,
    name: Option<&str>,
) -> String {
    if style == SectionMarkerStyle::None {
        return String::new();
    }
    let n = idx + 1;
    match format {
        FormatType::Pptx => match name.filter(|s| !s.is_empty()) {
            Some(name) => format!("<!-- slide {}: {} -->", n, name),
            None => format!("<!-- slide {} -->", n),
        },
        FormatType::Xlsx => match name.filter(|s| !s.is_empty()) {
            Some(name) => format!("<!-- sheet {}: {} -->", n, name),
            None => format!("<!-- sheet {} -->", n),
        },
        FormatType::Docx => String::new(),
    }
}

/// Core per-section render, shared by the batch and streaming paths.
fn render_section_impl(
    section: &crate::model::Section,
    section_index: usize,
    doc_format: FormatType,
    options: &RenderOptions,
    resource_map: &ResourceMap,
    output: &mut String,
) {
    let marker = section_marker(
        doc_format,
        options.section_markers,
        section_index,
        section.name.as_deref(),
    );
    if !marker.is_empty() {
        output.push_str(&marker);
        output.push_str("\n\n");
    }

    if let Some(ref name) = section.name {
        if section_index > 0 {
            output.push_str("\n---\n\n");
        }
        output.push_str(&format!("## {}\n\n", name));
    }

    if options.include_headers_footers {
        if let Some(ref header) = section.header {
            render_header_footer("Header", header, options, resource_map, output);
        }
    }

    for (block_idx, block) in section.content.iter().enumerate() {
        match block {
            Block::Paragraph(para) => {
                let md = render_paragraph(para, options, None, resource_map);
                if !md.is_empty() || options.include_empty_paragraphs {
                    output.push_str(&md);
                    let in_list = para.list_info.is_some();
                    let tight = in_list && next_block_continues_list(&section.content, block_idx);
                    if tight {
                        output.push('\n');
                    } else if options.paragraph_spacing {
                        output.push_str("\n\n");
                    } else {
                        output.push('\n');
                    }
                }
            }
            Block::Table(table) => {
                output.push_str(&render_table(table, options, resource_map));
                output.push_str("\n\n");
            }
            Block::PageBreak => {
                if options.emit_page_breaks {
                    push_thematic_break(output);
                }
            }
            Block::SectionBreak => {
                push_thematic_break(output);
            }
            Block::Image {
                resource_id,
                alt_text,
                ..
            } => {
                let alt = alt_text.as_deref().unwrap_or("image");
                let path =
                    resolve_image_path(resource_id, resource_map, &options.image_path_prefix);
                output.push_str(&format!("![{}]({})\n\n", alt, path));
            }
        }
    }

    if options.include_headers_footers {
        if let Some(ref footer) = section.footer {
            render_header_footer("Footer", footer, options, resource_map, output);
        }
    }

    if let Some(ref notes) = section.notes {
        if !notes.is_empty() {
            output.push_str("\n> **Notes:**\n");
            for note in notes {
                let text = render_paragraph(note, options, None, resource_map);
                if !text.is_empty() {
                    output.push_str(&format!("> {}\n", text));
                }
            }
            output.push('\n');
        }
    }
}

/// Standard markdown conversion (without heading analyzer).
fn to_markdown_standard(doc: &Document, options: &RenderOptions) -> Result<String> {
    let mut output = String::new();
    let resource_map = build_resource_map(doc);

    if options.include_frontmatter {
        output.push_str(&render_frontmatter(doc));
    }

    for (i, section) in doc.sections.iter().enumerate() {
        render_section_impl(section, i, doc.format, options, &resource_map, &mut output);
    }

    Ok(finalize(output, options))
}

/// Finish a rendered document: apply the configured cleanup, if any.
///
/// `RenderOptions::cleanup == None` means no post-processing, and nothing here may quietly
/// contradict that. Blank-line collapsing used to run unconditionally on the grounds that it
/// is lossless under CommonMark. That holds for the *rendered* result, but not for the
/// Markdown itself: a consumer diffing the output, citing it by line number, or chunking it
/// for retrieval sees different text. Collapsing is part of cleanup's whitespace
/// normalization now, which is also what makes Markdown and text output agree.
///
/// Trimming the document's own ends stays unconditional — those blank lines are the
/// renderer's scaffolding between sections, not content it was given.
fn finalize(output: String, options: &RenderOptions) -> String {
    let processed = match options.cleanup {
        Some(ref cleanup) => super::cleanup::clean_text(&output, cleanup),
        None => output,
    };

    processed.trim().to_string()
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
        // Section boundary marker (PPTX/XLSX only, opt-in)
        let marker = section_marker(
            doc.format,
            options.section_markers,
            section_idx,
            section.name.as_deref(),
        );
        if !marker.is_empty() {
            output.push_str(&marker);
            output.push_str("\n\n");
        }

        // Add section name as heading if present
        if let Some(ref name) = section.name {
            if section_idx > 0 {
                output.push_str("\n---\n\n");
            }
            output.push_str(&format!("## {}\n\n", name));
        }

        // Render header if present (DOCX)
        if options.include_headers_footers {
            if let Some(ref header) = section.header {
                render_header_footer("Header", header, options, &resource_map, &mut output);
            }
        }

        // Get decisions for this section
        let section_decisions = decisions.get(section_idx);

        // Track paragraph index within section (only count Paragraph blocks)
        let mut para_idx = 0;

        // Render content blocks
        for (block_idx, block) in section.content.iter().enumerate() {
            match block {
                Block::Paragraph(para) => {
                    // Get the pre-computed decision for this paragraph
                    let decision = section_decisions.and_then(|d| d.get(para_idx)).copied();
                    let md = render_paragraph(para, options, decision, &resource_map);

                    if !md.is_empty() || options.include_empty_paragraphs {
                        output.push_str(&md);
                        let in_list = para.list_info.is_some();
                        let tight =
                            in_list && next_block_continues_list(&section.content, block_idx);
                        if tight {
                            output.push('\n');
                        } else if options.paragraph_spacing {
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
                    if options.emit_page_breaks {
                        push_thematic_break(&mut output);
                    }
                }
                Block::SectionBreak => {
                    push_thematic_break(&mut output);
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

        // Render footer if present (DOCX)
        if options.include_headers_footers {
            if let Some(ref footer) = section.footer {
                render_header_footer("Footer", footer, options, &resource_map, &mut output);
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

    Ok(finalize(output, options))
}

/// Render YAML frontmatter from document metadata.
fn render_frontmatter(doc: &Document) -> String {
    let mut fm = String::from("---\n");
    let meta = &doc.metadata;

    // Document format (docx / xlsx / pptx)
    fm.push_str(&format!("format: {}\n", doc.format.extension()));

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

    // List items are never headings — suppress heading if list_info is set
    let effective_heading = if merged_para.list_info.is_some() {
        None
    } else {
        effective_heading
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

    // Decide whether the heading text is uniformly emphasized (a styling
    // artifact). Partial emphasis stays untouched.
    let suppress_heading_emphasis = effective_heading.is_some()
        && options.strip_redundant_emphasis_in_headings
        && all_runs_uniformly_bold(&merged_para);

    // Render text runs with smart spacing
    let run_ctx = RunContext {
        in_table_cell: false,
        suppress_emphasis: suppress_heading_emphasis,
    };
    for (i, run) in merged_para.runs.iter().enumerate() {
        let run_text = render_run(run, options, run_ctx);

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

/// True when the next block is another list paragraph — used to decide
/// whether the separator after the current list item should be tight
/// (`\n`) or loose (`\n\n`).
fn next_block_continues_list(blocks: &[Block], idx: usize) -> bool {
    matches!(
        blocks.get(idx + 1),
        Some(Block::Paragraph(p)) if p.list_info.is_some()
    )
}

/// True when every non-empty run in the paragraph carries `bold`. This
/// signals that the bold is a styling artifact (e.g. a Heading style with a
/// blanket `<w:b/>` run property), not author-applied emphasis on a slice.
fn all_runs_uniformly_bold(para: &Paragraph) -> bool {
    let mut saw_text = false;
    for run in &para.runs {
        if run.text.trim().is_empty() {
            continue;
        }
        saw_text = true;
        if !run.style.bold {
            return false;
        }
    }
    saw_text
}

/// Check if a character should NOT have a space before it.
fn is_no_space_before(c: char) -> bool {
    matches!(
        c,
        '.' | ',' | ':' | ';' | '!' | '?' | ')' | ']' | '}' | '"' | '\'' | '…'
    )
}

/// Context flags passed to [`render_run`].
#[derive(Debug, Clone, Copy, Default)]
struct RunContext {
    /// True when rendering inside a markdown table cell (`|` must be escaped).
    in_table_cell: bool,
    /// True when surrounding container (heading, header cell) is uniformly
    /// emphasized — bold/italic on individual runs is treated as a styling
    /// artifact and stripped.
    suppress_emphasis: bool,
}

/// Render a text run to Markdown.
/// Append a thematic break, separated from what precedes it by exactly one blank line.
///
/// A block-level element cannot see what came before it, so prepending a newline
/// unconditionally is a guess — and when the previous paragraph already ended with a blank
/// line, that guess produced a run of three newlines which a later pass had to collapse.
/// Normalizing the separator here means nothing downstream has to repair it.
fn push_thematic_break(output: &mut String) {
    while output.ends_with('\n') {
        output.pop();
    }
    if !output.is_empty() {
        output.push_str("\n\n");
    }
    output.push_str("---\n\n");
}

/// The break marker a run ends with, if the options ask for it.
///
/// A page break wins over a line break, and both are suffixes: they follow whatever the run
/// rendered, including a run that rendered nothing at all.
fn break_marker(run: &TextRun, options: &RenderOptions) -> &'static str {
    if run.page_break && options.emit_page_breaks {
        "\n\n---\n\n"
    } else if run.line_break && options.preserve_line_breaks {
        "  \n"
    } else {
        ""
    }
}

fn render_run(run: &TextRun, options: &RenderOptions, ctx: RunContext) -> String {
    // Handle tracked changes based on revision_handling option
    match (&run.revision, &options.revision_handling) {
        // AcceptAll: show inserted text, hide deleted text
        (RevisionType::Deleted, RevisionHandling::AcceptAll) => {
            // Hidden text still yields its break marker, if any.
            return break_marker(run, options).to_string();
        }
        // RejectAll: show deleted text, hide inserted text
        (RevisionType::Inserted, RevisionHandling::RejectAll) => {
            return break_marker(run, options).to_string();
        }
        // ShowMarkup or normal text: continue with rendering
        _ => {}
    }

    // A run with nothing to say — empty, or whitespace only — gets no markup: emphasis or a
    // link wrapped around nothing produces delimiters with no content between them. Its
    // whitespace is still emitted, because that whitespace is what separates its neighbours.
    let core = run.text.trim();
    if core.is_empty() {
        return format!("{}{}", run.text, break_marker(run, options));
    }

    // OOXML stores the whitespace around a word in the runs themselves — often in a run of
    // its own, or flagged with `xml:space="preserve"`. That whitespace sits *between* runs,
    // not inside the markup wrapping this one: `[label ](url)` puts the space inside the
    // link, and `**bold **` is not emphasis at all, because CommonMark refuses a closing
    // delimiter preceded by whitespace. So the markup goes around the trimmed text and the
    // whitespace is re-emitted outside it.
    let leading_len = run.text.len() - run.text.trim_start().len();
    let leading = &run.text[..leading_len];
    let trailing = &run.text[leading_len + core.len()..];

    let mut text = if options.escape_special_chars {
        escape_markdown(core, ctx.in_table_cell)
    } else {
        core.to_string()
    };

    // Apply formatting (innermost first)
    if run.style.code {
        text = format!("`{}`", text.replace('`', "\\`"));
    }
    if run.style.superscript {
        text = format!("<sup>{}</sup>", text);
    }
    if run.style.subscript {
        text = format!("<sub>{}</sub>", text);
    }
    if run.style.underline {
        text = format!("<u>{}</u>", text);
    }
    if run.style.strikethrough {
        text = format!("~~{}~~", text);
    }
    let effective_bold = run.style.bold && !ctx.suppress_emphasis;
    let effective_italic = run.style.italic && !ctx.suppress_emphasis;
    if effective_bold && effective_italic {
        text = format!("***{}***", text);
    } else if effective_bold {
        text = format!("**{}**", text);
    } else if effective_italic {
        text = format!("*{}*", text);
    }

    // Handle hyperlinks
    if let Some(ref url) = run.hyperlink {
        text = format!("[{}]({})", text, format_link_destination(url));
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

    format!("{leading}{text}{trailing}{}", break_marker(run, options))
}

/// Escape Markdown special characters.
///
/// Context-aware escaping - only escapes when the character could actually
/// trigger markdown formatting.
///
/// - `\` - always escape (escape character)
/// - `` ` `` - always escape (inline code)
/// - `|` - escape only inside table cells (where it is the column delimiter).
///   In regular paragraphs `|` is just a literal character.
/// - `*` - escape only when it could trigger emphasis (CommonMark flanking
///   rules). Intra-word `*` *can* still trigger emphasis, so it is escaped.
/// - `_` - same as `*`, **plus** an extra rule: CommonMark forbids intra-word
///   `_` from opening or closing emphasis, so identifiers like `snake_case`
///   are left intact.
///
/// Characters NOT escaped (only special in specific contexts):
/// - `()`, `[]`, `{}` - only special in link/image syntax `[text](url)`
/// - `#` - only special at start of line (headings)
/// - `+`, `-` - only special at start of line (lists) or `---` (rules)
/// - `!` - only special before `[` (images)
/// - `.` - only special in ordered lists at line start (e.g., "1.")
fn escape_markdown(s: &str, in_table_cell: bool) -> String {
    let mut result = String::with_capacity(s.len());
    let chars: Vec<char> = s.chars().collect();

    for (i, &c) in chars.iter().enumerate() {
        match c {
            '\\' | '`' => {
                result.push('\\');
                result.push(c);
            }
            '|' => {
                if in_table_cell {
                    result.push('\\');
                }
                result.push(c);
            }
            '*' | '_' => {
                let prev = if i > 0 { Some(chars[i - 1]) } else { None };
                let next = chars.get(i + 1).copied();

                // CommonMark: emphasis requires a matching pair of flanking
                // delimiters. If neither side can flank, no escape is needed.
                let after_opener = prev.is_none_or(|p| {
                    matches!(p, '(' | '[' | '{' | ':' | '-' | '/' | '\\') || p.is_whitespace()
                });
                let before_closer = next.is_none_or(|n| {
                    matches!(n, ')' | ']' | '}' | ':' | '-' | '/' | '\\') || n.is_whitespace()
                });

                // Underscore-only rule: CommonMark disallows intra-word `_`
                // from opening/closing emphasis, so identifiers like
                // `YESUNG_OMS_backup` never need escaping.
                let intra_word_underscore = c == '_'
                    && prev.is_some_and(|p| p.is_alphanumeric())
                    && next.is_some_and(|n| n.is_alphanumeric());

                if after_opener || before_closer || intra_word_underscore {
                    result.push(c);
                } else {
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
    is_header_cell: bool,
) -> String {
    let mut parts = Vec::new();

    for para in &cell.content {
        // Merge adjacent runs with same style (like render_paragraph does)
        let merged_para = para.with_merged_runs();
        let mut para_text = String::new();

        // Suppress uniform bold inside a header cell so the markdown header
        // row doesn't end up like `| **Field** |` when the docx applied a
        // blanket bold to every header cell.
        let suppress_emphasis = is_header_cell
            && options.strip_redundant_emphasis_in_headings
            && all_runs_uniformly_bold(&merged_para);
        let ctx = RunContext {
            in_table_cell: true,
            suppress_emphasis,
        };

        for (i, run) in merged_para.runs.iter().enumerate() {
            let run_text = render_run(run, options, ctx);

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
    let text = if options.preserve_line_breaks {
        text.replace("  \n", "<br>")
    } else {
        text
    };

    // Only replace line breaks - pipes are already escaped by escape_markdown
    // in render_run. Bare CR is treated as a line break too: a leaked CR
    // splits the pipe-table row in most Markdown renderers.
    text.replace("\r\n", "\n").replace(['\n', '\r'], " ")
}

/// Effective alignment for a cell.
///
/// Word stores cell alignment two places: directly on the cell
/// (`<w:tcPr>/<w:jc>`) and on the cell's paragraphs (`<w:pPr>/<w:jc>`).
/// Authors normally set only the latter, so when the explicit cell
/// alignment is `Left` (the default) we fall back to the first paragraph's
/// alignment to recover the visual intent.
fn effective_cell_alignment(cell: &crate::model::Cell) -> CellAlignment {
    if cell.alignment != CellAlignment::Left {
        return cell.alignment;
    }
    if let Some(para) = cell.content.first() {
        return match para.alignment {
            crate::model::TextAlignment::Center => CellAlignment::Center,
            crate::model::TextAlignment::Right => CellAlignment::Right,
            _ => CellAlignment::Left,
        };
    }
    CellAlignment::Left
}

/// Get column alignments from the first data row (or first row if no data rows).
/// Returns a vector of alignments for each column.
/// Column alignments for the separator row, read off the laid-out grid.
///
/// Taken from the grid rather than from a row's cell list so that a cell's alignment
/// lands on the column it actually occupies: a vertical merge from an earlier row
/// shifts a row's cells rightward, and counting cells would put every alignment after
/// it one column early.
fn get_column_alignments(
    table: &Table,
    grid: &[Vec<Option<&Cell>>],
    col_count: usize,
) -> Vec<CellAlignment> {
    // Prefer the first data row (non-header); fall back to the first row.
    let source = table.rows.iter().position(|r| !r.is_header).unwrap_or(0);

    let mut alignments = vec![CellAlignment::Left; col_count];
    if let Some(slots) = grid.get(source) {
        for (col, slot) in slots.iter().enumerate().take(col_count) {
            if let Some(cell) = slot {
                alignments[col] = effective_cell_alignment(cell);
            }
        }
        // A column covered by a merge takes the alignment of the cell covering it —
        // otherwise a centred merged cell gets a left-aligned tail.
        let mut carried = CellAlignment::Left;
        for (col, slot) in slots.iter().enumerate().take(col_count) {
            match slot {
                Some(cell) => carried = effective_cell_alignment(cell),
                None => alignments[col] = carried,
            }
        }
    }

    alignments
}

/// If `table` is a single-row, single-column table whose content is fully
/// emphasized (every non-empty run is bold), render it as a blockquote
/// callout: `> **...**`. Returns `None` when the heuristic does not apply.
fn render_callout_blockquote(
    table: &Table,
    options: &RenderOptions,
    resource_map: &ResourceMap,
) -> Option<String> {
    if table.rows.len() != 1 {
        return None;
    }
    let row = &table.rows[0];
    if row.cells.len() != 1 {
        return None;
    }
    let cell = &row.cells[0];
    if cell.content.is_empty() {
        return None;
    }
    let all_bold = cell.content.iter().all(|p| {
        p.runs
            .iter()
            .all(|r| r.text.trim().is_empty() || r.style.bold)
    });
    let any_text = cell
        .content
        .iter()
        .any(|p| p.runs.iter().any(|r| !r.text.trim().is_empty()));
    if !(all_bold && any_text) {
        return None;
    }

    // Render the cell with bold suppressed, then prefix every line with `> `.
    let inner = render_cell_content(cell, options, resource_map, true);
    let inner = inner.replace("<br>", "\n");
    let mut out = String::new();
    for line in inner.lines() {
        if line.trim().is_empty() {
            out.push_str(">\n");
        } else {
            out.push_str("> **");
            out.push_str(line.trim());
            out.push_str("**\n");
        }
    }
    Some(out)
}

/// Render a table to Markdown.
fn render_table(table: &Table, options: &RenderOptions, resource_map: &ResourceMap) -> String {
    if table.is_empty() {
        return String::new();
    }

    // Optional: a 1×1 emphasized table is almost always a callout box,
    // not tabular data. Render it as a blockquote when the user opts in.
    if options.callout_blockquote {
        if let Some(quote) = render_callout_blockquote(table, options, resource_map) {
            return quote;
        }
    }

    // Check if we need HTML fallback
    if table.has_merged_cells() && matches!(options.table_fallback, super::TableFallback::Html) {
        return render_table_html(table);
    }

    let mut output = String::new();
    let mut nested_tables: Vec<&Table> = Vec::new();

    // Merges are placed on a flat grid first. Markdown has no colspan/rowspan, so a
    // merged cell must be anchored at the column it starts in with the columns it
    // covers left empty — otherwise the row is narrower than the table and every
    // value in it shifts out from under its heading.
    let grid = super::grid::lay_out(table);
    let col_count = grid.first().map_or(0, Vec::len);
    if col_count == 0 {
        return String::new();
    }

    for (i, slots) in grid.iter().enumerate() {
        output.push('|');

        // The first row is always rendered as a header row in markdown,
        // regardless of `row.is_header`. Treat it as a header cell for
        // emphasis-suppression purposes too.
        let is_header_row = i == 0 || table.rows[i].is_header;
        for slot in slots {
            match slot {
                Some(cell) => {
                    let text = render_cell_content(cell, options, resource_map, is_header_row);
                    output.push_str(&format!(" {} |", text));

                    // Collect nested tables for rendering after the main table
                    for nested in &cell.nested_tables {
                        nested_tables.push(nested);
                    }
                }
                // A column covered by a merge, or one this row does not reach.
                None => output.push_str("  |"),
            }
        }
        output.push('\n');

        // Add separator after first row (markdown tables always need header separator)
        // In markdown, the first row is always treated as header regardless of source formatting
        if i == 0 {
            output.push('|');
            // Collect alignments from cells, filling with Left for missing columns
            let alignments = get_column_alignments(table, &grid, col_count);
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
            let text = escape_html(&cell.plain_text());
            html.push_str(&format!("    <{}{}>{}</{}>\n", tag, attrs, text, tag));
        }
        html.push_str("  </tr>\n");
    }

    html.push_str("</table>");
    html
}

fn escape_html(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::detect::FormatType;
    use crate::model::{Cell, HeadingLevel, RevisionType, Row, Section, TextStyle};
    use crate::render::options::SectionMarkerStyle;

    fn two_section_doc(format: FormatType, names: [&str; 2]) -> Document {
        let mut doc = Document::new();
        doc.format = format;
        let mut s0 = Section::new(0);
        s0.name = Some(names[0].to_string());
        let mut s1 = Section::new(1);
        s1.name = Some(names[1].to_string());
        doc.sections.push(s0);
        doc.sections.push(s1);
        doc
    }

    #[test]
    fn test_pptx_section_markers_comment() {
        let doc = two_section_doc(FormatType::Pptx, ["Introduction", "Conclusion"]);
        let opts = RenderOptions::new().with_section_markers(SectionMarkerStyle::Comment);
        let md = to_markdown(&doc, &opts).unwrap();
        assert!(
            md.contains("<!-- slide 1: Introduction -->"),
            "slide 1 marker missing\n{}",
            md
        );
        assert!(
            md.contains("<!-- slide 2: Conclusion -->"),
            "slide 2 marker missing\n{}",
            md
        );
    }

    #[test]
    fn test_pptx_section_markers_default_off() {
        let doc = two_section_doc(FormatType::Pptx, ["Introduction", "Conclusion"]);
        let opts = RenderOptions::new();
        let md = to_markdown(&doc, &opts).unwrap();
        assert!(
            !md.contains("<!-- slide"),
            "markers must be absent by default\n{}",
            md
        );
    }

    #[test]
    fn test_xlsx_section_markers_comment() {
        let doc = two_section_doc(FormatType::Xlsx, ["Revenue", "Costs"]);
        let opts = RenderOptions::new().with_section_markers(SectionMarkerStyle::Comment);
        let md = to_markdown(&doc, &opts).unwrap();
        assert!(
            md.contains("<!-- sheet 1: Revenue -->"),
            "sheet 1 marker missing\n{}",
            md
        );
        assert!(
            md.contains("<!-- sheet 2: Costs -->"),
            "sheet 2 marker missing\n{}",
            md
        );
    }

    #[test]
    fn test_docx_section_markers_never_emitted() {
        let doc = two_section_doc(FormatType::Docx, ["Chapter 1", "Chapter 2"]);
        let opts = RenderOptions::new().with_section_markers(SectionMarkerStyle::Comment);
        let md = to_markdown(&doc, &opts).unwrap();
        assert!(!md.contains("<!-- "), "DOCX must not emit markers\n{}", md);
    }

    #[test]
    fn test_pptx_nameless_section_marker() {
        let mut doc = Document::new();
        doc.format = FormatType::Pptx;
        let mut s = Section::new(0);
        s.name = None;
        doc.sections.push(s);
        let opts = RenderOptions::new().with_section_markers(SectionMarkerStyle::Comment);
        let md = to_markdown(&doc, &opts).unwrap();
        assert!(
            md.contains("<!-- slide 1 -->"),
            "nameless slide must use number-only marker\n{}",
            md
        );
    }

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
    fn test_hyperlink_destination_with_a_space_is_angle_wrapped() {
        // A bare `(dest with space)` is not valid CommonMark syntax at all --
        // pulldown-cmark never produces a Link event for it, so a consumer sees the
        // brackets as literal text instead of a link.
        let mut para = Paragraph::new();
        para.runs.push(TextRun::link("doc", "my folder/file.docx"));

        let options = RenderOptions::default();
        let md = render_paragraph(&para, &options, None, &empty_resource_map());
        assert!(
            md.contains("[doc](<my folder/file.docx>)"),
            "destination not angle-wrapped: {md:?}"
        );
    }

    #[test]
    fn test_hyperlink_destination_with_angle_brackets_is_backslash_escaped() {
        let mut para = Paragraph::new();
        para.runs.push(TextRun::link("doc", "a<b>c d"));

        let options = RenderOptions::default();
        let md = render_paragraph(&para, &options, None, &empty_resource_map());
        assert!(
            md.contains("[doc](<a\\<b\\>c d>)"),
            "angle brackets not backslash-escaped: {md:?}"
        );
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
    fn test_escape_pipe_only_in_table_cells() {
        // Outside a table cell, `|` is just a literal character.
        assert_eq!(
            escape_markdown("v1.0 | 2026-04-27", false),
            "v1.0 | 2026-04-27"
        );
        // Inside a table cell, `|` must be escaped to avoid breaking columns.
        assert_eq!(escape_markdown("a | b", true), "a \\| b");
    }

    #[test]
    fn test_escape_intra_word_underscore_not_escaped() {
        // CommonMark forbids intra-word `_` from triggering emphasis, so
        // identifiers like `snake_case` should pass through verbatim.
        assert_eq!(
            escape_markdown("YESUNG_OMS_backup_2026", false),
            "YESUNG_OMS_backup_2026"
        );
        assert_eq!(escape_markdown("in_house", false), "in_house");
        // Word-boundary `_` (next to whitespace) still needs escaping when it
        // could open an emphasis run.
        assert_eq!(escape_markdown("a _foo_ b", false), "a _foo_ b");
    }

    #[test]
    fn test_escape_backslash_and_backtick_always() {
        assert_eq!(escape_markdown("a`b\\c", false), "a\\`b\\\\c");
    }

    #[test]
    fn test_escape_star_intra_word_still_escaped() {
        // Unlike `_`, intra-word `*` *can* trigger emphasis, so it stays
        // escaped to be safe.
        assert_eq!(escape_markdown("foo*bar*baz", false), "foo\\*bar\\*baz");
    }

    #[test]
    fn test_heading_strips_uniform_bold_artifact() {
        // A heading whose every run is bold (Word's typical "Heading" style)
        // should not render as `# **text**` — the bold is a styling artifact.
        let mut para = Paragraph::heading(HeadingLevel::H1, "");
        para.runs.push(TextRun::styled("Title", TextStyle::bold()));
        let options = RenderOptions::default();
        let md = render_paragraph(&para, &options, None, &empty_resource_map());
        assert_eq!(md, "# Title", "heading bold should be stripped, got {md:?}");
    }

    #[test]
    fn test_heading_preserves_partial_bold_intent() {
        // If only part of the heading is bold, that's author intent — keep it.
        let mut para = Paragraph::heading(HeadingLevel::H2, "");
        para.runs.push(TextRun::plain("Section 2: "));
        para.runs
            .push(TextRun::styled("Required", TextStyle::bold()));
        let options = RenderOptions::default();
        let md = render_paragraph(&para, &options, None, &empty_resource_map());
        assert!(
            md.contains("**Required**"),
            "partial bold must be preserved, got {md:?}"
        );
        assert!(md.starts_with("## "));
    }

    #[test]
    fn test_cell_alignment_falls_back_to_paragraph_alignment() {
        use crate::model::TextAlignment;
        let mut table = Table::new();
        // Header row (Left default).
        table.add_row(Row::header(vec![
            Cell::header("L"),
            Cell::header("C"),
            Cell::header("R"),
        ]));
        // Data row: paragraphs carry alignment, cells do not.
        let mut data_row = Row {
            cells: vec![
                Cell::with_text("left-text"),
                Cell::with_text("center-text"),
                Cell::with_text("right-text"),
            ],
            is_header: false,
            height: None,
        };
        data_row.cells[1].content[0].alignment = TextAlignment::Center;
        data_row.cells[2].content[0].alignment = TextAlignment::Right;
        table.add_row(data_row);

        let md = render_table(&table, &RenderOptions::default(), &empty_resource_map());
        assert!(md.contains("| --- | :---: | ---: |"), "got {md:?}");
    }

    #[test]
    fn test_callout_blockquote_when_enabled() {
        // 1×1 fully-bold tables become `> **...**` when the option is on.
        let mut table = Table::new();
        let mut para = Paragraph::new();
        para.runs
            .push(TextRun::styled("Important note", TextStyle::bold()));
        let cell = Cell {
            content: vec![para],
            ..Cell::with_text("")
        };
        table.add_row(Row {
            cells: vec![cell],
            is_header: false,
            height: None,
        });

        let options = RenderOptions::default().with_callout_blockquote(true);
        let md = render_table(&table, &options, &empty_resource_map());
        assert!(md.starts_with("> **Important note**"), "got {md:?}");
        assert!(!md.contains("|"), "should not render as table: {md:?}");
    }

    #[test]
    fn test_callout_blockquote_off_by_default() {
        let mut table = Table::new();
        let mut para = Paragraph::new();
        para.runs.push(TextRun::styled("X", TextStyle::bold()));
        let cell = Cell {
            content: vec![para],
            ..Cell::with_text("")
        };
        table.add_row(Row {
            cells: vec![cell],
            is_header: false,
            height: None,
        });

        let md = render_table(&table, &RenderOptions::default(), &empty_resource_map());
        assert!(md.contains("|"), "default should keep table form: {md:?}");
    }

    #[test]
    fn test_table_cell_carriage_returns_do_not_split_rows() {
        // Defense-in-depth: even if model text carries a bare CR (e.g. a
        // Document built programmatically), a pipe-table row must stay on
        // one physical line (issue #7).
        let mut table = Table::new();
        table.add_row(Row {
            cells: vec![
                Cell::with_text("line one\rline two"),
                Cell::with_text("a\r\nb"),
            ],
            is_header: false,
            height: None,
        });

        let md = render_table(&table, &RenderOptions::default(), &empty_resource_map());
        assert!(!md.contains('\r'), "bare CR leaked into table: {md:?}");
        let row_line = md
            .lines()
            .find(|l| l.contains("line one"))
            .expect("row missing");
        assert!(
            row_line.contains("line two") && row_line.ends_with('|'),
            "row split across lines: {md:?}"
        );
    }

    #[test]
    fn test_page_break_default_off() {
        // PageBreak no longer emits a horizontal rule unless opted in.
        let mut doc = Document::new();
        let mut section = Section::new(0);
        section.add_paragraph(Paragraph::with_text("Before"));
        section.content.push(Block::PageBreak);
        section.add_paragraph(Paragraph::with_text("After"));
        doc.add_section(section);

        let md = to_markdown(&doc, &RenderOptions::default()).unwrap();
        assert!(
            !md.contains("---"),
            "page break should not emit ---: {md:?}"
        );

        let md_lossless = to_markdown(&doc, &RenderOptions::lossless()).unwrap();
        assert!(
            md_lossless.contains("---"),
            "lossless preset should emit ---"
        );
    }

    #[test]
    fn test_headers_footers_default_off() {
        let mut doc = Document::new();
        let mut section = Section::new(0);
        section.header = Some(vec![Paragraph::with_text("Page header text")]);
        section.footer = Some(vec![Paragraph::with_text("Page footer text")]);
        section.add_paragraph(Paragraph::with_text("Body."));
        doc.add_section(section);

        let md = to_markdown(&doc, &RenderOptions::default()).unwrap();
        assert!(!md.contains("Page header text"), "got {md:?}");
        assert!(!md.contains("Page footer text"), "got {md:?}");
    }

    #[test]
    fn test_consecutive_list_items_are_tight() {
        use crate::model::{ListInfo, ListType};
        let mut doc = Document::new();
        let mut section = Section::new(0);

        let mk_item = |text: &str| -> Paragraph {
            let mut p = Paragraph::with_text(text);
            p.list_info = Some(ListInfo {
                list_type: ListType::Bullet,
                level: 0,
                number: None,
            });
            p
        };
        section.add_paragraph(mk_item("Alpha"));
        section.add_paragraph(mk_item("Bravo"));
        section.add_paragraph(mk_item("Charlie"));
        section.add_paragraph(Paragraph::with_text("After list."));
        doc.add_section(section);

        let options = RenderOptions::default();
        let md = to_markdown(&doc, &options).unwrap();

        // Tight list: items separated by single newline (no blank line between).
        assert!(
            md.contains("- Alpha\n- Bravo\n- Charlie"),
            "expected tight list, got {md:?}"
        );
        // Blank line still separates the list from the following paragraph.
        assert!(md.contains("- Charlie\n\nAfter list."), "got {md:?}");
    }

    /// `cleanup: None` has to mean no post-processing. Blank runs that the *document* asked
    /// for are content — a consumer that diffs the output, cites it by line number, or
    /// chunks it for retrieval is entitled to get them back unchanged.
    #[test]
    fn test_no_cleanup_leaves_document_blank_lines_alone() {
        let mut doc = Document::new();
        let mut section = Section::new(0);
        // A paragraph holding nothing but a space is an ordinary artifact of authored
        // documents, and it is the document's own vertical spacing — not the renderer's.
        section.add_paragraph(Paragraph::with_text("First."));
        section.add_paragraph(Paragraph::with_text(" "));
        section.add_paragraph(Paragraph::with_text(" "));
        section.add_paragraph(Paragraph::with_text("Second."));
        doc.add_section(section);

        let options = RenderOptions::lossless();
        assert!(options.cleanup.is_none());
        let raw = to_markdown(&doc, &options).unwrap();

        let mut collapsing = RenderOptions::lossless();
        collapsing.cleanup = Some(crate::render::options::CleanupOptions::minimal());
        let cleaned = to_markdown(&doc, &collapsing).unwrap();

        assert!(
            raw.matches('\n').count() > cleaned.matches('\n').count(),
            "cleanup must be the only thing that collapses blank lines\n  raw: {raw:?}\n  cleaned: {cleaned:?}"
        );
    }

    #[test]
    fn test_blank_lines_collapsed_without_cleanup() {
        // A thematic break follows a paragraph that already ends with a blank line, so a
        // break that prepends its own newline yields three in a row. The renderer emits the
        // separator it actually needs instead, which is what lets this hold with no cleanup
        // configured — repairing it afterwards would mean post-processing that `cleanup:
        // None` promised not to do.
        let mut doc = Document::new();
        let mut section = Section::new(0);
        section.add_paragraph(Paragraph::with_text("Before break."));
        section.content.push(Block::PageBreak);
        section.add_paragraph(Paragraph::with_text("After break."));
        doc.add_section(section);

        // Page breaks are now opt-in (F8). Use the lossless preset so the
        // PageBreak block actually emits the `---` rule whose surrounding
        // newlines we want to verify get collapsed.
        let options = RenderOptions::lossless();
        assert!(options.cleanup.is_none());

        let md = to_markdown(&doc, &options).unwrap();

        assert!(
            !md.contains("\n\n\n"),
            "output must not contain 3+ consecutive newlines: {:?}",
            md
        );
        assert!(md.contains("Before break."));
        assert!(md.contains("After break."));
        assert!(md.contains("---"));
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

    /// A run's trailing space separates it from the next run — it is not part of the link
    /// label. Inside the brackets it renders in the wrong place and the raw Markdown reads
    /// `](url)`, which downstream tokenizers see as one glued token.
    #[test]
    fn test_hyperlink_trailing_space_stays_outside_the_link() {
        let mut para = Paragraph::new();
        para.runs.push(TextRun {
            text: "See the docs ".to_string(),
            hyperlink: Some("https://example.com".to_string()),
            ..Default::default()
        });
        para.runs.push(TextRun::plain("and continue"));

        let md = render_paragraph(
            &para,
            &RenderOptions::default(),
            None,
            &empty_resource_map(),
        );
        assert_eq!(md, "[See the docs](https://example.com) and continue");
    }

    /// CommonMark refuses a closing emphasis delimiter that is preceded by whitespace, so
    /// `**bold **` is not bold at all — the asterisks render literally.
    #[test]
    fn test_emphasis_trailing_space_stays_outside_the_delimiters() {
        let mut para = Paragraph::new();
        para.runs.push(TextRun {
            text: "bold ".to_string(),
            style: TextStyle {
                bold: true,
                ..Default::default()
            },
            ..Default::default()
        });
        para.runs.push(TextRun::plain("after"));

        let md = render_paragraph(
            &para,
            &RenderOptions::default(),
            None,
            &empty_resource_map(),
        );
        assert_eq!(md, "**bold** after");
    }

    /// A run that is only whitespace is the separator between its neighbours. Wrapping it in
    /// markup would emit delimiters around nothing.
    #[test]
    fn test_whitespace_only_run_is_a_separator_not_content() {
        let mut para = Paragraph::new();
        para.runs.push(TextRun::plain("word"));
        para.runs.push(TextRun {
            text: " ".to_string(),
            style: TextStyle {
                bold: true,
                ..Default::default()
            },
            ..Default::default()
        });
        para.runs.push(TextRun::plain("next"));

        let md = render_paragraph(
            &para,
            &RenderOptions::default(),
            None,
            &empty_resource_map(),
        );
        assert_eq!(md, "word next");
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
    fn test_table_cell_line_break_rendering_with_preserve_breaks() {
        let mut table = Table::new();
        table.add_row(Row::header(vec![Cell::header("Notes")]));
        table.add_row(Row {
            cells: vec![Cell {
                content: vec![Paragraph {
                    runs: vec![
                        TextRun {
                            text: "First line".to_string(),
                            line_break: true,
                            ..Default::default()
                        },
                        TextRun::plain("Second line"),
                    ],
                    ..Default::default()
                }],
                nested_tables: Vec::new(),
                col_span: 1,
                row_span: 1,
                alignment: CellAlignment::Left,
                vertical_alignment: crate::model::VerticalAlignment::Top,
                is_header: false,
                background: None,
            }],
            is_header: false,
            height: None,
        });

        let options = RenderOptions::new().with_preserve_breaks(true);
        let md = render_table(&table, &options, &empty_resource_map());

        assert!(
            md.contains("First line<br>Second line"),
            "Expected preserved line break in table cell, got: {md}"
        );
    }

    #[test]
    fn test_html_table_fallback_escapes_special_chars() {
        let mut table = Table::new();
        table.add_row(Row::header(vec![Cell::header("Header")]));
        table.add_row(Row {
            cells: vec![Cell {
                content: vec![Paragraph::with_text("<unsafe> & value")],
                nested_tables: Vec::new(),
                col_span: 2,
                row_span: 1,
                alignment: CellAlignment::Left,
                vertical_alignment: crate::model::VerticalAlignment::Top,
                is_header: false,
                background: None,
            }],
            is_header: false,
            height: None,
        });

        let options = RenderOptions::new().with_table_fallback(crate::render::TableFallback::Html);
        let md = render_table(&table, &options, &empty_resource_map());

        assert!(
            md.contains("&lt;unsafe&gt; &amp; value"),
            "Expected HTML-escaped table cell content, got: {md}"
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

    #[test]
    fn test_render_header_footer() {
        let mut doc = Document::new();
        let mut section = Section::new(0);
        section.header = Some(vec![Paragraph::with_text("My Header")]);
        section.footer = Some(vec![Paragraph::with_text("Page 1 of 10")]);
        section.add_paragraph(Paragraph::with_text("Body content"));
        doc.add_section(section);

        // Header/Footer rendering is now opt-in (F10).
        let options = RenderOptions::lossless();
        let md = to_markdown(&doc, &options).unwrap();

        assert!(
            md.contains("> *Header: My Header*"),
            "Expected header in output, got: {}",
            md
        );
        assert!(
            md.contains("> *Footer: Page 1 of 10*"),
            "Expected footer in output, got: {}",
            md
        );
        // Verify ordering: header before body, footer after body
        let header_pos = md.find("> *Header:").unwrap();
        let body_pos = md.find("Body content").unwrap();
        let footer_pos = md.find("> *Footer:").unwrap();
        assert!(header_pos < body_pos, "Header should appear before body");
        assert!(body_pos < footer_pos, "Footer should appear after body");
    }

    #[test]
    fn test_render_header_footer_multiple_paragraphs() {
        let mut doc = Document::new();
        let mut section = Section::new(0);
        section.header = Some(vec![
            Paragraph::with_text("Company"),
            Paragraph::with_text("Department"),
        ]);
        section.add_paragraph(Paragraph::with_text("Content"));
        doc.add_section(section);

        let options = RenderOptions::lossless();
        let md = to_markdown(&doc, &options).unwrap();

        assert!(
            md.contains("> *Header: Company | Department*"),
            "Multiple header paragraphs should be joined with ' | ', got: {}",
            md
        );
    }

    #[test]
    fn test_render_no_header_footer() {
        let mut doc = Document::new();
        let mut section = Section::new(0);
        section.add_paragraph(Paragraph::with_text("Body only"));
        doc.add_section(section);

        let options = RenderOptions::default();
        let md = to_markdown(&doc, &options).unwrap();

        assert!(!md.contains("Header:"), "No header should be rendered");
        assert!(!md.contains("Footer:"), "No footer should be rendered");
    }

    #[test]
    fn test_html_table_vmerge_rowspan() {
        // BUG-2: vMerge origin cells must carry correct rowspan in HTML fallback
        let mut table = Table::new();

        // Row 0: [A(rowspan=2), B]
        table.add_row(Row {
            cells: vec![
                Cell {
                    content: vec![Paragraph::with_text("A")],
                    nested_tables: vec![],
                    col_span: 1,
                    row_span: 2,
                    alignment: CellAlignment::Left,
                    vertical_alignment: crate::model::VerticalAlignment::Top,
                    is_header: false,
                    background: None,
                },
                Cell {
                    content: vec![Paragraph::with_text("B")],
                    nested_tables: vec![],
                    col_span: 1,
                    row_span: 1,
                    alignment: CellAlignment::Left,
                    vertical_alignment: crate::model::VerticalAlignment::Top,
                    is_header: false,
                    background: None,
                },
            ],
            is_header: false,
            height: None,
        });
        // Row 1: [C] (A's continuation is absent from cells, C is at col 1)
        table.add_row(Row {
            cells: vec![Cell {
                content: vec![Paragraph::with_text("C")],
                nested_tables: vec![],
                col_span: 1,
                row_span: 1,
                alignment: CellAlignment::Left,
                vertical_alignment: crate::model::VerticalAlignment::Top,
                is_header: false,
                background: None,
            }],
            is_header: false,
            height: None,
        });

        let options = RenderOptions::new().with_table_fallback(crate::render::TableFallback::Html);
        let html = render_table(&table, &options, &empty_resource_map());

        assert!(
            html.contains("rowspan=\"2\""),
            "Expected rowspan=2 on origin cell, got: {html}"
        );
    }

    #[test]
    fn test_list_item_with_heading_style_renders_as_list_not_heading() {
        // BUG-1: paragraph with both list_info and heading style must not produce "# -"
        use crate::model::{ListInfo, ListType};

        let mut para = Paragraph::new();
        para.runs.push(TextRun::plain("Item text"));
        para.heading = HeadingLevel::from_number(1);
        para.list_info = Some(ListInfo {
            level: 0,
            number: None,
            list_type: ListType::Bullet,
        });

        let mut doc = Document::new();
        let mut section = Section::new(0);
        section.add_paragraph(para);
        doc.add_section(section);

        let options = RenderOptions::default();
        let md = to_markdown(&doc, &options).unwrap();

        assert!(
            !md.contains("# "),
            "Heading marker must not appear for list item, got: {md}"
        );
        assert!(
            md.contains("- Item text"),
            "List marker must appear, got: {md}"
        );
    }

    #[test]
    fn test_frontmatter_format_field() {
        // FEAT-3: frontmatter must include format field
        use crate::detect::FormatType;

        let mut doc = Document::new();
        doc.format = FormatType::Docx;
        let mut section = Section::new(0);
        section.add_paragraph(Paragraph::with_text("Content"));
        doc.add_section(section);

        let options = RenderOptions::new().with_frontmatter(true);
        let md = to_markdown(&doc, &options).unwrap();

        assert!(
            md.contains("format: docx"),
            "Expected 'format: docx' in frontmatter, got: {md}"
        );
    }
}
