//! Document model structures.

use super::{Paragraph, Resource, Table};
use crate::detect::FormatType;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;

/// Document metadata extracted from docProps/core.xml and docProps/app.xml.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Metadata {
    /// Document title
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,

    /// Document author/creator
    #[serde(skip_serializing_if = "Option::is_none")]
    pub author: Option<String>,

    /// Document subject
    #[serde(skip_serializing_if = "Option::is_none")]
    pub subject: Option<String>,

    /// Document description
    #[serde(skip_serializing_if = "Option::is_none")]
    pub description: Option<String>,

    /// Keywords/tags
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub keywords: Vec<String>,

    /// Creation date (ISO 8601)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub created: Option<String>,

    /// Last modification date (ISO 8601)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub modified: Option<String>,

    /// Last modified by
    #[serde(skip_serializing_if = "Option::is_none")]
    pub last_modified_by: Option<String>,

    /// Application that created the document
    #[serde(skip_serializing_if = "Option::is_none")]
    pub application: Option<String>,

    /// Number of pages (DOCX), sheets (XLSX), or slides (PPTX)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub page_count: Option<u32>,

    /// Word count (DOCX only)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub word_count: Option<u32>,
}

/// A content block within a section.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "type")]
pub enum Block {
    /// A paragraph of text
    Paragraph(Paragraph),
    /// A table
    Table(Table),
    /// A page break
    PageBreak,
    /// A section break
    SectionBreak,
    /// A footnote or endnote, placed where the document collects its notes (the end of the
    /// section for .docx and .doc). The runs that refer to it carry its `label` in
    /// [`TextRun::note`](crate::model::TextRun::note).
    Note {
        /// The label its references carry: `1`, `2`, … for footnotes, `e1`, `e2`, … for endnotes
        label: String,
        /// The note's text
        content: Vec<Paragraph>,
    },
    /// An image (standalone, not inline)
    Image {
        /// Resource ID for the image
        resource_id: String,
        /// Alt text
        #[serde(skip_serializing_if = "Option::is_none")]
        alt_text: Option<String>,
        /// Width in EMUs (English Metric Units)
        #[serde(skip_serializing_if = "Option::is_none")]
        width: Option<u32>,
        /// Height in EMUs
        #[serde(skip_serializing_if = "Option::is_none")]
        height: Option<u32>,
    },
}

/// A document section (DOCX) or worksheet (XLSX) or slide (PPTX).
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Section {
    /// Section index (0-based)
    pub index: usize,

    /// Section name (sheet name for XLSX, slide title for PPTX)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,

    /// Content blocks
    #[serde(default)]
    pub content: Vec<Block>,

    /// Header content (DOCX only)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header: Option<Vec<Paragraph>>,

    /// Footer content (DOCX only)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub footer: Option<Vec<Paragraph>>,

    /// Speaker notes (PPTX only)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub notes: Option<Vec<Paragraph>>,

    /// Resource id of the picture that fills the background of this slide, sheet or
    /// document. A background is not content in reading order: it has no block and no
    /// Markdown, only this reference, so the listed image can be traced to where it is
    /// drawn.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub background_image: Option<String>,
}

impl Section {
    /// Create a new section with the given index.
    pub fn new(index: usize) -> Self {
        Self {
            index,
            ..Default::default()
        }
    }

    /// Create a new section with a name.
    pub fn with_name(index: usize, name: impl Into<String>) -> Self {
        Self {
            index,
            name: Some(name.into()),
            ..Default::default()
        }
    }

    /// Add a content block to this section.
    pub fn add_block(&mut self, block: Block) {
        self.content.push(block);
    }

    /// Add a paragraph to this section.
    pub fn add_paragraph(&mut self, para: Paragraph) {
        self.content.push(Block::Paragraph(para));
    }

    /// Add a table to this section.
    pub fn add_table(&mut self, table: Table) {
        self.content.push(Block::Table(table));
    }

    /// Check if this section is empty.
    pub fn is_empty(&self) -> bool {
        self.content.is_empty()
    }

    /// Get the number of content blocks.
    pub fn len(&self) -> usize {
        self.content.len()
    }
}

/// A parsed Office document.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Document {
    /// Document format (DOCX, XLSX, PPTX).
    pub format: FormatType,

    /// Document metadata
    pub metadata: Metadata,

    /// Document sections/sheets/slides
    #[serde(default)]
    pub sections: Vec<Section>,

    /// Extracted resources (images, media)
    #[serde(skip_serializing_if = "HashMap::is_empty", default)]
    pub resources: HashMap<String, Resource>,
}

impl Document {
    /// Create a new empty document.
    pub fn new() -> Self {
        Self::default()
    }

    /// Add a section to the document.
    pub fn add_section(&mut self, section: Section) {
        self.sections.push(section);
    }

    /// Add a resource to the document.
    pub fn add_resource(&mut self, id: impl Into<String>, resource: Resource) {
        self.resources.insert(id.into(), resource);
    }

    /// Get a resource by ID.
    pub fn get_resource(&self, id: &str) -> Option<&Resource> {
        self.resources.get(id)
    }

    /// Get the total number of content blocks across all sections.
    pub fn total_blocks(&self) -> usize {
        self.sections.iter().map(|s| s.len()).sum()
    }

    /// Check if the document is empty.
    pub fn is_empty(&self) -> bool {
        self.sections.is_empty() || self.sections.iter().all(|s| s.is_empty())
    }

    /// Every table of the document in reading order, with where it is: the section's number
    /// (from 1 — a sheet, a slide, a document section) and the table's place among that
    /// section's tables (from 1) — the names `undoc tables` writes them under
    /// (`s<section>-t<n>`). A table nested in a cell is a table of its own, right after the
    /// one that holds it.
    pub fn tables(&self) -> impl Iterator<Item = (usize, usize, &Table)> {
        self.sections.iter().enumerate().flat_map(|(s, section)| {
            let mut found = Vec::new();
            for block in &section.content {
                if let Block::Table(table) = block {
                    with_nested_tables(table, &mut found);
                }
            }
            found
                .into_iter()
                .enumerate()
                .map(move |(t, table)| (s + 1, t + 1, table))
        })
    }

    /// Extract all text content as a single string.
    pub fn plain_text(&self) -> String {
        let mut text = String::new();
        for section in &self.sections {
            for block in &section.content {
                match block {
                    Block::Paragraph(para) => {
                        text.push_str(&para.plain_text());
                        text.push('\n');
                    }
                    Block::Table(table) => {
                        text.push_str(&table.plain_text());
                        text.push('\n');
                    }
                    _ => {}
                }
            }
            text.push('\n');
        }
        text.trim_matches('\n').to_string()
    }

    /// Convert to JSON string.
    pub fn to_json(&self) -> Result<String, serde_json::Error> {
        serde_json::to_string_pretty(self)
    }

    /// Convert to JSON string (compact).
    pub fn to_json_compact(&self) -> Result<String, serde_json::Error> {
        serde_json::to_string(self)
    }
}

/// `table`, then every table nested in its cells, depth first — the order they are read in.
fn with_nested_tables<'a>(table: &'a Table, out: &mut Vec<&'a Table>) {
    out.push(table);
    for row in &table.rows {
        for cell in &row.cells {
            for nested in &cell.nested_tables {
                with_nested_tables(nested, out);
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::detect::FormatType;
    use crate::model::{Cell, RevisionType, Row, TextRun, TextStyle};

    #[test]
    fn test_document_default_format_is_docx() {
        let doc = Document::new();
        assert_eq!(doc.format, FormatType::Docx);
    }

    #[test]
    fn test_document_creation() {
        let mut doc = Document::new();
        assert!(doc.is_empty());

        let mut section = Section::new(0);
        let para = Paragraph {
            runs: vec![TextRun::plain("Hello, World!")],
            ..Default::default()
        };
        section.add_paragraph(para);
        doc.add_section(section);

        assert!(!doc.is_empty());
        assert_eq!(doc.total_blocks(), 1);
    }

    #[test]
    fn test_plain_text_extraction() {
        let mut doc = Document::new();
        let mut section = Section::new(0);

        section.add_paragraph(Paragraph {
            runs: vec![
                TextRun::plain("Hello, "),
                TextRun {
                    text: "World".to_string(),
                    style: TextStyle {
                        bold: true,
                        ..Default::default()
                    },
                    hyperlink: None,
                    line_break: false,
                    page_break: false,
                    revision: RevisionType::None,
                    note: None,
                },
                TextRun::plain("!"),
            ],
            ..Default::default()
        });

        doc.add_section(section);
        assert_eq!(doc.plain_text(), "Hello, World!");
    }

    #[test]
    fn test_plain_text_preserves_boundary_spaces() {
        let mut doc = Document::new();
        let mut section = Section::new(0);
        section.add_paragraph(Paragraph::with_text("  padded text  "));
        doc.add_section(section);

        assert_eq!(doc.plain_text(), "  padded text  ");
    }

    #[test]
    fn test_metadata_serialization() {
        let meta = Metadata {
            title: Some("Test Document".to_string()),
            author: Some("Test Author".to_string()),
            ..Default::default()
        };

        let json = serde_json::to_string(&meta).unwrap();
        assert!(json.contains("Test Document"));
        assert!(json.contains("Test Author"));
        // Empty fields should not be serialized
        assert!(!json.contains("subject"));
    }

    /// A one-cell table holding `text`, its cell holding `nested`.
    fn table(text: &str, nested: Vec<Table>) -> Table {
        let mut cell = Cell::with_text(text);
        cell.nested_tables = nested;
        let mut row = Row::new();
        row.add_cell(cell);
        let mut table = Table::new();
        table.add_row(row);
        table
    }

    #[test]
    fn tables_are_numbered_per_section_with_nested_tables_after_their_holder() {
        let mut doc = Document::new();
        let mut first = Section::new(0);
        first.add_table(table(
            "a",
            vec![
                table("a.1", vec![table("a.1.1", vec![])]),
                table("a.2", vec![]),
            ],
        ));
        first.add_paragraph(Paragraph::with_text("between"));
        first.add_table(table("b", vec![]));
        doc.add_section(first);
        doc.add_section(Section::new(1));
        let mut third = Section::new(2);
        third.add_table(table("c", vec![]));
        doc.add_section(third);

        let found: Vec<(usize, usize, String)> = doc
            .tables()
            .map(|(s, t, table)| (s, t, table.rows[0].cells[0].plain_text()))
            .collect();
        let expected = [
            (1, 1, "a"),
            (1, 2, "a.1"),
            (1, 3, "a.1.1"),
            (1, 4, "a.2"),
            (1, 5, "b"),
            (3, 1, "c"),
        ]
        .map(|(s, t, text)| (s, t, text.to_string()));
        assert_eq!(found, expected);
        assert_eq!(Document::new().tables().count(), 0);
    }

    #[test]
    fn test_section_with_name() {
        let section = Section::with_name(0, "Sheet1");
        assert_eq!(section.name, Some("Sheet1".to_string()));
        assert_eq!(section.index, 0);
    }
}
