//! Document model structures.

use super::{Paragraph, Resource, Table};
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
        text.trim().to_string()
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{TextRun, TextStyle};

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
                },
                TextRun::plain("!"),
            ],
            ..Default::default()
        });

        doc.add_section(section);
        assert_eq!(doc.plain_text(), "Hello, World!");
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

    #[test]
    fn test_section_with_name() {
        let section = Section::with_name(0, "Sheet1");
        assert_eq!(section.name, Some("Sheet1".to_string()));
        assert_eq!(section.index, 0);
    }
}
