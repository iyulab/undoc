//! JSON renderer implementation.

use crate::error::Result;
use crate::model::Document;

use super::options::RenderOptions;

/// JSON output format options.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum JsonFormat {
    /// Compact single-line JSON
    Compact,
    /// Pretty-printed with 2-space indentation
    #[default]
    Pretty,
}

/// Convert a Document to JSON.
pub fn to_json(doc: &Document, format: JsonFormat) -> Result<String> {
    match format {
        JsonFormat::Compact => serde_json::to_string(doc)
            .map_err(|e| crate::error::Error::XmlParse(format!("JSON serialization error: {}", e))),
        JsonFormat::Pretty => serde_json::to_string_pretty(doc)
            .map_err(|e| crate::error::Error::XmlParse(format!("JSON serialization error: {}", e))),
    }
}

/// Convert a Document to JSON with default formatting.
pub fn to_json_default(doc: &Document) -> Result<String> {
    to_json(doc, JsonFormat::Pretty)
}

/// Convert a Document to JSON with render options (for consistency).
pub fn to_json_with_options(doc: &Document, _options: &RenderOptions) -> Result<String> {
    // RenderOptions doesn't affect JSON output directly,
    // but we may add JSON-specific options in the future
    to_json(doc, JsonFormat::Pretty)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{HeadingLevel, Paragraph, Section};

    #[test]
    fn test_to_json_pretty() {
        let mut doc = Document::new();
        doc.metadata.title = Some("Test".to_string());
        let mut section = Section::new(0);
        section.add_paragraph(Paragraph::with_text("Hello"));
        doc.add_section(section);

        let json = to_json(&doc, JsonFormat::Pretty).unwrap();
        assert!(json.contains("\"title\": \"Test\""));
        assert!(json.contains("\"text\": \"Hello\""));
    }

    #[test]
    fn test_to_json_compact() {
        let mut doc = Document::new();
        doc.metadata.title = Some("Test".to_string());

        let json = to_json(&doc, JsonFormat::Compact).unwrap();
        assert!(!json.contains('\n')); // Compact has no newlines
        assert!(json.contains("\"title\":\"Test\""));
    }

    #[test]
    fn test_to_json_default() {
        let doc = Document::new();
        let json = to_json_default(&doc).unwrap();
        assert!(json.contains('\n')); // Default is pretty-printed
    }

    #[test]
    fn test_document_roundtrip() {
        let mut doc = Document::new();
        doc.metadata.title = Some("Roundtrip Test".to_string());
        doc.metadata.author = Some("Test Author".to_string());

        let mut section = Section::new(0);
        section.name = Some("Section 1".to_string());
        section.add_paragraph(Paragraph::heading(HeadingLevel::H1, "Heading"));
        section.add_paragraph(Paragraph::with_text("Content."));
        doc.add_section(section);

        let json = to_json(&doc, JsonFormat::Pretty).unwrap();
        let parsed: Document = serde_json::from_str(&json).unwrap();

        assert_eq!(parsed.metadata.title, doc.metadata.title);
        assert_eq!(parsed.metadata.author, doc.metadata.author);
        assert_eq!(parsed.sections.len(), 1);
    }
}
