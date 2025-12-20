//! DOCX parser implementation.

use crate::container::OoxmlContainer;
use crate::error::Result;
use crate::model::Document;
use std::path::Path;

/// Parser for DOCX (Word) documents.
pub struct DocxParser {
    container: OoxmlContainer,
}

impl DocxParser {
    /// Open a DOCX file for parsing.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let container = OoxmlContainer::open(path)?;
        Ok(Self { container })
    }

    /// Create a parser from bytes.
    pub fn from_bytes(data: Vec<u8>) -> Result<Self> {
        let container = OoxmlContainer::from_bytes(data)?;
        Ok(Self { container })
    }

    /// Parse the document and return a Document model.
    pub fn parse(&mut self) -> Result<Document> {
        // TODO: Implement in Phase 2
        let doc = Document::new();
        Ok(doc)
    }

    /// Get a reference to the container.
    pub fn container(&self) -> &OoxmlContainer {
        &self.container
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_open_docx() {
        let path = "test-files/file-sample_1MB.docx";
        if std::path::Path::new(path).exists() {
            let parser = DocxParser::open(path);
            assert!(parser.is_ok());
        }
    }
}
