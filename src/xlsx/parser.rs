//! XLSX parser implementation.

use crate::container::OoxmlContainer;
use crate::error::Result;
use crate::model::Document;
use std::path::Path;

/// Parser for XLSX (Excel) workbooks.
pub struct XlsxParser {
    container: OoxmlContainer,
}

impl XlsxParser {
    /// Open an XLSX file for parsing.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let container = OoxmlContainer::open(path)?;
        Ok(Self { container })
    }

    /// Create a parser from bytes.
    pub fn from_bytes(data: Vec<u8>) -> Result<Self> {
        let container = OoxmlContainer::from_bytes(data)?;
        Ok(Self { container })
    }

    /// Parse the workbook and return a Document model.
    pub fn parse(&mut self) -> Result<Document> {
        // TODO: Implement in Phase 3
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
    fn test_open_xlsx() {
        let path = "test-files/file_example_XLSX_5000.xlsx";
        if std::path::Path::new(path).exists() {
            let parser = XlsxParser::open(path);
            assert!(parser.is_ok());
        }
    }
}
