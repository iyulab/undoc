//! # undoc
//!
//! High-performance Microsoft Office document extraction to Markdown.
//!
//! This library provides tools for parsing DOCX, XLSX, and PPTX files
//! and converting them to Markdown, plain text, or structured JSON.
//!
//! ## Quick Start
//!
//! ```no_run
//! use undoc::{parse_file, to_markdown};
//!
//! // Simple text extraction
//! let text = undoc::extract_text("document.docx")?;
//! println!("{}", text);
//!
//! // Convert to Markdown
//! let markdown = to_markdown("document.docx")?;
//! std::fs::write("output.md", markdown)?;
//!
//! // Full parsing with access to structure
//! let doc = parse_file("document.docx")?;
//! println!("Sections: {}", doc.sections.len());
//! println!("Resources: {}", doc.resources.len());
//! # Ok::<(), undoc::Error>(())
//! ```
//!
//! ## Format-Specific APIs
//!
//! ```no_run
//! use undoc::docx::DocxParser;
//! use undoc::xlsx::XlsxParser;
//! use undoc::pptx::PptxParser;
//!
//! // Word documents
//! let doc = DocxParser::open("report.docx")?.parse()?;
//!
//! // Excel spreadsheets
//! let workbook = XlsxParser::open("data.xlsx")?.parse()?;
//!
//! // PowerPoint presentations
//! let presentation = PptxParser::open("slides.pptx")?.parse()?;
//! # Ok::<(), undoc::Error>(())
//! ```
//!
//! ## Features
//!
//! - `docx` (default): Word document support
//! - `xlsx` (default): Excel spreadsheet support
//! - `pptx` (default): PowerPoint presentation support
//! - `async`: Async I/O support with Tokio
//! - `ffi`: C-ABI bindings for foreign language integration

pub mod container;
pub mod detect;
pub mod error;
pub mod model;

#[cfg(feature = "docx")]
pub mod docx;

#[cfg(feature = "xlsx")]
pub mod xlsx;

#[cfg(feature = "pptx")]
pub mod pptx;

pub mod render;

// Re-exports
pub use container::{OoxmlContainer, Relationship, Relationships};
pub use detect::{detect_format_from_bytes, detect_format_from_path, FormatType};
pub use error::{Error, Result};
pub use model::{
    Block, Cell, CellAlignment, Document, HeadingLevel, ListInfo, ListType, Metadata, Paragraph,
    Resource, ResourceType, Row, Section, Table, TextAlignment, TextRun, TextStyle,
};

use std::path::Path;

/// Parse a document file and return a Document model.
///
/// This function auto-detects the file format and uses the appropriate parser.
///
/// # Example
///
/// ```no_run
/// use undoc::parse_file;
///
/// let doc = parse_file("document.docx")?;
/// println!("Sections: {}", doc.sections.len());
/// # Ok::<(), undoc::Error>(())
/// ```
pub fn parse_file(path: impl AsRef<Path>) -> Result<Document> {
    let path = path.as_ref();
    let format = detect_format_from_path(path)?;

    match format {
        #[cfg(feature = "docx")]
        FormatType::Docx => {
            let mut parser = docx::DocxParser::open(path)?;
            parser.parse()
        }
        #[cfg(feature = "xlsx")]
        FormatType::Xlsx => {
            let mut parser = xlsx::XlsxParser::open(path)?;
            parser.parse()
        }
        #[cfg(feature = "pptx")]
        FormatType::Pptx => {
            let mut parser = pptx::PptxParser::open(path)?;
            parser.parse()
        }
        #[cfg(not(all(feature = "docx", feature = "xlsx", feature = "pptx")))]
        _ => Err(Error::UnsupportedFormat(format!("{:?}", format))),
    }
}

/// Parse a document from bytes.
///
/// # Example
///
/// ```no_run
/// use undoc::parse_bytes;
///
/// let data = std::fs::read("document.docx")?;
/// let doc = parse_bytes(&data)?;
/// # Ok::<(), undoc::Error>(())
/// ```
pub fn parse_bytes(data: &[u8]) -> Result<Document> {
    let format = detect_format_from_bytes(data)?;

    match format {
        #[cfg(feature = "docx")]
        FormatType::Docx => {
            let mut parser = docx::DocxParser::from_bytes(data.to_vec())?;
            parser.parse()
        }
        #[cfg(feature = "xlsx")]
        FormatType::Xlsx => {
            let mut parser = xlsx::XlsxParser::from_bytes(data.to_vec())?;
            parser.parse()
        }
        #[cfg(feature = "pptx")]
        FormatType::Pptx => {
            let mut parser = pptx::PptxParser::from_bytes(data.to_vec())?;
            parser.parse()
        }
        #[cfg(not(all(feature = "docx", feature = "xlsx", feature = "pptx")))]
        _ => Err(Error::UnsupportedFormat(format!("{:?}", format))),
    }
}

/// Extract plain text from a document.
///
/// # Example
///
/// ```no_run
/// use undoc::extract_text;
///
/// let text = extract_text("document.docx")?;
/// println!("{}", text);
/// # Ok::<(), undoc::Error>(())
/// ```
pub fn extract_text(path: impl AsRef<Path>) -> Result<String> {
    let doc = parse_file(path)?;
    Ok(doc.plain_text())
}

/// Convert a document to Markdown.
///
/// # Example
///
/// ```no_run
/// use undoc::to_markdown;
///
/// let markdown = to_markdown("document.docx")?;
/// std::fs::write("output.md", markdown)?;
/// # Ok::<(), undoc::Error>(())
/// ```
pub fn to_markdown(path: impl AsRef<Path>) -> Result<String> {
    let doc = parse_file(path)?;
    render::to_markdown(&doc, &render::RenderOptions::default())
}

/// Convert a document to Markdown with options.
///
/// # Example
///
/// ```no_run
/// use undoc::{to_markdown_with_options, render::RenderOptions};
///
/// let options = RenderOptions::default()
///     .with_frontmatter(true)
///     .with_image_dir("assets");
///
/// let markdown = to_markdown_with_options("document.docx", &options)?;
/// # Ok::<(), undoc::Error>(())
/// ```
pub fn to_markdown_with_options(path: impl AsRef<Path>, options: &render::RenderOptions) -> Result<String> {
    let doc = parse_file(path)?;
    render::to_markdown(&doc, options)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_format_detection_docx() {
        let path = "test-files/file-sample_1MB.docx";
        if Path::new(path).exists() {
            let format = detect_format_from_path(path).unwrap();
            assert_eq!(format, FormatType::Docx);
        }
    }

    #[test]
    fn test_format_detection_xlsx() {
        let path = "test-files/file_example_XLSX_5000.xlsx";
        if Path::new(path).exists() {
            let format = detect_format_from_path(path).unwrap();
            assert_eq!(format, FormatType::Xlsx);
        }
    }

    #[test]
    fn test_format_detection_pptx() {
        let path = "test-files/file_example_PPT_1MB.pptx";
        if Path::new(path).exists() {
            let format = detect_format_from_path(path).unwrap();
            assert_eq!(format, FormatType::Pptx);
        }
    }
}
