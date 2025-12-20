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

#[cfg(feature = "ffi")]
pub mod ffi;

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
pub fn to_markdown_with_options(
    path: impl AsRef<Path>,
    options: &render::RenderOptions,
) -> Result<String> {
    let doc = parse_file(path)?;
    render::to_markdown(&doc, options)
}

/// Convert a document to plain text with render options.
///
/// # Example
///
/// ```no_run
/// use undoc::{to_text, render::RenderOptions};
///
/// let text = to_text("document.docx", &RenderOptions::default())?;
/// # Ok::<(), undoc::Error>(())
/// ```
pub fn to_text(path: impl AsRef<Path>, options: &render::RenderOptions) -> Result<String> {
    let doc = parse_file(path)?;
    render::to_text(&doc, options)
}

/// Convert a document to JSON.
///
/// # Example
///
/// ```no_run
/// use undoc::{to_json, render::JsonFormat};
///
/// let json = to_json("document.docx", JsonFormat::Pretty)?;
/// std::fs::write("output.json", json)?;
/// # Ok::<(), undoc::Error>(())
/// ```
pub fn to_json(path: impl AsRef<Path>, format: render::JsonFormat) -> Result<String> {
    let doc = parse_file(path)?;
    render::to_json(&doc, format)
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

    #[test]
    fn test_pptx_to_markdown_with_table() {
        let path = "test-files/file_example_PPT_1MB.pptx";
        if Path::new(path).exists() {
            let md = to_markdown(path).unwrap();

            // Should contain table markdown
            assert!(md.contains("|"), "Output should contain markdown table");

            // Print a portion for inspection
            println!("=== PPTX Markdown Output (Table Section) ===");
            // Find the Table slide and print it
            if let Some(table_start) = md.find("## Slide 3") {
                let table_section = &md[table_start..];
                if let Some(next_slide) = table_section.find("## Slide 4") {
                    println!("{}", &table_section[..next_slide]);
                } else {
                    println!("{}", table_section);
                }
            }
        }
    }

    #[test]
    fn test_pptx_korean_text_spacing() {
        let path = "test-files/고객 응대 시간.pptx";
        if Path::new(path).exists() {
            let md = to_markdown(path).unwrap();

            // Print Slide 1 content
            println!("=== Korean PPTX Output ===");
            if let Some(slide1) = md.find("## Slide 1") {
                let section = &md[slide1..];
                if let Some(next) = section.find("## Slide 2") {
                    println!("{}", &section[..next]);
                } else {
                    println!("{}", section);
                }
            }

            // Proper spacing: "고객 응대 평균 시간 4 시간 ~12 시간 소요"
            // Bad spacing: "고객 응대 평균 시간4시간~12시간 소요"
            assert!(
                !md.contains("시간4시간"),
                "Should have space between '시간' and '4'"
            );
        }
    }

    #[test]
    fn test_all_docx_files() {
        let files = [
            "test-files/file-sample_1MB.docx",
            "test-files/BT-B-24-0017 시험합의서_v0.2.docx",
            "test-files/CJ대한통운_ClusterPlex 로그 분석 보고서_240927-1.docx",
            "test-files/한국농어촌공사(체험마을정보)_OpenAPI활용가이드_v1.0.docx",
        ];

        println!("\n=== DOCX Conversion Report ===\n");
        for path in files {
            if Path::new(path).exists() {
                match parse_file(path) {
                    Ok(doc) => {
                        let md =
                            render::to_markdown(&doc, &render::RenderOptions::default()).unwrap();
                        let text = doc.plain_text();
                        println!("✓ {}", path);
                        println!(
                            "  Sections: {}, Resources: {}",
                            doc.sections.len(),
                            doc.resources.len()
                        );
                        println!(
                            "  Text length: {} chars, MD length: {} chars",
                            text.len(),
                            md.len()
                        );

                        // Check for common issues
                        if md.contains("DOCPROPERTY") || md.contains("HYPERLINK") {
                            println!("  ⚠ Contains field codes (DOCPROPERTY/HYPERLINK)");
                        }
                        if md.contains("\\*") {
                            println!("  ⚠ Contains escaped asterisks (over-escaping)");
                        }
                    }
                    Err(e) => {
                        println!("✗ {}: {}", path, e);
                    }
                }
                println!();
            }
        }
    }

    #[test]
    fn test_all_pptx_files() {
        let files = [
            "test-files/file_example_PPT_1MB.pptx",
            "test-files/고객 응대 시간.pptx",
            "test-files/1. 현장점검  보고서_우수 샘플.pptx",
            "test-files/2차 게이트웨이_20200831 인트세인 현황.pptx",
        ];

        println!("\n=== PPTX Conversion Report ===\n");
        for path in files {
            if Path::new(path).exists() {
                match parse_file(path) {
                    Ok(doc) => {
                        let md =
                            render::to_markdown(&doc, &render::RenderOptions::default()).unwrap();
                        let text = doc.plain_text();
                        println!("✓ {}", path);
                        println!(
                            "  Slides: {}, Resources: {}",
                            doc.sections.len(),
                            doc.resources.len()
                        );
                        println!(
                            "  Text length: {} chars, MD length: {} chars",
                            text.len(),
                            md.len()
                        );

                        // Count tables
                        let table_count = md.matches("| ---").count();
                        if table_count > 0 {
                            println!("  Tables: {}", table_count);
                        }
                    }
                    Err(e) => {
                        println!("✗ {}: {}", path, e);
                    }
                }
                println!();
            }
        }
    }

    #[test]
    fn test_docx_over_escaping() {
        // Test first file
        let path = "test-files/BT-B-24-0017 시험합의서_v0.2.docx";
        if Path::new(path).exists() {
            let md = to_markdown(path).unwrap();
            println!("\n=== Over-escaping Analysis: {} ===\n", path);
            for line in md.lines() {
                if line.contains("\\*") {
                    println!("Found: {}", line);
                }
            }
            assert!(!md.contains("\\*"), "Should not have escaped asterisks");
        }

        // Test second file
        let path2 = "test-files/CJ대한통운_ClusterPlex 로그 분석 보고서_240927-1.docx";
        if Path::new(path2).exists() {
            let md = to_markdown(path2).unwrap();
            println!("\n=== Over-escaping Analysis: {} ===\n", path2);
            for line in md.lines() {
                if line.contains("\\*") {
                    println!("Found: {}", line);
                }
            }
        }
    }

    #[test]
    fn test_all_xlsx_files() {
        let files = [
            "test-files/file_example_XLSX_5000.xlsx",
            "test-files/Auto Expense Report.xlsx",
            "test-files/Basic Invoice.xlsx",
        ];

        println!("\n=== XLSX Conversion Report ===\n");
        for path in files {
            if Path::new(path).exists() {
                match parse_file(path) {
                    Ok(doc) => {
                        let md =
                            render::to_markdown(&doc, &render::RenderOptions::default()).unwrap();
                        let text = doc.plain_text();
                        println!("✓ {}", path);
                        println!("  Sheets: {}", doc.sections.len());
                        println!(
                            "  Text length: {} chars, MD length: {} chars",
                            text.len(),
                            md.len()
                        );

                        // Count table rows
                        let row_count = md.matches("\n|").count();
                        println!("  Table rows: ~{}", row_count);
                    }
                    Err(e) => {
                        println!("✗ {}: {}", path, e);
                    }
                }
                println!();
            }
        }
    }
}
