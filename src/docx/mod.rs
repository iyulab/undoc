//! DOCX (Word) document parser.
//!
//! This module provides parsing for Microsoft Word documents in the
//! Office Open XML (.docx) format.
//!
//! # Example
//!
//! ```no_run
//! use undoc::docx::DocxParser;
//!
//! let mut parser = DocxParser::open("document.docx")?;
//! let doc = parser.parse()?;
//!
//! println!("Title: {:?}", doc.metadata.title);
//! println!("Content: {}", doc.plain_text());
//! # Ok::<(), undoc::Error>(())
//! ```

mod numbering;
mod parser;
mod styles;

pub use parser::DocxParser;
