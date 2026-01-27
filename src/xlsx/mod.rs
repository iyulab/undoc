//! XLSX (Excel) spreadsheet parser.
//!
//! This module provides parsing for Microsoft Excel workbooks in the
//! Office Open XML (.xlsx) format.
//!
//! # Example
//!
//! ```no_run
//! use undoc::xlsx::XlsxParser;
//!
//! let mut parser = XlsxParser::open("spreadsheet.xlsx")?;
//! let doc = parser.parse()?;
//!
//! for section in &doc.sections {
//!     println!("Sheet: {}", section.name.as_deref().unwrap_or("Unnamed"));
//! }
//! # Ok::<(), undoc::Error>(())
//! ```

mod parser;
mod shared_strings;
mod styles;

pub use parser::XlsxParser;
