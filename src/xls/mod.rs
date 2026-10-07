//! Excel 97-2003 workbook (.xls) reader.
//!
//! Reads the BIFF8 record stream ([MS-XLS]): the workbook globals (sheets, shared strings,
//! number formats) and each worksheet's cell records, assembled into one table per sheet by
//! the same rules as the .xlsx reader. Excel 5.0/95 (BIFF5) workbooks are reported as
//! unsupported; encrypted ones as [`Error::Encrypted`](crate::Error::Encrypted).
//!
//! # Example
//!
//! ```no_run
//! use undoc::xls::XlsParser;
//!
//! let mut parser = XlsParser::open("figures.xls")?;
//! let doc = parser.parse()?;
//! for section in &doc.sections {
//!     println!("Sheet: {}", section.name.as_deref().unwrap_or(""));
//! }
//! # Ok::<(), undoc::Error>(())
//! ```

mod parser;
mod records;

pub use parser::XlsParser;

#[cfg(test)]
mod tests;
