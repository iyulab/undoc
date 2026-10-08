//! Excel binary workbook (.xls) reader.
//!
//! Reads the BIFF record stream of every Excel version that wrote one: Excel 97-2003 (BIFF8,
//! [MS-XLS]) and Excel 5.0/95 (BIFF5/BIFF7) in a compound file, and the bare streams of
//! Excel 2.x, 3.0 and 4.0 (BIFF2–BIFF4, including the BIFF4 multi-sheet workbook). The
//! workbook globals (sheets, shared strings, number formats) and each worksheet's cell
//! records are assembled into one table per sheet by the same rules as the .xlsx reader.
//!
//! Text before BIFF8 is stored in the workbook's code page: the one its `CODEPAGE` record
//! names, else the character set of its fonts, else Windows-1252. Encrypted workbooks are
//! reported as [`Error::Encrypted`](crate::Error::Encrypted).
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
#[cfg(test)]
mod tests_legacy;
