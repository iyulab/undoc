//! Word 97-2003 binary document (.doc) reader.
//!
//! Reads the main document text through the piece table, and the paragraph properties that
//! decide structure — tables, headings — from the formatted disk pages and the style sheet
//! ([MS-DOC]). Word 6.0 and Word 95 files, which use an earlier layout, are reported as
//! unsupported; encrypted files as [`Error::Encrypted`](crate::Error::Encrypted).
//!
//! # Example
//!
//! ```no_run
//! use undoc::doc::DocParser;
//!
//! let mut parser = DocParser::open("report.doc")?;
//! let doc = parser.parse()?;
//! println!("{}", doc.plain_text());
//! # Ok::<(), undoc::Error>(())
//! ```

mod chars;
mod fib;
mod lists;
mod parser;
mod pictures;
mod props;
mod symbol;
mod text;

pub use parser::DocParser;

#[cfg(test)]
mod tests;
