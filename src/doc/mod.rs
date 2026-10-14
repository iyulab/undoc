//! Word 97-2003 binary document (.doc) reader.
//!
//! Reads the main document text through the piece table, and the paragraph properties that
//! decide structure — tables, headings — from the formatted disk pages and the style sheet
//! ([MS-DOC]). Word 6.0 and Word 95 files, which use an earlier layout, are read for their
//! text -- paragraphs, headings, tables, character formatting, breaks and field results, in the
//! document's code page -- without their notes, headers or pictures. Encrypted files are
//! reported as [`Error::Encrypted`](crate::Error::Encrypted).
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
mod word6;

pub use parser::DocParser;
pub use pictures::Inflate;

#[cfg(test)]
mod tests;
