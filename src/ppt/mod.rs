//! PowerPoint 97-2003 presentation (.ppt) reader.
//!
//! Follows the user edits to the persist object directory ([MS-PPT]), takes the slides in
//! presentation order from the slide list, and reads each slide's text in drawing order —
//! placeholders through the outline text they reference, text boxes from their own text
//! atoms — with titles as headings and the slide's notes page as its notes. An encrypted
//! presentation is reported as [`Error::Encrypted`](crate::Error::Encrypted).
//!
//! # Example
//!
//! ```no_run
//! use undoc::ppt::PptParser;
//!
//! let mut parser = PptParser::open("deck.ppt")?;
//! let doc = parser.parse()?;
//! println!("{} slides", doc.sections.len());
//! # Ok::<(), undoc::Error>(())
//! ```

mod parser;

pub use parser::PptParser;

#[cfg(test)]
mod tests;
