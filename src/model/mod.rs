//! Intermediate document model for Office documents.
//!
//! This module defines the data structures that represent parsed Office documents
//! in a format-agnostic way. Parsers convert format-specific XML into these structures,
//! and renderers convert them to output formats like Markdown.

mod document;
mod paragraph;
mod resource;
mod table;

pub use document::*;
pub use paragraph::*;
pub use resource::*;
pub use table::*;
