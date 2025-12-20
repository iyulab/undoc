//! Output rendering for documents.
//!
//! This module provides renderers for converting Document models
//! to various output formats: Markdown, plain text, and JSON.

mod options;
mod markdown;

pub use options::{RenderOptions, TableFallback, CleanupOptions, CleanupPreset};
pub use markdown::to_markdown;
