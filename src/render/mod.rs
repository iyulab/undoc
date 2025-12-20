//! Output rendering for documents.
//!
//! This module provides renderers for converting Document models
//! to various output formats: Markdown, plain text, and JSON.
//!
//! # Example
//!
//! ```no_run
//! use undoc::{parse_file, render::*};
//!
//! let doc = parse_file("document.docx")?;
//!
//! // Render to Markdown
//! let md = to_markdown(&doc, &RenderOptions::default())?;
//!
//! // Render to plain text
//! let text = to_text(&doc, &RenderOptions::default())?;
//!
//! // Render to JSON
//! let json = to_json(&doc, JsonFormat::Pretty)?;
//! # Ok::<(), undoc::Error>(())
//! ```

mod cleanup;
mod json;
mod markdown;
mod options;
mod text;

pub use cleanup::{clean_text, detect_mojibake};
pub use json::{to_json, to_json_default, to_json_with_options, JsonFormat};
pub use markdown::to_markdown;
pub use options::{CleanupOptions, CleanupPreset, RenderOptions, TableFallback};
pub use text::to_text;
