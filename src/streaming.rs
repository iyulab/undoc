//! Streaming document parsing API.
//!
//! This module provides [`parse_file_streaming`], a public API for processing
//! large OOXML documents with bounded memory. Instead of materializing the
//! entire [`Document`](crate::model::Document) in memory, it emits events for
//! each section as it is parsed, allowing the caller to process and discard
//! each section before the next one is loaded.
//!
//! ## Supported formats
//!
//! - **PPTX**: each slide is a separate event.
//! - **XLSX**: each sheet is a separate event.
//! - **DOCX**: not yet supported (returns [`Error::UnsupportedFormat`]).
//!
//! ## Event order
//!
//! ```text
//! DocumentStart → (SectionParsed | SectionFailed)* → DocumentEnd → ResourceExtracted*
//! ```
//!
//! `ResourceExtracted` events are emitted after `DocumentEnd` so that section
//! memory is fully freed before any large binary data arrives.
//!
//! ## Early termination
//!
//! Return [`ControlFlow::Break(())`](std::ops::ControlFlow::Break) from the
//! callback to stop parsing early. No `DocumentEnd` event is emitted on early
//! break.

use crate::detect::{detect_format_from_path, FormatType};
use crate::error::Result;
use crate::model::Metadata;
use crate::Error;
use std::collections::HashMap;
use std::ops::ControlFlow;
use std::path::Path;

/// An event emitted during streaming document parsing.
///
/// Events are always ordered:
/// ```text
/// DocumentStart → (SectionParsed | SectionFailed)* → DocumentEnd → ResourceExtracted*
/// ```
pub enum ParseEvent<'doc> {
    /// Emitted once before any section events.
    ///
    /// `metadata` is valid for the lifetime of the entire stream (until
    /// `DocumentEnd` or early termination).
    DocumentStart {
        /// Document metadata (title, author, dates, etc.)
        metadata: &'doc Metadata,
        /// Number of sections detected (slides for PPTX, sheets for XLSX).
        section_count: usize,
        /// Maps each resource ID to its filename with extension.
        /// Built from the document manifest before any sections are emitted
        /// so streaming renderers can produce correct image paths without
        /// waiting for `ResourceExtracted` events.
        image_map: HashMap<String, String>,
    },

    /// A section was successfully parsed.
    ///
    /// The section is dropped at the end of the callback invocation — its
    /// memory is freed before the next event is emitted.
    SectionParsed(&'doc crate::model::Section),

    /// A section failed to parse.
    ///
    /// Only emitted when [`SectionStreamOptions::lenient`] is `true`.
    /// In strict mode, the stream terminates with `Err` instead.
    SectionFailed {
        /// Zero-based section index
        index: usize,
        /// The parse error
        error: Error,
    },

    /// Emitted once after all section events and before any `ResourceExtracted`
    /// events.
    DocumentEnd,

    /// A binary resource (image, media) extracted from the document.
    ///
    /// Emitted after `DocumentEnd` when
    /// [`SectionStreamOptions::extract_resources`] is `true`.
    ResourceExtracted {
        /// Resource identifier / filename (e.g., `"image1.png"`)
        name: String,
        /// Raw binary data
        data: Vec<u8>,
    },
}

/// Options for streaming document parsing.
#[derive(Debug, Clone)]
pub struct SectionStreamOptions {
    /// When `true`, per-section parse errors emit [`ParseEvent::SectionFailed`]
    /// and parsing continues. When `false` (default), any section error
    /// terminates the stream with `Err`.
    pub lenient: bool,

    /// Whether to emit [`ParseEvent::ResourceExtracted`] events after
    /// `DocumentEnd`. Default: `true`.
    pub extract_resources: bool,
}

impl Default for SectionStreamOptions {
    fn default() -> Self {
        Self {
            lenient: false,
            extract_resources: true,
        }
    }
}

/// Parses a document from a file, emitting events for each section.
///
/// `f` is called once per event in strict order. Return
/// [`ControlFlow::Break(())`](std::ops::ControlFlow::Break) to stop parsing
/// early (no `DocumentEnd` is emitted on early break).
///
/// ## Example
///
/// ```no_run
/// use std::ops::ControlFlow;
/// use undoc::{parse_file_streaming, ParseEvent, SectionStreamOptions};
///
/// parse_file_streaming("slides.pptx", SectionStreamOptions::default(), |event| {
///     match event {
///         ParseEvent::DocumentStart { metadata, section_count, .. } => {
///             println!("Title: {:?}, Sections: {}", metadata.title, section_count);
///         }
///         ParseEvent::SectionParsed(section) => {
///             println!("Section {}: {} blocks", section.index, section.content.len());
///         }
///         ParseEvent::DocumentEnd => {}
///         ParseEvent::SectionFailed { index, error } => {
///             eprintln!("Section {} failed: {}", index, error);
///         }
///         ParseEvent::ResourceExtracted { name, .. } => {
///             println!("Resource: {}", name);
///         }
///     }
///     ControlFlow::Continue(())
/// })?;
/// # Ok::<(), undoc::Error>(())
/// ```
pub fn parse_file_streaming<F>(
    path: impl AsRef<Path>,
    opts: SectionStreamOptions,
    f: F,
) -> Result<()>
where
    F: FnMut(ParseEvent<'_>) -> ControlFlow<()>,
{
    let path = path.as_ref();
    let format = detect_format_from_path(path)?;

    match format {
        #[cfg(feature = "pptx")]
        FormatType::Pptx => {
            let mut parser = crate::pptx::PptxParser::open(path)?;
            parser.for_each_section(opts, f)
        }
        #[cfg(feature = "xlsx")]
        FormatType::Xlsx => {
            let mut parser = crate::xlsx::XlsxParser::open(path)?;
            parser.for_each_section(opts, f)
        }
        #[cfg(feature = "docx")]
        FormatType::Docx => Err(Error::UnsupportedFormat(
            "streaming not yet supported for DOCX — use parse_file() instead".into(),
        )),
        #[allow(unreachable_patterns)]
        _ => Err(Error::UnsupportedFormat(format!("{:?}", format))),
    }
}
