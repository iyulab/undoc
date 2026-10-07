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
//! - **DOCX**, **DOC** and **XLS**: the entire document is parsed and its sections are
//!   yielded as events.
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

#[cfg(not(target_arch = "wasm32"))]
use crate::detect::{detect_format_from_path, FormatType};
#[cfg(not(target_arch = "wasm32"))]
use crate::error::Result;
use crate::model::Metadata;
use crate::Error;
use std::collections::HashMap;
#[cfg(not(target_arch = "wasm32"))]
use std::ops::ControlFlow;
#[cfg(not(target_arch = "wasm32"))]
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
#[cfg(not(target_arch = "wasm32"))]
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
        FormatType::Docx => {
            let mut parser = crate::docx::DocxParser::open(path)?;
            parser.for_each_section(opts, f)
        }
        #[cfg(feature = "doc")]
        FormatType::Doc => {
            let mut parser = crate::doc::DocParser::open(path)?;
            emit_parsed_document(parser.parse(), None, opts, f)
        }
        #[cfg(feature = "xls")]
        FormatType::Xls => {
            let mut parser = crate::xls::XlsParser::open(path)?;
            emit_parsed_document(parser.parse(), None, opts, f)
        }
        #[allow(unreachable_patterns)]
        _ => Err(Error::UnsupportedFormat(format!("{:?}", format))),
    }
}

/// Stream a document that a format reads whole: `DocumentStart`, one `SectionParsed` per
/// section, `DocumentEnd`, then its resources.
///
/// `metadata` is announced in `DocumentStart`; `None` announces the document's own. A failed
/// parse is the error itself, or under `lenient` a degenerate stream carrying it as
/// `SectionFailed` at index 0. `Break` ends the stream at any event, as it does for the
/// formats that stream section by section.
#[cfg(not(target_arch = "wasm32"))]
pub(crate) fn emit_parsed_document<F>(
    parsed: Result<crate::model::Document>,
    metadata: Option<Metadata>,
    opts: SectionStreamOptions,
    mut f: F,
) -> Result<()>
where
    F: FnMut(ParseEvent<'_>) -> ControlFlow<()>,
{
    let doc = match parsed {
        Ok(doc) => doc,
        Err(e) if opts.lenient => {
            let metadata = metadata.unwrap_or_default();
            if f(ParseEvent::DocumentStart {
                metadata: &metadata,
                section_count: 0,
                image_map: HashMap::new(),
            })
            .is_break()
            {
                return Ok(());
            }
            if f(ParseEvent::SectionFailed { index: 0, error: e }).is_break() {
                return Ok(());
            }
            // The last event: nothing follows that a `Break` could stop.
            let _ = f(ParseEvent::DocumentEnd);
            return Ok(());
        }
        Err(e) => return Err(e),
    };

    let image_map: HashMap<String, String> = doc
        .resources
        .iter()
        .filter_map(|(id, r)| r.filename.as_ref().map(|name| (id.clone(), name.clone())))
        .collect();
    let metadata = metadata.unwrap_or_else(|| doc.metadata.clone());

    if f(ParseEvent::DocumentStart {
        metadata: &metadata,
        section_count: doc.sections.len(),
        image_map,
    })
    .is_break()
    {
        return Ok(());
    }
    for section in &doc.sections {
        if f(ParseEvent::SectionParsed(section)).is_break() {
            return Ok(());
        }
    }
    if f(ParseEvent::DocumentEnd).is_break() {
        return Ok(());
    }
    if opts.extract_resources {
        for (id, resource) in doc.resources {
            let name = resource.filename.clone().unwrap_or(id);
            if f(ParseEvent::ResourceExtracted {
                name,
                data: resource.data,
            })
            .is_break()
            {
                return Ok(());
            }
        }
    }
    Ok(())
}
