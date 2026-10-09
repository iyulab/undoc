//! C-ABI Foreign Function Interface for undoc.
//!
//! This module provides C-compatible bindings for using undoc from other languages
//! such as C, C++, C#, Python, and any language with C FFI support.
//!
//! # Memory Management
//!
//! All strings returned by this library must be freed using `undoc_free_string`.
//! Byte buffers (`undoc_get_resource_data`, `undoc_render_section`) are freed with
//! `undoc_free_bytes`. All document handles must be freed using `undoc_free_document`.
//!
//! # Error Handling
//!
//! Functions that can fail return a null pointer on error. Use `undoc_last_error`
//! to retrieve the error message and `undoc_last_error_kind` to classify it without
//! parsing that message. Both read thread-local state written by the immediately
//! preceding call on the same thread, and are always written and cleared together.
//!
//! The kind values are a stable ABI contract: a new failure reason takes the next
//! free number and existing numbers are never reused or renumbered, so a caller may
//! safely treat an unrecognised value as a generic failure.
//!
//! # Example (C)
//!
//! ```c
//! #include <stdio.h>
//! #include "undoc.h"
//!
//! int main() {
//!     UndocDocument* doc = undoc_parse_file("document.docx");
//!     if (!doc) {
//!         const char* error = undoc_last_error();
//!         fprintf(stderr, "Error: %s\n", error);
//!         return 1;
//!     }
//!
//!     char* markdown = undoc_to_markdown(doc, 0);
//!     if (markdown) {
//!         printf("%s\n", markdown);
//!         undoc_free_string(markdown);
//!     }
//!
//!     undoc_free_document(doc);
//!     return 0;
//! }
//! ```
//!
//! # Example (C#)
//!
//! ```csharp
//! using System;
//! using System.Runtime.InteropServices;
//!
//! public class Undoc {
//!     [DllImport("undoc")]
//!     public static extern IntPtr undoc_parse_file(string path);
//!
//!     [DllImport("undoc")]
//!     public static extern IntPtr undoc_to_markdown(IntPtr doc, int flags);
//!
//!     [DllImport("undoc")]
//!     public static extern void undoc_free_string(IntPtr str);
//!
//!     [DllImport("undoc")]
//!     public static extern void undoc_free_document(IntPtr doc);
//! }
//! ```

use std::ffi::{c_char, c_int, CString};
use std::ptr;

use unparser_shared::ffi::{self, invalid_argument, FfiError, LastErrorSlot};

use crate::error::ErrorKind;
use crate::model::Document;
use crate::pptx::PptxParser;
use crate::render::{JsonFormat, RenderOptions};

// Thread-local storage for the last error message and its classification. Declared
// here rather than in `unparser-shared` — see that crate's `ffi` module docs for why the slot
// must live in the consuming crate.
thread_local! {
    static LAST_ERROR: LastErrorSlot = const { LastErrorSlot::new() };
}

unparser_shared::export_last_error_abi!(LAST_ERROR, undoc_last_error, undoc_last_error_kind);

/// `undoc_last_error_kind` value when no error is recorded on this thread.
pub const UNDOC_ERROR_NONE: c_int = unparser_shared::kind::NONE;

// Values 1..=13 and 300..=399 are [`ErrorKind`] discriminants — core failure reasons.
// Values 100..=199 are FFI-boundary reasons with no core `Error` counterpart.

/// An argument was null or not valid UTF-8.
pub const UNDOC_ERROR_INVALID_ARGUMENT: c_int = unparser_shared::kind::INVALID_ARGUMENT;
/// A panic was caught at the FFI boundary.
pub const UNDOC_ERROR_PANIC: c_int = unparser_shared::kind::PANIC;
/// The produced output contains an interior NUL byte and cannot cross the C ABI.
pub const UNDOC_ERROR_INVALID_OUTPUT: c_int = unparser_shared::kind::INVALID_OUTPUT;

/// Classify a core error and render its message, for return from a closure.
fn ffi_err(e: crate::Error) -> FfiError {
    (e.kind() as c_int, e.to_string())
}

/// Classify a JSON serialization failure — producing output is rendering.
fn json_err(e: serde_json::Error) -> FfiError {
    (ErrorKind::Render as c_int, e.to_string())
}

/// A parsed document, and — for a presentation — the parser that read it, kept so a
/// slide can be painted ([`undoc_render_section`]) without reading the file again. The
/// parser holds the package in memory, as it did while parsing. Reads as the [`Document`].
pub struct HeldDocument {
    document: Document,
    presentation: Option<PptxParser>,
}

impl HeldDocument {
    fn parse_file(path: &str) -> crate::Result<Self> {
        if crate::detect_format_from_path(path)? == crate::FormatType::Pptx {
            return Self::presentation(PptxParser::open(path)?);
        }
        crate::parse_file(path).map(Self::from)
    }

    fn parse_bytes(data: &[u8]) -> crate::Result<Self> {
        if crate::detect_format_from_bytes(data)? == crate::FormatType::Pptx {
            return Self::presentation(PptxParser::from_bytes(data.to_vec())?);
        }
        crate::parse_bytes(data).map(Self::from)
    }

    fn presentation(mut parser: PptxParser) -> crate::Result<Self> {
        let document = parser.parse()?;
        Ok(Self {
            document,
            presentation: Some(parser),
        })
    }
}

// Every entry point runs inside `catch_unwind`, which asks whether a handle observed after
// a caught panic could be in a broken state. The document is read-only. The presentation
// parser is only read after parsing: its slide list and relationships do not change, and
// its archive sits behind a `RefCell` whose borrow is released as a panic unwinds — a
// later read seeks to the entry it wants, whatever the last read left behind.
impl std::panic::RefUnwindSafe for HeldDocument {}
impl std::panic::UnwindSafe for HeldDocument {}

impl From<Document> for HeldDocument {
    fn from(document: Document) -> Self {
        Self {
            document,
            presentation: None,
        }
    }
}

impl std::ops::Deref for HeldDocument {
    type Target = Document;

    fn deref(&self) -> &Document {
        &self.document
    }
}

unparser_shared::export_handle! {
    /// Opaque handle to a parsed document.
    handle UndocDocument { inner: HeldDocument },

    /// Free a document handle.
    ///
    /// # Safety
    ///
    /// - `doc` must be a valid pointer returned by `undoc_parse_file` or `undoc_parse_bytes`.
    /// - After calling this function, the handle is invalid and must not be used.
    free undoc_free_document,
}

/// Flags for markdown rendering.
pub const UNDOC_FLAG_FRONTMATTER: u32 = 1;
pub const UNDOC_FLAG_ESCAPE_SPECIAL: u32 = 2;
pub const UNDOC_FLAG_PARAGRAPH_SPACING: u32 = 4;
pub const UNDOC_FLAG_REFINE: u32 = 8;

/// JSON format options.
pub const UNDOC_JSON_PRETTY: c_int = 0;
pub const UNDOC_JSON_COMPACT: c_int = 1;

/// Get the version of the library.
///
/// # Safety
///
/// Returns a static string that must not be freed.
#[no_mangle]
pub extern "C" fn undoc_version() -> *const c_char {
    concat!(env!("CARGO_PKG_VERSION"), "\0").as_ptr() as *const c_char
}

/// Parse a document from a file path.
///
/// # Safety
///
/// - `path` must be a valid null-terminated UTF-8 string.
/// - Returns null on error. Use `undoc_last_error` to get the error message.
/// - The returned handle must be freed with `undoc_free_document`.
#[no_mangle]
pub unsafe extern "C" fn undoc_parse_file(path: *const c_char) -> *mut UndocDocument {
    LAST_ERROR.with(|slot| slot.clear());

    let result: Result<*mut UndocDocument, FfiError> = ffi::catch(|| {
        let path_str = unparser_shared::with_c_str!(path)?;

        HeldDocument::parse_file(path_str)
            .map(|inner| Box::into_raw(Box::new(UndocDocument { inner })))
            .map_err(ffi_err)
    });

    match result {
        Ok(doc) => doc,
        Err(error) => {
            LAST_ERROR.with(|slot| slot.set_error(&error));
            ptr::null_mut()
        }
    }
}

/// Parse a document from a byte buffer.
///
/// # Safety
///
/// - `data` must be a valid pointer to a byte buffer of at least `len` bytes.
/// - Returns null on error. Use `undoc_last_error` to get the error message.
/// - The returned handle must be freed with `undoc_free_document`.
#[no_mangle]
pub unsafe extern "C" fn undoc_parse_bytes(data: *const u8, len: usize) -> *mut UndocDocument {
    LAST_ERROR.with(|slot| slot.clear());

    if data.is_null() {
        LAST_ERROR.with(|slot| slot.set_error(&invalid_argument("data is null")));
        return ptr::null_mut();
    }

    let result: Result<*mut UndocDocument, FfiError> = ffi::catch(|| {
        let bytes = std::slice::from_raw_parts(data, len);

        HeldDocument::parse_bytes(bytes)
            .map(|inner| Box::into_raw(Box::new(UndocDocument { inner })))
            .map_err(ffi_err)
    });

    match result {
        Ok(doc) => doc,
        Err(error) => {
            LAST_ERROR.with(|slot| slot.set_error(&error));
            ptr::null_mut()
        }
    }
}

unparser_shared::export_string_getter!(
    /// Convert a document to Markdown.
    ///
    /// # Safety
    ///
    /// - `doc` must be a valid document handle.
    /// - `flags` is a bitwise OR of `UNDOC_FLAG_*` constants.
    /// - Returns null on error. Use `undoc_last_error` to get the error message.
    /// - The returned string must be freed with `undoc_free_string`.
    LAST_ERROR,
    undoc_to_markdown(doc: UndocDocument, flags: u32),
    {
        let document = &(*doc).inner;

        let mut options = RenderOptions::new();

        if flags & UNDOC_FLAG_FRONTMATTER != 0 {
            options.include_frontmatter = true;
        }
        if flags & UNDOC_FLAG_ESCAPE_SPECIAL != 0 {
            options.escape_special_chars = true;
        }
        if flags & UNDOC_FLAG_PARAGRAPH_SPACING != 0 {
            options.paragraph_spacing = true;
        }
        #[cfg(feature = "refine")]
        if flags & UNDOC_FLAG_REFINE != 0 {
            options = options.with_refine();
        }

        crate::render::to_markdown(document, &options).map_err(ffi_err)
    }
);

unparser_shared::export_string_getter!(
    /// Convert a document to plain text.
    ///
    /// # Safety
    ///
    /// - `doc` must be a valid document handle.
    /// - Returns null on error. Use `undoc_last_error` to get the error message.
    /// - The returned string must be freed with `undoc_free_string`.
    LAST_ERROR,
    undoc_to_text(doc: UndocDocument),
    {
        let document = &(*doc).inner;
        let options = RenderOptions::default();
        crate::render::to_text(document, &options).map_err(ffi_err)
    }
);

unparser_shared::export_string_getter!(
    /// Convert a document to JSON.
    ///
    /// # Safety
    ///
    /// - `doc` must be a valid document handle.
    /// - `format` is one of `UNDOC_JSON_PRETTY` or `UNDOC_JSON_COMPACT`.
    /// - Returns null on error. Use `undoc_last_error` to get the error message.
    /// - The returned string must be freed with `undoc_free_string`.
    LAST_ERROR,
    undoc_to_json(doc: UndocDocument, format: c_int),
    {
        let document = &(*doc).inner;
        let json_format = if format == UNDOC_JSON_COMPACT {
            JsonFormat::Compact
        } else {
            JsonFormat::Pretty
        };
        crate::render::to_json(document, json_format).map_err(ffi_err)
    }
);

unparser_shared::export_string_getter!(
    /// Get the plain text content of a document.
    ///
    /// # Safety
    ///
    /// - `doc` must be a valid document handle.
    /// - Returns null on error; use `undoc_last_error` to get the error message.
    /// - A valid document with no text returns a non-null, empty (`""`) string.
    /// - The returned string must be freed with `undoc_free_string`.
    LAST_ERROR,
    undoc_plain_text(doc: UndocDocument),
    {
        let document = &(*doc).inner;
        Ok(document.plain_text())
    }
);

unparser_shared::export_count_getter!(
    /// Get the number of sections in a document.
    ///
    /// # Safety
    ///
    /// - `doc` must be a valid document handle.
    /// - Returns -1 on error.
    LAST_ERROR,
    undoc_section_count(doc: UndocDocument),
    {
        let document = &(*doc).inner;
        Ok(document.sections.len() as c_int)
    }
);

unparser_shared::export_count_getter!(
    /// Get the number of resources in a document.
    ///
    /// # Safety
    ///
    /// - `doc` must be a valid document handle.
    /// - Returns -1 on error.
    LAST_ERROR,
    undoc_resource_count(doc: UndocDocument),
    {
        let document = &(*doc).inner;
        Ok(document.resources.len() as c_int)
    }
);

unparser_shared::export_optional_string_getter!(
    /// Get the document title.
    ///
    /// # Safety
    ///
    /// - `doc` must be a valid document handle.
    /// - Returns null if no title is set — with `undoc_last_error_kind` left at
    ///   `UNDOC_ERROR_NONE`, since an absent title is not a failure. A null return paired
    ///   with a non-zero kind means the title could not be produced (for instance
    ///   `UNDOC_ERROR_INVALID_OUTPUT` when it holds an interior NUL byte).
    /// - The returned string must be freed with `undoc_free_string`.
    LAST_ERROR,
    undoc_get_title(doc: UndocDocument),
    {
        let document = &(*doc).inner;
        Ok(document.metadata.title.clone())
    }
);

unparser_shared::export_optional_string_getter!(
    /// Get the document author.
    ///
    /// # Safety
    ///
    /// - `doc` must be a valid document handle.
    /// - Returns null if no author is set — with `undoc_last_error_kind` left at
    ///   `UNDOC_ERROR_NONE`, since an absent author is not a failure. A null return paired
    ///   with a non-zero kind means the author could not be produced (for instance
    ///   `UNDOC_ERROR_INVALID_OUTPUT` when it holds an interior NUL byte).
    /// - The returned string must be freed with `undoc_free_string`.
    LAST_ERROR,
    undoc_get_author(doc: UndocDocument),
    {
        let document = &(*doc).inner;
        Ok(document.metadata.author.clone())
    }
);

unparser_shared::export_free_string!(
    /// Free a string allocated by this library.
    ///
    /// # Safety
    ///
    /// - `s` must be a pointer returned by an undoc function, or null.
    /// - After calling this function, the pointer is invalid and must not be used.
    undoc_free_string
);

unparser_shared::export_string_getter!(
    /// Get all resource IDs as a JSON array.
    ///
    /// # Safety
    ///
    /// - `doc` must be a valid document handle.
    /// - Returns null on error. Use `undoc_last_error` to get the error message.
    /// - A valid document with no resources returns a non-null `"[]"` string.
    /// - The returned string must be freed with `undoc_free_string`.
    ///
    /// # Returns
    ///
    /// A JSON array of resource IDs, e.g., `["rId1", "rId2", "rId3"]`
    LAST_ERROR,
    undoc_get_resource_ids(doc: UndocDocument),
    {
        let document = &(*doc).inner;
        let ids: Vec<&String> = document.resources.keys().collect();
        serde_json::to_string(&ids).map_err(json_err)
    }
);

unparser_shared::export_string_getter!(
    /// Get resource metadata as JSON (without binary data).
    ///
    /// # Safety
    ///
    /// - `doc` must be a valid document handle.
    /// - `resource_id` must be a valid null-terminated UTF-8 string.
    /// - Returns null if resource not found or on error.
    /// - The returned string must be freed with `undoc_free_string`.
    ///
    /// # Returns
    ///
    /// JSON object with resource metadata:
    /// `{"id":"rId1","type":"image","filename":"image1.png","mime_type":"image/png","size":1024,"width":800,"height":600,"alt_text":"Description","role":"primary","companion_of":null}`
    ///
    /// `role` is `"primary"` for an image the document shows (or media it plays),
    /// `"alternate"` for another encoding of one (a picture's SVG original) and `"layer"` for
    /// a layer composited onto one (an HD Photo effects layer); `companion_of` is the id of
    /// the primary an alternate or layer belongs to.
    LAST_ERROR,
    undoc_get_resource_info(doc: UndocDocument, resource_id: *const c_char),
    {
        let id_str = unparser_shared::with_c_str!(resource_id)?;

        let document = &(*doc).inner;

        match document.resources.get(id_str) {
            Some(resource) => {
                let info = serde_json::json!({
                    "id": id_str,
                    "type": resource.resource_type,
                    "filename": resource.filename,
                    "mime_type": resource.mime_type,
                    "size": resource.size,
                    "width": resource.width,
                    "height": resource.height,
                    "alt_text": resource.alt_text,
                    "role": resource.role,
                    "companion_of": resource.companion_of
                });
                serde_json::to_string(&info).map_err(json_err)
            }
            None => Err(ffi_err(crate::Error::ResourceNotFound(id_str.to_string()))),
        }
    }
);

/// Deserializable mirror of [`SlideRasterOptions`](crate::raster::SlideRasterOptions) for
/// `undoc_render_section`: `{"dpi": 150, "font_dirs": ["..."], "system_fonts": true}`.
/// Every field is optional.
#[derive(serde::Deserialize, Default)]
#[serde(default, deny_unknown_fields)]
struct FfiSlideRasterOptions {
    dpi: Option<f32>,
    font_dirs: Vec<String>,
    system_fonts: Option<bool>,
}

/// Rasterize a section to a PNG. A section of a presentation is a slide.
///
/// The slide is painted by the parser the handle keeps — the same package, and no second
/// read of the file. What the rasterizer cannot paint yet (charts, tables and other
/// graphic frames, custom geometry, pictures in formats other than PNG and JPEG, text no
/// face covers) is left out and counted in `out_info`; the rest of the slide is painted.
///
/// `index` is 0-based, in presentation order.
///
/// `options_json`: null, or `{"dpi": 150, "font_dirs": ["..."], "system_fonts": true}` —
/// resolution in dots per inch (default 150), directories searched with their
/// subdirectories for font files, and whether the system's font directories are searched
/// too (default true). No face is bundled: on a host without fonts, pass `font_dirs`, or
/// text is reported as a gap.
///
/// `out_info`, when not null, receives `{"width":N,"height":N,"gaps":{"shapes":N,
/// "images":N,"text_runs":N,"charts":N,"graphic_frames":N,"approximated_fills":N},
/// "substituted_text_runs":N}` — free it with `undoc_free_string`. `approximated_fills`
/// counts gradients and patterns painted in one of their colors;
/// `substituted_text_runs` counts text drawn in a face standing in for the one it asks
/// for — readable, not the slide's own typeface, and not a gap.
///
/// # Safety
///
/// - `doc` must be a valid document handle from `undoc_parse_file` or `undoc_parse_bytes`.
/// - `options_json` must be null or a valid null-terminated UTF-8 string.
/// - `out_len` must be a valid pointer; `out_info` must be null or a valid pointer.
/// - Returns null on error (`SECTION_OUT_OF_RANGE` for an index the presentation does not
///   have, `UNSUPPORTED_FORMAT` for a document that is not a `.pptx` presentation, `RENDER`
///   for a resolution that is not positive or would make the slide too large,
///   `INVALID_ARGUMENT` for options that do not parse); see `undoc_last_error`.
/// - The returned PNG must be freed with `undoc_free_bytes`.
#[no_mangle]
pub unsafe extern "C" fn undoc_render_section(
    doc: *const UndocDocument,
    index: c_int,
    options_json: *const c_char,
    out_len: *mut usize,
    out_info: *mut *mut c_char,
) -> *mut u8 {
    LAST_ERROR.with(|slot| slot.clear());
    if doc.is_null() || out_len.is_null() {
        LAST_ERROR
            .with(|slot| slot.set_error(&invalid_argument("doc and out_len must not be null")));
        return ptr::null_mut();
    }
    *out_len = 0;
    if !out_info.is_null() {
        *out_info = ptr::null_mut();
    }

    let result: Result<(Vec<u8>, String), FfiError> = ffi::catch(|| {
        let options: FfiSlideRasterOptions = if options_json.is_null() {
            FfiSlideRasterOptions::default()
        } else {
            let json = unparser_shared::ffi::c_str_utf8(options_json)?;
            serde_json::from_str(json).map_err(|e| invalid_argument(e.to_string()))?
        };
        let dpi = options.dpi.unwrap_or(150.0);
        let index = usize::try_from(index)
            .map_err(|_| invalid_argument(format!("index must not be negative, got {index}")))?;
        let presentation = (*doc).inner.presentation.as_ref().ok_or_else(|| {
            ffi_err(crate::Error::UnsupportedFormat(
                "rendering a section is supported for .pptx presentations".to_string(),
            ))
        })?;
        let raster_options = crate::raster::SlideRasterOptions {
            dpi,
            font_dirs: options.font_dirs.into_iter().map(Into::into).collect(),
            system_fonts: options.system_fonts.unwrap_or(true),
            ..Default::default()
        };
        let slide = presentation
            .render_slide(index, &raster_options)
            .map_err(ffi_err)?;
        let g = slide.gaps;
        let info = serde_json::json!({
            "width": slide.width,
            "height": slide.height,
            "gaps": {
                "shapes": g.shapes,
                "images": g.images,
                "text_runs": g.text_runs,
                "charts": g.charts,
                "graphic_frames": g.graphic_frames,
                "approximated_fills": g.approximated_fills,
            },
            "substituted_text_runs": slide.substituted_text_runs,
        });
        Ok((slide.to_png(), info.to_string()))
    });

    match result {
        Ok((png, info)) => {
            if !out_info.is_null() {
                *out_info = CString::new(info).map_or(ptr::null_mut(), CString::into_raw);
            }
            *out_len = png.len();
            Box::into_raw(png.into_boxed_slice()) as *mut u8
        }
        Err(error) => {
            LAST_ERROR.with(|slot| slot.set_error(&error));
            ptr::null_mut()
        }
    }
}

unparser_shared::export_bytes_getter!(
    /// Get resource binary data.
    ///
    /// # Safety
    ///
    /// - `doc` must be a valid document handle.
    /// - `resource_id` must be a valid null-terminated UTF-8 string.
    /// - `out_len` must be a valid pointer to receive the data length.
    /// - Returns null if resource not found or on error.
    /// - The returned pointer must be freed with `undoc_free_bytes`.
    LAST_ERROR,
    undoc_get_resource_data(doc: UndocDocument, resource_id, out out_len),
    {
        let id_str = unparser_shared::ffi::c_str_utf8(resource_id)?;

        let document = &(*doc).inner;

        match document.resources.get(id_str) {
            Some(resource) => Ok(resource.data.clone()),
            None => Err(ffi_err(crate::Error::ResourceNotFound(id_str.to_string()))),
        }
    }
);

unparser_shared::export_free_bytes!(
    /// Free binary data allocated by `undoc_get_resource_data`.
    ///
    /// # Safety
    ///
    /// - `data` must be a pointer returned by `undoc_get_resource_data`, or null.
    /// - `len` must be the length returned by `undoc_get_resource_data`.
    /// - After calling this function, the pointer is invalid and must not be used.
    undoc_free_bytes
);

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::Document;
    use std::ffi::{CStr, CString};

    #[test]
    fn test_version() {
        let version = undoc_version();
        assert!(!version.is_null());
        let version_str = unsafe { CStr::from_ptr(version) }.to_str().unwrap();
        assert!(!version_str.is_empty());
    }

    #[test]
    fn test_parse_null_path() {
        let doc = unsafe { undoc_parse_file(ptr::null()) };
        assert!(doc.is_null());

        let error = undoc_last_error();
        assert!(!error.is_null());
    }

    #[test]
    fn test_parse_invalid_path() {
        let path = CString::new("nonexistent.docx").unwrap();
        let doc = unsafe { undoc_parse_file(path.as_ptr()) };
        assert!(doc.is_null());

        let error = undoc_last_error();
        assert!(!error.is_null());
    }

    const HELLO: &str = "Hello from the C ABI";

    /// A one-paragraph DOCX, assembled in memory.
    fn hello_docx() -> Vec<u8> {
        use std::io::Write;
        let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#).unwrap();
        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#).unwrap();
        zip.start_file("word/document.xml", options).unwrap();
        write!(
            zip,
            r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>{HELLO}</w:t></w:r></w:p></w:body>
</w:document>"#
        )
        .unwrap();
        zip.finish().unwrap().into_inner()
    }

    /// Takes an owned copy of a returned string and frees the original.
    fn take_string(ptr: *mut c_char) -> String {
        assert!(!ptr.is_null());
        let owned = unsafe { CStr::from_ptr(ptr) }.to_str().unwrap().to_owned();
        unsafe { undoc_free_string(ptr) };
        owned
    }

    #[test]
    fn test_parse_and_convert() {
        let data = hello_docx();
        let doc = unsafe { undoc_parse_bytes(data.as_ptr(), data.len()) };
        assert!(!doc.is_null());
        assert_eq!(undoc_last_error_kind(), UNDOC_ERROR_NONE);

        assert_eq!(take_string(unsafe { undoc_to_text(doc) }).trim(), HELLO);
        let md = take_string(unsafe { undoc_to_markdown(doc, 0) });
        assert!(md.contains(HELLO), "markdown: {md}");
        let json = take_string(unsafe { undoc_to_json(doc, UNDOC_JSON_PRETTY) });
        assert!(json.contains(HELLO), "json: {json}");
        assert_eq!(unsafe { undoc_section_count(doc) }, 1);

        unsafe { undoc_free_document(doc) };
    }

    /// The file-path entry point -- otherwise reached only by the missing-file failure.
    #[test]
    fn test_parse_file_reads_a_document_from_disk() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("hello.docx");
        std::fs::write(&path, hello_docx()).unwrap();

        let c_path = CString::new(path.to_str().unwrap()).unwrap();
        let doc = unsafe { undoc_parse_file(c_path.as_ptr()) };
        assert!(!doc.is_null());
        assert_eq!(take_string(unsafe { undoc_to_text(doc) }).trim(), HELLO);

        unsafe { undoc_free_document(doc) };
    }

    #[test]
    fn test_null_document_operations() {
        let md = unsafe { undoc_to_markdown(ptr::null(), 0) };
        assert!(md.is_null());

        let text = unsafe { undoc_to_text(ptr::null()) };
        assert!(text.is_null());

        let json = unsafe { undoc_to_json(ptr::null(), 0) };
        assert!(json.is_null());

        let count = unsafe { undoc_section_count(ptr::null()) };
        assert_eq!(count, -1);

        let res_count = unsafe { undoc_resource_count(ptr::null()) };
        assert_eq!(res_count, -1);
    }

    #[test]
    fn test_plain_text_empty_document_returns_non_null_empty_string() {
        let doc = Box::into_raw(Box::new(UndocDocument {
            inner: Document::new().into(),
        }));

        let text = unsafe { undoc_plain_text(doc) };
        assert!(!text.is_null());
        let text_str = unsafe { CStr::from_ptr(text) }.to_str().unwrap();
        assert_eq!(text_str, "");

        unsafe {
            undoc_free_string(text);
            undoc_free_document(doc);
        }
    }

    #[test]
    fn test_last_error_kind_is_none_after_success() {
        let doc = Box::into_raw(Box::new(UndocDocument {
            inner: Document::new().into(),
        }));

        let text = unsafe { undoc_plain_text(doc) };
        assert!(!text.is_null());
        assert_eq!(undoc_last_error_kind(), UNDOC_ERROR_NONE);

        unsafe {
            undoc_free_string(text);
            undoc_free_document(doc);
        }
    }

    #[test]
    fn test_null_argument_kind_is_invalid_argument() {
        let doc = unsafe { undoc_parse_file(ptr::null()) };
        assert!(doc.is_null());
        assert_eq!(undoc_last_error_kind(), UNDOC_ERROR_INVALID_ARGUMENT);
    }

    #[test]
    fn test_missing_file_kind_is_io() {
        let path = CString::new("nonexistent-for-kind-test.docx").unwrap();
        let doc = unsafe { undoc_parse_file(path.as_ptr()) };
        assert!(doc.is_null());
        assert_eq!(undoc_last_error_kind(), ErrorKind::Io as c_int);
    }

    #[test]
    fn test_non_zip_bytes_kind_is_unknown_format() {
        let data = b"not an office document at all";
        let doc = unsafe { undoc_parse_bytes(data.as_ptr(), data.len()) };
        assert!(doc.is_null());
        assert_eq!(undoc_last_error_kind(), ErrorKind::UnknownFormat as c_int);
    }

    /// Regression fixture for the corrupted-archive case: ZIP magic present, contents
    /// unreadable. A consumer must learn "the container is damaged" from the kind
    /// alone, without matching on message text — and the message must still be there
    /// for a human reading a log.
    #[test]
    fn test_corrupted_archive_bytes_kind_is_zip_archive() {
        let mut data = vec![0x50, 0x4B, 0x03, 0x04];
        data.extend_from_slice(b"truncated garbage with no central directory");

        let doc = unsafe { undoc_parse_bytes(data.as_ptr(), data.len()) };
        assert!(doc.is_null());
        assert_eq!(undoc_last_error_kind(), ErrorKind::ZipArchive as c_int);

        let msg = unsafe { CStr::from_ptr(undoc_last_error()) }
            .to_str()
            .unwrap();
        assert!(
            !msg.is_empty(),
            "the human-readable message must survive alongside the kind"
        );
    }

    #[test]
    fn test_resource_not_found_kind() {
        let doc = Box::into_raw(Box::new(UndocDocument {
            inner: Document::new().into(),
        }));
        let id = CString::new("rIdMissing").unwrap();

        let info = unsafe { undoc_get_resource_info(doc, id.as_ptr()) };
        assert!(info.is_null());
        assert_eq!(
            undoc_last_error_kind(),
            ErrorKind::ResourceNotFound as c_int
        );

        unsafe { undoc_free_document(doc) };
    }

    /// The count accessors return a value rather than a pointer, so a caller cannot
    /// tell from the return value alone whether the recorded error is theirs. They must
    /// therefore clear the previous failure like every other entry point does.
    #[test]
    fn test_count_accessors_clear_a_previous_failure() {
        let path = CString::new("nonexistent-for-count-clear-test.docx").unwrap();
        let failed = unsafe { undoc_parse_file(path.as_ptr()) };
        assert!(failed.is_null());
        assert_ne!(undoc_last_error_kind(), UNDOC_ERROR_NONE);

        let doc = Box::into_raw(Box::new(UndocDocument {
            inner: Document::new().into(),
        }));

        assert_eq!(unsafe { undoc_section_count(doc) }, 0);
        assert_eq!(undoc_last_error_kind(), UNDOC_ERROR_NONE);

        let path = CString::new("nonexistent-for-count-clear-test.docx").unwrap();
        assert!(unsafe { undoc_parse_file(path.as_ptr()) }.is_null());
        assert_ne!(undoc_last_error_kind(), UNDOC_ERROR_NONE);

        assert_eq!(unsafe { undoc_resource_count(doc) }, 0);
        assert_eq!(undoc_last_error_kind(), UNDOC_ERROR_NONE);

        unsafe { undoc_free_document(doc) };
    }

    /// A null return does not always mean failure: an absent title is not an error.
    /// The kind channel is what lets a caller tell the two apart.
    #[test]
    fn test_absent_metadata_is_not_reported_as_a_failure() {
        let doc = Box::into_raw(Box::new(UndocDocument {
            inner: Document::new().into(),
        }));

        assert!(unsafe { undoc_get_title(doc) }.is_null());
        assert_eq!(undoc_last_error_kind(), UNDOC_ERROR_NONE);

        assert!(unsafe { undoc_get_author(doc) }.is_null());
        assert_eq!(undoc_last_error_kind(), UNDOC_ERROR_NONE);

        unsafe { undoc_free_document(doc) };
    }

    /// The counterpart to the test above: metadata that *exists* but cannot cross the
    /// ABI must not be reported as absent. Both cases return null, so the kind is the
    /// only thing that tells a caller "there is nothing" from "we could not give it
    /// to you".
    #[test]
    fn test_unrepresentable_metadata_is_not_reported_as_absent() {
        let mut document = Document::new();
        document.metadata.title = Some("has\0interior nul".to_string());
        document.metadata.author = Some("also\0bad".to_string());
        let doc = Box::into_raw(Box::new(UndocDocument {
            inner: document.into(),
        }));

        assert!(unsafe { undoc_get_title(doc) }.is_null());
        assert_eq!(undoc_last_error_kind(), UNDOC_ERROR_INVALID_OUTPUT);

        assert!(unsafe { undoc_get_author(doc) }.is_null());
        assert_eq!(undoc_last_error_kind(), UNDOC_ERROR_INVALID_OUTPUT);

        unsafe { undoc_free_document(doc) };
    }

    /// A failure must not leave its kind behind for the next call to be misread as
    /// still-failing — message and kind are cleared together.
    #[test]
    fn test_kind_is_cleared_by_the_next_call() {
        let path = CString::new("nonexistent-for-clear-test.docx").unwrap();
        let failed = unsafe { undoc_parse_file(path.as_ptr()) };
        assert!(failed.is_null());
        assert_ne!(undoc_last_error_kind(), UNDOC_ERROR_NONE);

        let doc = Box::into_raw(Box::new(UndocDocument {
            inner: Document::new().into(),
        }));
        let text = unsafe { undoc_plain_text(doc) };
        assert!(!text.is_null());
        assert_eq!(undoc_last_error_kind(), UNDOC_ERROR_NONE);

        unsafe {
            undoc_free_string(text);
            undoc_free_document(doc);
        }
    }

    #[test]
    fn test_get_resource_ids_empty_document_returns_non_null_empty_json() {
        let doc = Box::into_raw(Box::new(UndocDocument {
            inner: Document::new().into(),
        }));

        let ids = unsafe { undoc_get_resource_ids(doc) };
        assert!(!ids.is_null());
        let ids_str = unsafe { CStr::from_ptr(ids) }.to_str().unwrap();
        assert_eq!(ids_str, "[]");

        unsafe {
            undoc_free_string(ids);
            undoc_free_document(doc);
        }
    }

    /// `out_len` is written only once the call has reached the point of producing a
    /// buffer. A rejected argument leaves the caller's variable alone, so a caller that
    /// seeded it can tell "not attempted" from "attempted and produced nothing".
    #[test]
    fn test_rejected_arguments_leave_out_len_untouched() {
        let doc = Box::into_raw(Box::new(UndocDocument {
            inner: Document::new().into(),
        }));
        let id = CString::new("rId1").unwrap();
        const SEEDED: usize = 0xDEAD;

        let mut out_len: usize = SEEDED;
        assert!(
            unsafe { undoc_get_resource_data(ptr::null(), id.as_ptr(), &mut out_len) }.is_null()
        );
        assert_eq!(out_len, SEEDED, "a null document must not write out_len");

        assert!(unsafe { undoc_get_resource_data(doc, ptr::null(), &mut out_len) }.is_null());
        assert_eq!(out_len, SEEDED, "a null resource_id must not write out_len");

        // A resource that is merely absent *is* looked up, so the length is zeroed.
        assert!(unsafe { undoc_get_resource_data(doc, id.as_ptr(), &mut out_len) }.is_null());
        assert_eq!(out_len, 0, "a lookup that failed reports zero length");
        assert_eq!(
            undoc_last_error_kind(),
            ErrorKind::ResourceNotFound as c_int
        );

        unsafe { undoc_free_document(doc) };
    }

    #[test]
    fn test_free_null() {
        // Should not crash
        unsafe {
            undoc_free_document(ptr::null_mut());
            undoc_free_string(ptr::null_mut());
        }
    }

    // -----------------------------------------------------------------------------------------
    // undoc_render_section

    use crate::pptx::raster_fixtures::{deck, shape, solid};

    /// Parses `data`, panicking with the recorded error if that fails.
    fn parse(data: &[u8]) -> *mut UndocDocument {
        let doc = unsafe { undoc_parse_bytes(data.as_ptr(), data.len()) };
        assert!(
            !doc.is_null(),
            "parse failed: kind {}",
            undoc_last_error_kind()
        );
        doc
    }

    /// A 100 × 50 pt slide with a red rectangle in its left half.
    fn red_deck() -> Vec<u8> {
        deck(&shape("rect", (0, 0, 50, 50), "", &solid("FF0000"), ""), "")
    }

    #[test]
    fn test_render_section_paints_a_slide_to_png_with_its_report() {
        let doc = parse(&red_deck());
        let options = CString::new(r#"{"dpi": 72, "system_fonts": false}"#).unwrap();
        let mut len = 0usize;
        let mut info: *mut c_char = ptr::null_mut();
        let png = unsafe { undoc_render_section(doc, 0, options.as_ptr(), &mut len, &mut info) };
        assert!(!png.is_null(), "kind {}", undoc_last_error_kind());
        assert_eq!(undoc_last_error_kind(), UNDOC_ERROR_NONE);

        let bytes = unsafe { std::slice::from_raw_parts(png, len) }.to_vec();
        unsafe { undoc_free_bytes(png, len) };
        assert_eq!(&bytes[..8], b"\x89PNG\r\n\x1a\n");
        // The IHDR chunk: 100 × 50 pixels at 72 dpi, where a point is a pixel.
        assert_eq!(u32::from_be_bytes(bytes[16..20].try_into().unwrap()), 100);
        assert_eq!(u32::from_be_bytes(bytes[20..24].try_into().unwrap()), 50);

        let info: serde_json::Value = serde_json::from_str(&take_string(info)).unwrap();
        assert_eq!(info["width"], 100);
        assert_eq!(info["height"], 50);
        for gap in [
            "shapes",
            "images",
            "text_runs",
            "charts",
            "graphic_frames",
            "approximated_fills",
        ] {
            assert_eq!(info["gaps"][gap], 0, "{gap}: {info}");
        }
        assert_eq!(info["substituted_text_runs"], 0);

        unsafe { undoc_free_document(doc) };
    }

    /// The default resolution is 150 dpi, and `out_info` may be null.
    #[test]
    fn test_render_section_defaults_and_optional_report() {
        let doc = parse(&red_deck());
        let mut len = 0usize;
        let png = unsafe { undoc_render_section(doc, 0, ptr::null(), &mut len, ptr::null_mut()) };
        assert!(!png.is_null(), "kind {}", undoc_last_error_kind());
        let bytes = unsafe { std::slice::from_raw_parts(png, len) };
        // 100 pt at 150 dpi.
        assert_eq!(u32::from_be_bytes(bytes[16..20].try_into().unwrap()), 208);
        unsafe { undoc_free_bytes(png, len) };
        unsafe { undoc_free_document(doc) };
    }

    #[test]
    fn test_render_section_reports_an_index_the_presentation_does_not_have() {
        let doc = parse(&red_deck());
        let mut len = 7usize;
        let mut info: *mut c_char = ptr::null_mut();
        let png = unsafe { undoc_render_section(doc, 1, ptr::null(), &mut len, &mut info) };
        assert!(png.is_null());
        assert_eq!(len, 0);
        assert!(info.is_null());
        assert_eq!(
            undoc_last_error_kind(),
            ErrorKind::SectionOutOfRange as c_int
        );
        assert_eq!(undoc_last_error_kind(), 300);
        let message = unsafe { CStr::from_ptr(undoc_last_error()) }
            .to_str()
            .unwrap();
        assert!(message.contains("section 1"), "{message}");

        let png = unsafe { undoc_render_section(doc, -1, ptr::null(), &mut len, &mut info) };
        assert!(png.is_null());
        assert_eq!(undoc_last_error_kind(), UNDOC_ERROR_INVALID_ARGUMENT);
        unsafe { undoc_free_document(doc) };
    }

    #[test]
    fn test_render_section_of_a_document_that_is_not_a_presentation() {
        let doc = parse(&hello_docx());
        let mut len = 0usize;
        let png = unsafe { undoc_render_section(doc, 0, ptr::null(), &mut len, ptr::null_mut()) };
        assert!(png.is_null());
        assert_eq!(
            undoc_last_error_kind(),
            ErrorKind::UnsupportedFormat as c_int
        );
        unsafe { undoc_free_document(doc) };
    }

    #[test]
    fn test_render_section_rejects_bad_options_and_null_arguments() {
        let doc = parse(&red_deck());
        let mut len = 0usize;
        // A resolution the library cannot draw is a rendering failure, as in unpdf; options
        // that do not parse are the caller's argument.
        for (json, kind) in [
            (r#"{"dpi": 0}"#, ErrorKind::Render as c_int),
            (r#"{"dpi": -3}"#, ErrorKind::Render as c_int),
            (r#"{"fonts": []}"#, UNDOC_ERROR_INVALID_ARGUMENT),
            ("not json", UNDOC_ERROR_INVALID_ARGUMENT),
        ] {
            let options = CString::new(json).unwrap();
            let png = unsafe {
                undoc_render_section(doc, 0, options.as_ptr(), &mut len, ptr::null_mut())
            };
            assert!(png.is_null(), "{json}");
            assert_eq!(undoc_last_error_kind(), kind, "{json}");
        }
        let png =
            unsafe { undoc_render_section(doc, 0, ptr::null(), ptr::null_mut(), ptr::null_mut()) };
        assert!(png.is_null());
        assert_eq!(undoc_last_error_kind(), UNDOC_ERROR_INVALID_ARGUMENT);
        let png =
            unsafe { undoc_render_section(ptr::null(), 0, ptr::null(), &mut len, ptr::null_mut()) };
        assert!(png.is_null());
        assert_eq!(undoc_last_error_kind(), UNDOC_ERROR_INVALID_ARGUMENT);
        unsafe { undoc_free_document(doc) };
    }

    /// The file-path entry point keeps the presentation too.
    #[test]
    fn test_render_section_after_parsing_a_file() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("deck.pptx");
        std::fs::write(&path, red_deck()).unwrap();
        let c_path = CString::new(path.to_str().unwrap()).unwrap();
        let doc = unsafe { undoc_parse_file(c_path.as_ptr()) };
        assert!(!doc.is_null());
        let mut len = 0usize;
        let png = unsafe { undoc_render_section(doc, 0, ptr::null(), &mut len, ptr::null_mut()) };
        assert!(!png.is_null(), "kind {}", undoc_last_error_kind());
        unsafe { undoc_free_bytes(png, len) };
        unsafe { undoc_free_document(doc) };
    }
}
