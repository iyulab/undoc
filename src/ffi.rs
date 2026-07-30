//! C-ABI Foreign Function Interface for undoc.
//!
//! This module provides C-compatible bindings for using undoc from other languages
//! such as C, C++, C#, Python, and any language with C FFI support.
//!
//! # Memory Management
//!
//! All strings returned by this library must be freed using `undoc_free_string`.
//! All document handles must be freed using `undoc_free_document`.
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

use std::ffi::{c_char, c_int, CStr, CString};
use std::ptr;

use uncore::ffi::{self, invalid_argument, FfiError, LastErrorSlot};

use crate::error::ErrorKind;
use crate::model::Document;
use crate::render::{JsonFormat, RenderOptions};

// Thread-local storage for the last error message and its classification. Declared
// here rather than in `uncore` — see that crate's `ffi` module docs for why the slot
// must live in the consuming crate.
thread_local! {
    static LAST_ERROR: LastErrorSlot = const { LastErrorSlot::new() };
}

uncore::export_last_error_abi!(LAST_ERROR, undoc_last_error, undoc_last_error_kind);

/// `undoc_last_error_kind` value when no error is recorded on this thread.
pub const UNDOC_ERROR_NONE: c_int = uncore::kind::NONE;

// Values 1..=13 are [`ErrorKind`] discriminants — core failure reasons.
// Values 100+ are FFI-boundary reasons with no core `Error` counterpart.

/// An argument was null or not valid UTF-8.
pub const UNDOC_ERROR_INVALID_ARGUMENT: c_int = uncore::kind::INVALID_ARGUMENT;
/// A panic was caught at the FFI boundary.
pub const UNDOC_ERROR_PANIC: c_int = uncore::kind::PANIC;
/// The produced output contains an interior NUL byte and cannot cross the C ABI.
pub const UNDOC_ERROR_INVALID_OUTPUT: c_int = uncore::kind::INVALID_OUTPUT;

/// Classify a core error and render its message, for return from a closure.
fn ffi_err(e: crate::Error) -> FfiError {
    (e.kind() as c_int, e.to_string())
}

/// Classify a JSON serialization failure — producing output is rendering.
fn json_err(e: serde_json::Error) -> FfiError {
    (ErrorKind::Render as c_int, e.to_string())
}

/// Classify a non-UTF-8 string argument received across the ABI.
fn utf8_err(e: std::str::Utf8Error) -> FfiError {
    invalid_argument(e.to_string())
}

/// Opaque handle to a parsed document.
#[repr(C)]
pub struct UndocDocument {
    inner: Document,
}

/// Flags for markdown rendering.
pub const UNDOC_FLAG_FRONTMATTER: u32 = 1;
pub const UNDOC_FLAG_ESCAPE_SPECIAL: u32 = 2;
pub const UNDOC_FLAG_PARAGRAPH_SPACING: u32 = 4;

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

    if path.is_null() {
        LAST_ERROR.with(|slot| slot.set_error(&invalid_argument("path is null")));
        return ptr::null_mut();
    }

    let result: Result<*mut UndocDocument, FfiError> = ffi::catch(|| {
        let path_str = CStr::from_ptr(path).to_str().map_err(utf8_err)?;

        crate::parse_file(path_str)
            .map(|doc| Box::into_raw(Box::new(UndocDocument { inner: doc })))
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

        crate::parse_bytes(bytes)
            .map(|doc| Box::into_raw(Box::new(UndocDocument { inner: doc })))
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

/// Free a document handle.
///
/// # Safety
///
/// - `doc` must be a valid pointer returned by `undoc_parse_file` or `undoc_parse_bytes`.
/// - After calling this function, the handle is invalid and must not be used.
#[no_mangle]
pub unsafe extern "C" fn undoc_free_document(doc: *mut UndocDocument) {
    if !doc.is_null() {
        let _ = Box::from_raw(doc);
    }
}

/// Convert a document to Markdown.
///
/// # Safety
///
/// - `doc` must be a valid document handle.
/// - `flags` is a bitwise OR of `UNDOC_FLAG_*` constants.
/// - Returns null on error. Use `undoc_last_error` to get the error message.
/// - The returned string must be freed with `undoc_free_string`.
#[no_mangle]
pub unsafe extern "C" fn undoc_to_markdown(doc: *const UndocDocument, flags: u32) -> *mut c_char {
    LAST_ERROR.with(|slot| slot.clear());

    if doc.is_null() {
        LAST_ERROR.with(|slot| slot.set_error(&invalid_argument("document is null")));
        return ptr::null_mut();
    }

    let result: Result<String, FfiError> = ffi::catch(|| {
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

        crate::render::to_markdown(document, &options).map_err(ffi_err)
    });

    match result {
        Ok(md) => match CString::new(md) {
            Ok(s) => s.into_raw(),
            Err(_) => {
                LAST_ERROR.with(|slot| slot.set_error(&ffi::invalid_output()));
                ptr::null_mut()
            }
        },
        Err(error) => {
            LAST_ERROR.with(|slot| slot.set_error(&error));
            ptr::null_mut()
        }
    }
}

/// Convert a document to plain text.
///
/// # Safety
///
/// - `doc` must be a valid document handle.
/// - Returns null on error. Use `undoc_last_error` to get the error message.
/// - The returned string must be freed with `undoc_free_string`.
#[no_mangle]
pub unsafe extern "C" fn undoc_to_text(doc: *const UndocDocument) -> *mut c_char {
    LAST_ERROR.with(|slot| slot.clear());

    if doc.is_null() {
        LAST_ERROR.with(|slot| slot.set_error(&invalid_argument("document is null")));
        return ptr::null_mut();
    }

    let result: Result<String, FfiError> = ffi::catch(|| {
        let document = &(*doc).inner;
        let options = RenderOptions::default();
        crate::render::to_text(document, &options).map_err(ffi_err)
    });

    match result {
        Ok(text) => match CString::new(text) {
            Ok(s) => s.into_raw(),
            Err(_) => {
                LAST_ERROR.with(|slot| slot.set_error(&ffi::invalid_output()));
                ptr::null_mut()
            }
        },
        Err(error) => {
            LAST_ERROR.with(|slot| slot.set_error(&error));
            ptr::null_mut()
        }
    }
}

/// Convert a document to JSON.
///
/// # Safety
///
/// - `doc` must be a valid document handle.
/// - `format` is one of `UNDOC_JSON_PRETTY` or `UNDOC_JSON_COMPACT`.
/// - Returns null on error. Use `undoc_last_error` to get the error message.
/// - The returned string must be freed with `undoc_free_string`.
#[no_mangle]
pub unsafe extern "C" fn undoc_to_json(doc: *const UndocDocument, format: c_int) -> *mut c_char {
    LAST_ERROR.with(|slot| slot.clear());

    if doc.is_null() {
        LAST_ERROR.with(|slot| slot.set_error(&invalid_argument("document is null")));
        return ptr::null_mut();
    }

    let result: Result<String, FfiError> = ffi::catch(|| {
        let document = &(*doc).inner;
        let json_format = if format == UNDOC_JSON_COMPACT {
            JsonFormat::Compact
        } else {
            JsonFormat::Pretty
        };
        crate::render::to_json(document, json_format).map_err(ffi_err)
    });

    match result {
        Ok(json) => match CString::new(json) {
            Ok(s) => s.into_raw(),
            Err(_) => {
                LAST_ERROR.with(|slot| slot.set_error(&ffi::invalid_output()));
                ptr::null_mut()
            }
        },
        Err(error) => {
            LAST_ERROR.with(|slot| slot.set_error(&error));
            ptr::null_mut()
        }
    }
}

/// Get the plain text content of a document.
///
/// # Safety
///
/// - `doc` must be a valid document handle.
/// - Returns null on error; use `undoc_last_error` to get the error message.
/// - A valid document with no text returns a non-null, empty (`""`) string.
/// - The returned string must be freed with `undoc_free_string`.
#[no_mangle]
pub unsafe extern "C" fn undoc_plain_text(doc: *const UndocDocument) -> *mut c_char {
    LAST_ERROR.with(|slot| slot.clear());

    if doc.is_null() {
        LAST_ERROR.with(|slot| slot.set_error(&invalid_argument("document is null")));
        return ptr::null_mut();
    }

    let result: Result<String, FfiError> = ffi::catch(|| {
        let document = &(*doc).inner;
        Ok(document.plain_text())
    });

    match result {
        Ok(text) => match CString::new(text) {
            Ok(s) => s.into_raw(),
            Err(_) => {
                LAST_ERROR.with(|slot| slot.set_error(&ffi::invalid_output()));
                ptr::null_mut()
            }
        },
        Err(error) => {
            LAST_ERROR.with(|slot| slot.set_error(&error));
            ptr::null_mut()
        }
    }
}

/// Get the number of sections in a document.
///
/// # Safety
///
/// - `doc` must be a valid document handle.
/// - Returns -1 on error.
#[no_mangle]
pub unsafe extern "C" fn undoc_section_count(doc: *const UndocDocument) -> c_int {
    LAST_ERROR.with(|slot| slot.clear());

    if doc.is_null() {
        LAST_ERROR.with(|slot| slot.set_error(&invalid_argument("document is null")));
        return -1;
    }

    match ffi::catch(|| Ok((*doc).inner.sections.len() as c_int)) {
        Ok(count) => count,
        Err(error) => {
            LAST_ERROR.with(|slot| slot.set_error(&error));
            -1
        }
    }
}

/// Get the number of resources in a document.
///
/// # Safety
///
/// - `doc` must be a valid document handle.
/// - Returns -1 on error.
#[no_mangle]
pub unsafe extern "C" fn undoc_resource_count(doc: *const UndocDocument) -> c_int {
    LAST_ERROR.with(|slot| slot.clear());

    if doc.is_null() {
        LAST_ERROR.with(|slot| slot.set_error(&invalid_argument("document is null")));
        return -1;
    }

    match ffi::catch(|| Ok((*doc).inner.resources.len() as c_int)) {
        Ok(count) => count,
        Err(error) => {
            LAST_ERROR.with(|slot| slot.set_error(&error));
            -1
        }
    }
}

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
#[no_mangle]
pub unsafe extern "C" fn undoc_get_title(doc: *const UndocDocument) -> *mut c_char {
    LAST_ERROR.with(|slot| slot.clear());

    if doc.is_null() {
        LAST_ERROR.with(|slot| slot.set_error(&invalid_argument("document is null")));
        return ptr::null_mut();
    }

    let result: Result<Option<CString>, FfiError> =
        ffi::catch(|| match (*doc).inner.metadata.title.as_ref() {
            Some(title) => CString::new(title.as_str())
                .map(Some)
                .map_err(|_| ffi::invalid_output()),
            None => Ok(None),
        });

    match result {
        Ok(Some(s)) => s.into_raw(),
        Ok(None) => ptr::null_mut(),
        Err(error) => {
            LAST_ERROR.with(|slot| slot.set_error(&error));
            ptr::null_mut()
        }
    }
}

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
#[no_mangle]
pub unsafe extern "C" fn undoc_get_author(doc: *const UndocDocument) -> *mut c_char {
    LAST_ERROR.with(|slot| slot.clear());

    if doc.is_null() {
        LAST_ERROR.with(|slot| slot.set_error(&invalid_argument("document is null")));
        return ptr::null_mut();
    }

    let result: Result<Option<CString>, FfiError> =
        ffi::catch(|| match (*doc).inner.metadata.author.as_ref() {
            Some(author) => CString::new(author.as_str())
                .map(Some)
                .map_err(|_| ffi::invalid_output()),
            None => Ok(None),
        });

    match result {
        Ok(Some(s)) => s.into_raw(),
        Ok(None) => ptr::null_mut(),
        Err(error) => {
            LAST_ERROR.with(|slot| slot.set_error(&error));
            ptr::null_mut()
        }
    }
}

/// Free a string allocated by this library.
///
/// # Safety
///
/// - `s` must be a pointer returned by an undoc function, or null.
/// - After calling this function, the pointer is invalid and must not be used.
#[no_mangle]
pub unsafe extern "C" fn undoc_free_string(s: *mut c_char) {
    if !s.is_null() {
        let _ = CString::from_raw(s);
    }
}

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
#[no_mangle]
pub unsafe extern "C" fn undoc_get_resource_ids(doc: *const UndocDocument) -> *mut c_char {
    LAST_ERROR.with(|slot| slot.clear());

    if doc.is_null() {
        LAST_ERROR.with(|slot| slot.set_error(&invalid_argument("document is null")));
        return ptr::null_mut();
    }

    let result: Result<String, FfiError> = ffi::catch(|| {
        let document = &(*doc).inner;
        let ids: Vec<&String> = document.resources.keys().collect();
        serde_json::to_string(&ids).map_err(json_err)
    });

    match result {
        Ok(json) => match CString::new(json) {
            Ok(s) => s.into_raw(),
            Err(_) => {
                LAST_ERROR.with(|slot| slot.set_error(&ffi::invalid_output()));
                ptr::null_mut()
            }
        },
        Err(error) => {
            LAST_ERROR.with(|slot| slot.set_error(&error));
            ptr::null_mut()
        }
    }
}

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
/// `{"id":"rId1","type":"image","filename":"image1.png","mime_type":"image/png","size":1024,"width":800,"height":600,"alt_text":"Description"}`
#[no_mangle]
pub unsafe extern "C" fn undoc_get_resource_info(
    doc: *const UndocDocument,
    resource_id: *const c_char,
) -> *mut c_char {
    LAST_ERROR.with(|slot| slot.clear());

    if doc.is_null() {
        LAST_ERROR.with(|slot| slot.set_error(&invalid_argument("document is null")));
        return ptr::null_mut();
    }

    if resource_id.is_null() {
        LAST_ERROR.with(|slot| slot.set_error(&invalid_argument("resource_id is null")));
        return ptr::null_mut();
    }

    let result: Result<String, FfiError> = ffi::catch(|| {
        let id_str = CStr::from_ptr(resource_id).to_str().map_err(utf8_err)?;

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
                    "alt_text": resource.alt_text
                });
                serde_json::to_string(&info).map_err(json_err)
            }
            None => Err(ffi_err(crate::Error::ResourceNotFound(id_str.to_string()))),
        }
    });

    match result {
        Ok(json) => match CString::new(json) {
            Ok(s) => s.into_raw(),
            Err(_) => {
                LAST_ERROR.with(|slot| slot.set_error(&ffi::invalid_output()));
                ptr::null_mut()
            }
        },
        Err(error) => {
            LAST_ERROR.with(|slot| slot.set_error(&error));
            ptr::null_mut()
        }
    }
}

/// Get resource binary data.
///
/// # Safety
///
/// - `doc` must be a valid document handle.
/// - `resource_id` must be a valid null-terminated UTF-8 string.
/// - `out_len` must be a valid pointer to receive the data length.
/// - Returns null if resource not found or on error.
/// - The returned pointer must be freed with `undoc_free_bytes`.
#[no_mangle]
pub unsafe extern "C" fn undoc_get_resource_data(
    doc: *const UndocDocument,
    resource_id: *const c_char,
    out_len: *mut usize,
) -> *mut u8 {
    LAST_ERROR.with(|slot| slot.clear());

    if doc.is_null() {
        LAST_ERROR.with(|slot| slot.set_error(&invalid_argument("document is null")));
        return ptr::null_mut();
    }

    if resource_id.is_null() {
        LAST_ERROR.with(|slot| slot.set_error(&invalid_argument("resource_id is null")));
        return ptr::null_mut();
    }

    if out_len.is_null() {
        LAST_ERROR.with(|slot| slot.set_error(&invalid_argument("out_len is null")));
        return ptr::null_mut();
    }

    let result: Result<(*mut u8, usize), FfiError> = ffi::catch(|| {
        let id_str = CStr::from_ptr(resource_id).to_str().map_err(utf8_err)?;

        let document = &(*doc).inner;

        match document.resources.get(id_str) {
            Some(resource) => {
                let data = resource.data.clone();
                let len = data.len();
                let boxed = data.into_boxed_slice();
                let ptr = Box::into_raw(boxed) as *mut u8;
                Ok((ptr, len))
            }
            None => Err(ffi_err(crate::Error::ResourceNotFound(id_str.to_string()))),
        }
    });

    match result {
        Ok((ptr, len)) => {
            *out_len = len;
            ptr
        }
        Err(error) => {
            LAST_ERROR.with(|slot| slot.set_error(&error));
            *out_len = 0;
            ptr::null_mut()
        }
    }
}

/// Free binary data allocated by `undoc_get_resource_data`.
///
/// # Safety
///
/// - `data` must be a pointer returned by `undoc_get_resource_data`, or null.
/// - `len` must be the length returned by `undoc_get_resource_data`.
/// - After calling this function, the pointer is invalid and must not be used.
#[no_mangle]
pub unsafe extern "C" fn undoc_free_bytes(data: *mut u8, len: usize) {
    if !data.is_null() && len > 0 {
        let _ = Box::from_raw(std::ptr::slice_from_raw_parts_mut(data, len));
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::Document;
    use std::ffi::CString;
    use std::path::Path;

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

    #[test]
    fn test_parse_and_convert() {
        let path = "test-files/file-sample_1MB.docx";
        if !Path::new(path).exists() {
            return;
        }

        let path_cstr = CString::new(path).unwrap();
        let doc = unsafe { undoc_parse_file(path_cstr.as_ptr()) };
        assert!(!doc.is_null());

        // Test markdown conversion
        let md = unsafe { undoc_to_markdown(doc, 0) };
        assert!(!md.is_null());
        unsafe { undoc_free_string(md) };

        // Test text conversion
        let text = unsafe { undoc_to_text(doc) };
        assert!(!text.is_null());
        unsafe { undoc_free_string(text) };

        // Test JSON conversion
        let json = unsafe { undoc_to_json(doc, UNDOC_JSON_PRETTY) };
        assert!(!json.is_null());
        unsafe { undoc_free_string(json) };

        // Test section count
        let count = unsafe { undoc_section_count(doc) };
        assert!(count >= 0);

        // Free document
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
            inner: Document::new(),
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
            inner: Document::new(),
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
            inner: Document::new(),
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
            inner: Document::new(),
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
            inner: Document::new(),
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
        let doc = Box::into_raw(Box::new(UndocDocument { inner: document }));

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
            inner: Document::new(),
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
            inner: Document::new(),
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
            inner: Document::new(),
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
}
