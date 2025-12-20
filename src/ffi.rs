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
//! to retrieve the error message.
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

use std::cell::RefCell;
use std::ffi::{c_char, c_int, CStr, CString};
use std::panic::catch_unwind;
use std::ptr;

use crate::model::Document;
use crate::render::{JsonFormat, RenderOptions};

// Thread-local storage for the last error message.
thread_local! {
    static LAST_ERROR: RefCell<Option<CString>> = RefCell::new(None);
}

/// Set the last error message.
fn set_last_error(msg: &str) {
    LAST_ERROR.with(|e| {
        *e.borrow_mut() = CString::new(msg).ok();
    });
}

/// Clear the last error message.
fn clear_last_error() {
    LAST_ERROR.with(|e| {
        *e.borrow_mut() = None;
    });
}

/// Opaque handle to a parsed document.
#[repr(C)]
pub struct UndocDocument {
    inner: Document,
}

/// Flags for markdown rendering.
pub const UNDOC_FLAG_FRONTMATTER: c_int = 1;
pub const UNDOC_FLAG_ESCAPE_SPECIAL: c_int = 2;
pub const UNDOC_FLAG_PARAGRAPH_SPACING: c_int = 4;

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

/// Get the last error message.
///
/// # Safety
///
/// Returns a pointer to a thread-local error string. The pointer is valid until
/// the next call to any undoc function on the same thread.
#[no_mangle]
pub extern "C" fn undoc_last_error() -> *const c_char {
    LAST_ERROR.with(|e| {
        e.borrow()
            .as_ref()
            .map(|s| s.as_ptr())
            .unwrap_or(ptr::null())
    })
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
    clear_last_error();

    if path.is_null() {
        set_last_error("path is null");
        return ptr::null_mut();
    }

    let result = catch_unwind(|| {
        let path_str = CStr::from_ptr(path).to_str().map_err(|e| e.to_string())?;

        crate::parse_file(path_str)
            .map(|doc| Box::into_raw(Box::new(UndocDocument { inner: doc })))
            .map_err(|e| e.to_string())
    });

    match result {
        Ok(Ok(doc)) => doc,
        Ok(Err(e)) => {
            set_last_error(&e);
            ptr::null_mut()
        }
        Err(_) => {
            set_last_error("panic occurred during parsing");
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
    clear_last_error();

    if data.is_null() {
        set_last_error("data is null");
        return ptr::null_mut();
    }

    let result = catch_unwind(|| {
        let bytes = std::slice::from_raw_parts(data, len);

        crate::parse_bytes(bytes)
            .map(|doc| Box::into_raw(Box::new(UndocDocument { inner: doc })))
            .map_err(|e| e.to_string())
    });

    match result {
        Ok(Ok(doc)) => doc,
        Ok(Err(e)) => {
            set_last_error(&e);
            ptr::null_mut()
        }
        Err(_) => {
            set_last_error("panic occurred during parsing");
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
pub unsafe extern "C" fn undoc_to_markdown(doc: *const UndocDocument, flags: c_int) -> *mut c_char {
    clear_last_error();

    if doc.is_null() {
        set_last_error("document is null");
        return ptr::null_mut();
    }

    let result = catch_unwind(|| {
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

        crate::render::to_markdown(document, &options).map_err(|e| e.to_string())
    });

    match result {
        Ok(Ok(md)) => match CString::new(md) {
            Ok(s) => s.into_raw(),
            Err(_) => {
                set_last_error("output contains null byte");
                ptr::null_mut()
            }
        },
        Ok(Err(e)) => {
            set_last_error(&e);
            ptr::null_mut()
        }
        Err(_) => {
            set_last_error("panic occurred during rendering");
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
    clear_last_error();

    if doc.is_null() {
        set_last_error("document is null");
        return ptr::null_mut();
    }

    let result = catch_unwind(|| {
        let document = &(*doc).inner;
        let options = RenderOptions::default();
        crate::render::to_text(document, &options).map_err(|e| e.to_string())
    });

    match result {
        Ok(Ok(text)) => match CString::new(text) {
            Ok(s) => s.into_raw(),
            Err(_) => {
                set_last_error("output contains null byte");
                ptr::null_mut()
            }
        },
        Ok(Err(e)) => {
            set_last_error(&e);
            ptr::null_mut()
        }
        Err(_) => {
            set_last_error("panic occurred during rendering");
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
    clear_last_error();

    if doc.is_null() {
        set_last_error("document is null");
        return ptr::null_mut();
    }

    let result = catch_unwind(|| {
        let document = &(*doc).inner;
        let json_format = if format == UNDOC_JSON_COMPACT {
            JsonFormat::Compact
        } else {
            JsonFormat::Pretty
        };
        crate::render::to_json(document, json_format).map_err(|e| e.to_string())
    });

    match result {
        Ok(Ok(json)) => match CString::new(json) {
            Ok(s) => s.into_raw(),
            Err(_) => {
                set_last_error("output contains null byte");
                ptr::null_mut()
            }
        },
        Ok(Err(e)) => {
            set_last_error(&e);
            ptr::null_mut()
        }
        Err(_) => {
            set_last_error("panic occurred during rendering");
            ptr::null_mut()
        }
    }
}

/// Get the plain text content of a document.
///
/// # Safety
///
/// - `doc` must be a valid document handle.
/// - Returns null on error.
/// - The returned string must be freed with `undoc_free_string`.
#[no_mangle]
pub unsafe extern "C" fn undoc_plain_text(doc: *const UndocDocument) -> *mut c_char {
    clear_last_error();

    if doc.is_null() {
        set_last_error("document is null");
        return ptr::null_mut();
    }

    let result = catch_unwind(|| {
        let document = &(*doc).inner;
        document.plain_text()
    });

    match result {
        Ok(text) => match CString::new(text) {
            Ok(s) => s.into_raw(),
            Err(_) => {
                set_last_error("output contains null byte");
                ptr::null_mut()
            }
        },
        Err(_) => {
            set_last_error("panic occurred");
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
    if doc.is_null() {
        set_last_error("document is null");
        return -1;
    }

    match catch_unwind(|| (*doc).inner.sections.len() as c_int) {
        Ok(count) => count,
        Err(_) => {
            set_last_error("panic occurred");
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
    if doc.is_null() {
        set_last_error("document is null");
        return -1;
    }

    match catch_unwind(|| (*doc).inner.resources.len() as c_int) {
        Ok(count) => count,
        Err(_) => {
            set_last_error("panic occurred");
            -1
        }
    }
}

/// Get the document title.
///
/// # Safety
///
/// - `doc` must be a valid document handle.
/// - Returns null if no title is set.
/// - The returned string must be freed with `undoc_free_string`.
#[no_mangle]
pub unsafe extern "C" fn undoc_get_title(doc: *const UndocDocument) -> *mut c_char {
    clear_last_error();

    if doc.is_null() {
        set_last_error("document is null");
        return ptr::null_mut();
    }

    let result = catch_unwind(|| {
        (*doc)
            .inner
            .metadata
            .title
            .as_ref()
            .map(|t| CString::new(t.as_str()).ok())
            .flatten()
    });

    match result {
        Ok(Some(s)) => s.into_raw(),
        Ok(None) => ptr::null_mut(),
        Err(_) => {
            set_last_error("panic occurred");
            ptr::null_mut()
        }
    }
}

/// Get the document author.
///
/// # Safety
///
/// - `doc` must be a valid document handle.
/// - Returns null if no author is set.
/// - The returned string must be freed with `undoc_free_string`.
#[no_mangle]
pub unsafe extern "C" fn undoc_get_author(doc: *const UndocDocument) -> *mut c_char {
    clear_last_error();

    if doc.is_null() {
        set_last_error("document is null");
        return ptr::null_mut();
    }

    let result = catch_unwind(|| {
        (*doc)
            .inner
            .metadata
            .author
            .as_ref()
            .map(|a| CString::new(a.as_str()).ok())
            .flatten()
    });

    match result {
        Ok(Some(s)) => s.into_raw(),
        Ok(None) => ptr::null_mut(),
        Err(_) => {
            set_last_error("panic occurred");
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

#[cfg(test)]
mod tests {
    use super::*;
    use std::ffi::CString;
    use std::path::Path;

    #[test]
    fn test_version() {
        let version = unsafe { undoc_version() };
        assert!(!version.is_null());
        let version_str = unsafe { CStr::from_ptr(version) }.to_str().unwrap();
        assert!(!version_str.is_empty());
    }

    #[test]
    fn test_parse_null_path() {
        let doc = unsafe { undoc_parse_file(ptr::null()) };
        assert!(doc.is_null());

        let error = unsafe { undoc_last_error() };
        assert!(!error.is_null());
    }

    #[test]
    fn test_parse_invalid_path() {
        let path = CString::new("nonexistent.docx").unwrap();
        let doc = unsafe { undoc_parse_file(path.as_ptr()) };
        assert!(doc.is_null());

        let error = unsafe { undoc_last_error() };
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
    fn test_free_null() {
        // Should not crash
        unsafe {
            undoc_free_document(ptr::null_mut());
            undoc_free_string(ptr::null_mut());
        }
    }
}
