//! Lenient XML entity decoding.
//!
//! Preserves malformed entity tokens (e.g. `&bogus;`) as raw text while
//! decoding legitimate entities (standard five, numeric refs) in the same
//! node. Complements `quick_xml::escape::unescape`'s all-or-nothing behavior.

use std::borrow::Cow;

use quick_xml::events::BytesText;

use crate::error::{Error, Result};

/// Maximum bytes after a stray `&` to search for a closing `;` in the slow
/// path. Covers the longest XML-spec-supported reference
/// (`&#x10FFFF;` = 10 bytes) with headroom. HTML5 named entities are
/// intentionally not covered — `quick_xml::escape::unescape` does not
/// decode them, so they fail per-token and are preserved raw regardless.
const MAX_ENTITY_LEN: usize = 16;
