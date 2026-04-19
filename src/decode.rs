//! Lenient XML entity decoding.
//!
//! Preserves malformed entity tokens (e.g. `&bogus;`) as raw text while
//! decoding legitimate entities (standard five, numeric refs) in the same
//! node. Complements `quick_xml::escape::unescape`'s all-or-nothing behavior.

use std::borrow::Cow;

/// Maximum bytes after a stray `&` to search for a closing `;` in the slow
/// path. Covers the longest XML-spec-supported reference
/// (`&#x10FFFF;` = 10 bytes) with headroom. HTML5 named entities are
/// intentionally not covered — `quick_xml::escape::unescape` does not
/// decode them, so they fail per-token and are preserved raw regardless.
#[allow(dead_code)] // consumed by slow path in Task 3
const MAX_ENTITY_LEN: usize = 16;

/// Decode XML entity references with graceful handling of malformed tokens.
///
/// Fast path: if `quick_xml::escape::unescape` succeeds on the whole input,
/// its result is returned directly — `Cow::Borrowed` when no substitution
/// was needed, `Cow::Owned` when at least one legitimate entity was decoded.
///
/// Slow path (only when the fast path errors on an unknown/malformed entity):
/// scans `&...;` tokens linearly and decodes each independently, preserving
/// undecodable tokens as raw text. See `MAX_ENTITY_LEN` for the lookahead
/// bound.
#[allow(dead_code)] // wired into parsers in Tasks 4/5
pub(crate) fn lenient_unescape(input: &str) -> Cow<'_, str> {
    match quick_xml::escape::unescape(input) {
        Ok(cow) => cow,
        Err(_) => Cow::Owned(lenient_slow_path(input)),
    }
}

#[allow(dead_code)] // real body lands in Task 3
fn lenient_slow_path(input: &str) -> String {
    // Placeholder — real implementation lands in Task 3. Fast path is the
    // only reachable branch until slow-path tests exist.
    input.to_string()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn fast_path_no_entities_borrows() {
        let input = "plain text without references";
        let out = lenient_unescape(input);
        assert_eq!(out, "plain text without references");
        assert!(matches!(out, Cow::Borrowed(_)), "expected Borrowed, got Owned");
    }

    #[test]
    fn fast_path_legitimate_only_decodes() {
        assert_eq!(lenient_unescape("A &amp; B"), "A & B");
        assert_eq!(lenient_unescape("&lt;tag&gt;"), "<tag>");
        assert_eq!(lenient_unescape("&quot;x&quot;"), "\"x\"");
        assert_eq!(lenient_unescape("don&apos;t"), "don't");
    }

    #[test]
    fn fast_path_numeric_refs_decode() {
        assert_eq!(lenient_unescape("&#65;&#x41;"), "AA");
        assert_eq!(lenient_unescape("&#128512;"), "\u{1F600}");
    }
}
