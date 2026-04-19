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
pub(crate) fn lenient_unescape(input: &str) -> Cow<'_, str> {
    match quick_xml::escape::unescape(input) {
        Ok(cow) => cow,
        Err(_) => Cow::Owned(lenient_slow_path(input)),
    }
}

fn lenient_slow_path(input: &str) -> String {
    let bytes = input.as_bytes();
    let mut out = String::with_capacity(input.len());
    let mut i = 0;

    while i < bytes.len() {
        // Find the next '&' from the cursor. ASCII-safe: '&' (0x26) cannot
        // appear as a UTF-8 continuation byte, so every index is a char
        // boundary and `&input[..]` slicing is valid.
        match bytes[i..].iter().position(|&b| b == b'&') {
            None => {
                out.push_str(&input[i..]);
                break;
            }
            Some(rel) => {
                let amp = i + rel;
                out.push_str(&input[i..amp]);

                let window_end = (amp + MAX_ENTITY_LEN).min(bytes.len());
                match bytes[amp + 1..window_end].iter().position(|&b| b == b';') {
                    None => {
                        // Stray '&' without a closing ';' in range.
                        out.push('&');
                        i = amp + 1;
                    }
                    Some(srel) => {
                        let semi = amp + 1 + srel;
                        let token = &input[amp..=semi]; // "&...;"
                        match quick_xml::escape::unescape(token) {
                            Ok(decoded) => {
                                out.push_str(&decoded);
                                i = semi + 1;
                            }
                            Err(_) => {
                                // Preserve only the leading '&' raw; let the
                                // next iteration re-scan the span so any
                                // legitimate entity inside still decodes.
                                out.push('&');
                                i = amp + 1;
                            }
                        }
                    }
                }
            }
        }
    }

    out
}

/// Decode a `BytesText` event into an owned `String` using lossy UTF-8
/// substitution and the lenient entity path.
///
/// Intended for content text (paragraphs, cells, runs, chart labels) where
/// invalid UTF-8 bytes should be replaced with U+FFFD rather than surface
/// as an error.
pub(crate) fn decode_text_lossy(text: &BytesText<'_>) -> String {
    let raw = String::from_utf8_lossy(text.as_ref());
    lenient_unescape(raw.as_ref()).into_owned()
}

/// Decode a `BytesText` event into an owned `String`, requiring valid UTF-8.
///
/// Intended for metadata and other paths where invalid UTF-8 must surface as
/// `Error::XmlParse` with a location context rather than be silently
/// replaced.
pub(crate) fn decode_text_strict(text: &BytesText<'_>, location: &str) -> Result<String> {
    let raw = std::str::from_utf8(text.as_ref())
        .map_err(|err| Error::xml_parse_with_context(err.to_string(), location))?;
    Ok(lenient_unescape(raw).into_owned())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn fast_path_no_entities_borrows() {
        let input = "plain text without references";
        let out = lenient_unescape(input);
        assert_eq!(out, "plain text without references");
        assert!(
            matches!(out, Cow::Borrowed(_)),
            "expected Borrowed, got Owned"
        );
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

    #[test]
    fn slow_path_mixed_legitimate_and_malformed() {
        assert_eq!(lenient_unescape("A &amp; B &bogus; C"), "A & B &bogus; C");
    }

    #[test]
    fn slow_path_multiple_malformed_preserved() {
        assert_eq!(lenient_unescape("&foo;&bar;"), "&foo;&bar;");
    }

    #[test]
    fn slow_path_stray_ampersand_then_malformed_entity() {
        assert_eq!(
            lenient_unescape("R&D and &bogus; tail"),
            "R&D and &bogus; tail"
        );
    }

    #[test]
    fn slow_path_stray_ampersand_then_legitimate_entity() {
        // The `&` in "R&D" is a stray, and the `;` of the following
        // `&amp;` falls within its lookahead window. A naive greedy
        // scanner would consume "&D &amp;" as one failing token and
        // lose the legitimate decoding. Correct behavior: preserve the
        // stray `&` raw and re-scan so `&amp;` decodes on its own.
        assert_eq!(lenient_unescape("R&D &amp; tail"), "R&D & tail");
    }

    #[test]
    fn slow_path_adjacent_ampersand_before_legitimate_entity() {
        // Adjacent case: stray `&` immediately followed by a legitimate
        // entity. The `&&lt;` span fails as a whole; recovery must push
        // the first `&` and re-scan so `&lt;` decodes.
        assert_eq!(lenient_unescape("&&lt;"), "&<");
    }

    #[test]
    fn slow_path_unterminated_ampersand_bounded() {
        // No `;` within MAX_ENTITY_LEN bytes after the stray `&`: the `&`
        // is pushed raw and scanning continues. The trailing `&bogus;`
        // triggers the slow path (fast-path unescape fails on either
        // malformed token); both are preserved verbatim, so output equals
        // input.
        let padding = "x".repeat(MAX_ENTITY_LEN + 8);
        let input = format!("&{padding}&bogus;");
        assert_eq!(lenient_unescape(&input), input);
    }

    #[test]
    fn slow_path_numeric_mixed_with_malformed() {
        assert_eq!(lenient_unescape("&#65;&bogus;&#x42;"), "A&bogus;B");
    }

    #[test]
    fn slow_path_preserves_non_ascii_between_tokens() {
        assert_eq!(
            lenient_unescape("한글 &amp; &bogus; 日本語"),
            "한글 & &bogus; 日本語"
        );
    }

    #[test]
    fn slow_path_empty_entity_reference() {
        assert_eq!(lenient_unescape("a &; b &bogus;"), "a &; b &bogus;");
    }
}
