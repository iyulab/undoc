//! Shared quick-xml decoding helpers.
//!
//! ## quick-xml 0.40+ entity model
//!
//! quick-xml 0.40 redesigned entity handling. A text node such as `A &amp; B`
//! is no longer delivered as a single [`Event::Text`] with the entity resolved
//! inline. Instead the reader splits it into `Text("A ")`, `GeneralRef("amp")`,
//! `Text(" B")`. Every loop that accumulates text therefore MUST handle
//! [`Event::GeneralRef`] via [`resolve_general_ref`], or every
//! `&amp; &lt; &gt; &#NN;` in a document silently disappears.
//!
//! A raw (dangling) `&` that does not form a reference is *ill-formed* XML and,
//! by default, aborts the whole part with `Error::IllFormed(UnclosedReference)`.
//! undoc's robustness goal is graceful degradation on malformed input, so every
//! reader is built via [`reader_for`], which enables `allow_dangling_amp` — the
//! stray `&` then arrives as raw [`Event::Text`] bytes and is preserved verbatim
//! (matching quick-xml 0.37's lenient behavior).
//!
//! ## End-of-line normalization
//!
//! quick-xml does not perform XML §2.11 EOL normalization, and CRs also re-enter
//! via `&#13;`/`&#xD;` character references (how Excel stores in-cell line breaks,
//! issue #7). Because `&#13;&#10;` now arrives as *two* separate `GeneralRef`
//! events, a lone CR cannot be collapsed at the event level. [`decode_text_lossy`]
//! still normalizes literal CR/CRLF inside a single text node, but callers that
//! reconstruct in-cell text from character references (XLSX) must additionally
//! run [`normalize_line_endings`] once on the fully accumulated string at the
//! flush point. The function is idempotent (`\r`-free input returns borrowed), so
//! the double pass is safe.
//!
//! Scope note (deliberate): only the XLSX paths add flush-point normalization,
//! because character-reference CRs are an Excel-specific convention. DOCX/PPTX
//! encode line breaks as elements (`<w:br/>`, `<w:cr/>`, `<a:br/>`), not `&#13;`,
//! so their direct-run text (a `GeneralRef` that becomes its own run) is not
//! re-normalized. A literal CR inside a single text node is still handled by
//! [`decode_text_lossy`]; a hypothetical `&#13;` inside a `<w:t>`/`<a:t>` run
//! would survive as a raw CR — accepted, as no conforming producer emits it.
//! DOCX body/table/textbox paragraphs are unaffected regardless, since they
//! round-trip through re-serialized XML that is normalized on the second parse.
//!
//! CDATA (`Event::CData`) is not handled anywhere: no loop handled it before
//! this migration and OOXML does not use CDATA sections, so there is no
//! regression. Add a `CData` arm alongside the `GeneralRef` arms if that ever
//! changes.
//!
//! [`Event::Text`]: quick_xml::events::Event::Text
//! [`Event::GeneralRef`]: quick_xml::events::Event::GeneralRef

use std::borrow::Cow;

use quick_xml::escape::resolve_predefined_entity;
use quick_xml::events::{BytesRef, BytesText};
use quick_xml::Reader;

use crate::error::{Error, Result};

/// Build a [`Reader`] with undoc's shared leniency policy.
///
/// Enables `allow_dangling_amp` so a stray `&` (ill-formed XML from a
/// non-conforming producer) is delivered as raw [`Event::Text`] rather than
/// aborting the entire part with `Error::IllFormed(UnclosedReference)`. This
/// preserves undoc's graceful-degradation contract under quick-xml 0.40+, where
/// the check is on by default.
///
/// Callers set `trim_text` themselves — it differs per call site (text content
/// must preserve `xml:space="preserve"` whitespace; attribute-only loops may
/// trim).
///
/// [`Event::Text`]: quick_xml::events::Event::Text
pub(crate) fn reader_for(xml: &str) -> Reader<&[u8]> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().allow_dangling_amp = true;
    reader
}

/// Resolve an [`Event::GeneralRef`] entity reference to its string value.
///
/// Handles numeric character references (`&#48;`, `&#x30;`) and the five
/// predefined XML entities (`&amp; &lt; &gt; &quot; &apos;`). An unrecognized
/// *named* entity (custom DTD or an empty `&;` token) is preserved literally as
/// `&name;` to avoid silent data loss.
///
/// A *malformed numeric* reference (`&#xZZ;`, a codepoint above U+10FFFF, a lone
/// surrogate, or an overflowing value) yields the empty string — it is dropped,
/// not preserved. This is a deliberate, minor divergence from quick-xml 0.37's
/// lenient path, which surfaced such tokens as raw `&#xZZ;` text. No conforming
/// OOXML producer emits an out-of-range character reference, so the input is
/// unreachable in practice; dropping keeps the numeric branch panic-free and
/// mirrors the sibling unhwp migration.
///
/// CR (`&#13;`) is intentionally NOT normalized here — normalization must happen
/// once at the flush point on the fully accumulated string, because a CRLF pair
/// arrives as two separate references (see module docs and
/// [`normalize_line_endings`]).
///
/// [`Event::GeneralRef`]: quick_xml::events::Event::GeneralRef
pub(crate) fn resolve_general_ref(r: &BytesRef<'_>) -> String {
    let name = match r.decode() {
        Ok(n) => n,
        Err(_) => return String::new(),
    };

    if let Some(num) = name.strip_prefix('#') {
        let codepoint = match num.strip_prefix(['x', 'X']) {
            Some(hex) => u32::from_str_radix(hex, 16).ok(),
            None => num.parse::<u32>().ok(),
        };
        return codepoint
            .and_then(char::from_u32)
            .map(String::from)
            .unwrap_or_default();
    }

    if let Some(value) = resolve_predefined_entity(&name) {
        return value.to_string();
    }

    // Unknown (custom DTD) or empty/malformed token: preserve the literal
    // reference rather than dropping content. OOXML does not define custom
    // entities, so this path is reached only by non-conforming producers.
    format!("&{name};")
}

/// Normalize line endings to LF.
///
/// Covers two layers quick-xml leaves untouched: XML §2.11 end-of-line
/// normalization for literal CR/CRLF from non-conforming producers, and CRs
/// that re-enter via `&#13;`/`&#xD;` character references — how Excel stores
/// in-cell line breaks. undoc extracts text, where a CR carries no meaning
/// distinct from LF but breaks Markdown pipe tables, so both collapse to LF.
///
/// Idempotent: input without `\r` is returned borrowed and unchanged.
pub(crate) fn normalize_line_endings(input: Cow<'_, str>) -> Cow<'_, str> {
    if !input.contains('\r') {
        return input;
    }
    let mut out = String::with_capacity(input.len());
    let mut chars = input.chars().peekable();
    while let Some(c) = chars.next() {
        if c == '\r' {
            if chars.peek() == Some(&'\n') {
                chars.next();
            }
            out.push('\n');
        } else {
            out.push(c);
        }
    }
    Cow::Owned(out)
}

/// Decode a [`Event::Text`] event into an owned `String` using lossy UTF-8
/// substitution.
///
/// Under quick-xml 0.40+ text events no longer carry entity references (those
/// arrive as [`Event::GeneralRef`] — see [`resolve_general_ref`]), so this only
/// decodes raw bytes and normalizes literal CR/CRLF within the node. Invalid
/// UTF-8 is replaced with U+FFFD rather than surfacing as an error, matching the
/// graceful-degradation goal for content text (paragraphs, cells, runs, labels).
///
/// [`Event::Text`]: quick_xml::events::Event::Text
/// [`Event::GeneralRef`]: quick_xml::events::Event::GeneralRef
pub(crate) fn decode_text_lossy(text: &BytesText<'_>) -> String {
    let raw = String::from_utf8_lossy(text.as_ref());
    normalize_line_endings(raw).into_owned()
}

/// Decode a [`Event::Text`] event into an owned `String`, requiring valid UTF-8.
///
/// Intended for metadata paths where invalid UTF-8 must surface as
/// `Error::XmlParse` with a location context rather than be silently replaced.
///
/// [`Event::Text`]: quick_xml::events::Event::Text
pub(crate) fn decode_text_strict(text: &BytesText<'_>, location: &str) -> Result<String> {
    let raw = std::str::from_utf8(text.as_ref())
        .map_err(|err| Error::xml_parse_with_context(err.to_string(), location))?;
    Ok(normalize_line_endings(Cow::Borrowed(raw)).into_owned())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn resolve(name: &str) -> String {
        resolve_general_ref(&BytesRef::new(name))
    }

    #[test]
    fn resolves_predefined_entities() {
        assert_eq!(resolve("amp"), "&");
        assert_eq!(resolve("lt"), "<");
        assert_eq!(resolve("gt"), ">");
        assert_eq!(resolve("quot"), "\"");
        assert_eq!(resolve("apos"), "'");
    }

    #[test]
    fn resolves_numeric_character_references() {
        assert_eq!(resolve("#65"), "A"); // decimal
        assert_eq!(resolve("#x41"), "A"); // lower-case hex
        assert_eq!(resolve("#X41"), "A"); // upper-case hex
        assert_eq!(resolve("#128512"), "\u{1F600}"); // astral plane
        assert_eq!(resolve("#44032"), "가"); // multi-byte (U+AC00)
    }

    #[test]
    fn unknown_entity_is_preserved_literally() {
        assert_eq!(resolve("nbsp"), "&nbsp;");
        assert_eq!(resolve("custom"), "&custom;");
    }

    #[test]
    fn empty_reference_is_preserved_literally() {
        // A `&;` token arrives from the reader as GeneralRef("") — preserve it
        // verbatim rather than dropping it (matches 0.37 lenient behavior).
        assert_eq!(resolve(""), "&;");
    }

    #[test]
    fn malformed_references_degrade_without_panicking() {
        // Adversarial inputs must never panic — the parser's robustness goal.
        assert_eq!(resolve("#"), ""); // empty numeric ref
        assert_eq!(resolve("#x"), ""); // empty hex ref
        assert_eq!(resolve("#xZZ"), ""); // non-hex digits
        assert_eq!(resolve("#x110000"), ""); // above the Unicode maximum
        assert_eq!(resolve("#xD800"), ""); // lone surrogate (not a scalar value)
        assert_eq!(resolve("#xFFFFFFFFFFFF"), ""); // overflows u32 parse
    }

    #[test]
    fn lossy_normalizes_literal_carriage_returns() {
        // quick-xml does not perform XML §2.11 EOL normalization, so literal
        // CR/CRLF from non-conforming producers must be normalized here. Bare
        // CR breaks Markdown pipe tables (issue #7).
        let text = BytesText::from_escaped("one\rtwo\r\nthree");
        assert_eq!(decode_text_lossy(&text), "one\ntwo\nthree");
    }

    #[test]
    fn strict_normalizes_literal_carriage_returns() {
        let text = BytesText::from_escaped("a\r\nb\rc");
        assert_eq!(decode_text_strict(&text, "test").unwrap(), "a\nb\nc");
    }

    #[test]
    fn normalize_is_idempotent_and_borrows_without_cr() {
        let clean = normalize_line_endings(Cow::Borrowed("no carriage returns"));
        assert!(matches!(clean, Cow::Borrowed(_)));
        // Double application collapses nothing further.
        let once = normalize_line_endings(Cow::Borrowed("a\r\nb"));
        let twice = normalize_line_endings(once.clone());
        assert_eq!(once, twice);
        assert_eq!(once, "a\nb");
    }
}
