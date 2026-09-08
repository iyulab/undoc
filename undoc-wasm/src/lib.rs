mod document;

pub use document::OfficeDocument;

use wasm_bindgen::prelude::*;

/// The formats this package can parse, as JSON.
///
/// Returns `[{"extension":"docx","name":"Word Document"}, ...]`.
///
/// A consumer that decides which files to hand to this package otherwise keeps its own copy
/// of the extension list, and that copy goes stale the moment this package learns a new
/// format -- silently, because nothing compares the two. Asking the package removes the
/// second copy. The name is included so a consumer can label a file without inventing its
/// own wording for a format this package already names.
#[wasm_bindgen(js_name = supportedFormats)]
pub fn supported_formats() -> Result<String, JsValue> {
    let formats: Vec<_> = undoc::FormatType::ALL
        .iter()
        .map(|format| serde_json::json!({ "extension": format.extension(), "name": format.name() }))
        .collect();
    serde_json::to_string(&formats).map_err(json_error)
}

#[wasm_bindgen]
pub fn parse(data: &[u8]) -> Result<OfficeDocument, JsValue> {
    undoc::parse_bytes(data)
        .map(|inner| OfficeDocument { inner })
        .map_err(undoc_error)
}

/// Build the JS error for a failed call, carrying both its message and its reason.
///
/// A bare string would force callers to match on message text. Throwing a real `Error`
/// with a numeric `kind` gives JavaScript the same contract the C ABI offers through
/// `undoc_last_error_kind`: branch on the reason, and treat an unrecognised number as a
/// generic failure, since new reasons take new numbers and existing ones never change.
///
/// Every fallible entry point in this crate goes through here, so no throw site can
/// quietly drop the classification and leave JavaScript with a bare string.
fn js_error(message: String, kind: i32) -> JsValue {
    let error = js_sys::Error::new(&message);
    let assigned = js_sys::Reflect::set(
        &error,
        &JsValue::from_str("kind"),
        &JsValue::from_f64(kind as f64),
    );
    debug_assert!(assigned.is_ok(), "kind must be assignable on a fresh Error");
    error.into()
}

/// A failure that came from the library itself, reported with its own reason.
pub(crate) fn undoc_error(e: undoc::Error) -> JsValue {
    js_error(e.to_string(), e.kind() as i32)
}

/// A failure to serialise a result. Producing output is rendering, and it stays
/// rendering when the last step of producing it is serialisation — so this is
/// [`undoc::ErrorKind::Render`] rather than a generic failure.
pub(crate) fn json_error(e: serde_json::Error) -> JsValue {
    js_error(e.to_string(), undoc::ErrorKind::Render as i32)
}

#[cfg(test)]
mod tests {
    use super::*;
    use wasm_bindgen_test::*;

    wasm_bindgen_test_configure!(run_in_node_experimental);

    #[wasm_bindgen_test]
    fn test_parse_invalid_returns_error() {
        let result = parse(b"garbage data");
        assert!(result.is_err());
    }

    /// The documented contract is that a caller can branch on the reason instead of the
    /// message. That only holds if the property is actually there.
    #[wasm_bindgen_test]
    fn test_parse_error_carries_its_kind() {
        let error = parse(b"garbage data")
            .err()
            .expect("garbage must not parse");
        let kind = js_sys::Reflect::get(&error, &JsValue::from_str("kind"))
            .expect("a thrown error must expose kind");

        assert_eq!(
            kind.as_f64(),
            Some(undoc::ErrorKind::UnknownFormat as i32 as f64),
            "bytes that are not an Office container are an unknown format"
        );
    }

    /// Plain `#[test]`, not `#[wasm_bindgen_test]`: this is pure serialisation with no JS
    /// interop, so running it on the host under `cargo test --workspace` covers it in CI's
    /// ordinary test job as well as being runnable locally.
    ///
    /// The exact JSON is pinned on purpose. A silent key rename would leave every consumer's
    /// lookup returning nothing, with no error anywhere to notice it.
    #[test]
    fn supported_formats_reports_every_format_the_library_parses() {
        let expected = concat!(
            r#"[{"extension":"docx","name":"Word Document"},"#,
            r#"{"extension":"xlsx","name":"Excel Workbook"},"#,
            r#"{"extension":"pptx","name":"PowerPoint Presentation"}]"#
        );
        assert_eq!(
            supported_formats().expect("serialising a fixed list cannot fail"),
            expected
        );
    }

    /// The point of the call is that a consumer stops keeping its own copy of the list, so
    /// the reported set has to track the parsed set rather than being a second hand-written one.
    #[test]
    fn supported_formats_stays_in_step_with_the_library() {
        let json = supported_formats().unwrap();
        for format in undoc::FormatType::ALL {
            assert!(
                json.contains(format.extension()),
                "{} is parsed but not reported as supported",
                format.extension()
            );
        }
    }
}
