mod document;

pub use document::OfficeDocument;

use wasm_bindgen::prelude::*;

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
}
