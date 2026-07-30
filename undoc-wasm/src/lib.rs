mod document;

pub use document::OfficeDocument;

use wasm_bindgen::prelude::*;

#[wasm_bindgen]
pub fn parse(data: &[u8]) -> Result<OfficeDocument, JsValue> {
    undoc::parse_bytes(data)
        .map(|inner| OfficeDocument { inner })
        .map_err(parse_error)
}

/// Build the JS error for a failed parse, carrying both its message and its reason.
///
/// A bare string would force callers to match on message text. Throwing a real `Error`
/// with a numeric `kind` gives JavaScript the same contract the C ABI offers through
/// `undoc_last_error_kind`: branch on the reason, and treat an unrecognised number as a
/// generic failure, since new reasons take new numbers and existing ones never change.
fn parse_error(e: undoc::Error) -> JsValue {
    let error = js_sys::Error::new(&e.to_string());
    let assigned = js_sys::Reflect::set(
        &error,
        &JsValue::from_str("kind"),
        &JsValue::from_f64(e.kind() as i32 as f64),
    );
    debug_assert!(assigned.is_ok(), "kind must be assignable on a fresh Error");
    error.into()
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
}
