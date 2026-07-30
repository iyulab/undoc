use crate::{json_error, undoc_error};
use undoc::render::{JsonFormat, RenderOptions};
use wasm_bindgen::prelude::*;

#[wasm_bindgen]
pub struct OfficeDocument {
    #[allow(dead_code)]
    pub(crate) inner: undoc::Document,
}

#[wasm_bindgen]
impl OfficeDocument {
    #[wasm_bindgen(js_name = fromBytes)]
    pub fn from_bytes(data: &[u8]) -> Result<OfficeDocument, JsValue> {
        undoc::parse_bytes(data)
            .map(|inner| OfficeDocument { inner })
            .map_err(undoc_error)
    }

    #[wasm_bindgen(js_name = toMarkdown)]
    pub fn to_markdown(&self) -> Result<String, JsValue> {
        let opts = RenderOptions::default();
        undoc::render::to_markdown(&self.inner, &opts).map_err(undoc_error)
    }

    #[wasm_bindgen(js_name = toText)]
    pub fn to_text(&self) -> Result<String, JsValue> {
        let opts = RenderOptions::default();
        undoc::render::to_text(&self.inner, &opts).map_err(undoc_error)
    }

    #[wasm_bindgen(js_name = toJson)]
    pub fn to_json(&self) -> Result<String, JsValue> {
        undoc::render::to_json(&self.inner, JsonFormat::Compact).map_err(undoc_error)
    }

    pub fn format(&self) -> String {
        self.inner.format.extension().to_string()
    }

    pub fn metadata(&self) -> Result<String, JsValue> {
        serde_json::to_string(&self.inner.metadata).map_err(json_error)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use wasm_bindgen_test::*;

    wasm_bindgen_test_configure!(run_in_node_experimental);

    #[wasm_bindgen_test]
    fn test_from_bytes_invalid_returns_err() {
        let result = OfficeDocument::from_bytes(b"not an office file");
        assert!(result.is_err());
    }

    /// `fromBytes` is the entry point the npm package documents, so it owes callers the
    /// same reason channel `parse` gives them — a bare string would force them back to
    /// matching on message text.
    #[wasm_bindgen_test]
    fn test_from_bytes_error_carries_its_kind() {
        let error = OfficeDocument::from_bytes(b"not an office file")
            .err()
            .expect("a non-Office file must not parse");
        let kind = js_sys::Reflect::get(&error, &JsValue::from_str("kind"))
            .expect("a thrown error must expose kind");

        assert_eq!(
            kind.as_f64(),
            Some(undoc::ErrorKind::UnknownFormat as i32 as f64)
        );
    }
}
