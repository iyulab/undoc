//! XLSX shared strings parsing.

use crate::error::{Error, Result};

/// Shared strings table.
#[derive(Debug, Clone, Default)]
pub struct SharedStrings {
    /// All strings in order
    strings: Vec<String>,
}

impl SharedStrings {
    /// Parse shared strings from XML content.
    pub fn parse(xml: &str) -> Result<Self> {
        let mut strings = Vec::new();
        let mut reader = quick_xml::Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut in_si = false;
        let mut in_t = false;
        let mut current_text = String::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(e)) => match e.name().as_ref() {
                    b"si" => {
                        in_si = true;
                        current_text.clear();
                    }
                    b"t" if in_si => {
                        in_t = true;
                    }
                    _ => {}
                },
                Ok(quick_xml::events::Event::Text(e)) => {
                    if in_t {
                        let text = e.unescape().unwrap_or_default();
                        current_text.push_str(&text);
                    }
                }
                Ok(quick_xml::events::Event::End(e)) => match e.name().as_ref() {
                    b"si" => {
                        strings.push(current_text.clone());
                        in_si = false;
                    }
                    b"t" => {
                        in_t = false;
                    }
                    _ => {}
                },
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => return Err(Error::XmlParse(e.to_string())),
                _ => {}
            }
            buf.clear();
        }

        Ok(Self { strings })
    }

    /// Get a string by index.
    pub fn get(&self, index: usize) -> Option<&str> {
        self.strings.get(index).map(|s| s.as_str())
    }

    /// Get the count of shared strings.
    #[allow(dead_code)]
    pub fn len(&self) -> usize {
        self.strings.len()
    }

    /// Check if empty.
    #[allow(dead_code)]
    pub fn is_empty(&self) -> bool {
        self.strings.is_empty()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_shared_strings() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="5" uniqueCount="3">
    <si><t>Hello</t></si>
    <si><t>World</t></si>
    <si><t>Test</t></si>
</sst>"#;

        let ss = SharedStrings::parse(xml).unwrap();
        assert_eq!(ss.len(), 3);
        assert_eq!(ss.get(0), Some("Hello"));
        assert_eq!(ss.get(1), Some("World"));
        assert_eq!(ss.get(2), Some("Test"));
        assert_eq!(ss.get(3), None);
    }

    #[test]
    fn test_rich_text() {
        // Rich text with runs - note: t element must include any trailing spaces
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <si>
        <r><t>Hello</t></r>
        <r><t>World</t></r>
    </si>
</sst>"#;

        let ss = SharedStrings::parse(xml).unwrap();
        assert_eq!(ss.len(), 1);
        // Rich text runs are concatenated as-is
        assert_eq!(ss.get(0), Some("HelloWorld"));
    }
}
