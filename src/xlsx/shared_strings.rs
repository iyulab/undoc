//! XLSX shared strings parsing.

use std::borrow::Cow;

use crate::decode::{decode_text_lossy, normalize_line_endings, resolve_general_ref};
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
        let mut reader = crate::decode::reader_for(xml);
        // IMPORTANT: Don't trim text - preserve whitespace from xml:space="preserve" elements
        // Excel cells may contain significant leading/trailing spaces
        reader.config_mut().trim_text(false);

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
                Ok(quick_xml::events::Event::Text(e)) if in_t => {
                    let text = decode_text_lossy(&e);
                    current_text.push_str(&text);
                }
                // quick-xml 0.40+ emits entity refs (&amp;, &#13;, …) as separate
                // events; without this arm every entity in a shared string vanishes.
                Ok(quick_xml::events::Event::GeneralRef(e)) if in_t => {
                    current_text.push_str(&resolve_general_ref(&e));
                }
                Ok(quick_xml::events::Event::End(e)) => match e.name().as_ref() {
                    b"si" => {
                        // Collapse CR that re-entered via &#13;/&#xD; refs (Excel
                        // in-cell breaks); a CRLF pair arrives as two refs, so the
                        // whole accumulated string is normalized once here.
                        strings.push(
                            normalize_line_endings(Cow::Borrowed(&current_text)).into_owned(),
                        );
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

    #[test]
    fn test_malformed_shared_string_preserves_raw_entity_text() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <si><t>Hello &bogus; World</t></si>
</sst>"#;

        let ss = SharedStrings::parse(xml).unwrap();
        assert_eq!(ss.get(0), Some("Hello &bogus; World"));
    }

    #[test]
    fn test_shared_strings_mixed_entities_preserve_legitimate_and_malformed() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>A &amp; B &bogus; C</t></si>
</sst>"#;

        let table = SharedStrings::parse(xml).expect("parse succeeds");
        let s = table.get(0).expect("index 0 exists");
        assert_eq!(s, "A & B &bogus; C");
    }

    #[test]
    fn test_shared_strings_cr_character_refs_collapse_to_lf() {
        // Excel stores in-cell line breaks as &#13;&#10; (CRLF) or bare &#13;.
        // Under quick-xml 0.40+ these arrive as separate GeneralRef events, so
        // the CRLF pair must be collapsed at the flush point, not per-event.
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <si><t>line1&#13;&#10;line2&#13;line3&#xD;line4</t></si>
</sst>"#;

        let ss = SharedStrings::parse(xml).expect("parse succeeds");
        assert_eq!(ss.get(0), Some("line1\nline2\nline3\nline4"));
    }

    #[test]
    fn test_shared_strings_numeric_and_predefined_refs_round_trip() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <si><t>&#48;&#x30; &lt;a&gt; &amp; &quot;q&quot; &apos;p&apos;</t></si>
</sst>"#;

        let ss = SharedStrings::parse(xml).expect("parse succeeds");
        assert_eq!(ss.get(0), Some("00 <a> & \"q\" 'p'"));
    }
}
