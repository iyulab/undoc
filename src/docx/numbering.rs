//! DOCX numbering (list) parsing.

use std::collections::HashMap;
use crate::error::{Error, Result};
use crate::model::ListType;

/// Abstract numbering definition.
#[derive(Debug, Clone)]
pub struct AbstractNum {
    /// Abstract numbering ID
    pub id: String,
    /// Levels (0-8)
    pub levels: Vec<NumLevel>,
}

/// A numbering level definition.
#[derive(Debug, Clone)]
pub struct NumLevel {
    /// Level index (0-8)
    pub level: u8,
    /// Start value
    pub start: u32,
    /// Number format (decimal, bullet, lowerLetter, etc.)
    pub num_fmt: String,
    /// Level text (e.g., "%1.", "%1.%2.")
    pub level_text: String,
}

impl NumLevel {
    /// Get the list type for this level.
    pub fn list_type(&self) -> ListType {
        match self.num_fmt.as_str() {
            "bullet" => ListType::Bullet,
            "decimal" | "lowerLetter" | "upperLetter" | "lowerRoman" | "upperRoman" => {
                ListType::Numbered
            }
            _ => ListType::Bullet, // Default to bullet
        }
    }
}

/// Concrete numbering instance.
#[derive(Debug, Clone)]
pub struct NumInstance {
    /// Numbering ID
    #[allow(dead_code)]
    pub num_id: String,
    /// Abstract numbering ID
    pub abstract_num_id: String,
}

/// Collection of numbering definitions.
#[derive(Debug, Clone, Default)]
pub struct NumberingMap {
    /// Abstract numbering definitions
    pub abstract_nums: HashMap<String, AbstractNum>,
    /// Numbering instances
    pub instances: HashMap<String, NumInstance>,
    /// Current count for each numId+level combination
    counters: HashMap<(String, u8), u32>,
}

impl NumberingMap {
    /// Parse numbering from XML content.
    pub fn parse(xml: &str) -> Result<Self> {
        let mut map = NumberingMap::default();
        let mut reader = quick_xml::Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut current_abstract: Option<AbstractNum> = None;
        let mut current_level: Option<NumLevel> = None;
        let mut in_abstract_num = false;
        let mut in_lvl = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(e)) => {
                    match e.name().as_ref() {
                        b"w:abstractNum" => {
                            let mut abstract_num = AbstractNum {
                                id: String::new(),
                                levels: Vec::new(),
                            };
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"w:abstractNumId" {
                                    abstract_num.id =
                                        String::from_utf8_lossy(&attr.value).to_string();
                                }
                            }
                            current_abstract = Some(abstract_num);
                            in_abstract_num = true;
                        }
                        b"w:lvl" if in_abstract_num => {
                            let mut level = NumLevel {
                                level: 0,
                                start: 1,
                                num_fmt: "bullet".to_string(),
                                level_text: String::new(),
                            };
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"w:ilvl" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    level.level = val.parse().unwrap_or(0);
                                }
                            }
                            current_level = Some(level);
                            in_lvl = true;
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Empty(e)) => {
                    match e.name().as_ref() {
                        b"w:start" if in_lvl => {
                            if let Some(ref mut level) = current_level {
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"w:val" {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        level.start = val.parse().unwrap_or(1);
                                    }
                                }
                            }
                        }
                        b"w:numFmt" if in_lvl => {
                            if let Some(ref mut level) = current_level {
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"w:val" {
                                        level.num_fmt =
                                            String::from_utf8_lossy(&attr.value).to_string();
                                    }
                                }
                            }
                        }
                        b"w:lvlText" if in_lvl => {
                            if let Some(ref mut level) = current_level {
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"w:val" {
                                        level.level_text =
                                            String::from_utf8_lossy(&attr.value).to_string();
                                    }
                                }
                            }
                        }
                        b"w:abstractNumId" => {
                            // This is in w:num element
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::End(e)) => {
                    match e.name().as_ref() {
                        b"w:abstractNum" => {
                            if let Some(abstract_num) = current_abstract.take() {
                                map.abstract_nums
                                    .insert(abstract_num.id.clone(), abstract_num);
                            }
                            in_abstract_num = false;
                        }
                        b"w:lvl" => {
                            if let Some(level) = current_level.take() {
                                if let Some(ref mut abstract_num) = current_abstract {
                                    abstract_num.levels.push(level);
                                }
                            }
                            in_lvl = false;
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => return Err(Error::XmlParse(e.to_string())),
                _ => {}
            }
            buf.clear();
        }

        // Parse w:num elements for instances
        map.parse_num_instances(xml)?;

        Ok(map)
    }

    /// Parse w:num elements.
    fn parse_num_instances(&mut self, xml: &str) -> Result<()> {
        let mut reader = quick_xml::Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut current_num_id: Option<String> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(e)) => {
                    if e.name().as_ref() == b"w:num" {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"w:numId" {
                                current_num_id =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                }
                Ok(quick_xml::events::Event::Empty(e)) => {
                    if e.name().as_ref() == b"w:abstractNumId" {
                        if let Some(ref num_id) = current_num_id {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"w:val" {
                                    let abstract_id =
                                        String::from_utf8_lossy(&attr.value).to_string();
                                    self.instances.insert(
                                        num_id.clone(),
                                        NumInstance {
                                            num_id: num_id.clone(),
                                            abstract_num_id: abstract_id,
                                        },
                                    );
                                }
                            }
                        }
                    }
                }
                Ok(quick_xml::events::Event::End(e)) => {
                    if e.name().as_ref() == b"w:num" {
                        current_num_id = None;
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(())
    }

    /// Get list info for a paragraph.
    pub fn get_list_info(&mut self, num_id: &str, level: u8) -> Option<(ListType, u32)> {
        let instance = self.instances.get(num_id)?;
        let abstract_num = self.abstract_nums.get(&instance.abstract_num_id)?;
        let num_level = abstract_num.levels.iter().find(|l| l.level == level)?;

        let list_type = num_level.list_type();

        // Get or initialize counter
        let key = (num_id.to_string(), level);
        let counter = self.counters.entry(key).or_insert(num_level.start);
        let number = *counter;

        // Increment counter for next use
        *self.counters.get_mut(&(num_id.to_string(), level)).unwrap() += 1;

        Some((list_type, number))
    }

    /// Reset counters (e.g., at start of document).
    #[allow(dead_code)]
    pub fn reset_counters(&mut self) {
        self.counters.clear();
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_numbering() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0">
            <w:start w:val="1"/>
            <w:numFmt w:val="decimal"/>
            <w:lvlText w:val="%1."/>
        </w:lvl>
        <w:lvl w:ilvl="1">
            <w:start w:val="1"/>
            <w:numFmt w:val="bullet"/>
            <w:lvlText w:val="•"/>
        </w:lvl>
    </w:abstractNum>
    <w:num w:numId="1">
        <w:abstractNumId w:val="0"/>
    </w:num>
</w:numbering>"#;

        let mut map = NumberingMap::parse(xml).unwrap();

        assert!(map.abstract_nums.contains_key("0"));
        assert!(map.instances.contains_key("1"));

        let abstract_num = map.abstract_nums.get("0").unwrap();
        assert_eq!(abstract_num.levels.len(), 2);
        assert_eq!(abstract_num.levels[0].num_fmt, "decimal");

        // Test list info
        let (list_type, num) = map.get_list_info("1", 0).unwrap();
        assert_eq!(list_type, ListType::Numbered);
        assert_eq!(num, 1);

        // Second call increments
        let (_, num) = map.get_list_info("1", 0).unwrap();
        assert_eq!(num, 2);
    }
}
