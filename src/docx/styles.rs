//! DOCX styles parsing.

use crate::error::{Error, Result};
use crate::model::{HeadingLevel, TextStyle};
use std::collections::HashMap;

/// Style type (paragraph, character, table, etc.)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StyleType {
    Paragraph,
    Character,
    Table,
    Numbering,
}

/// A parsed style definition.
#[derive(Debug, Clone, Default)]
pub struct Style {
    /// Style ID (e.g., "Heading1")
    pub id: String,
    /// Style name (e.g., "Heading 1")
    pub name: String,
    /// Style type
    pub style_type: Option<StyleType>,
    /// Based on another style
    pub based_on: Option<String>,
    /// Paragraph properties
    pub paragraph_props: ParagraphProps,
    /// Run (text) properties
    pub run_props: RunProps,
    /// Outline level (for headings)
    pub outline_level: Option<u8>,
}

/// Paragraph-level properties from a style.
#[derive(Debug, Clone, Default)]
#[allow(dead_code)]
pub struct ParagraphProps {
    pub justification: Option<String>,
    pub spacing_before: Option<i32>,
    pub spacing_after: Option<i32>,
    pub line_spacing: Option<i32>,
    pub indent_left: Option<i32>,
    pub indent_right: Option<i32>,
    pub indent_hanging: Option<i32>,
    pub indent_first_line: Option<i32>,
}

/// Run-level (character) properties from a style.
#[derive(Debug, Clone, Default)]
pub struct RunProps {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    pub strike: Option<bool>,
    pub font_name: Option<String>,
    pub font_size: Option<u32>,
    pub color: Option<String>,
    pub highlight: Option<String>,
}

impl RunProps {
    /// Convert to TextStyle.
    #[allow(dead_code)]
    pub fn to_text_style(&self) -> TextStyle {
        TextStyle {
            bold: self.bold.unwrap_or(false),
            italic: self.italic.unwrap_or(false),
            underline: self.underline.unwrap_or(false),
            strikethrough: self.strike.unwrap_or(false),
            font: self.font_name.clone(),
            size: self.font_size,
            color: self.color.clone(),
            highlight: self.highlight.clone(),
            ..Default::default()
        }
    }

    /// Merge with another RunProps (other takes precedence).
    pub fn merge(&mut self, other: &RunProps) {
        if other.bold.is_some() {
            self.bold = other.bold;
        }
        if other.italic.is_some() {
            self.italic = other.italic;
        }
        if other.underline.is_some() {
            self.underline = other.underline;
        }
        if other.strike.is_some() {
            self.strike = other.strike;
        }
        if other.font_name.is_some() {
            self.font_name = other.font_name.clone();
        }
        if other.font_size.is_some() {
            self.font_size = other.font_size;
        }
        if other.color.is_some() {
            self.color = other.color.clone();
        }
        if other.highlight.is_some() {
            self.highlight = other.highlight.clone();
        }
    }
}

/// Collection of styles from styles.xml.
#[derive(Debug, Clone, Default)]
pub struct StyleMap {
    /// Styles by ID
    pub styles: HashMap<String, Style>,
    /// Default paragraph style
    pub default_paragraph: Option<String>,
    /// Default character style
    pub default_character: Option<String>,
}

impl StyleMap {
    /// Parse styles from XML content.
    pub fn parse(xml: &str) -> Result<Self> {
        // Handle empty content
        if xml.trim().is_empty() {
            return Ok(Self::default());
        }

        let mut map = StyleMap::default();
        let mut reader = quick_xml::Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut current_style: Option<Style> = None;
        let mut in_style = false;
        let mut in_ppr = false;
        let mut in_rpr = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(e)) => {
                    let name = e.name();
                    match name.as_ref() {
                        b"w:style" => {
                            let mut style = Style::default();
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"w:styleId" => {
                                        style.id = String::from_utf8_lossy(&attr.value).to_string();
                                    }
                                    b"w:type" => {
                                        let t = String::from_utf8_lossy(&attr.value);
                                        style.style_type = match t.as_ref() {
                                            "paragraph" => Some(StyleType::Paragraph),
                                            "character" => Some(StyleType::Character),
                                            "table" => Some(StyleType::Table),
                                            "numbering" => Some(StyleType::Numbering),
                                            _ => None,
                                        };
                                    }
                                    b"w:default" => {
                                        let is_default =
                                            String::from_utf8_lossy(&attr.value) == "1";
                                        if is_default {
                                            if let Some(ref style_type) = style.style_type {
                                                match style_type {
                                                    StyleType::Paragraph => {
                                                        map.default_paragraph =
                                                            Some(style.id.clone());
                                                    }
                                                    StyleType::Character => {
                                                        map.default_character =
                                                            Some(style.id.clone());
                                                    }
                                                    _ => {}
                                                }
                                            }
                                        }
                                    }
                                    _ => {}
                                }
                            }
                            current_style = Some(style);
                            in_style = true;
                        }
                        b"w:pPr" if in_style => {
                            in_ppr = true;
                        }
                        b"w:rPr" if in_style => {
                            in_rpr = true;
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Empty(e)) => {
                    let name = e.name();
                    if let Some(ref mut style) = current_style {
                        match name.as_ref() {
                            b"w:name" => {
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"w:val" {
                                        style.name =
                                            String::from_utf8_lossy(&attr.value).to_string();
                                    }
                                }
                            }
                            b"w:basedOn" => {
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"w:val" {
                                        style.based_on =
                                            Some(String::from_utf8_lossy(&attr.value).to_string());
                                    }
                                }
                            }
                            b"w:outlineLvl" if in_ppr => {
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"w:val" {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        style.outline_level = val.parse().ok();
                                    }
                                }
                            }
                            b"w:jc" if in_ppr => {
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"w:val" {
                                        style.paragraph_props.justification =
                                            Some(String::from_utf8_lossy(&attr.value).to_string());
                                    }
                                }
                            }
                            b"w:b" if in_rpr => {
                                let val = get_bool_attr(&e, b"w:val");
                                style.run_props.bold = Some(val.unwrap_or(true));
                            }
                            b"w:i" if in_rpr => {
                                let val = get_bool_attr(&e, b"w:val");
                                style.run_props.italic = Some(val.unwrap_or(true));
                            }
                            b"w:u" if in_rpr => {
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"w:val" {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        style.run_props.underline = Some(val != "none");
                                    }
                                }
                            }
                            b"w:strike" if in_rpr => {
                                let val = get_bool_attr(&e, b"w:val");
                                style.run_props.strike = Some(val.unwrap_or(true));
                            }
                            b"w:sz" if in_rpr => {
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"w:val" {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        style.run_props.font_size = val.parse().ok();
                                    }
                                }
                            }
                            b"w:color" if in_rpr => {
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"w:val" {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        if val != "auto" {
                                            style.run_props.color = Some(val.to_string());
                                        }
                                    }
                                }
                            }
                            b"w:rFonts" if in_rpr => {
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"w:ascii" {
                                        style.run_props.font_name =
                                            Some(String::from_utf8_lossy(&attr.value).to_string());
                                        break;
                                    }
                                }
                            }
                            _ => {}
                        }
                    }
                }
                Ok(quick_xml::events::Event::End(e)) => match e.name().as_ref() {
                    b"w:style" => {
                        if let Some(style) = current_style.take() {
                            map.styles.insert(style.id.clone(), style);
                        }
                        in_style = false;
                        in_ppr = false;
                        in_rpr = false;
                    }
                    b"w:pPr" => {
                        in_ppr = false;
                    }
                    b"w:rPr" => {
                        in_rpr = false;
                    }
                    _ => {}
                },
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => return Err(Error::XmlParse(e.to_string())),
                _ => {}
            }
            buf.clear();
        }

        Ok(map)
    }

    /// Get a style by ID, resolving inheritance.
    pub fn get_resolved(&self, id: &str) -> Option<Style> {
        let mut style = self.styles.get(id)?.clone();

        // Resolve inheritance chain (max 10 levels to prevent infinite loops)
        let mut depth = 0;
        let mut current_based_on = style.based_on.clone();
        while let Some(ref base_id) = current_based_on {
            if depth > 10 {
                break;
            }
            if let Some(base) = self.styles.get(base_id) {
                // Merge base properties (base first, then override)
                let mut merged_run = base.run_props.clone();
                merged_run.merge(&style.run_props);
                style.run_props = merged_run;

                if style.outline_level.is_none() {
                    style.outline_level = base.outline_level;
                }

                current_based_on = base.based_on.clone();
            } else {
                break;
            }
            depth += 1;
        }

        Some(style)
    }

    /// Get the heading level for a style ID.
    pub fn get_heading_level(&self, style_id: &str) -> HeadingLevel {
        if let Some(style) = self.get_resolved(style_id) {
            if let Some(level) = style.outline_level {
                return HeadingLevel::from_number(level + 1);
            }
            // Also check for Title/Subtitle styles
            let name_lower = style.name.to_lowercase();
            if name_lower == "title" {
                return HeadingLevel::H1;
            }
            if name_lower == "subtitle" {
                return HeadingLevel::H2;
            }
        }
        HeadingLevel::None
    }
}

/// Helper to get a boolean attribute value.
fn get_bool_attr(e: &quick_xml::events::BytesStart, key: &[u8]) -> Option<bool> {
    for attr in e.attributes().flatten() {
        if attr.key.as_ref() == key {
            let val = String::from_utf8_lossy(&attr.value);
            return Some(val != "0" && val != "false");
        }
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_styles() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:style w:type="paragraph" w:styleId="Heading1">
        <w:name w:val="Heading 1"/>
        <w:basedOn w:val="Normal"/>
        <w:pPr>
            <w:outlineLvl w:val="0"/>
        </w:pPr>
        <w:rPr>
            <w:b/>
            <w:sz w:val="32"/>
        </w:rPr>
    </w:style>
</w:styles>"#;

        let map = StyleMap::parse(xml).unwrap();
        assert!(map.styles.contains_key("Heading1"));

        let style = map.styles.get("Heading1").unwrap();
        assert_eq!(style.name, "Heading 1");
        assert_eq!(style.outline_level, Some(0));
        assert_eq!(style.run_props.bold, Some(true));
        assert_eq!(style.run_props.font_size, Some(32));
    }

    #[test]
    fn test_heading_level() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:style w:type="paragraph" w:styleId="Title">
        <w:name w:val="Title"/>
    </w:style>
    <w:style w:type="paragraph" w:styleId="Heading2">
        <w:name w:val="Heading 2"/>
        <w:pPr>
            <w:outlineLvl w:val="1"/>
        </w:pPr>
    </w:style>
</w:styles>"#;

        let map = StyleMap::parse(xml).unwrap();
        assert_eq!(map.get_heading_level("Title"), HeadingLevel::H1);
        assert_eq!(map.get_heading_level("Heading2"), HeadingLevel::H2);
        assert_eq!(map.get_heading_level("Unknown"), HeadingLevel::None);
    }
}
