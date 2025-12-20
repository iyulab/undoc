//! Paragraph and text run models.

use serde::{Deserialize, Serialize};

/// Text alignment within a paragraph.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum TextAlignment {
    #[default]
    Left,
    Center,
    Right,
    Justify,
}

/// Heading level (h1-h6 or none).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
pub enum HeadingLevel {
    #[default]
    None,
    H1,
    H2,
    H3,
    H4,
    H5,
    H6,
}

impl HeadingLevel {
    /// Create a heading level from a number (1-6).
    pub fn from_number(n: u8) -> Self {
        match n {
            1 => HeadingLevel::H1,
            2 => HeadingLevel::H2,
            3 => HeadingLevel::H3,
            4 => HeadingLevel::H4,
            5 => HeadingLevel::H5,
            6 => HeadingLevel::H6,
            _ => HeadingLevel::None,
        }
    }

    /// Get the numeric level (0 for none, 1-6 for headings).
    pub fn level(&self) -> u8 {
        match self {
            HeadingLevel::None => 0,
            HeadingLevel::H1 => 1,
            HeadingLevel::H2 => 2,
            HeadingLevel::H3 => 3,
            HeadingLevel::H4 => 4,
            HeadingLevel::H5 => 5,
            HeadingLevel::H6 => 6,
        }
    }

    /// Check if this is a heading (not None).
    pub fn is_heading(&self) -> bool {
        !matches!(self, HeadingLevel::None)
    }
}

/// List type for paragraphs.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum ListType {
    #[default]
    None,
    /// Unordered (bulleted) list
    Bullet,
    /// Ordered (numbered) list
    Numbered,
}

/// List information for a paragraph.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ListInfo {
    /// Type of list
    pub list_type: ListType,
    /// Nesting level (0 = top level)
    pub level: u8,
    /// Item number (for numbered lists)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub number: Option<u32>,
}

/// Text style properties.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
pub struct TextStyle {
    /// Bold text
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub bold: bool,

    /// Italic text
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub italic: bool,

    /// Underlined text
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub underline: bool,

    /// Strikethrough text
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub strikethrough: bool,

    /// Superscript
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub superscript: bool,

    /// Subscript
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub subscript: bool,

    /// Code/monospace font
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub code: bool,

    /// Font name
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font: Option<String>,

    /// Font size in half-points (e.g., 24 = 12pt)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub size: Option<u32>,

    /// Text color (hex, e.g., "FF0000")
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,

    /// Background/highlight color
    #[serde(skip_serializing_if = "Option::is_none")]
    pub highlight: Option<String>,
}

impl TextStyle {
    /// Create a new default style.
    pub fn new() -> Self {
        Self::default()
    }

    /// Create a bold style.
    pub fn bold() -> Self {
        Self {
            bold: true,
            ..Default::default()
        }
    }

    /// Create an italic style.
    pub fn italic() -> Self {
        Self {
            italic: true,
            ..Default::default()
        }
    }

    /// Check if style has any formatting.
    pub fn has_formatting(&self) -> bool {
        self.bold
            || self.italic
            || self.underline
            || self.strikethrough
            || self.superscript
            || self.subscript
            || self.code
    }
}

/// A run of text with consistent styling.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TextRun {
    /// The text content
    pub text: String,

    /// Text styling
    #[serde(default, skip_serializing_if = "is_default_style")]
    pub style: TextStyle,

    /// Hyperlink URL (if this run is a link)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hyperlink: Option<String>,
}

fn is_default_style(style: &TextStyle) -> bool {
    *style == TextStyle::default()
}

impl TextRun {
    /// Create a plain text run with no styling.
    pub fn plain(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            style: TextStyle::default(),
            hyperlink: None,
        }
    }

    /// Create a styled text run.
    pub fn styled(text: impl Into<String>, style: TextStyle) -> Self {
        Self {
            text: text.into(),
            style,
            hyperlink: None,
        }
    }

    /// Create a hyperlink text run.
    pub fn link(text: impl Into<String>, url: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            style: TextStyle::default(),
            hyperlink: Some(url.into()),
        }
    }

    /// Check if this run is a hyperlink.
    pub fn is_link(&self) -> bool {
        self.hyperlink.is_some()
    }

    /// Check if this run is empty.
    pub fn is_empty(&self) -> bool {
        self.text.is_empty()
    }
}

/// An inline image within text.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct InlineImage {
    /// Resource ID for the image
    pub resource_id: String,

    /// Alt text
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alt_text: Option<String>,

    /// Width in EMUs
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<u32>,

    /// Height in EMUs
    #[serde(skip_serializing_if = "Option::is_none")]
    pub height: Option<u32>,
}

/// An element within a paragraph (text run or inline image).
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum ParagraphElement {
    Text(TextRun),
    Image(InlineImage),
}

/// A paragraph of text.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Paragraph {
    /// Text runs in this paragraph
    #[serde(default)]
    pub runs: Vec<TextRun>,

    /// Heading level
    #[serde(default, skip_serializing_if = "HeadingLevel::is_none")]
    pub heading: HeadingLevel,

    /// Text alignment
    #[serde(default, skip_serializing_if = "is_default_alignment")]
    pub alignment: TextAlignment,

    /// List information
    #[serde(skip_serializing_if = "Option::is_none")]
    pub list_info: Option<ListInfo>,

    /// Style ID reference
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style_id: Option<String>,

    /// Indentation level
    #[serde(default, skip_serializing_if = "is_zero")]
    pub indent_level: u8,
}

fn is_default_alignment(a: &TextAlignment) -> bool {
    *a == TextAlignment::Left
}

fn is_zero(n: &u8) -> bool {
    *n == 0
}

impl HeadingLevel {
    fn is_none(&self) -> bool {
        matches!(self, HeadingLevel::None)
    }
}

impl Paragraph {
    /// Create a new empty paragraph.
    pub fn new() -> Self {
        Self::default()
    }

    /// Create a paragraph with the given text.
    pub fn with_text(text: impl Into<String>) -> Self {
        Self {
            runs: vec![TextRun::plain(text)],
            ..Default::default()
        }
    }

    /// Create a heading paragraph.
    pub fn heading(level: HeadingLevel, text: impl Into<String>) -> Self {
        Self {
            runs: vec![TextRun::plain(text)],
            heading: level,
            ..Default::default()
        }
    }

    /// Add a text run to this paragraph.
    pub fn add_run(&mut self, run: TextRun) {
        self.runs.push(run);
    }

    /// Get the plain text content.
    pub fn plain_text(&self) -> String {
        self.runs.iter().map(|r| r.text.as_str()).collect()
    }

    /// Check if this paragraph is empty.
    pub fn is_empty(&self) -> bool {
        self.runs.is_empty() || self.runs.iter().all(|r| r.is_empty())
    }

    /// Check if this paragraph is a heading.
    pub fn is_heading(&self) -> bool {
        self.heading.is_heading()
    }

    /// Check if this paragraph is a list item.
    pub fn is_list_item(&self) -> bool {
        self.list_info.is_some()
    }

    /// Merge consecutive runs with the same style.
    ///
    /// This is useful for documents where each character or word is in a separate run
    /// with the same styling (common in Word documents with letter spacing).
    ///
    /// Example: `**시** **험**` becomes `**시험**` after merging.
    pub fn merge_adjacent_runs(&mut self) {
        if self.runs.len() <= 1 {
            return;
        }

        let mut merged: Vec<TextRun> = Vec::with_capacity(self.runs.len());

        for run in self.runs.drain(..) {
            // Check if we can merge with the last run
            let should_merge = merged.last().map_or(false, |last: &TextRun| {
                // Same style and same hyperlink (both None or both Some with same URL)
                last.style == run.style && last.hyperlink == run.hyperlink
            });

            if should_merge {
                // Merge text with the last run
                if let Some(last) = merged.last_mut() {
                    last.text.push_str(&run.text);
                }
            } else {
                // Start a new run
                merged.push(run);
            }
        }

        self.runs = merged;
    }

    /// Get a version of this paragraph with merged adjacent runs.
    ///
    /// This is a non-mutating version of `merge_adjacent_runs`.
    pub fn with_merged_runs(&self) -> Self {
        let mut para = self.clone();
        para.merge_adjacent_runs();
        para
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_heading_level() {
        assert_eq!(HeadingLevel::from_number(1), HeadingLevel::H1);
        assert_eq!(HeadingLevel::from_number(6), HeadingLevel::H6);
        assert_eq!(HeadingLevel::from_number(7), HeadingLevel::None);
        assert_eq!(HeadingLevel::from_number(0), HeadingLevel::None);

        assert_eq!(HeadingLevel::H3.level(), 3);
        assert!(HeadingLevel::H1.is_heading());
        assert!(!HeadingLevel::None.is_heading());
    }

    #[test]
    fn test_text_run() {
        let plain = TextRun::plain("Hello");
        assert_eq!(plain.text, "Hello");
        assert!(!plain.is_link());

        let link = TextRun::link("Click here", "https://example.com");
        assert!(link.is_link());
        assert_eq!(link.hyperlink, Some("https://example.com".to_string()));
    }

    #[test]
    fn test_text_style() {
        let style = TextStyle::bold();
        assert!(style.bold);
        assert!(style.has_formatting());

        let plain = TextStyle::default();
        assert!(!plain.has_formatting());
    }

    #[test]
    fn test_paragraph() {
        let para = Paragraph::with_text("Hello, World!");
        assert_eq!(para.plain_text(), "Hello, World!");
        assert!(!para.is_heading());
        assert!(!para.is_empty());

        let heading = Paragraph::heading(HeadingLevel::H1, "Title");
        assert!(heading.is_heading());
        assert_eq!(heading.heading.level(), 1);
    }

    #[test]
    fn test_paragraph_serialization() {
        let para = Paragraph::with_text("Test");
        let json = serde_json::to_string(&para).unwrap();
        // Default values should not be serialized
        assert!(!json.contains("heading"));
        assert!(!json.contains("alignment"));
    }
}
