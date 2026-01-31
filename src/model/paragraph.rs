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

/// Revision type for tracked changes support.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum RevisionType {
    /// Normal text (not a tracked change)
    #[default]
    None,
    /// Inserted text (addition)
    Inserted,
    /// Deleted text (deletion)
    Deleted,
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

    /// Whether this run ends with a line break (<w:br/>)
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub line_break: bool,

    /// Whether this run ends with a page break (<w:br w:type="page"/>)
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub page_break: bool,

    /// Revision type for tracked changes (inserted/deleted)
    #[serde(default, skip_serializing_if = "is_default_revision")]
    pub revision: RevisionType,
}

fn is_default_style(style: &TextStyle) -> bool {
    *style == TextStyle::default()
}

fn is_default_revision(revision: &RevisionType) -> bool {
    *revision == RevisionType::None
}

impl TextRun {
    /// Create a plain text run with no styling.
    pub fn plain(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            style: TextStyle::default(),
            hyperlink: None,
            line_break: false,
            page_break: false,
            revision: RevisionType::None,
        }
    }

    /// Create a styled text run.
    pub fn styled(text: impl Into<String>, style: TextStyle) -> Self {
        Self {
            text: text.into(),
            style,
            hyperlink: None,
            line_break: false,
            page_break: false,
            revision: RevisionType::None,
        }
    }

    /// Create a hyperlink text run.
    pub fn link(text: impl Into<String>, url: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            style: TextStyle::default(),
            hyperlink: Some(url.into()),
            line_break: false,
            page_break: false,
            revision: RevisionType::None,
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

    /// Inline images in this paragraph
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub images: Vec<InlineImage>,

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

    /// Style name (human-readable, from styles.xml)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style_name: Option<String>,

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
    ///
    /// Smart spacing is applied when merging runs - spaces are added between
    /// CJK text and ASCII alphanumeric characters, and between ASCII words.
    pub fn merge_adjacent_runs(&mut self) {
        if self.runs.len() <= 1 {
            return;
        }

        let mut merged: Vec<TextRun> = Vec::with_capacity(self.runs.len());

        for run in self.runs.drain(..) {
            // Check if we can merge with the last run
            let should_merge = merged.last().is_some_and(|last: &TextRun| {
                // Same style and same hyperlink (both None or both Some with same URL)
                // Don't merge if the previous run has a line break (preserve the break)
                last.style == run.style && last.hyperlink == run.hyperlink && !last.line_break
            });

            if should_merge {
                // Merge text with the last run, with smart spacing
                if let Some(last) = merged.last_mut() {
                    // Check if we need to add a space between the runs
                    let needs_space = Self::needs_space_between(&last.text, &run.text);
                    if needs_space {
                        last.text.push(' ');
                    }
                    last.text.push_str(&run.text);
                    // Preserve line_break from the merged run
                    if run.line_break {
                        last.line_break = true;
                    }
                }
            } else {
                // Start a new run
                merged.push(run);
            }
        }

        self.runs = merged;
    }

    /// Determine if a space is needed between two text fragments when merging.
    fn needs_space_between(prev: &str, next: &str) -> bool {
        let last_char = match prev.chars().last() {
            Some(c) => c,
            None => return false,
        };
        let first_char = match next.chars().next() {
            Some(c) => c,
            None => return false,
        };

        // No space needed if either side already has whitespace
        if last_char.is_whitespace() || first_char.is_whitespace() {
            return false;
        }

        // No space before punctuation
        if matches!(
            first_char,
            '.' | ',' | ':' | ';' | '!' | '?' | ')' | ']' | '}' | '"' | '\'' | '…' | '~'
        ) {
            return false;
        }

        // No space after opening brackets/quotes
        if matches!(last_char, '(' | '[' | '{' | '"' | '\'') {
            return false;
        }

        // Same-style runs = same word (artificially split by Word)
        // This applies to ALL scripts: ASCII, CJK, Hangul, mixed scripts, etc.
        //
        // Examples of runs that should NOT have spaces added:
        // - "DRBD" stored as ["DRB", "D"] → "DRBD"
        // - "정의" stored as ["정", "의"] → "정의"
        // - "CJ대한통운" stored as ["C", "J", "대한통운"] → "CJ대한통운" (brand name)
        //
        // Key insight: Word splits runs for various reasons (formatting, editing history),
        // but same-style consecutive runs are ALWAYS part of the same word.
        // If they were different words, they would have explicit whitespace between them.
        //
        // Note: Korean DOES use spaces between words, but those spaces exist in the
        // source document (with xml:space="preserve"). We don't invent spaces.
        //
        // Previously this function added spaces at ASCII↔CJK boundaries, but this was
        // incorrect for Korean brand names like "CJ대한통운" where the intent is no space.
        false
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

    #[test]
    fn test_merge_adjacent_runs_ascii_no_split() {
        // Test fix for word splitting bug: "DRBD" stored as ["DRB", "D"] should merge to "DRBD"
        let mut para = Paragraph::new();
        para.runs.push(TextRun::plain("DRB"));
        para.runs.push(TextRun::plain("D"));
        para.merge_adjacent_runs();

        assert_eq!(para.runs.len(), 1);
        assert_eq!(para.runs[0].text, "DRBD"); // NOT "DRB D"
    }

    #[test]
    fn test_merge_adjacent_runs_ping() {
        // Test fix: "PING" stored as ["P", "ING"] should merge to "PING"
        let mut para = Paragraph::new();
        para.runs.push(TextRun::plain("P"));
        para.runs.push(TextRun::plain("ING"));
        para.merge_adjacent_runs();

        assert_eq!(para.runs.len(), 1);
        assert_eq!(para.runs[0].text, "PING"); // NOT "P ING"
    }

    #[test]
    fn test_merge_adjacent_runs_tcp() {
        // Test fix: "TCP" stored as ["T", "CP"] should merge to "TCP"
        let mut para = Paragraph::new();
        para.runs.push(TextRun::plain("T"));
        para.runs.push(TextRun::plain("CP"));
        para.merge_adjacent_runs();

        assert_eq!(para.runs.len(), 1);
        assert_eq!(para.runs[0].text, "TCP"); // NOT "T CP"
    }

    #[test]
    fn test_merge_adjacent_runs_cjk_ascii_no_space() {
        // Same-style runs merge WITHOUT space - even across script boundaries
        // Example: "VIP리소스" is a valid Korean compound where VIP is a brand/term
        let mut para = Paragraph::new();
        para.runs.push(TextRun::plain("리소스"));
        para.runs.push(TextRun::plain("DRBD"));
        para.merge_adjacent_runs();

        assert_eq!(para.runs.len(), 1);
        assert_eq!(para.runs[0].text, "리소스DRBD"); // No space - same-style runs = same word
    }

    #[test]
    fn test_merge_adjacent_runs_ascii_cjk_no_space() {
        // Same-style runs merge WITHOUT space - even across script boundaries
        // Example: "CJ대한통운" is a brand name where CJ is part of the Korean name
        let mut para = Paragraph::new();
        para.runs.push(TextRun::plain("CJ"));
        para.runs.push(TextRun::plain("대한통운"));
        para.merge_adjacent_runs();

        assert_eq!(para.runs.len(), 1);
        assert_eq!(para.runs[0].text, "CJ대한통운"); // No space - brand name
    }

    #[test]
    fn test_merge_adjacent_runs_korean_no_space() {
        // Korean: Same-style runs merge WITHOUT space (like ASCII, Chinese, Japanese)
        // Word may split a single word into multiple runs for various reasons.
        // Example: "정의" stored as ["정", "의"] should become "정의", not "정 의"
        let mut para = Paragraph::new();
        para.runs.push(TextRun::plain("네트워크"));
        para.runs.push(TextRun::plain("카드"));
        para.merge_adjacent_runs();

        assert_eq!(para.runs.len(), 1);
        assert_eq!(para.runs[0].text, "네트워크카드"); // No space - same word or compound
    }

    #[test]
    fn test_merge_adjacent_runs_korean_syllables() {
        // When Word splits Korean syllables, they should merge without space
        // This was the bug: "정의" → "정 의" (wrong) instead of "정의" (correct)
        let mut para = Paragraph::new();
        para.runs.push(TextRun::plain("정"));
        para.runs.push(TextRun::plain("의"));
        para.merge_adjacent_runs();

        assert_eq!(para.runs.len(), 1);
        assert_eq!(para.runs[0].text, "정의"); // Syllables merge without space
    }

    #[test]
    fn test_merge_adjacent_runs_korean_with_explicit_space() {
        // When source has explicit space, it's preserved in the run text
        let mut para = Paragraph::new();
        para.runs.push(TextRun::plain("서버 ")); // Note: space is in the text
        para.runs.push(TextRun::plain("리부팅"));
        para.merge_adjacent_runs();

        assert_eq!(para.runs.len(), 1);
        assert_eq!(para.runs[0].text, "서버 리부팅"); // Space preserved from source
    }

    #[test]
    fn test_merge_adjacent_runs_chinese_no_space() {
        // Chinese: Same-style runs merge without space
        let mut para = Paragraph::new();
        para.runs.push(TextRun::plain("中文"));
        para.runs.push(TextRun::plain("测试"));
        para.merge_adjacent_runs();

        assert_eq!(para.runs.len(), 1);
        assert_eq!(para.runs[0].text, "中文测试"); // No space between Chinese
    }

    #[test]
    fn test_merge_adjacent_runs_japanese_no_space() {
        // Japanese: Same-style runs merge without space
        let mut para = Paragraph::new();
        para.runs.push(TextRun::plain("日本語"));
        para.runs.push(TextRun::plain("テスト"));
        para.merge_adjacent_runs();

        assert_eq!(para.runs.len(), 1);
        assert_eq!(para.runs[0].text, "日本語テスト"); // No space between Japanese
    }

    #[test]
    fn test_merge_adjacent_runs_different_styles_not_merged() {
        // Runs with different styles should NOT be merged
        let mut para = Paragraph::new();
        para.runs.push(TextRun::plain("normal"));
        para.runs.push(TextRun::styled("bold", TextStyle::bold()));
        para.merge_adjacent_runs();

        assert_eq!(para.runs.len(), 2); // Still 2 runs
        assert_eq!(para.runs[0].text, "normal");
        assert_eq!(para.runs[1].text, "bold");
    }

    #[test]
    fn test_merge_preserves_existing_spaces() {
        // Existing spaces should be preserved
        let mut para = Paragraph::new();
        para.runs.push(TextRun::plain("Hello "));
        para.runs.push(TextRun::plain("World"));
        para.merge_adjacent_runs();

        assert_eq!(para.runs.len(), 1);
        assert_eq!(para.runs[0].text, "Hello World"); // Original space preserved
    }
}
