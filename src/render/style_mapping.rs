//! Style name to heading level mapping.
//!
//! This module provides a configurable mapping from style names (both style ID and style name)
//! to heading levels. It supports both English and Korean style names commonly used in documents.

use std::collections::HashMap;

use crate::model::HeadingLevel;

/// Mapping from style names/IDs to heading levels.
#[derive(Debug, Clone, Default)]
pub struct StyleMapping {
    /// Mapping from style name (case-insensitive) to heading level
    name_to_heading: HashMap<String, HeadingLevel>,
    /// Mapping from style ID to heading level
    id_to_heading: HashMap<String, HeadingLevel>,
}

impl StyleMapping {
    /// Create a new empty style mapping.
    pub fn new() -> Self {
        Self::default()
    }

    /// Create a style mapping with default patterns for English and Korean.
    pub fn with_defaults() -> Self {
        let mut mapping = Self::new();

        // English heading patterns
        mapping.add_name_mapping("Heading 1", HeadingLevel::H1);
        mapping.add_name_mapping("Heading 2", HeadingLevel::H2);
        mapping.add_name_mapping("Heading 3", HeadingLevel::H3);
        mapping.add_name_mapping("Heading 4", HeadingLevel::H4);
        mapping.add_name_mapping("Heading 5", HeadingLevel::H5);
        mapping.add_name_mapping("Heading 6", HeadingLevel::H6);
        mapping.add_name_mapping("Title", HeadingLevel::H1);
        mapping.add_name_mapping("Subtitle", HeadingLevel::H2);
        mapping.add_name_mapping("Chapter", HeadingLevel::H1);

        // Korean heading patterns (제목)
        mapping.add_name_mapping("제목 1", HeadingLevel::H1);
        mapping.add_name_mapping("제목 2", HeadingLevel::H2);
        mapping.add_name_mapping("제목 3", HeadingLevel::H3);
        mapping.add_name_mapping("제목 4", HeadingLevel::H4);
        mapping.add_name_mapping("제목 5", HeadingLevel::H5);
        mapping.add_name_mapping("제목 6", HeadingLevel::H6);
        mapping.add_name_mapping("제목1", HeadingLevel::H1);
        mapping.add_name_mapping("제목2", HeadingLevel::H2);
        mapping.add_name_mapping("제목3", HeadingLevel::H3);
        mapping.add_name_mapping("제목4", HeadingLevel::H4);
        mapping.add_name_mapping("제목5", HeadingLevel::H5);
        mapping.add_name_mapping("제목6", HeadingLevel::H6);

        // Common style IDs
        mapping.add_id_mapping("Heading1", HeadingLevel::H1);
        mapping.add_id_mapping("Heading2", HeadingLevel::H2);
        mapping.add_id_mapping("Heading3", HeadingLevel::H3);
        mapping.add_id_mapping("Heading4", HeadingLevel::H4);
        mapping.add_id_mapping("Heading5", HeadingLevel::H5);
        mapping.add_id_mapping("Heading6", HeadingLevel::H6);
        mapping.add_id_mapping("Title", HeadingLevel::H1);
        mapping.add_id_mapping("Subtitle", HeadingLevel::H2);

        mapping
    }

    /// Add a name-based mapping (case-insensitive).
    pub fn add_name_mapping(&mut self, name: impl Into<String>, level: HeadingLevel) {
        self.name_to_heading
            .insert(name.into().to_lowercase(), level);
    }

    /// Add an ID-based mapping (exact match).
    pub fn add_id_mapping(&mut self, id: impl Into<String>, level: HeadingLevel) {
        self.id_to_heading.insert(id.into(), level);
    }

    /// Get heading level by style name (case-insensitive).
    pub fn get_by_name(&self, name: &str) -> Option<HeadingLevel> {
        self.name_to_heading.get(&name.to_lowercase()).copied()
    }

    /// Get heading level by style ID (exact match).
    pub fn get_by_id(&self, id: &str) -> Option<HeadingLevel> {
        self.id_to_heading.get(id).copied()
    }

    /// Get heading level by either style name or ID.
    /// Style name takes precedence.
    pub fn get(&self, style_id: Option<&str>, style_name: Option<&str>) -> Option<HeadingLevel> {
        // Try style name first (more human-readable)
        if let Some(name) = style_name {
            if let Some(level) = self.get_by_name(name) {
                return Some(level);
            }
        }

        // Fall back to style ID
        if let Some(id) = style_id {
            if let Some(level) = self.get_by_id(id) {
                return Some(level);
            }
        }

        None
    }

    /// Check if the mapping is empty.
    pub fn is_empty(&self) -> bool {
        self.name_to_heading.is_empty() && self.id_to_heading.is_empty()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_default_english_mappings() {
        let mapping = StyleMapping::with_defaults();

        assert_eq!(mapping.get_by_name("Heading 1"), Some(HeadingLevel::H1));
        assert_eq!(mapping.get_by_name("heading 1"), Some(HeadingLevel::H1));
        assert_eq!(mapping.get_by_name("HEADING 1"), Some(HeadingLevel::H1));
        assert_eq!(mapping.get_by_name("Title"), Some(HeadingLevel::H1));
    }

    #[test]
    fn test_default_korean_mappings() {
        let mapping = StyleMapping::with_defaults();

        assert_eq!(mapping.get_by_name("제목 1"), Some(HeadingLevel::H1));
        assert_eq!(mapping.get_by_name("제목1"), Some(HeadingLevel::H1));
        assert_eq!(mapping.get_by_name("제목 2"), Some(HeadingLevel::H2));
    }

    #[test]
    fn test_id_mappings() {
        let mapping = StyleMapping::with_defaults();

        assert_eq!(mapping.get_by_id("Heading1"), Some(HeadingLevel::H1));
        assert_eq!(mapping.get_by_id("Title"), Some(HeadingLevel::H1));
        assert_eq!(mapping.get_by_id("Unknown"), None);
    }

    #[test]
    fn test_combined_lookup() {
        let mapping = StyleMapping::with_defaults();

        // Name takes precedence
        assert_eq!(
            mapping.get(Some("Heading1"), Some("제목 1")),
            Some(HeadingLevel::H1)
        );

        // Fall back to ID
        assert_eq!(mapping.get(Some("Heading2"), None), Some(HeadingLevel::H2));

        // Neither matches
        assert_eq!(mapping.get(Some("Unknown"), Some("Unknown")), None);
    }

    #[test]
    fn test_custom_mapping() {
        let mut mapping = StyleMapping::new();
        mapping.add_name_mapping("Custom Title", HeadingLevel::H1);
        mapping.add_id_mapping("CustomID", HeadingLevel::H3);

        assert_eq!(mapping.get_by_name("custom title"), Some(HeadingLevel::H1));
        assert_eq!(mapping.get_by_id("CustomID"), Some(HeadingLevel::H3));
    }
}
