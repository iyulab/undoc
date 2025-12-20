//! Rendering options configuration.

use std::path::PathBuf;

/// How to render complex tables.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum TableFallback {
    /// Use Markdown pipe tables (may break with merged cells)
    #[default]
    Markdown,
    /// Fall back to HTML tables for complex layouts
    Html,
    /// Use ASCII art tables
    Ascii,
}

/// Cleanup preset for LLM training data preparation.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum CleanupPreset {
    /// Minimal cleanup: only essential normalization
    Minimal,
    /// Default: balanced normalization
    #[default]
    Default,
    /// Aggressive: maximum purification
    Aggressive,
}

/// Cleanup options for post-processing.
#[derive(Debug, Clone, Default)]
pub struct CleanupOptions {
    /// Normalize Unicode strings (NFC), standardize bullets
    pub normalize_strings: bool,

    /// Remove headers, footers, page numbers, TOC markers
    pub clean_lines: bool,

    /// Filter empty paragraphs, orphaned elements
    pub filter_structure: bool,

    /// Final whitespace normalization
    pub final_normalize: bool,

    /// Remove Private Use Area characters
    pub remove_pua: bool,

    /// Detect and flag potential mojibake
    pub detect_mojibake: bool,

    /// Preserve YAML frontmatter during cleanup
    pub preserve_frontmatter: bool,
}

impl CleanupOptions {
    /// Create cleanup options from a preset.
    pub fn from_preset(preset: CleanupPreset) -> Self {
        match preset {
            CleanupPreset::Minimal => Self {
                normalize_strings: true,
                final_normalize: true,
                ..Default::default()
            },
            CleanupPreset::Default => Self {
                normalize_strings: true,
                clean_lines: true,
                final_normalize: true,
                preserve_frontmatter: true,
                ..Default::default()
            },
            CleanupPreset::Aggressive => Self {
                normalize_strings: true,
                clean_lines: true,
                filter_structure: true,
                final_normalize: true,
                remove_pua: true,
                detect_mojibake: true,
                preserve_frontmatter: true,
            },
        }
    }

    /// Create minimal cleanup options.
    pub fn minimal() -> Self {
        Self::from_preset(CleanupPreset::Minimal)
    }

    /// Create default cleanup options.
    pub fn standard() -> Self {
        Self::from_preset(CleanupPreset::Default)
    }

    /// Create aggressive cleanup options.
    pub fn aggressive() -> Self {
        Self::from_preset(CleanupPreset::Aggressive)
    }
}

/// Options for rendering documents.
#[derive(Debug, Clone)]
pub struct RenderOptions {
    /// Directory to save extracted images
    pub image_dir: Option<PathBuf>,

    /// Prefix for image paths in markdown (e.g., "assets/")
    pub image_path_prefix: String,

    /// How to handle complex tables
    pub table_fallback: TableFallback,

    /// Maximum heading level (1-6)
    pub max_heading_level: u8,

    /// Include YAML frontmatter with metadata
    pub include_frontmatter: bool,

    /// Preserve line breaks within paragraphs
    pub preserve_line_breaks: bool,

    /// Include empty paragraphs in output
    pub include_empty_paragraphs: bool,

    /// Character for unordered list markers
    pub list_marker: char,

    /// Use ATX-style headers (# instead of underlines)
    pub use_atx_headers: bool,

    /// Add blank line between paragraphs
    pub paragraph_spacing: bool,

    /// Escape special Markdown characters
    pub escape_special_chars: bool,

    /// Cleanup options (None = no cleanup)
    pub cleanup: Option<CleanupOptions>,
}

impl Default for RenderOptions {
    fn default() -> Self {
        Self {
            image_dir: None,
            image_path_prefix: String::new(),
            table_fallback: TableFallback::Markdown,
            max_heading_level: 6,
            include_frontmatter: false,
            preserve_line_breaks: false,
            include_empty_paragraphs: false,
            list_marker: '-',
            use_atx_headers: true,
            paragraph_spacing: true,
            escape_special_chars: true,
            cleanup: None,
        }
    }
}

impl RenderOptions {
    /// Create new render options.
    pub fn new() -> Self {
        Self::default()
    }

    /// Set the image output directory.
    pub fn with_image_dir(mut self, dir: impl Into<PathBuf>) -> Self {
        self.image_dir = Some(dir.into());
        self
    }

    /// Set the image path prefix for markdown references.
    pub fn with_image_prefix(mut self, prefix: impl Into<String>) -> Self {
        self.image_path_prefix = prefix.into();
        self
    }

    /// Set table fallback mode.
    pub fn with_table_fallback(mut self, fallback: TableFallback) -> Self {
        self.table_fallback = fallback;
        self
    }

    /// Enable YAML frontmatter.
    pub fn with_frontmatter(mut self, include: bool) -> Self {
        self.include_frontmatter = include;
        self
    }

    /// Enable cleanup with default options.
    pub fn with_cleanup(mut self) -> Self {
        self.cleanup = Some(CleanupOptions::standard());
        self
    }

    /// Enable cleanup with specific preset.
    pub fn with_cleanup_preset(mut self, preset: CleanupPreset) -> Self {
        self.cleanup = Some(CleanupOptions::from_preset(preset));
        self
    }

    /// Enable cleanup with custom options.
    pub fn with_cleanup_options(mut self, options: CleanupOptions) -> Self {
        self.cleanup = Some(options);
        self
    }

    /// Set maximum heading level.
    pub fn with_max_heading(mut self, level: u8) -> Self {
        self.max_heading_level = level.clamp(1, 6);
        self
    }

    /// Preserve line breaks within paragraphs.
    pub fn with_preserve_breaks(mut self, preserve: bool) -> Self {
        self.preserve_line_breaks = preserve;
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_default_options() {
        let opts = RenderOptions::default();
        assert!(opts.image_dir.is_none());
        assert!(opts.include_frontmatter == false);
        assert_eq!(opts.table_fallback, TableFallback::Markdown);
    }

    #[test]
    fn test_builder_pattern() {
        let opts = RenderOptions::new()
            .with_image_dir("assets")
            .with_frontmatter(true)
            .with_cleanup();

        assert_eq!(opts.image_dir, Some(PathBuf::from("assets")));
        assert!(opts.include_frontmatter);
        assert!(opts.cleanup.is_some());
    }

    #[test]
    fn test_cleanup_presets() {
        let minimal = CleanupOptions::minimal();
        assert!(minimal.normalize_strings);
        assert!(!minimal.clean_lines);

        let aggressive = CleanupOptions::aggressive();
        assert!(aggressive.normalize_strings);
        assert!(aggressive.clean_lines);
        assert!(aggressive.filter_structure);
        assert!(aggressive.remove_pua);
    }

    #[test]
    fn test_max_heading_clamp() {
        let opts = RenderOptions::new().with_max_heading(10);
        assert_eq!(opts.max_heading_level, 6);

        let opts = RenderOptions::new().with_max_heading(0);
        assert_eq!(opts.max_heading_level, 1);
    }
}
