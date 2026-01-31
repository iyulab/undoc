//! Text cleanup and normalization utilities.
//!
//! This module provides text cleaning functionality optimized for
//! LLM training data preparation.

use super::options::CleanupOptions;
use unicode_normalization::UnicodeNormalization;

/// Clean text according to the provided options.
pub fn clean_text(text: &str, options: &CleanupOptions) -> String {
    let mut result = text.to_string();

    if options.normalize_strings {
        result = normalize_unicode(&result);
    }

    if options.remove_pua {
        result = remove_private_use_area(&result);
    }

    if options.clean_lines {
        result = clean_lines(&result, options.preserve_frontmatter);
    }

    if options.filter_structure {
        result = filter_structure(&result);
    }

    if options.final_normalize {
        result = final_normalize(&result);
    }

    result
}

/// Normalize Unicode strings to NFC form and standardize common elements.
fn normalize_unicode(text: &str) -> String {
    let normalized: String = text.nfc().collect();

    // Standardize bullets and dashes
    normalized
        // Various bullet characters
        .replace(['•', '◦', '▪', '▫', '●', '○', '■', '□'], "•")
        // Various dashes (en-dash, em-dash, minus sign, figure dash)
        .replace(['\u{2013}', '\u{2014}', '\u{2212}', '\u{2012}'], "-")
        // Various single quotes (left single, right single)
        .replace(['\u{2018}', '\u{2019}'], "'")
        // Various double quotes (left, right, low-9, left guillemet, right guillemet)
        .replace(
            ['\u{201C}', '\u{201D}', '\u{201E}', '\u{00AB}', '\u{00BB}'],
            "\"",
        )
        // Various spaces (non-breaking, en, em, thin, hair, narrow no-break)
        .replace(
            [
                '\u{00A0}', '\u{2002}', '\u{2003}', '\u{2009}', '\u{200A}', '\u{202F}',
            ],
            " ",
        )
        // Zero-width characters (remove: zero-width space, non-joiner, joiner, BOM)
        .replace(['\u{200B}', '\u{200C}', '\u{200D}', '\u{FEFF}'], "")
}

/// Remove Private Use Area (PUA) characters.
fn remove_private_use_area(text: &str) -> String {
    text.chars()
        .filter(|c| {
            let code = *c as u32;
            // Private Use Area ranges:
            // U+E000 - U+F8FF (BMP PUA)
            // U+F0000 - U+FFFFD (Supplementary PUA-A)
            // U+100000 - U+10FFFD (Supplementary PUA-B)
            !((0xE000..=0xF8FF).contains(&code)
                || (0xF0000..=0xFFFFD).contains(&code)
                || (0x100000..=0x10FFFD).contains(&code))
        })
        .collect()
}

/// Clean lines - remove headers, footers, page numbers, TOC markers.
fn clean_lines(text: &str, preserve_frontmatter: bool) -> String {
    let lines: Vec<&str> = text.lines().collect();
    let mut result = Vec::new();
    let mut in_frontmatter = false;

    for (i, line) in lines.iter().enumerate() {
        // Handle YAML frontmatter
        if preserve_frontmatter {
            if i == 0 && line.trim() == "---" {
                in_frontmatter = true;
                result.push(*line);
                continue;
            }
            if in_frontmatter {
                result.push(*line);
                if line.trim() == "---" {
                    in_frontmatter = false;
                }
                continue;
            }
        }

        // Skip likely header/footer patterns
        if should_skip_line(line) {
            continue;
        }

        result.push(*line);
    }

    result.join("\n")
}

/// Check if a line should be skipped (header, footer, page number, etc.).
fn should_skip_line(line: &str) -> bool {
    let trimmed = line.trim();

    // Empty lines are not skipped
    if trimmed.is_empty() {
        return false;
    }

    // Page number patterns
    if is_page_number(trimmed) {
        return true;
    }

    // Common header/footer patterns
    if is_header_footer(trimmed) {
        return true;
    }

    // TOC marker patterns
    if is_toc_marker(trimmed) {
        return true;
    }

    false
}

/// Check if line is a page number.
fn is_page_number(line: &str) -> bool {
    // Simple page number patterns
    let patterns = ["Page ", "page ", "- ", "— "];

    for pattern in patterns {
        if let Some(rest) = line.strip_prefix(pattern) {
            if rest.trim().chars().all(|c| c.is_ascii_digit()) {
                return true;
            }
        }
        if let Some(rest) = line.strip_suffix(pattern.trim()) {
            if rest.trim().chars().all(|c| c.is_ascii_digit()) {
                return true;
            }
        }
    }

    // Just a number alone (potential page number)
    if line.len() <= 5 && line.chars().all(|c| c.is_ascii_digit()) {
        return true;
    }

    false
}

/// Check if line is a common header/footer.
fn is_header_footer(line: &str) -> bool {
    let lower = line.to_lowercase();

    // Common footer phrases
    let footer_patterns = [
        "all rights reserved",
        "confidential",
        "proprietary",
        "copyright ©",
        "copyright (c)",
        "© ",
        "(c) ",
    ];

    for pattern in footer_patterns {
        if lower.contains(pattern) {
            return true;
        }
    }

    false
}

/// Check if line is a TOC marker.
fn is_toc_marker(line: &str) -> bool {
    let lower = line.to_lowercase();

    // TOC patterns - lines with lots of dots (leader dots)
    if line.contains("...") || line.contains("…") {
        // If it has dots followed by a number, likely TOC entry
        let dot_count = line.chars().filter(|c| *c == '.').count();
        if dot_count > 3 {
            return true;
        }
    }

    // Explicit TOC headers
    if lower == "table of contents" || lower == "contents" {
        return true;
    }

    false
}

/// Filter structural elements - remove empty paragraphs, orphaned elements.
fn filter_structure(text: &str) -> String {
    let lines: Vec<&str> = text.lines().collect();
    let mut result = Vec::new();
    let mut prev_blank = false;

    for line in lines {
        let is_blank = line.trim().is_empty();

        // Skip consecutive blank lines
        if is_blank && prev_blank {
            continue;
        }

        // Skip lines that are just whitespace with special characters
        if !is_blank && line.trim().len() == 1 {
            if let Some(c) = line.trim().chars().next() {
                if matches!(c, '|' | '-' | '_' | '=' | '*' | '#' | '~') {
                    continue;
                }
            }
        }

        result.push(line);
        prev_blank = is_blank;
    }

    result.join("\n")
}

/// Final whitespace normalization.
fn final_normalize(text: &str) -> String {
    let mut result = String::new();

    for line in text.lines() {
        let mut normalized_line = String::new();
        let mut prev_space = false;

        for c in line.chars() {
            if c.is_whitespace() {
                if !prev_space {
                    normalized_line.push(' ');
                    prev_space = true;
                }
            } else {
                normalized_line.push(c);
                prev_space = false;
            }
        }

        // Trim trailing whitespace from each line
        let trimmed = normalized_line.trim_end();
        if !result.is_empty() {
            result.push('\n');
        }
        result.push_str(trimmed);
    }

    // Trim leading/trailing whitespace from entire document
    result.trim().to_string()
}

/// Detect potential mojibake patterns (for reporting, not fixing).
#[allow(dead_code)]
pub fn detect_mojibake(text: &str) -> Vec<(usize, String)> {
    let mut issues = Vec::new();

    // Common mojibake patterns (UTF-8 decoded as Windows-1252, etc.)
    // These are byte sequences that result from mis-encoding
    let patterns: &[(&str, &str)] = &[
        ("\u{00E2}\u{20AC}\u{201C}", "em-dash"),
        ("\u{00E2}\u{20AC}\u{2122}", "apostrophe"),
        ("\u{00E2}\u{20AC}\u{0153}", "left quote"),
        ("\u{00C3}\u{00A9}", "e-acute"),
        ("\u{00C3}\u{00A8}", "e-grave"),
        ("\u{00C3}\u{00A0}", "a-grave"),
        ("\u{00C3}\u{00A2}", "a-circumflex"),
        ("\u{00C2}\u{00A0}", "non-breaking space"),
        ("\u{00C3}\u{00A7}", "c-cedilla"),
    ];

    for (i, line) in text.lines().enumerate() {
        for (pattern, desc) in patterns {
            if line.contains(pattern) {
                issues.push((i + 1, format!("Possible mojibake: {} ({})", pattern, desc)));
            }
        }
    }

    issues
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_normalize_unicode() {
        // en-dash (\u{2013}) and em-dash (\u{2014})
        let input = "Hello \u{2013} World \u{2014} Test";
        let result = normalize_unicode(input);
        assert_eq!(result, "Hello - World - Test");
    }

    #[test]
    fn test_normalize_quotes() {
        // smart quotes: \u{201C} left, \u{201D} right double; \u{2018} left, \u{2019} right single
        let input = "\u{201C}Smart quotes\u{201D} and \u{2018}apostrophes\u{2019}";
        let result = normalize_unicode(input);
        assert_eq!(result, "\"Smart quotes\" and 'apostrophes'");
    }

    #[test]
    fn test_remove_pua() {
        let input = "Normal text\u{E001}hidden\u{F000}text";
        let result = remove_private_use_area(input);
        assert_eq!(result, "Normal texthiddentext");
    }

    #[test]
    fn test_clean_lines_page_numbers() {
        let input = "Content here\nPage 1\nMore content\n15";
        let result = clean_lines(input, false);
        assert!(!result.contains("Page 1"));
        assert!(!result.contains("\n15"));
    }

    #[test]
    fn test_clean_lines_preserve_frontmatter() {
        let input = "---\ntitle: Test\n---\nContent\nPage 1";
        let result = clean_lines(input, true);
        assert!(result.contains("title: Test"));
        assert!(!result.contains("Page 1"));
    }

    #[test]
    fn test_filter_structure() {
        let input = "Line 1\n\n\n\nLine 2";
        let result = filter_structure(input);
        assert!(!result.contains("\n\n\n")); // No triple blank lines
    }

    #[test]
    fn test_final_normalize() {
        let input = "Multiple   spaces   here";
        let result = final_normalize(input);
        assert_eq!(result, "Multiple spaces here");
    }

    #[test]
    fn test_clean_text_full() {
        let options = CleanupOptions {
            normalize_strings: true,
            clean_lines: true,
            filter_structure: true,
            final_normalize: true,
            remove_pua: true,
            detect_mojibake: false,
            preserve_frontmatter: true,
        };

        let input = "---\ntitle: Test\n---\n\nHello – World\n\n\n\nPage 1\nContent.";
        let result = clean_text(input, &options);

        assert!(result.contains("Hello - World")); // Normalized dash
        assert!(!result.contains("Page 1")); // Removed page number
        assert!(!result.contains("\n\n\n")); // No excess blank lines
    }

    #[test]
    fn test_detect_mojibake() {
        // Mojibake pattern for em-dash: \u{00E2}\u{20AC}\u{201C}
        let input = "This has \u{00E2}\u{20AC}\u{201C} some issues";
        let issues = detect_mojibake(input);
        assert!(!issues.is_empty());
    }
}
