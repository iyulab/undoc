//! Quality Report Generator for undoc
//!
//! This test module generates a comprehensive quality report for all test files,
//! measuring extraction success, document statistics, and identifying issues.
//!
//! Run with: cargo test --test quality_report -- --nocapture

use std::collections::HashMap;
use std::fs;
use std::path::Path;
use undoc::{parse_bytes, Block, Document};

/// Statistics for a single document
#[derive(Debug, Default)]
struct DocumentStats {
    /// File path
    path: String,
    /// File format (docx, xlsx, pptx)
    format: String,
    /// File size in bytes
    file_size: usize,
    /// Parse result
    success: bool,
    /// Error message if failed
    error: Option<String>,
    /// Number of sections (sheets/slides)
    section_count: usize,
    /// Number of paragraphs
    paragraph_count: usize,
    /// Number of tables
    table_count: usize,
    /// Total number of cells across all tables
    cell_count: usize,
    /// Number of merged cells (col_span > 1 or row_span > 1)
    merged_cell_count: usize,
    /// Number of hyperlinks
    hyperlink_count: usize,
    /// Number of images/resources
    image_count: usize,
    /// Number of headings
    heading_count: usize,
    /// Total text length
    text_length: usize,
    /// Warnings (partial failures, missing elements, etc.)
    warnings: Vec<String>,
}

impl DocumentStats {
    fn from_document(path: &str, doc: &Document, file_size: usize) -> Self {
        let format = Path::new(path)
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("unknown")
            .to_lowercase();

        let mut stats = Self {
            path: path.to_string(),
            format,
            file_size,
            success: true,
            error: None,
            section_count: doc.sections.len(),
            image_count: doc.resources.len(),
            text_length: doc.plain_text().len(),
            ..Default::default()
        };

        for section in &doc.sections {
            for block in &section.content {
                match block {
                    Block::Paragraph(para) => {
                        stats.paragraph_count += 1;
                        if para.heading != undoc::HeadingLevel::None {
                            stats.heading_count += 1;
                        }
                        for run in &para.runs {
                            if run.hyperlink.is_some() {
                                stats.hyperlink_count += 1;
                            }
                        }
                    }
                    Block::Table(table) => {
                        stats.table_count += 1;
                        for row in &table.rows {
                            for cell in &row.cells {
                                stats.cell_count += 1;
                                if cell.col_span > 1 || cell.row_span > 1 {
                                    stats.merged_cell_count += 1;
                                }
                                // Check hyperlinks in table cells
                                for para in &cell.content {
                                    for run in &para.runs {
                                        if run.hyperlink.is_some() {
                                            stats.hyperlink_count += 1;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
        }

        // Add warnings for potential issues
        if stats.section_count == 0 {
            stats
                .warnings
                .push("Empty document (no sections)".to_string());
        }
        if stats.format == "xlsx" && stats.table_count == 0 {
            stats.warnings.push("XLSX with no tables".to_string());
        }
        if stats.format == "pptx" && stats.paragraph_count == 0 {
            stats.warnings.push("PPTX with no text content".to_string());
        }

        stats
    }

    fn from_error(path: &str, error: &str, file_size: usize) -> Self {
        let format = Path::new(path)
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("unknown")
            .to_lowercase();

        Self {
            path: path.to_string(),
            format,
            file_size,
            success: false,
            error: Some(error.to_string()),
            ..Default::default()
        }
    }
}

/// Aggregate statistics by format
#[derive(Debug, Default)]
struct FormatStats {
    total: usize,
    success: usize,
    failed: usize,
    total_paragraphs: usize,
    total_tables: usize,
    total_cells: usize,
    total_merged_cells: usize,
    total_hyperlinks: usize,
    total_images: usize,
    total_headings: usize,
}

/// Scan a directory recursively for Office files
fn scan_directory(dir: &Path, stats: &mut Vec<DocumentStats>) {
    if let Ok(entries) = fs::read_dir(dir) {
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                scan_directory(&path, stats);
            } else {
                let ext = path.extension().and_then(|e| e.to_str());
                if matches!(ext, Some("docx" | "xlsx" | "pptx")) {
                    let path_str = path.to_string_lossy().to_string();

                    // Known expected failures
                    let filename = path.file_name().unwrap().to_str().unwrap();
                    let expected_failure = matches!(
                        filename,
                        "badcrc.docx" | "testascii.docx" | "testutf16.docx"
                    );

                    match fs::read(&path) {
                        Ok(data) => {
                            let file_size = data.len();
                            match parse_bytes(&data) {
                                Ok(doc) => {
                                    stats.push(DocumentStats::from_document(
                                        &path_str, &doc, file_size,
                                    ));
                                }
                                Err(e) => {
                                    let mut doc_stats = DocumentStats::from_error(
                                        &path_str,
                                        &e.to_string(),
                                        file_size,
                                    );
                                    if expected_failure {
                                        doc_stats.warnings.push("Expected failure".to_string());
                                    }
                                    stats.push(doc_stats);
                                }
                            }
                        }
                        Err(e) => {
                            stats.push(DocumentStats::from_error(
                                &path_str,
                                &format!("IO error: {}", e),
                                0,
                            ));
                        }
                    }
                }
            }
        }
    }
}

/// Generate quality report
#[test]
fn generate_quality_report() {
    let test_dirs = ["test-files", "test-files/officedissector/test"];

    let mut all_stats: Vec<DocumentStats> = Vec::new();

    for dir in test_dirs {
        if Path::new(dir).exists() {
            scan_directory(Path::new(dir), &mut all_stats);
        }
    }

    if all_stats.is_empty() {
        println!("No test files found. Please ensure test-files directory exists.");
        return;
    }

    // Sort by path for consistent output
    all_stats.sort_by(|a, b| a.path.cmp(&b.path));

    // Calculate aggregate stats by format
    let mut by_format: HashMap<String, FormatStats> = HashMap::new();

    for stat in &all_stats {
        let format_stats = by_format.entry(stat.format.clone()).or_default();
        format_stats.total += 1;
        if stat.success {
            format_stats.success += 1;
            format_stats.total_paragraphs += stat.paragraph_count;
            format_stats.total_tables += stat.table_count;
            format_stats.total_cells += stat.cell_count;
            format_stats.total_merged_cells += stat.merged_cell_count;
            format_stats.total_hyperlinks += stat.hyperlink_count;
            format_stats.total_images += stat.image_count;
            format_stats.total_headings += stat.heading_count;
        } else {
            format_stats.failed += 1;
        }
    }

    // Print report
    println!("\n{}", "=".repeat(80));
    println!("{:^80}", "UNDOC QUALITY REPORT");
    println!("{}\n", "=".repeat(80));

    // Summary by format
    println!("## Summary by Format\n");
    println!(
        "{:<8} {:>8} {:>8} {:>8} {:>10}",
        "Format", "Total", "Success", "Failed", "Rate"
    );
    println!("{:-<50}", "");

    let total_files: usize = all_stats.len();
    let total_success: usize = all_stats.iter().filter(|s| s.success).count();

    for (format, stats) in by_format.iter() {
        let rate = if stats.total > 0 {
            (stats.success as f64 / stats.total as f64) * 100.0
        } else {
            0.0
        };
        println!(
            "{:<8} {:>8} {:>8} {:>8} {:>9.1}%",
            format.to_uppercase(),
            stats.total,
            stats.success,
            stats.failed,
            rate
        );
    }

    println!("{:-<50}", "");
    let total_rate = (total_success as f64 / total_files as f64) * 100.0;
    println!(
        "{:<8} {:>8} {:>8} {:>8} {:>9.1}%\n",
        "TOTAL",
        total_files,
        total_success,
        total_files - total_success,
        total_rate
    );

    // Statistics summary
    println!("## Content Statistics\n");
    println!(
        "{:<10} {:>12} {:>10} {:>10} {:>12} {:>10} {:>10}",
        "Format", "Paragraphs", "Tables", "Cells", "MergedCells", "Links", "Images"
    );
    println!("{:-<80}", "");

    for (format, stats) in by_format.iter() {
        println!(
            "{:<10} {:>12} {:>10} {:>10} {:>12} {:>10} {:>10}",
            format.to_uppercase(),
            stats.total_paragraphs,
            stats.total_tables,
            stats.total_cells,
            stats.total_merged_cells,
            stats.total_hyperlinks,
            stats.total_images
        );
    }
    println!();

    // Failures
    let failures: Vec<_> = all_stats.iter().filter(|s| !s.success).collect();
    if !failures.is_empty() {
        println!("## Failed Files ({})\n", failures.len());
        for stat in failures {
            let expected = stat.warnings.iter().any(|w| w.contains("Expected"));
            let marker = if expected { "[expected]" } else { "" };
            println!("- {} {}", stat.path, marker);
            if let Some(ref err) = stat.error {
                println!("  Error: {}", err);
            }
        }
        println!();
    }

    // Warnings
    let with_warnings: Vec<_> = all_stats
        .iter()
        .filter(|s| s.success && !s.warnings.is_empty())
        .collect();
    if !with_warnings.is_empty() {
        println!("## Warnings ({})\n", with_warnings.len());
        for stat in with_warnings {
            println!("- {}", stat.path);
            for warning in &stat.warnings {
                println!("  ⚠ {}", warning);
            }
        }
        println!();
    }

    // Feature coverage analysis
    println!("## Feature Coverage Analysis\n");

    // Check merged cells
    let xlsx_with_merged: usize = all_stats
        .iter()
        .filter(|s| s.format == "xlsx" && s.merged_cell_count > 0)
        .count();
    let xlsx_total: usize = all_stats
        .iter()
        .filter(|s| s.format == "xlsx" && s.success)
        .count();
    println!(
        "- XLSX merged cells: {}/{} files have detected merged cells",
        xlsx_with_merged, xlsx_total
    );

    // Check hyperlinks
    let with_hyperlinks: usize = all_stats
        .iter()
        .filter(|s| s.success && s.hyperlink_count > 0)
        .count();
    println!(
        "- Hyperlinks: {}/{} files have detected hyperlinks",
        with_hyperlinks, total_success
    );

    // Check headings
    let with_headings: usize = all_stats
        .iter()
        .filter(|s| s.success && s.heading_count > 0)
        .count();
    println!(
        "- Headings: {}/{} files have detected headings",
        with_headings, total_success
    );

    // PPTX headings specifically
    let pptx_with_headings: usize = all_stats
        .iter()
        .filter(|s| s.format == "pptx" && s.heading_count > 0)
        .count();
    let pptx_total: usize = all_stats
        .iter()
        .filter(|s| s.format == "pptx" && s.success)
        .count();
    println!(
        "- PPTX headings: {}/{} files have detected headings",
        pptx_with_headings, pptx_total
    );

    println!();

    // Detailed file list (optional, can be verbose)
    println!("## Detailed Results\n");
    println!(
        "{:<60} {:>6} {:>8} {:>8} {:>6} {:>6}",
        "File", "Status", "Sections", "Tables", "Links", "Imgs"
    );
    println!("{:-<100}", "");

    for stat in &all_stats {
        // Use char indices to avoid breaking UTF-8 boundaries
        let short_path = if stat.path.chars().count() > 58 {
            let chars: Vec<char> = stat.path.chars().collect();
            let start = chars.len().saturating_sub(55);
            format!("...{}", chars[start..].iter().collect::<String>())
        } else {
            stat.path.clone()
        };

        let status = if stat.success { "OK" } else { "FAIL" };
        println!(
            "{:<60} {:>6} {:>8} {:>8} {:>6} {:>6}",
            short_path,
            status,
            stat.section_count,
            stat.table_count,
            stat.hyperlink_count,
            stat.image_count
        );
    }

    // Assert minimum success rate
    assert!(
        total_rate >= 95.0,
        "Success rate should be at least 95%, got {:.1}%",
        total_rate
    );
}

/// Individual file quality check for debugging
#[test]
fn check_specific_file() {
    let files = [
        "test-files/Basic Invoice.xlsx",
        "test-files/file_example_PPT_1MB.pptx",
        "test-files/officedissector/test/unit_test/testdocs/testutf16.docx",
    ];

    for path in files {
        if !Path::new(path).exists() {
            continue;
        }

        println!("\n=== {} ===\n", path);

        match fs::read(path) {
            Ok(data) => match parse_bytes(&data) {
                Ok(doc) => {
                    let stats = DocumentStats::from_document(path, &doc, data.len());
                    println!("Sections: {}", stats.section_count);
                    println!("Paragraphs: {}", stats.paragraph_count);
                    println!("Tables: {}", stats.table_count);
                    println!("Cells: {}", stats.cell_count);
                    println!("Merged cells: {}", stats.merged_cell_count);
                    println!("Hyperlinks: {}", stats.hyperlink_count);
                    println!("Images: {}", stats.image_count);
                    println!("Headings: {}", stats.heading_count);
                    println!("Text length: {} chars", stats.text_length);

                    if !stats.warnings.is_empty() {
                        println!("\nWarnings:");
                        for w in &stats.warnings {
                            println!("  - {}", w);
                        }
                    }
                }
                Err(e) => {
                    println!("Parse error: {}", e);
                }
            },
            Err(e) => {
                println!("Read error: {}", e);
            }
        }
    }
}
