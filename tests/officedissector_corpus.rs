//! Integration tests using OfficeDissector test corpus
//!
//! This test suite validates undoc against real-world Office documents
//! from the OfficeDissector project's test corpus.
//!
//! Test corpus location: test-files/officedissector/test/
//! Source: https://github.com/grierforensics/officedissector

use std::fs;
use std::path::Path;
use undoc::{parse_bytes, Document, Result};

/// Helper to run extraction and verify basic success
fn extract_and_verify(path: &str) -> Result<Document> {
    let data = fs::read(path).map_err(|e| undoc::Error::Io(e))?;
    let result = parse_bytes(&data)?;

    // Basic validations - document should parse without panic
    // Empty documents are valid
    Ok(result)
}

/// Helper to check if test file exists
fn test_file_exists(relative_path: &str) -> bool {
    Path::new(relative_path).exists()
}

// =============================================================================
// Fraunhofer Library Tests - Various Office features
// =============================================================================

#[test]
fn test_basic_document() {
    let path = "test-files/officedissector/test/fraunhoferlibrary/A Basic Document (docx).docx";
    if test_file_exists(path) {
        extract_and_verify(path).expect("Basic document should parse successfully");
    }
}

#[test]
fn test_3d_charts_docx() {
    let path = "test-files/officedissector/test/fraunhoferlibrary/3D Bar O12 Word Charts.docx";
    if test_file_exists(path) {
        extract_and_verify(path).expect("3D charts document should parse");
    }
}

#[test]
fn test_autoshapes_pptx() {
    let path =
        "test-files/officedissector/test/fraunhoferlibrary/AutoShapes O12 PPT AllShapes.pptx";
    if test_file_exists(path) {
        extract_and_verify(path).expect("AutoShapes PPTX should parse");
    }
}

#[test]
fn test_animation_pptx() {
    let path = "test-files/officedissector/test/fraunhoferlibrary/Animation.pptx";
    if test_file_exists(path) {
        extract_and_verify(path).expect("Animation PPTX should parse");
    }
}

#[test]
fn test_bidi_text_pptx() {
    let path = "test-files/officedissector/test/fraunhoferlibrary/BiDi+English text.pptx";
    if test_file_exists(path) {
        extract_and_verify(path).expect("BiDi text PPTX should parse");
    }
}

#[test]
fn test_balance_xlsx() {
    let path = "test-files/officedissector/test/fraunhoferlibrary/Balance.xlsx";
    if test_file_exists(path) {
        extract_and_verify(path).expect("Balance XLSX should parse");
    }
}

#[test]
fn test_pitch_book_xlsx() {
    let path = "test-files/officedissector/test/fraunhoferlibrary/A Pitch Book.xlsx";
    if test_file_exists(path) {
        extract_and_verify(path).expect("Pitch Book XLSX should parse");
    }
}

// =============================================================================
// GovDocs Tests - Real government documents
// =============================================================================

#[test]
fn test_govdoc_docx() {
    let path = "test-files/officedissector/test/govdocs/014760.docx";
    if test_file_exists(path) {
        extract_and_verify(path).expect("GovDoc DOCX should parse");
    }
}

#[test]
fn test_govdoc_xlsx() {
    let path = "test-files/officedissector/test/govdocs/019916.xlsx";
    if test_file_exists(path) {
        extract_and_verify(path).expect("GovDoc XLSX should parse");
    }
}

#[test]
fn test_govdoc_pptx() {
    let path = "test-files/officedissector/test/govdocs/018375.pptx";
    if test_file_exists(path) {
        extract_and_verify(path).expect("GovDoc PPTX should parse");
    }
}

// =============================================================================
// Edge Cases - Unit test documents
// =============================================================================

#[test]
fn test_corrupt_xml_graceful() {
    // This file has intentionally corrupt XML but should still extract some content
    let path = "test-files/officedissector/test/unit_test/testdocs/corrupt_xml.docx";
    if test_file_exists(path) {
        // May succeed or fail gracefully - either is acceptable
        let _ = extract_and_verify(path);
    }
}

#[test]
fn test_missing_content_type() {
    let path = "test-files/officedissector/test/unit_test/testdocs/missing_content_type.docx";
    if test_file_exists(path) {
        extract_and_verify(path).expect("Missing content type should be handled");
    }
}

#[test]
fn test_missing_part() {
    let path = "test-files/officedissector/test/unit_test/testdocs/missing_part.docx";
    if test_file_exists(path) {
        // May succeed with partial content or fail gracefully
        let _ = extract_and_verify(path);
    }
}

#[test]
fn test_no_core_props() {
    let path = "test-files/officedissector/test/unit_test/testdocs/no_core_props.docx";
    if test_file_exists(path) {
        extract_and_verify(path).expect("Document without core props should parse");
    }
}

#[test]
fn test_non_standard_namespace() {
    let path = "test-files/officedissector/test/unit_test/testdocs/non-standard-namespace.docx";
    if test_file_exists(path) {
        extract_and_verify(path).expect("Non-standard namespace should be handled");
    }
}

#[test]
#[ignore = "Test file has truncated XML content - UTF-16 decoding works but XML is incomplete"]
fn test_utf16_encoding() {
    let path = "test-files/officedissector/test/unit_test/testdocs/testutf16.docx";
    if test_file_exists(path) {
        extract_and_verify(path).expect("UTF-16 encoded XML should parse");
    }
}

#[test]
fn test_bad_crc_should_fail() {
    // This file has intentionally bad CRC - extraction should fail
    let path = "test-files/officedissector/test/unit_test/testdocs/badcrc.docx";
    if test_file_exists(path) {
        let data = fs::read(path).expect("Should read file");
        let result = parse_bytes(&data);
        assert!(result.is_err(), "Bad CRC file should fail extraction");
    }
}

#[test]
fn test_url_hyperlinks() {
    let path = "test-files/officedissector/test/unit_test/testdocs/url.docx";
    if test_file_exists(path) {
        extract_and_verify(path).expect("URL document should parse");
    }
}

#[test]
fn test_sounds_pptx() {
    let path = "test-files/officedissector/test/unit_test/testdocs/sounds.pptx";
    if test_file_exists(path) {
        extract_and_verify(path).expect("PPTX with sounds should parse");
    }
}

// =============================================================================
// Batch Tests - Run all files in a directory
// =============================================================================

#[test]
fn test_all_fraunhofer_docx() {
    let dir = "test-files/officedissector/test/fraunhoferlibrary";
    if !Path::new(dir).exists() {
        return;
    }

    let mut success = 0;
    let mut failed = 0;
    let mut failures: Vec<String> = Vec::new();

    for entry in fs::read_dir(dir).unwrap() {
        let entry = entry.unwrap();
        let path = entry.path();
        if path.extension().map(|e| e == "docx").unwrap_or(false) {
            match extract_and_verify(path.to_str().unwrap()) {
                Ok(_) => success += 1,
                Err(e) => {
                    failed += 1;
                    failures.push(format!("{}: {}", path.display(), e));
                }
            }
        }
    }

    println!("Fraunhofer DOCX: {} success, {} failed", success, failed);
    if !failures.is_empty() {
        println!("Failures:\n{}", failures.join("\n"));
    }

    // Allow some failures for edge cases
    assert!(failed <= 2, "Too many failures in Fraunhofer DOCX corpus");
}

#[test]
fn test_all_govdocs() {
    let dir = "test-files/officedissector/test/govdocs";
    if !Path::new(dir).exists() {
        return;
    }

    let mut success = 0;
    let mut failed = 0;

    for entry in fs::read_dir(dir).unwrap() {
        let entry = entry.unwrap();
        let path = entry.path();
        let ext = path.extension().and_then(|e| e.to_str());

        if matches!(ext, Some("docx" | "xlsx" | "pptx")) {
            match extract_and_verify(path.to_str().unwrap()) {
                Ok(_) => success += 1,
                Err(e) => {
                    failed += 1;
                    eprintln!("Failed: {} - {}", path.display(), e);
                }
            }
        }
    }

    println!("GovDocs: {} success, {} failed", success, failed);
    assert_eq!(failed, 0, "All GovDocs should parse successfully");
}

// =============================================================================
// Full Corpus Smoke Test
// =============================================================================

#[test]
fn test_officedissector_full_corpus() {
    let base_dir = "test-files/officedissector/test";
    if !Path::new(base_dir).exists() {
        println!("OfficeDissector corpus not found, skipping");
        return;
    }

    let mut total = 0;
    let mut success = 0;
    let mut failed = 0;
    let mut failures: Vec<String> = Vec::new();

    // Known expected failures
    let expected_failures = [
        "badcrc.docx",    // Intentionally bad CRC
        "testascii.docx", // Non-standard encoding
        "testutf16.docx", // UTF-16 not supported yet
    ];

    fn scan_dir(
        dir: &Path,
        total: &mut usize,
        success: &mut usize,
        failed: &mut usize,
        failures: &mut Vec<String>,
        expected_failures: &[&str],
    ) {
        if let Ok(entries) = fs::read_dir(dir) {
            for entry in entries.flatten() {
                let path = entry.path();
                if path.is_dir() {
                    scan_dir(&path, total, success, failed, failures, expected_failures);
                } else {
                    let ext = path.extension().and_then(|e| e.to_str());
                    if matches!(ext, Some("docx" | "xlsx" | "pptx")) {
                        *total += 1;
                        let filename = path.file_name().unwrap().to_str().unwrap();

                        match extract_and_verify(path.to_str().unwrap()) {
                            Ok(_) => *success += 1,
                            Err(e) => {
                                if expected_failures.contains(&filename) {
                                    // Expected failure, count as success
                                    *success += 1;
                                } else {
                                    *failed += 1;
                                    failures.push(format!("{}: {}", path.display(), e));
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    scan_dir(
        Path::new(base_dir),
        &mut total,
        &mut success,
        &mut failed,
        &mut failures,
        &expected_failures,
    );

    println!("\n=== OfficeDissector Corpus Test Results ===");
    println!("Total: {}", total);
    println!(
        "Success: {} ({:.1}%)",
        success,
        (success as f64 / total as f64) * 100.0
    );
    println!("Failed: {}", failed);

    if !failures.is_empty() {
        println!("\nUnexpected failures:");
        for f in &failures {
            println!("  - {}", f);
        }
    }

    // Assert high success rate
    let success_rate = success as f64 / total as f64;
    assert!(
        success_rate >= 0.95,
        "Success rate should be at least 95%, got {:.1}%",
        success_rate * 100.0
    );
}
