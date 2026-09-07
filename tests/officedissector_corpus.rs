//! Integration tests using the OfficeDissector test corpus.
//!
//! These validate undoc against real-world Office documents rather than the minimal
//! packages the rest of the suite builds. The corpus is large and lives outside this
//! repository, so every test here is `#[ignore]`d by default:
//!
//! ```text
//! git clone https://github.com/grierforensics/officedissector test-files/officedissector
//! cargo test --test officedissector_corpus -- --ignored
//! ```
//!
//! Expected corpus root: `test-files/officedissector/test/`.
//!
//! They were previously guarded by an `if file_exists` check instead, which reported
//! them as *passing* when the corpus was absent — so a plain `cargo test` claimed
//! twenty-one real-world documents had been validated while reading none of them.
//! `#[ignore]` makes the same situation show up as ignored, and a missing file under
//! `--ignored` now fails rather than passing quietly.

use std::fs;
use std::path::Path;
use undoc::{parse_bytes, Document, Result};

/// Read a corpus file and parse it.
///
/// A missing file is a panic, not an `Err`. Several tests here accept either parse
/// outcome — they only pin that the parser does not panic — and if "file not found"
/// could reach them as an ordinary error, an absent corpus would look exactly like a
/// document that was handled gracefully.
fn extract_and_verify(path: &str) -> Result<Document> {
    let data = fs::read(path).unwrap_or_else(|e| {
        panic!("corpus file {path} could not be read ({e}); see the module docs")
    });
    parse_bytes(&data)
}

// =============================================================================
// Fraunhofer Library Tests - Various Office features
// =============================================================================

#[test]
#[ignore = "requires the external OfficeDissector corpus; see the module docs"]
fn test_basic_document() {
    let path = "test-files/officedissector/test/fraunhoferlibrary/A Basic Document (docx).docx";
    extract_and_verify(path).expect("Basic document should parse successfully");
}

#[test]
#[ignore = "requires the external OfficeDissector corpus; see the module docs"]
fn test_3d_charts_docx() {
    let path = "test-files/officedissector/test/fraunhoferlibrary/3D Bar O12 Word Charts.docx";
    extract_and_verify(path).expect("3D charts document should parse");
}

#[test]
#[ignore = "requires the external OfficeDissector corpus; see the module docs"]
fn test_autoshapes_pptx() {
    let path =
        "test-files/officedissector/test/fraunhoferlibrary/AutoShapes O12 PPT AllShapes.pptx";
    extract_and_verify(path).expect("AutoShapes PPTX should parse");
}

#[test]
#[ignore = "requires the external OfficeDissector corpus; see the module docs"]
fn test_animation_pptx() {
    let path = "test-files/officedissector/test/fraunhoferlibrary/Animation.pptx";
    extract_and_verify(path).expect("Animation PPTX should parse");
}

#[test]
#[ignore = "requires the external OfficeDissector corpus; see the module docs"]
fn test_bidi_text_pptx() {
    let path = "test-files/officedissector/test/fraunhoferlibrary/BiDi+English text.pptx";
    extract_and_verify(path).expect("BiDi text PPTX should parse");
}

#[test]
#[ignore = "requires the external OfficeDissector corpus; see the module docs"]
fn test_balance_xlsx() {
    let path = "test-files/officedissector/test/fraunhoferlibrary/Balance.xlsx";
    extract_and_verify(path).expect("Balance XLSX should parse");
}

#[test]
#[ignore = "requires the external OfficeDissector corpus; see the module docs"]
fn test_pitch_book_xlsx() {
    let path = "test-files/officedissector/test/fraunhoferlibrary/A Pitch Book.xlsx";
    extract_and_verify(path).expect("Pitch Book XLSX should parse");
}

// =============================================================================
// GovDocs Tests - Real government documents
// =============================================================================

#[test]
#[ignore = "requires the external OfficeDissector corpus; see the module docs"]
fn test_govdoc_docx() {
    let path = "test-files/officedissector/test/govdocs/014760.docx";
    extract_and_verify(path).expect("GovDoc DOCX should parse");
}

#[test]
#[ignore = "requires the external OfficeDissector corpus; see the module docs"]
fn test_govdoc_xlsx() {
    let path = "test-files/officedissector/test/govdocs/019916.xlsx";
    extract_and_verify(path).expect("GovDoc XLSX should parse");
}

#[test]
#[ignore = "requires the external OfficeDissector corpus; see the module docs"]
fn test_govdoc_pptx() {
    let path = "test-files/officedissector/test/govdocs/018375.pptx";
    extract_and_verify(path).expect("GovDoc PPTX should parse");
}

// =============================================================================
// Edge Cases - Unit test documents
// =============================================================================

#[test]
#[ignore = "requires the external OfficeDissector corpus; see the module docs"]
fn test_corrupt_xml_graceful() {
    // This file has intentionally corrupt XML but should still extract some content
    let path = "test-files/officedissector/test/unit_test/testdocs/corrupt_xml.docx";
    // May succeed or fail gracefully - either is acceptable
    let _ = extract_and_verify(path);
}

#[test]
#[ignore = "requires the external OfficeDissector corpus; see the module docs"]
fn test_missing_content_type() {
    let path = "test-files/officedissector/test/unit_test/testdocs/missing_content_type.docx";
    extract_and_verify(path).expect("Missing content type should be handled");
}

#[test]
#[ignore = "requires the external OfficeDissector corpus; see the module docs"]
fn test_missing_part() {
    let path = "test-files/officedissector/test/unit_test/testdocs/missing_part.docx";
    // May succeed with partial content or fail gracefully
    let _ = extract_and_verify(path);
}

#[test]
#[ignore = "requires the external OfficeDissector corpus; see the module docs"]
fn test_no_core_props() {
    let path = "test-files/officedissector/test/unit_test/testdocs/no_core_props.docx";
    extract_and_verify(path).expect("Document without core props should parse");
}

#[test]
#[ignore = "requires the external OfficeDissector corpus; see the module docs"]
fn test_non_standard_namespace() {
    let path = "test-files/officedissector/test/unit_test/testdocs/non-standard-namespace.docx";
    extract_and_verify(path).expect("Non-standard namespace should be handled");
}

/// Ignored for this corpus file alone: its XML is truncated, so it cannot demonstrate a
/// successful parse. The UTF-16 guarantee is pinned without a fixture in
/// `utf16_xml_end_to_end_test.rs`.
#[test]
#[ignore = "this corpus file's XML is truncated; UTF-16 itself is covered by utf16_xml_end_to_end_test"]
fn test_utf16_encoding() {
    let path = "test-files/officedissector/test/unit_test/testdocs/testutf16.docx";
    extract_and_verify(path).expect("UTF-16 encoded XML should parse");
}

#[test]
#[ignore = "requires the external OfficeDissector corpus; see the module docs"]
fn test_bad_crc_should_fail() {
    // This file has intentionally bad CRC - extraction should fail
    let path = "test-files/officedissector/test/unit_test/testdocs/badcrc.docx";
    let data = fs::read(path).expect("Should read file");
    let result = parse_bytes(&data);
    assert!(result.is_err(), "Bad CRC file should fail extraction");
}

#[test]
#[ignore = "requires the external OfficeDissector corpus; see the module docs"]
fn test_url_hyperlinks() {
    let path = "test-files/officedissector/test/unit_test/testdocs/url.docx";
    extract_and_verify(path).expect("URL document should parse");
}

#[test]
#[ignore = "requires the external OfficeDissector corpus; see the module docs"]
fn test_sounds_pptx() {
    let path = "test-files/officedissector/test/unit_test/testdocs/sounds.pptx";
    extract_and_verify(path).expect("PPTX with sounds should parse");
}

// =============================================================================
// Batch Tests - Run all files in a directory
// =============================================================================

#[test]
#[ignore = "requires the external OfficeDissector corpus; see the module docs"]
fn test_all_fraunhofer_docx() {
    let dir = "test-files/officedissector/test/fraunhoferlibrary";
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
#[ignore = "requires the external OfficeDissector corpus; see the module docs"]
fn test_all_govdocs() {
    let dir = "test-files/officedissector/test/govdocs";
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
#[ignore = "requires the external OfficeDissector corpus; see the module docs"]
fn test_officedissector_full_corpus() {
    let base_dir = "test-files/officedissector/test";
    assert!(
        Path::new(base_dir).exists(),
        "corpus root {base_dir} not found; see the module docs"
    );

    let mut total = 0;
    let mut success = 0;
    let mut failed = 0;
    let mut failures: Vec<String> = Vec::new();

    // Known expected failures
    let expected_failures = [
        "badcrc.docx",    // Intentionally bad CRC
        "testascii.docx", // Non-standard encoding
        // Decoded fine; this particular file's XML is truncated. UTF-16 support itself is
        // covered without a fixture by `utf16_xml_end_to_end_test.rs`.
        "testutf16.docx",
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
