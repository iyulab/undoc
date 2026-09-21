//! Every image link `convert` renders must point at a file `convert` actually wrote.
//!
//! The CLI does both halves of this: it extracts the image bytes into a subdirectory and
//! it renders the markdown that refers to them. Nothing failed when the two disagreed --
//! the file was written, the link was rendered, the command reported success, and the
//! document was simply unusable. This test closes the loop by reading the produced
//! markdown and resolving each destination on disk.
//!
//! The DOCX is built here rather than read from a fixture directory, so the test runs
//! everywhere instead of quietly skipping when a fixture is absent.

use std::io::{Cursor, Write};
use std::path::Path;
use std::process::Command;

fn bin() -> &'static str {
    env!("CARGO_BIN_EXE_undoc")
}

const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;

const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;

const DOCUMENT_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>A paragraph, so the document is not empty.</w:t></w:r></w:p>
    <w:p><w:r><w:drawing><a:blip r:embed="rIdImg"/></w:drawing></w:r></w:p>
  </w:body>
</w:document>"#;

const DOCUMENT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>"#;

/// A one-pixel PNG. The bytes only have to survive the round trip; nothing decodes them.
const PNG: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
    0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
    0x42, 0x60, 0x82,
];

fn docx_with_one_image() -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);

    for (name, body) in [
        ("[Content_Types].xml", CONTENT_TYPES),
        ("_rels/.rels", ROOT_RELS),
        ("word/document.xml", DOCUMENT_XML),
        ("word/_rels/document.xml.rels", DOCUMENT_RELS),
    ] {
        zip.start_file(name, options).unwrap();
        zip.write_all(body.as_bytes()).unwrap();
    }
    zip.start_file("word/media/image1.png", options).unwrap();
    zip.write_all(PNG).unwrap();

    zip.finish().unwrap().into_inner()
}

/// Destinations of every `![alt](dest)` in the markdown, in order.
fn image_destinations(markdown: &str) -> Vec<String> {
    let mut found = Vec::new();
    let mut rest = markdown;
    while let Some(start) = rest.find("![") {
        rest = &rest[start..];
        let Some(close) = rest.find("](") else { break };
        let Some(end) = rest[close + 2..].find(')') else {
            break;
        };
        found.push(rest[close + 2..close + 2 + end].to_string());
        rest = &rest[close + 2 + end..];
    }
    found
}

#[test]
fn convert_renders_image_links_that_resolve_on_disk() {
    let tmp = tempfile::tempdir().unwrap();
    let input = tmp.path().join("with_image.docx");
    std::fs::write(&input, docx_with_one_image()).unwrap();
    let out = tmp.path().join("out");

    let status = Command::new(bin())
        .args(["convert"])
        .arg(&input)
        .arg("-o")
        .arg(&out)
        .arg("--quiet")
        .status()
        .unwrap();
    assert!(status.success(), "convert failed: {:?}", status);

    let markdown = std::fs::read_to_string(out.join("extract.md")).unwrap();
    let destinations = image_destinations(&markdown);

    // Without this the rest of the test passes vacuously on a document that rendered no
    // links at all -- the failure mode this file exists to catch would slip through.
    assert_eq!(
        destinations.len(),
        1,
        "expected exactly one image link, got {:?}\n--- markdown ---\n{}",
        destinations,
        markdown
    );

    for dest in &destinations {
        assert!(
            dest.starts_with("images/"),
            "image link must carry the subdirectory the CLI writes into, got {:?}",
            dest
        );
        let resolved = out.join(Path::new(dest));
        assert!(
            resolved.is_file(),
            "image link {:?} resolves to {}, which does not exist",
            dest,
            resolved.display()
        );
    }
}

#[test]
fn convert_without_images_renders_no_image_links() {
    let tmp = tempfile::tempdir().unwrap();
    let input = tmp.path().join("with_image.docx");
    std::fs::write(&input, docx_with_one_image()).unwrap();
    let out = tmp.path().join("out");

    let status = Command::new(bin())
        .args(["convert"])
        .arg(&input)
        .arg("-o")
        .arg(&out)
        .arg("--no-images")
        .arg("--quiet")
        .status()
        .unwrap();
    assert!(status.success(), "convert failed: {:?}", status);

    let markdown = std::fs::read_to_string(out.join("extract.md")).unwrap();
    assert_eq!(
        image_destinations(&markdown),
        Vec::<String>::new(),
        "--no-images must not leave links pointing at files it declined to write"
    );
    assert!(!out.join("images").exists());
}
