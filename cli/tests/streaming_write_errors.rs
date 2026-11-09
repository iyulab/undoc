//! A presentation is converted section by section, and a file the converter cannot write must
//! fail the command and say which one.
//!
//! The streaming path discarded the result of every write it made -- the JSON preamble, each
//! section, each extracted image -- and counted the image as written all the same, so a folder
//! that could not be created ended in a success exit code and a summary that named files which
//! did not exist. The deck is built here, so the test needs no fixture file.

use std::io::{Cursor, Write};
use std::process::Command;

fn bin() -> &'static str {
    env!("CARGO_BIN_EXE_undoc")
}

const PNG: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
    0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
    0x42, 0x60, 0x82,
];

const RELS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL_TYPE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn deck_with_one_picture() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>"#;
    let presentation = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>"#;
    let slide = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>Quarterly review</a:t></a:r></a:p></p:txBody></p:sp>
<p:pic><p:nvPicPr><p:cNvPr id="3" name="Chart"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdImg"/></p:blipFill><p:spPr/></p:pic>
</p:spTree></p:cSld></p:sld>"#;
    let rels = |items: &[(&str, &str, &str)]| -> String {
        let body: String = items
            .iter()
            .map(|(id, kind, target)| {
                format!(r#"<Relationship Id="{id}" Type="{REL_TYPE}/{kind}" Target="{target}"/>"#)
            })
            .collect();
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{RELS_NS}">{body}</Relationships>"#
        )
    };

    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    let parts: Vec<(&str, Vec<u8>)> = vec![
        ("[Content_Types].xml", content_types.as_bytes().to_vec()),
        (
            "_rels/.rels",
            rels(&[("rId1", "officeDocument", "ppt/presentation.xml")]).into_bytes(),
        ),
        ("ppt/presentation.xml", presentation.as_bytes().to_vec()),
        (
            "ppt/_rels/presentation.xml.rels",
            rels(&[("rId1", "slide", "slides/slide1.xml")]).into_bytes(),
        ),
        ("ppt/slides/slide1.xml", slide.as_bytes().to_vec()),
        (
            "ppt/slides/_rels/slide1.xml.rels",
            rels(&[("rIdImg", "image", "../media/image1.png")]).into_bytes(),
        ),
        ("ppt/media/image1.png", PNG.to_vec()),
    ];
    for (name, bytes) in parts {
        zip.start_file(name, options).unwrap();
        zip.write_all(&bytes).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn scratch(name: &str) -> std::path::PathBuf {
    std::env::temp_dir().join(format!("undoc-cli-stream-{}-{name}", std::process::id()))
}

#[test]
fn a_picture_folder_that_cannot_be_created_fails_the_command_and_names_it() {
    let dir = scratch("blocked");
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();
    let deck = dir.join("deck.pptx");
    std::fs::write(&deck, deck_with_one_picture()).unwrap();

    // A file where the pictures folder belongs: the folder cannot be created below it.
    let out = dir.join("out");
    std::fs::create_dir_all(&out).unwrap();
    std::fs::write(out.join("images"), b"in the way").unwrap();

    let result = Command::new(bin())
        .arg("convert")
        .arg(&deck)
        .arg("-o")
        .arg(&out)
        .arg("--quiet")
        .output()
        .expect("the CLI runs");
    let stderr = String::from_utf8_lossy(&result.stderr);
    let stdout = String::from_utf8_lossy(&result.stdout);

    assert!(
        !result.status.success(),
        "a picture that could not be written ended in success\nstdout: {stdout}\nstderr: {stderr}"
    );
    assert!(
        stderr.contains("cannot create output directory") && stderr.contains("images"),
        "the error does not name the folder: {stderr}"
    );
    std::fs::remove_dir_all(&dir).unwrap();
}

#[test]
fn a_deck_with_a_writable_output_still_converts() {
    let dir = scratch("ok");
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();
    let deck = dir.join("deck.pptx");
    std::fs::write(&deck, deck_with_one_picture()).unwrap();
    let out = dir.join("out");

    let result = Command::new(bin())
        .arg("convert")
        .arg(&deck)
        .arg("-o")
        .arg(&out)
        .arg("--quiet")
        .output()
        .expect("the CLI runs");
    assert!(
        result.status.success(),
        "{}",
        String::from_utf8_lossy(&result.stderr)
    );
    assert!(out.join("images").join("image1.png").is_file());
    std::fs::remove_dir_all(&dir).unwrap();
}
