//! Without `-o`, `convert` writes beside the input into a directory named by the
//! input's stem *and* extension.
//!
//! With the stem alone, `A2.docx` and `A2.pptx` shared one directory and the second
//! conversion silently overwrote the first one's files. The DOCX is built here so the
//! test carries its own fixture.

use std::io::{Cursor, Write};
use std::process::Command;

fn minimal_docx() -> Vec<u8> {
    const PARTS: [(&str, &str); 3] = [
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#,
        ),
        (
            "_rels/.rels",
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
        ),
        (
            "word/document.xml",
            r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>Hello.</w:t></w:r></w:p></w:body>
</w:document>"#,
        ),
    ];
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, body) in PARTS {
        zip.start_file(name, options).unwrap();
        zip.write_all(body.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

#[test]
fn convert_without_output_writes_next_to_the_input_named_by_its_extension() {
    let tmp = tempfile::tempdir().unwrap();
    let input = tmp.path().join("A2.docx");
    std::fs::write(&input, minimal_docx()).unwrap();
    let status = Command::new(env!("CARGO_BIN_EXE_undoc"))
        .args(["convert"])
        .arg(&input)
        .arg("--quiet")
        .current_dir(std::env::temp_dir())
        .status()
        .unwrap();
    assert!(status.success(), "convert failed: {status:?}");
    assert!(
        tmp.path()
            .join("A2_docx_output")
            .join("extract.md")
            .exists(),
        "default output must sit beside the input as A2_docx_output/"
    );
    assert!(!tmp.path().join("A2_output").exists());
}
