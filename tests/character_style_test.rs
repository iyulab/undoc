//! A run's character style (`w:rStyle`) carries formatting: Word's Strong and Emphasis make
//! text bold and italic, pandoc's VerbatimChar and Word's HTML Code make it code. The run's
//! own properties still win over the style's. Checked in body text and in a table cell — the
//! two places the reader parses runs.

use std::io::{Cursor, Write};

use undoc::parse_bytes;
use undoc::render::{to_markdown, RenderOptions};
use zip::write::SimpleFileOptions;

const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"#;

const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;

const DOC_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#;

const STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>
<w:style w:type="character" w:styleId="Emphasis"><w:name w:val="Emphasis"/><w:rPr><w:i/></w:rPr></w:style>
<w:style w:type="character" w:styleId="VerbatimChar"><w:name w:val="Verbatim Char"/><w:rPr><w:rFonts w:ascii="Consolas"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="MyCode"><w:name w:val="My Code"/><w:basedOn w:val="VerbatimChar"/></w:style>
</w:styles>"#;

fn run(style: &str, text: &str, own: &str) -> String {
    format!(
        r#"<w:r><w:rPr><w:rStyle w:val="{style}"/>{own}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>"#
    )
}

fn docx() -> Vec<u8> {
    let body = format!(
        r#"<w:p>{}{}{}{}{}{}</w:p>"#,
        run("Strong", "strong", ""),
        r#"<w:r><w:t xml:space="preserve"> and </w:t></w:r>"#,
        run("Emphasis", "emphasis", ""),
        r#"<w:r><w:t xml:space="preserve"> and </w:t></w:r>"#,
        run("MyCode", r"C:\dir", ""),
        run("Strong", " not bold", r#"<w:b w:val="0"/>"#),
    );
    let cell = format!(
        r#"<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Name</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p>{}</w:p></w:tc><w:tc><w:p><w:r><w:t>plain</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
        run("Strong", "cell", "")
    );
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}{cell}</w:body></w:document>"#
    );
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = SimpleFileOptions::default();
    for (name, data) in [
        ("[Content_Types].xml", CONTENT_TYPES),
        ("_rels/.rels", ROOT_RELS),
        ("word/_rels/document.xml.rels", DOC_RELS),
        ("word/styles.xml", STYLES),
        ("word/document.xml", document.as_str()),
    ] {
        zip.start_file(name, opts).unwrap();
        zip.write_all(data.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

#[test]
fn character_styles_format_their_runs() {
    let md = to_markdown(&parse_bytes(&docx()).unwrap(), &RenderOptions::default()).unwrap();
    // The last run names Strong, but its own `w:b w:val="0"` overrides the style's bold.
    assert!(
        md.contains("**strong** and *emphasis* and `C:\\dir` not bold"),
        "{md}"
    );
}

#[test]
fn a_character_style_in_a_table_cell_formats_its_run_too() {
    let md = to_markdown(&parse_bytes(&docx()).unwrap(), &RenderOptions::default()).unwrap();
    assert!(md.contains("| **cell** | plain |"), "{md}");
}
