//! Footnotes are structure, and brackets in the text are text. A footnote reference and its
//! definition come out as a GFM footnote pair; text that happens to read `[x](y)` or
//! `[x]: y` stays text instead of becoming a link or a link reference definition (which is
//! not printed at all).

use std::io::{Cursor, Write};

use pulldown_cmark::{Event, Options, Parser, Tag, TagEnd};
use undoc::render::{to_markdown, RenderOptions};
use undoc::{parse_bytes, Block};
use zip::write::SimpleFileOptions;

const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>"#;

const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;

const DOC_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>"#;

const DOCUMENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t xml:space="preserve">See [x](y) here</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:t>.</w:t></w:r></w:p>
<w:p><w:r><w:t>[ref]: https://example.com is how a reference looks</w:t></w:r></w:p>
</w:body>
</w:document>"#;

const FOOTNOTES: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
<w:footnote w:id="1"><w:p><w:r><w:t>Source: [a] table</w:t></w:r></w:p></w:footnote>
</w:footnotes>"#;

fn docx() -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = SimpleFileOptions::default();
    for (name, body) in [
        ("[Content_Types].xml", CONTENT_TYPES),
        ("_rels/.rels", ROOT_RELS),
        ("word/_rels/document.xml.rels", DOC_RELS),
        ("word/document.xml", DOCUMENT),
        ("word/footnotes.xml", FOOTNOTES),
    ] {
        zip.start_file(name, opts).unwrap();
        zip.write_all(body.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

/// What a GFM reader sees: footnote references, footnote definitions (label and text),
/// links, and the plain text outside any definition.
#[derive(Default, Debug)]
struct Read {
    references: Vec<String>,
    definitions: Vec<(String, String)>,
    links: usize,
    text: String,
}

fn read(markdown: &str) -> Read {
    let mut read = Read::default();
    let mut definition: Option<(String, String)> = None;
    for event in Parser::new_ext(markdown, Options::ENABLE_FOOTNOTES) {
        match event {
            Event::FootnoteReference(label) => read.references.push(label.to_string()),
            Event::Start(Tag::FootnoteDefinition(label)) => {
                definition = Some((label.to_string(), String::new()))
            }
            Event::End(TagEnd::FootnoteDefinition) => read.definitions.extend(definition.take()),
            Event::Start(Tag::Link { .. }) => read.links += 1,
            Event::Text(text) => match definition.as_mut() {
                Some((_, body)) => body.push_str(&text),
                None => read.text.push_str(&text),
            },
            _ => {}
        }
    }
    read
}

#[test]
fn a_footnote_is_a_reference_and_a_definition() {
    let doc = parse_bytes(&docx()).unwrap();
    let notes: Vec<_> = doc.sections[0]
        .content
        .iter()
        .filter_map(|block| match block {
            Block::Note { label, content } => Some((label.clone(), content.len())),
            _ => None,
        })
        .collect();
    assert_eq!(notes, [("1".to_string(), 1)]);

    let md = to_markdown(&doc, &RenderOptions::default()).unwrap();
    let read = read(&md);
    assert_eq!(read.references, ["1"], "{md}");
    assert_eq!(
        read.definitions,
        [("1".to_string(), "Source: [a] table".to_string())],
        "{md}"
    );
}

#[test]
fn brackets_in_the_text_stay_text() {
    let md = to_markdown(&parse_bytes(&docx()).unwrap(), &RenderOptions::default()).unwrap();
    let read = read(&md);
    assert_eq!(read.links, 0, "{md}");
    assert!(read.text.contains("See [x](y) here."), "{md}");
    // A link reference definition is not printed: the whole line used to vanish.
    assert!(
        read.text
            .contains("[ref]: https://example.com is how a reference looks"),
        "{md}"
    );
}
