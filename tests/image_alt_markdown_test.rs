//! A picture's description is document data, not Markdown. Whatever a person typed into the
//! description box -- several lines, blank lines, brackets, backslashes -- must come out as
//! one image whose alt text is that description, never as an image that fell apart into
//! literal text plus a stray paragraph.

use std::io::{Cursor, Write};

use pulldown_cmark::{Event, Parser, Tag, TagEnd};
use undoc::render::{to_markdown, RenderOptions};
use undoc::{parse_bytes, Block};
use zip::write::SimpleFileOptions;

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

const DOC_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>"#;

/// A one-paragraph DOCX holding one inline picture; `descr` is the raw (already
/// XML-escaped) attribute value.
fn docx_with_picture(descr: &str) -> Vec<u8> {
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body><w:p><w:r><w:drawing><wp:inline>
<wp:docPr id="1" name="Picture 1" descr="{descr}"/>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
<pic:pic><pic:blipFill><a:blip r:embed="rId5"/></pic:blipFill></pic:pic>
</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>"#
    );
    let png: &[u8] = &[
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 0, 0, 0, 0, 0,
    ];
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, bytes) in [
        ("[Content_Types].xml", CONTENT_TYPES.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/_rels/document.xml.rels", DOC_RELS.as_bytes()),
        ("word/document.xml", document.as_bytes()),
        ("word/media/image1.png", png),
    ] {
        zip.start_file(name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

/// Everything a CommonMark reader sees: the alt text of each image, and whether any text
/// or paragraph exists outside the single paragraph holding the image.
struct Read {
    images: Vec<String>,
    paragraphs: usize,
    text_outside_alt: String,
}

fn read(md: &str) -> Read {
    let mut r = Read {
        images: Vec::new(),
        paragraphs: 0,
        text_outside_alt: String::new(),
    };
    let mut in_image = false;
    for ev in Parser::new(md) {
        match ev {
            Event::Start(Tag::Image { .. }) => {
                in_image = true;
                r.images.push(String::new());
            }
            Event::End(TagEnd::Image) => in_image = false,
            Event::Start(Tag::Paragraph) => r.paragraphs += 1,
            Event::Text(t) | Event::Code(t) => {
                if in_image {
                    r.images.last_mut().unwrap().push_str(&t);
                } else {
                    r.text_outside_alt.push_str(&t);
                }
            }
            Event::SoftBreak | Event::HardBreak if in_image => {
                r.images.last_mut().unwrap().push(' ')
            }
            _ => {}
        }
    }
    r
}

fn markdown_of(descr: &str) -> String {
    let doc = parse_bytes(&docx_with_picture(descr)).expect("synthetic DOCX parses");
    to_markdown(&doc, &RenderOptions::default()).expect("renders")
}

#[test]
fn a_description_with_a_blank_line_stays_one_image() {
    // `&#10;` is a newline kept by the XML parser (a literal one would be normalised away).
    let md = markdown_of("First part.&#10;&#10;Second part after a blank line.");
    let r = read(&md);
    assert_eq!(
        r.images,
        vec!["First part. Second part after a blank line."],
        "{md:?}"
    );
    assert_eq!(r.paragraphs, 1, "stray paragraph in {md:?}");
    assert_eq!(r.text_outside_alt, "", "{md:?}");
}

#[test]
fn brackets_backslashes_and_block_markers_in_a_description_stay_one_image() {
    let md =
        markdown_of(r"a ] b [c] \ d&#10;# not a heading&#10;- not a list&#10;&gt; not a quote");
    let r = read(&md);
    assert_eq!(
        r.images,
        vec![r"a ] b [c] \ d # not a heading - not a list > not a quote"],
        "{md:?}"
    );
    assert_eq!(r.paragraphs, 1, "{md:?}");
    assert_eq!(r.text_outside_alt, "", "{md:?}");
}

#[test]
fn the_model_and_json_keep_the_description_unchanged() {
    let doc = parse_bytes(&docx_with_picture("one&#10;&#10;two [x]")).unwrap();
    let alt = doc
        .sections
        .iter()
        .flat_map(|s| s.content.iter())
        .find_map(|b| match b {
            Block::Paragraph(p) => p.images.first().and_then(|i| i.alt_text.clone()),
            Block::Image { alt_text, .. } => alt_text.clone(),
            _ => None,
        })
        .expect("the picture carries its description");
    assert_eq!(alt, "one\n\ntwo [x]");
    let json = doc.to_json().unwrap();
    assert!(json.contains(r"one\n\ntwo [x]"), "{json}");
}
