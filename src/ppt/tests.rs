//! End-to-end tests over presentations the tests assemble themselves: a `Current User`
//! stream and a `PowerPoint Document` stream holding the document container, slide and
//! notes containers, a persist directory and a user edit, in a CFB container.

use super::PptParser;
use crate::error::Error;
use crate::model::{Block, HeadingLevel};

fn record(ver: u8, instance: u16, kind: u16, data: &[u8]) -> Vec<u8> {
    let mut out = ((instance << 4) | ver as u16).to_le_bytes().to_vec();
    out.extend(kind.to_le_bytes());
    out.extend((data.len() as u32).to_le_bytes());
    out.extend(data);
    out
}

fn container(kind: u16, instance: u16, parts: &[Vec<u8>]) -> Vec<u8> {
    record(0x0F, instance, kind, &parts.concat())
}

fn text_header(kind: u32) -> Vec<u8> {
    record(0, 0, 0x0F9F, &kind.to_le_bytes())
}

fn chars(text: &str) -> Vec<u8> {
    let data: Vec<u8> = text.encode_utf16().flat_map(|u| u.to_le_bytes()).collect();
    record(0, 0, 0x0FA0, &data)
}

fn bytes(text: &str) -> Vec<u8> {
    record(0, 0, 0x0FA8, &text.bytes().collect::<Vec<_>>())
}

fn slide_persist(persist: u32, slide_id: u32) -> Vec<u8> {
    let mut d = persist.to_le_bytes().to_vec();
    d.extend([0u8; 8]);
    d.extend(slide_id.to_le_bytes());
    d.extend([0u8; 4]);
    record(0, 0, 0x03F3, &d)
}

/// A slide: its outline texts (kind, text), the shapes on it in drawing order, and notes.
struct Slide {
    outline: Vec<(u32, &'static str)>,
    shapes: Vec<Shape>,
    notes: Option<&'static str>,
    /// Hyperlinks over outline texts: (outline index, begin, end, target).
    links: Vec<(usize, u32, u32, &'static str)>,
}

enum Shape {
    Outline(u32),
    TextBox(&'static str),
}

fn build(slides: &[Slide], encrypted: bool) -> Vec<u8> {
    // Persist ids: 1 document, then for slide i: 2 + 2i (slide), 3 + 2i (notes).
    let mut objects: Vec<(u32, Vec<u8>)> = Vec::new();
    let mut hyperlinks = Vec::new();
    let mut slide_list = Vec::new();
    let mut notes_list = Vec::new();
    for (i, slide) in slides.iter().enumerate() {
        let (slide_ref, notes_ref) = (2 + 2 * i as u32, 3 + 2 * i as u32);
        let (slide_id, notes_id) = (256 + i as u32, 512 + i as u32);
        slide_list.push(slide_persist(slide_ref, slide_id));
        for (o, (kind, text)) in slide.outline.iter().enumerate() {
            slide_list.push(text_header(*kind));
            slide_list.push(bytes(text));
            for (_, begin, end, url) in slide.links.iter().filter(|l| l.0 == o) {
                let id = hyperlinks.len() as u32 + 1;
                let target: Vec<u8> = url.encode_utf16().flat_map(|u| u.to_le_bytes()).collect();
                hyperlinks.push(container(
                    0x0FD7,
                    0,
                    &[
                        record(0, 0, 0x0FD3, &id.to_le_bytes()),
                        record(0, 1, 0x0FBA, &target),
                    ],
                ));
                let mut info = vec![0u8; 16];
                info[4..8].copy_from_slice(&id.to_le_bytes());
                slide_list.push(container(0x0FF2, 0, &[record(0, 0, 0x0FF3, &info)]));
                let mut range = begin.to_le_bytes().to_vec();
                range.extend(end.to_le_bytes());
                slide_list.push(record(0, 0, 0x0FDF, &range));
            }
        }
        let mut atom = vec![0u8; 24];
        if slide.notes.is_some() {
            atom[16..20].copy_from_slice(&notes_id.to_le_bytes());
            notes_list.push(slide_persist(notes_ref, notes_id));
        }
        let shapes: Vec<Vec<u8>> = slide
            .shapes
            .iter()
            .map(|shape| match shape {
                Shape::Outline(i) => {
                    container(0xF00D, 0, &[record(0, 0, 0x0F9E, &i.to_le_bytes())])
                }
                Shape::TextBox(text) => container(0xF00D, 0, &[text_header(4), chars(text)]),
            })
            .collect();
        let drawing = container(0xF004, 0, &shapes);
        objects.push((
            slide_ref,
            container(0x03EE, 0, &[record(0, 0, 0x03EF, &atom), drawing]),
        ));
        if let Some(notes) = slide.notes {
            let body = container(0xF00D, 0, &[text_header(2), chars(notes)]);
            objects.push((notes_ref, container(0x03F0, 0, &[body])));
        }
    }
    let document = container(
        0x03E8,
        0,
        &[
            container(0x0409, 0, &hyperlinks),
            container(0x0FF0, 0, &slide_list),
            container(0x0FF0, 2, &notes_list),
        ],
    );
    objects.insert(0, (1, document));

    let mut stream = Vec::new();
    let mut offsets = Vec::new();
    for (id, data) in &objects {
        offsets.push((*id, stream.len() as u32));
        stream.extend(data);
    }
    offsets.sort();
    let mut dir = Vec::new();
    for (id, offset) in &offsets {
        dir.extend((id | (1 << 20)).to_le_bytes());
        dir.extend(offset.to_le_bytes());
    }
    let directory_at = stream.len() as u32;
    stream.extend(record(0, 0, 0x1772, &dir));
    let edit_at = stream.len() as u32;
    let mut edit = Vec::new();
    edit.extend(0u32.to_le_bytes()); // lastSlideIdRef
    edit.extend(0u32.to_le_bytes()); // version etc.
    edit.extend(0u32.to_le_bytes()); // offsetLastEdit: none
    edit.extend(directory_at.to_le_bytes());
    edit.extend(1u32.to_le_bytes()); // docPersistIdRef
    edit.extend([0u8; 8]);
    stream.extend(record(0, 0, 0x0FF5, &edit));

    let mut user = Vec::new();
    user.extend(0x14u32.to_le_bytes());
    user.extend(
        (if encrypted {
            0xF3D1_C4DFu32
        } else {
            0xE391_C05F
        })
        .to_le_bytes(),
    );
    user.extend(edit_at.to_le_bytes());
    user.extend([0u8; 8]);
    let current_user = record(0, 0, 0x0FF6, &user);

    let mut cfb = cfb::CompoundFile::create(std::io::Cursor::new(Vec::new())).unwrap();
    for (name, data) in [
        ("/Current User", &current_user),
        ("/PowerPoint Document", &stream),
    ] {
        let mut s = cfb.create_stream(name).unwrap();
        std::io::Write::write_all(&mut s, data).unwrap();
    }
    cfb.flush().unwrap();
    cfb.into_inner().into_inner()
}

fn texts(section: &crate::Section) -> Vec<(String, HeadingLevel)> {
    section
        .content
        .iter()
        .filter_map(|b| match b {
            Block::Paragraph(p) => Some((p.plain_text(), p.heading)),
            _ => None,
        })
        .collect()
}

#[test]
fn slides_read_in_order_with_titles_as_headings_and_text_boxes_in_drawing_order() {
    let doc = PptParser::from_bytes(build(
        &[
            Slide {
                outline: vec![(6, "Quarterly review"), (5, "Prepared for the board")],
                shapes: vec![Shape::Outline(0), Shape::Outline(1)],
                notes: None,
                links: vec![],
            },
            Slide {
                outline: vec![(0, "Results"), (1, "Revenue grew\rCosts fell")],
                shapes: vec![
                    Shape::Outline(0),
                    Shape::TextBox("Source: audit"),
                    Shape::Outline(1),
                ],
                notes: Some("Mention the one-off."),
                links: vec![],
            },
        ],
        false,
    ))
    .unwrap()
    .parse()
    .unwrap();

    assert_eq!(doc.format, crate::FormatType::Ppt);
    assert_eq!(doc.sections.len(), 2);
    assert_eq!(doc.sections[0].name.as_deref(), Some("Slide 1"));
    assert_eq!(
        texts(&doc.sections[0]),
        [
            ("Quarterly review".into(), HeadingLevel::H1),
            ("Prepared for the board".into(), HeadingLevel::H2),
        ]
    );
    assert_eq!(
        texts(&doc.sections[1]),
        [
            ("Results".into(), HeadingLevel::H1),
            ("Source: audit".into(), HeadingLevel::None),
            ("Revenue grew".into(), HeadingLevel::None),
            ("Costs fell".into(), HeadingLevel::None),
        ]
    );
    let notes = doc.sections[1].notes.as_ref().expect("notes");
    assert_eq!(notes[0].plain_text(), "Mention the one-off.");
    assert!(doc.sections[0].notes.is_none());
}

#[test]
fn outline_text_no_shape_references_is_still_read() {
    let doc = PptParser::from_bytes(build(
        &[Slide {
            outline: vec![(0, "Title only in the outline")],
            shapes: vec![],
            notes: None,
            links: vec![],
        }],
        false,
    ))
    .unwrap()
    .parse()
    .unwrap();
    assert_eq!(texts(&doc.sections[0])[0].0, "Title only in the outline");
}

#[test]
fn a_vertical_tab_breaks_a_line_within_a_paragraph() {
    let doc = PptParser::from_bytes(build(
        &[Slide {
            outline: vec![],
            shapes: vec![Shape::TextBox("one\u{0B}two")],
            notes: None,
            links: vec![],
        }],
        false,
    ))
    .unwrap()
    .parse()
    .unwrap();
    let Block::Paragraph(p) = &doc.sections[0].content[0] else {
        panic!()
    };
    assert_eq!(p.runs.len(), 2);
    assert!(p.runs[0].line_break);
    assert_eq!(
        (p.runs[0].text.as_str(), p.runs[1].text.as_str()),
        ("one", "two")
    );
}

#[test]
fn an_encrypted_presentation_is_reported_as_encrypted() {
    let err = PptParser::from_bytes(build(&[], true))
        .unwrap()
        .parse()
        .unwrap_err();
    assert!(matches!(err, Error::Encrypted), "{err}");
}

#[test]
fn the_library_entry_points_read_a_ppt() {
    let bytes = build(
        &[Slide {
            outline: vec![(0, "Agenda")],
            shapes: vec![Shape::Outline(0)],
            notes: None,
            links: vec![],
        }],
        false,
    );
    let doc = crate::parse_bytes(&bytes).unwrap();
    let md = crate::render::to_markdown(&doc, &crate::render::RenderOptions::default()).unwrap();
    assert!(md.contains("# Agenda"), "{md}");
}

#[test]
fn a_linked_range_becomes_a_link_run() {
    let doc = PptParser::from_bytes(build(
        &[Slide {
            outline: vec![(1, "See the site now\rNext")],
            shapes: vec![Shape::Outline(0)],
            notes: None,
            links: vec![(0, 4, 12, "https://example.com/")],
        }],
        false,
    ))
    .unwrap()
    .parse()
    .unwrap();
    let Block::Paragraph(p) = &doc.sections[0].content[0] else {
        panic!()
    };
    let runs: Vec<(&str, Option<&str>)> = p
        .runs
        .iter()
        .map(|r| (r.text.as_str(), r.hyperlink.as_deref()))
        .collect();
    assert_eq!(
        runs,
        [
            ("See ", None),
            ("the site", Some("https://example.com/")),
            (" now", None)
        ]
    );
    assert_eq!(texts(&doc.sections[0])[1].0, "Next");
}
