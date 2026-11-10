//! End-to-end tests over presentations the tests assemble themselves: a `Current User`
//! stream and a `PowerPoint Document` stream holding the document container, slide and
//! notes containers, a persist directory and a user edit, in a CFB container.

use super::PptParser;
use crate::error::Error;
use crate::model::{Block, HeadingLevel};
use crate::officeart::tests::{fbse, fopt, png_blip};

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
    /// A picture shape: its `pib` (1-based blip store index) and alternative text.
    Picture(u32, Option<&'static str>),
    /// A group of shapes, a table when `table` (its `tableProperties` flags and the number of
    /// rows it states) is given: the cells are rectangles `(left, top, right, bottom, text)`,
    /// stored in the order given.
    Group {
        table: Option<(u32, usize)>,
        cells: Vec<(i32, i32, i32, i32, &'static str)>,
    },
}

/// The tertiary property table of a table group's own shape: `tableProperties` and the array
/// `tableRowProperties` with one element per row.
fn table_properties(flags: u32, rows: usize) -> Vec<u8> {
    let mut array = (rows as u16).to_le_bytes().to_vec();
    array.extend((rows as u16).to_le_bytes());
    array.extend(4u16.to_le_bytes());
    for _ in 0..rows {
        array.extend(100u32.to_le_bytes());
    }
    let mut table = 0x039Fu16.to_le_bytes().to_vec();
    table.extend(flags.to_le_bytes());
    table.extend((0x03A0u16 | 0x8000).to_le_bytes());
    table.extend((array.len() as u32).to_le_bytes());
    table.extend(array);
    record(3, 2, 0xF122, &table)
}

/// An entry of the blip store the document's drawing group lists.
enum Entry {
    /// A picture kept in the `Pictures` stream, found through `foDelay`.
    Png(&'static [u8]),
    /// A slot the presentation emptied.
    Empty,
}

fn build(slides: &[Slide], encrypted: bool) -> Vec<u8> {
    build_with(slides, encrypted, &[])
}

/// The presentation, with a drawing group whose blip store holds `store` and the `Pictures`
/// stream the entries point into.
fn build_with(slides: &[Slide], encrypted: bool, store: &[Entry]) -> Vec<u8> {
    let mut pictures = Vec::new();
    let mut fbses = Vec::new();
    for entry in store {
        match entry {
            Entry::Png(png) => {
                let blip = png_blip(png);
                fbses.push(fbse(blip.len() as u32, pictures.len() as u32, &[]));
                pictures.extend(blip);
            }
            Entry::Empty => fbses.push(fbse(0, 0, &[])),
        }
    }
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
            .map(|shape| {
                // A shape is a container that starts with its `FSP` (type 75: a picture frame).
                let fsp = |kind| record(2, kind, 0xF00A, &[0; 8]);
                match shape {
                    Shape::Outline(i) => container(
                        0xF004,
                        0,
                        &[
                            fsp(202),
                            container(0xF00D, 0, &[record(0, 0, 0x0F9E, &i.to_le_bytes())]),
                        ],
                    ),
                    Shape::TextBox(text) => container(
                        0xF004,
                        0,
                        &[
                            fsp(202),
                            container(0xF00D, 0, &[text_header(4), chars(text)]),
                        ],
                    ),
                    Shape::Picture(pib, alt) => {
                        container(0xF004, 0, &[fsp(75), fopt(Some(*pib), *alt)])
                    }
                    Shape::Group { table, cells } => {
                        // The group's own shape first, then its children.
                        let mut own = vec![fsp(0)];
                        if let Some((flags, rows)) = table {
                            own.push(table_properties(*flags, *rows));
                        }
                        let mut parts = vec![container(0xF004, 0, &own)];
                        for (left, top, right, bottom, text) in cells {
                            let anchor: Vec<u8> = [left, top, right, bottom]
                                .into_iter()
                                .flat_map(|v| v.to_le_bytes())
                                .collect();
                            parts.push(container(
                                0xF004,
                                0,
                                &[
                                    fsp(1),
                                    record(0, 0, 0xF00F, &anchor),
                                    container(0xF00D, 0, &[text_header(4), chars(text)]),
                                ],
                            ));
                        }
                        container(0xF003, 0, &parts)
                    }
                }
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
    let mut parts = vec![
        container(0x0409, 0, &hyperlinks),
        container(0x0FF0, 0, &slide_list),
        container(0x0FF0, 2, &notes_list),
    ];
    if !store.is_empty() {
        let bstore = container(0xF001, store.len() as u16, &fbses);
        parts.push(container(0x040B, 0, &[container(0xF000, 0, &[bstore])]));
    }
    let document = container(0x03E8, 0, &parts);
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
    let mut streams = vec![
        ("/Current User", &current_user),
        ("/PowerPoint Document", &stream),
    ];
    if !pictures.is_empty() {
        streams.push(("/Pictures", &pictures));
    }
    for (name, data) in streams {
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

const PNG_A: &[u8] = b"\x89PNG\r\n\x1a\nfirst";
const PNG_B: &[u8] = b"\x89PNG\r\n\x1a\nsecond";

/// The image blocks of a section: (resource id, alt text).
fn images(section: &crate::Section) -> Vec<(String, Option<String>)> {
    section
        .content
        .iter()
        .filter_map(|b| match b {
            Block::Image {
                resource_id,
                alt_text,
                ..
            } => Some((resource_id.clone(), alt_text.clone())),
            _ => None,
        })
        .collect()
}

#[test]
fn a_picture_shape_becomes_an_image_between_the_texts_drawn_around_it() {
    let doc = PptParser::from_bytes(build_with(
        &[Slide {
            outline: vec![],
            shapes: vec![
                Shape::TextBox("before"),
                Shape::Picture(2, Some("A bar chart")),
                Shape::TextBox("after"),
            ],
            notes: None,
            links: vec![],
        }],
        false,
        &[Entry::Png(PNG_A), Entry::Png(PNG_B)],
    ))
    .unwrap()
    .parse()
    .unwrap();

    let kinds: Vec<&str> = doc.sections[0]
        .content
        .iter()
        .map(|b| match b {
            Block::Paragraph(_) => "text",
            Block::Image { .. } => "image",
            _ => "other",
        })
        .collect();
    assert_eq!(kinds, ["text", "image", "text"]);
    // `pib` 2 is the second entry of the store.
    let [(id, alt)] = &images(&doc.sections[0])[..] else {
        panic!()
    };
    assert_eq!(alt.as_deref(), Some("A bar chart"));
    let resource = doc.resources.get(id).expect("the resource");
    assert_eq!(resource.data, PNG_B);
    assert_eq!(resource.filename.as_deref(), Some("image1.png"));
}

#[test]
fn a_picture_on_several_slides_is_one_resource() {
    let slide = |alt| Slide {
        outline: vec![],
        shapes: vec![Shape::Picture(1, alt)],
        notes: None,
        links: vec![],
    };
    let doc = PptParser::from_bytes(build_with(
        &[slide(Some("Logo")), slide(None)],
        false,
        &[Entry::Png(PNG_A)],
    ))
    .unwrap()
    .parse()
    .unwrap();
    let (first, second) = (images(&doc.sections[0]), images(&doc.sections[1]));
    assert_eq!(first[0].0, second[0].0);
    assert_eq!(doc.resources.len(), 1);
    // The alternative text belongs to the shape, not to the picture.
    assert_eq!(
        (first[0].1.as_deref(), second[0].1.as_deref()),
        (Some("Logo"), None)
    );
}

#[test]
fn an_emptied_slot_or_a_missing_entry_shows_no_picture() {
    // `pib` 1 is an emptied slot (not the first blip), 3 is past the store, 0 names nothing.
    let doc = PptParser::from_bytes(build_with(
        &[Slide {
            outline: vec![],
            shapes: vec![
                Shape::Picture(1, None),
                Shape::Picture(3, None),
                Shape::Picture(0, None),
                Shape::TextBox("still read"),
            ],
            notes: None,
            links: vec![],
        }],
        false,
        &[Entry::Empty, Entry::Png(PNG_A)],
    ))
    .unwrap()
    .parse()
    .unwrap();
    assert!(images(&doc.sections[0]).is_empty());
    assert!(doc.resources.is_empty());
    assert_eq!(texts(&doc.sections[0])[0].0, "still read");
}

#[test]
fn a_picture_renders_as_an_image_in_markdown() {
    let bytes = build_with(
        &[Slide {
            outline: vec![],
            shapes: vec![Shape::Picture(1, Some("Org chart"))],
            notes: None,
            links: vec![],
        }],
        false,
        &[Entry::Png(PNG_A)],
    );
    let doc = crate::parse_bytes(&bytes).unwrap();
    let md = crate::render::to_markdown(&doc, &crate::render::RenderOptions::default()).unwrap();
    assert!(md.contains("![Org chart](image1.png)"), "{md}");
}

/// A row as read back: whether it is a header, and its cells as (text, row span, col span).
type ReadRow = (bool, Vec<(String, u32, u32)>);

/// The table of a section.
fn table_of(section: &crate::Section) -> Vec<ReadRow> {
    let Some(Block::Table(table)) = section
        .content
        .iter()
        .find(|b| matches!(b, Block::Table(_)))
    else {
        panic!("no table")
    };
    table
        .rows
        .iter()
        .map(|row| {
            (
                row.is_header,
                row.cells
                    .iter()
                    .map(|c| {
                        let text: Vec<String> = c.content.iter().map(|p| p.plain_text()).collect();
                        (text.join("|"), c.row_span, c.col_span)
                    })
                    .collect(),
            )
        })
        .collect()
}

fn slide_with(shapes: Vec<Shape>) -> Slide {
    Slide {
        outline: vec![],
        shapes,
        notes: None,
        links: vec![],
    }
}

fn parse_slide(shapes: Vec<Shape>) -> crate::Document {
    PptParser::from_bytes(build(&[slide_with(shapes)], false))
        .unwrap()
        .parse()
        .unwrap()
}

#[test]
fn a_group_that_states_table_properties_is_a_table_read_by_position() {
    // Two rows of three cells, stored last row and right-most cell first — the order the
    // sample files keep them in. The text boxes around the table keep their place.
    let doc = parse_slide(vec![
        Shape::TextBox("before"),
        Shape::Group {
            table: Some((3, 2)),
            cells: vec![
                (200, 100, 300, 200, "F"),
                (100, 100, 200, 200, "E"),
                (0, 100, 100, 200, "D"),
                (200, 0, 300, 100, "C"),
                (100, 0, 200, 100, "B"),
                (0, 0, 100, 100, "A"),
            ],
        },
        Shape::TextBox("after"),
    ]);
    let kinds: Vec<&str> = doc.sections[0]
        .content
        .iter()
        .map(|b| match b {
            Block::Paragraph(_) => "text",
            Block::Table(_) => "table",
            _ => "other",
        })
        .collect();
    assert_eq!(kinds, ["text", "table", "text"]);
    let row = |cells: [&str; 3]| cells.map(|t| (t.to_string(), 1, 1)).to_vec();
    assert_eq!(
        table_of(&doc.sections[0]),
        [(true, row(["A", "B", "C"])), (false, row(["D", "E", "F"]))]
    );
}

#[test]
fn a_cell_that_covers_several_grid_positions_is_one_cell_with_spans() {
    // A wide cell over two columns in the first row, a tall one over two rows in the last column.
    let doc = parse_slide(vec![Shape::Group {
        table: Some((1, 2)),
        cells: vec![
            (0, 0, 200, 100, "wide"),
            (200, 0, 300, 200, "tall"),
            (0, 100, 100, 200, "x"),
            (100, 100, 200, 200, "y"),
        ],
    }]);
    assert_eq!(
        table_of(&doc.sections[0]),
        [
            (
                false,
                vec![("wide".to_string(), 1, 2), ("tall".to_string(), 2, 1)]
            ),
            (
                false,
                vec![("x".to_string(), 1, 1), ("y".to_string(), 1, 1)]
            ),
        ]
    );
}

#[test]
fn cells_that_do_not_tile_a_grid_are_read_as_the_texts_they_are() {
    // The bottom-right position has no cell, and a group of two rows states three: neither is
    // a table, and neither loses a word.
    for (rows, cells) in [
        (
            2,
            vec![
                (0, 0, 100, 100, "a"),
                (100, 0, 200, 100, "b"),
                (0, 100, 100, 200, "c"),
            ],
        ),
        (
            3,
            vec![
                (0, 0, 100, 100, "a"),
                (100, 0, 200, 100, "b"),
                (0, 100, 100, 200, "c"),
                (100, 100, 200, 200, "d"),
            ],
        ),
    ] {
        let doc = parse_slide(vec![Shape::Group {
            table: Some((1, rows)),
            cells,
        }]);
        assert!(!doc.sections[0]
            .content
            .iter()
            .any(|b| matches!(b, Block::Table(_))));
        let words: Vec<String> = texts(&doc.sections[0]).into_iter().map(|t| t.0).collect();
        assert!(words.len() >= 3, "{words:?}");
        assert!(words.iter().any(|w| w == "a"));
    }
}

#[test]
fn a_group_without_table_properties_is_not_a_table() {
    let doc = parse_slide(vec![Shape::Group {
        table: None,
        cells: vec![(0, 0, 100, 100, "one"), (100, 0, 200, 100, "two")],
    }]);
    assert!(!doc.sections[0]
        .content
        .iter()
        .any(|b| matches!(b, Block::Table(_))));
    let words: Vec<String> = texts(&doc.sections[0]).into_iter().map(|t| t.0).collect();
    assert_eq!(words, ["one", "two"]);
}

#[test]
fn a_table_renders_as_a_markdown_table() {
    let bytes = build(
        &[slide_with(vec![Shape::Group {
            table: Some((3, 2)),
            cells: vec![
                (100, 100, 200, 200, "d"),
                (0, 100, 100, 200, "c"),
                (100, 0, 200, 100, "b"),
                (0, 0, 100, 100, "a"),
            ],
        }])],
        false,
    );
    let doc = crate::parse_bytes(&bytes).unwrap();
    let md = crate::render::to_markdown(&doc, &crate::render::RenderOptions::default()).unwrap();
    assert!(md.contains("| a | b |\n| --- | --- |\n| c | d |"), "{md}");
}
