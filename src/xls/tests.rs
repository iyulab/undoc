//! End-to-end tests over BIFF8 workbooks the tests assemble themselves: workbook globals and
//! worksheet substreams in a `Workbook` stream, in a CFB container.

use super::XlsParser;
use crate::error::Error;
use crate::model::Block;

pub(super) fn rec(out: &mut Vec<u8>, kind: u16, data: &[u8]) {
    out.extend_from_slice(&kind.to_le_bytes());
    out.extend_from_slice(&(data.len() as u16).to_le_bytes());
    out.extend_from_slice(data);
}

fn bof(kind: u16) -> Vec<u8> {
    let mut d = vec![0u8; 16];
    d[0..2].copy_from_slice(&0x0600u16.to_le_bytes());
    d[2..4].copy_from_slice(&kind.to_le_bytes());
    d
}

/// `XLUnicodeString`, compressed when every character fits in a byte.
fn xl_string(s: &str) -> Vec<u8> {
    let units: Vec<u16> = s.encode_utf16().collect();
    let mut d = (units.len() as u16).to_le_bytes().to_vec();
    if units.iter().all(|&u| u < 0x100) {
        d.push(0);
        d.extend(units.iter().map(|&u| u as u8));
    } else {
        d.push(1);
        d.extend(units.iter().flat_map(|u| u.to_le_bytes()));
    }
    d
}

fn cell_head(row: u16, col: u16, xf: u16) -> Vec<u8> {
    [row, col, xf]
        .iter()
        .flat_map(|v| v.to_le_bytes())
        .collect()
}

/// A cell record of a worksheet.
enum C {
    Sst(u16, u16, u32),
    Number(u16, u16, u16, f64),
    Rk(u16, u16, u16, u32),
    MulRk(u16, u16, Vec<(u16, u32)>),
    FormulaString(u16, u16, &'static str),
    Bool(u16, u16, bool),
    Error(u16, u16, u8),
    Merge(u16, u16, u16, u16),
    Link(u16, u16, &'static str),
    Comment(u16, u16, u16, &'static str),
    /// A cell with rich text runs (`RSTRING`).
    Rich(u16, u16, &'static str),
    /// A chart embedded in the sheet: a substream of its own, BOF to EOF.
    EmbeddedChart,
}

struct Sheet {
    name: &'static str,
    /// BOUNDSHEET8 `dt`: 0 worksheet, 2 chart.
    kind: u8,
    cells: Vec<C>,
}

#[derive(Default)]
struct Book {
    sheets: Vec<Sheet>,
    strings: Vec<&'static str>,
    /// Split the shared string table into a CONTINUE after this many bytes of its data.
    split_sst_at: Option<usize>,
    encrypted: bool,
    biff_version: Option<u16>,
}

fn build(book: Book) -> Vec<u8> {
    let mut globals = Vec::new();
    let mut first = bof(0x0005);
    if let Some(v) = book.biff_version {
        first[0..2].copy_from_slice(&v.to_le_bytes());
    }
    rec(&mut globals, 0x0809, &first);
    if book.encrypted {
        rec(&mut globals, 0x002F, &[0, 0]);
    }
    let mut format = 164u16.to_le_bytes().to_vec();
    format.extend(xl_string("yyyy\\-mm\\-dd"));
    rec(&mut globals, 0x041E, &format);
    // XF 0 general, XF 1 built-in date (14), XF 2 the custom date format.
    for ifmt in [0u16, 14, 164] {
        let mut xf = vec![0u8; 20];
        xf[2..4].copy_from_slice(&ifmt.to_le_bytes());
        rec(&mut globals, 0x00E0, &xf);
    }
    let mut patch_at = Vec::new();
    for sheet in &book.sheets {
        let mut d = vec![0u8; 4];
        d.push(0);
        d.push(sheet.kind);
        let name: Vec<u16> = sheet.name.encode_utf16().collect();
        d.push(name.len() as u8);
        d.push(0);
        d.extend(name.iter().map(|&u| u as u8));
        patch_at.push(globals.len() + 4);
        rec(&mut globals, 0x0085, &d);
    }
    let mut sst = Vec::new();
    sst.extend((book.strings.len() as u32).to_le_bytes());
    sst.extend((book.strings.len() as u32).to_le_bytes());
    for s in &book.strings {
        sst.extend(xl_string(s));
    }
    match book.split_sst_at {
        Some(at) => {
            rec(&mut globals, 0x00FC, &sst[..at]);
            // A string split mid-characters restarts with its flags byte.
            rec(&mut globals, 0x003C, &[&[0x01][..], &sst[at..]].concat());
        }
        None => rec(&mut globals, 0x00FC, &sst),
    }
    rec(&mut globals, 0x000A, &[]);

    let mut stream = globals;
    for (i, sheet) in book.sheets.iter().enumerate() {
        let offset = stream.len() as u32;
        stream[patch_at[i]..patch_at[i] + 4].copy_from_slice(&offset.to_le_bytes());
        rec(&mut stream, 0x0809, &bof(0x0010));
        let mut merges = Vec::new();
        let mut notes = Vec::new();
        for c in &sheet.cells {
            match c {
                C::Sst(r, col, i) => {
                    let mut d = cell_head(*r, *col, 0);
                    d.extend(i.to_le_bytes());
                    rec(&mut stream, 0x00FD, &d);
                }
                C::Number(r, col, xf, v) => {
                    let mut d = cell_head(*r, *col, *xf);
                    d.extend(v.to_le_bytes());
                    rec(&mut stream, 0x0203, &d);
                }
                C::Rk(r, col, xf, rk) => {
                    let mut d = cell_head(*r, *col, *xf);
                    d.extend(rk.to_le_bytes());
                    rec(&mut stream, 0x027E, &d);
                }
                C::MulRk(r, first, values) => {
                    let mut d: Vec<u8> =
                        [*r, *first].iter().flat_map(|v| v.to_le_bytes()).collect();
                    for (xf, rk) in values {
                        d.extend(xf.to_le_bytes());
                        d.extend(rk.to_le_bytes());
                    }
                    d.extend((first + values.len() as u16 - 1).to_le_bytes());
                    rec(&mut stream, 0x00BD, &d);
                }
                C::FormulaString(r, col, text) => {
                    let mut d = cell_head(*r, *col, 0);
                    d.extend([0, 0, 0, 0, 0, 0, 0xFF, 0xFF]);
                    d.extend([0u8; 6]);
                    rec(&mut stream, 0x0006, &d);
                    rec(&mut stream, 0x0207, &xl_string(text));
                }
                C::Bool(r, col, v) => {
                    let mut d = cell_head(*r, *col, 0);
                    d.extend([*v as u8, 0]);
                    rec(&mut stream, 0x0205, &d);
                }
                C::Error(r, col, code) => {
                    let mut d = cell_head(*r, *col, 0);
                    d.extend([*code, 1]);
                    rec(&mut stream, 0x0205, &d);
                }
                C::Rich(r, col, text) => {
                    let mut d = cell_head(*r, *col, 0);
                    d.extend(xl_string(text));
                    d.extend(1u16.to_le_bytes()); // one formatting run
                    d.extend([0u8; 4]);
                    rec(&mut stream, 0x00D6, &d);
                }
                C::EmbeddedChart => {
                    rec(&mut stream, 0x0809, &bof(0x0020));
                    rec(&mut stream, 0x1002, &[0u8; 16]); // CHART
                    rec(&mut stream, 0x000A, &[]);
                }
                C::Merge(top, bottom, left, right) => merges.push([*top, *bottom, *left, *right]),
                C::Link(r, col, url) => {
                    let mut d: Vec<u8> = [*r, *r, *col, *col]
                        .iter()
                        .flat_map(|v| v.to_le_bytes())
                        .collect();
                    d.extend([0u8; 16]); // hlinkClsid
                    d.extend(2u32.to_le_bytes());
                    d.extend(0x03u32.to_le_bytes()); // has moniker, absolute
                    d.extend([
                        0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA,
                        0x00, 0x4B, 0xA9, 0x0B,
                    ]);
                    let units: Vec<u16> = url.encode_utf16().chain([0]).collect();
                    d.extend(((units.len() * 2) as u32).to_le_bytes());
                    d.extend(units.iter().flat_map(|u| u.to_le_bytes()));
                    rec(&mut stream, 0x01B8, &d);
                }
                C::Comment(r, col, id, text) => {
                    let mut obj: Vec<u8> = [0x15u16, 0x12, 0x19, *id]
                        .iter()
                        .flat_map(|v| v.to_le_bytes())
                        .collect();
                    obj.extend([0u8; 14 + 4]);
                    rec(&mut stream, 0x005D, &obj);
                    let mut txo = vec![0u8; 18];
                    txo[10..12].copy_from_slice(&(text.len() as u16).to_le_bytes());
                    rec(&mut stream, 0x01B6, &txo);
                    let mut chars = vec![0u8];
                    chars.extend(text.bytes());
                    rec(&mut stream, 0x003C, &chars);
                    rec(&mut stream, 0x003C, &[0u8; 16]);
                    notes.push((*r, *col, *id));
                }
            }
        }
        for (r, col, id) in &notes {
            let mut d: Vec<u8> = [*r, *col, 0, *id]
                .iter()
                .flat_map(|v| v.to_le_bytes())
                .collect();
            d.extend(xl_string("A. Author"));
            rec(&mut stream, 0x001C, &d);
        }
        if !merges.is_empty() {
            let mut d = (merges.len() as u16).to_le_bytes().to_vec();
            for m in &merges {
                d.extend(m.iter().flat_map(|v| v.to_le_bytes()));
            }
            rec(&mut stream, 0x00E5, &d);
        }
        rec(&mut stream, 0x000A, &[]);
    }

    let mut container = cfb::CompoundFile::create(std::io::Cursor::new(Vec::new())).unwrap();
    let mut s = container.create_stream("/Workbook").unwrap();
    std::io::Write::write_all(&mut s, &stream).unwrap();
    drop(s);
    container.flush().unwrap();
    container.into_inner().into_inner()
}

fn parse(book: Book) -> crate::Result<crate::Document> {
    XlsParser::from_bytes(build(book))?.parse()
}

pub(super) fn grid(doc: &crate::Document, sheet: usize) -> Vec<Vec<String>> {
    let Some(Block::Table(table)) = doc.sections[sheet].content.first() else {
        panic!(
            "sheet {sheet} has no table: {:?}",
            doc.sections[sheet].content
        );
    };
    table
        .rows
        .iter()
        .map(|r| r.cells.iter().map(|c| c.plain_text()).collect())
        .collect()
}

fn rk_int(n: i32) -> u32 {
    ((n << 2) as u32) | 0x02
}

#[test]
fn cells_of_every_kind_land_in_their_grid_positions() {
    let doc = parse(Book {
        strings: vec!["Item", "Qty", "Bolt"],
        sheets: vec![Sheet {
            name: "Stock",
            kind: 0,
            cells: vec![
                C::Sst(0, 0, 0),
                C::Sst(0, 1, 1),
                C::Sst(1, 0, 2),
                C::Rk(1, 1, 0, rk_int(12)),
                C::MulRk(2, 1, vec![(0, rk_int(3)), (0, (125 << 2) | 0x03)]),
                C::FormulaString(3, 0, "computed"),
                C::Bool(3, 1, true),
                C::Error(3, 2, 0x07),
                C::Number(4, 0, 0, 0.1),
            ],
        }],
        ..Default::default()
    })
    .unwrap();
    assert_eq!(doc.format, crate::FormatType::Xls);
    assert_eq!(doc.sections[0].name.as_deref(), Some("Stock"));
    assert_eq!(
        grid(&doc, 0),
        [
            vec!["Item", "Qty"],
            vec!["Bolt", "12"],
            vec!["", "3", "1.25"],
            vec!["computed", "TRUE", "#ERROR:#DIV/0!"],
            vec!["0.1"],
        ]
    );
    let Some(Block::Table(table)) = doc.sections[0].content.first() else {
        panic!()
    };
    assert!(table.rows[0].is_header && !table.rows[1].is_header);
}

#[test]
fn date_formatted_numbers_read_as_dates() {
    let doc = parse(Book {
        sheets: vec![Sheet {
            name: "Dates",
            kind: 0,
            cells: vec![
                C::Number(0, 0, 1, 45292.0),
                C::Rk(0, 1, 2, rk_int(45293)),
                C::Number(0, 2, 0, 45292.0),
            ],
        }],
        ..Default::default()
    })
    .unwrap();
    assert_eq!(grid(&doc, 0), [["2024-01-01", "2024-01-02", "45292"]]);
}

#[test]
fn merged_cells_span_and_cover() {
    let doc = parse(Book {
        strings: vec!["Wide", "Tall", "x"],
        sheets: vec![Sheet {
            name: "M",
            kind: 0,
            cells: vec![
                C::Sst(0, 0, 0),
                C::Sst(1, 0, 1),
                C::Sst(1, 1, 2),
                C::Sst(2, 1, 2),
                C::Merge(0, 0, 0, 1),
                C::Merge(1, 2, 0, 0),
            ],
        }],
        ..Default::default()
    })
    .unwrap();
    let Some(Block::Table(table)) = doc.sections[0].content.first() else {
        panic!()
    };
    assert_eq!(table.rows[0].cells[0].col_span, 2);
    assert_eq!(table.rows[1].cells[0].row_span, 2);
    assert_eq!(
        table.rows[2].cells.len(),
        1,
        "the covered position has no cell"
    );
}

#[test]
fn a_shared_string_split_across_continue_records_reads_whole() {
    // "Grüße 한국" is stored as UTF-16 (it has a character above U+00FF); split it mid-string.
    let doc = parse(Book {
        strings: vec!["first", "Grüße 한국"],
        split_sst_at: Some(8 + 3 + 5 + 3 + 6),
        sheets: vec![Sheet {
            name: "S",
            kind: 0,
            cells: vec![C::Sst(0, 0, 0), C::Sst(0, 1, 1)],
        }],
        ..Default::default()
    })
    .unwrap();
    assert_eq!(grid(&doc, 0), [["first", "Grüße 한국"]]);
}

#[test]
fn chart_sheets_are_not_sections() {
    let doc = parse(Book {
        strings: vec!["v"],
        sheets: vec![
            Sheet {
                name: "Chart1",
                kind: 2,
                cells: vec![],
            },
            Sheet {
                name: "Data",
                kind: 0,
                cells: vec![C::Sst(0, 0, 0)],
            },
        ],
        ..Default::default()
    })
    .unwrap();
    assert_eq!(doc.sections.len(), 1);
    assert_eq!(doc.sections[0].name.as_deref(), Some("Data"));
}

#[test]
fn an_encrypted_workbook_is_reported_as_encrypted() {
    let err = parse(Book {
        encrypted: true,
        ..Default::default()
    })
    .unwrap_err();
    assert!(matches!(err, Error::Encrypted), "{err}");
}

#[test]
fn an_unknown_biff_version_is_named_as_unsupported() {
    let err = parse(Book {
        biff_version: Some(0x0400),
        ..Default::default()
    })
    .unwrap_err();
    assert!(
        matches!(err, Error::UnsupportedFormat(ref m) if m.contains("0x0400")),
        "{err}"
    );
}

/// An embedded chart is a substream inside the worksheet's, with an EOF of its own; the
/// sheet goes on after it — merges and comments are written at the sheet's end.
#[test]
fn an_embedded_chart_does_not_end_its_sheet() {
    let doc = parse(Book {
        strings: vec!["Wide", "after"],
        sheets: vec![Sheet {
            name: "S",
            kind: 0,
            cells: vec![
                C::Sst(0, 0, 0),
                C::EmbeddedChart,
                C::Sst(1, 0, 1),
                C::Comment(1, 0, 3, "kept"),
                C::Merge(0, 0, 0, 1),
            ],
        }],
        ..Default::default()
    })
    .unwrap();
    let Some(Block::Table(table)) = doc.sections[0].content.first() else {
        panic!()
    };
    assert_eq!(table.rows[0].cells[0].col_span, 2);
    assert_eq!(table.rows[1].cells[0].plain_text(), "after [Comment: kept]");
}

#[test]
fn a_rich_text_cell_reads_as_its_text() {
    let doc = parse(Book {
        sheets: vec![Sheet {
            name: "S",
            kind: 0,
            cells: vec![C::Rich(0, 0, "bold start, plain rest")],
        }],
        ..Default::default()
    })
    .unwrap();
    assert_eq!(grid(&doc, 0), [["bold start, plain rest"]]);
}

#[test]
fn the_library_entry_points_read_an_xls() {
    let bytes = build(Book {
        strings: vec!["Name", "Ada"],
        sheets: vec![Sheet {
            name: "People",
            kind: 0,
            cells: vec![C::Sst(0, 0, 0), C::Sst(1, 0, 1)],
        }],
        ..Default::default()
    });
    let doc = crate::parse_bytes(&bytes).unwrap();
    let markdown =
        crate::render::to_markdown(&doc, &crate::render::RenderOptions::default()).unwrap();
    assert!(markdown.contains("| Name |"), "{markdown}");
    assert!(markdown.contains("| Ada |"), "{markdown}");
}

#[test]
fn hyperlinks_and_comments_attach_to_their_cells() {
    let doc = parse(Book {
        strings: vec!["Site", "Note me"],
        sheets: vec![Sheet {
            name: "S",
            kind: 0,
            cells: vec![
                C::Sst(0, 0, 0),
                C::Sst(0, 1, 1),
                C::Link(0, 0, "https://example.com/"),
                C::Comment(0, 1, 7, "A. Author:\ncheck this"),
                C::Comment(1, 0, 8, "on an empty cell"),
            ],
        }],
        ..Default::default()
    })
    .unwrap();
    let Some(Block::Table(table)) = doc.sections[0].content.first() else {
        panic!()
    };
    let runs = |r: usize, c: usize| table.rows[r].cells[c].content[0].runs.clone();
    assert_eq!(
        runs(0, 0)[0].hyperlink.as_deref(),
        Some("https://example.com/")
    );
    let note = runs(0, 1);
    assert_eq!(note[1].text, " [Comment: A. Author:\ncheck this]");
    assert!(note[1].style.italic);
    assert_eq!(runs(1, 0)[1].text, " [Comment: on an empty cell]");
}
