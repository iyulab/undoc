//! Workbooks older than Excel 97, assembled record by record: Excel 5.0/95 (BIFF5) in a
//! compound file's `Book` stream, and the bare record streams of Excel 2.x–4.0 (BIFF2–BIFF4).

use super::tests::{grid, rec};
use super::XlsParser;
use crate::error::Error;
use crate::model::Block;

const EOF: u16 = 0x000A;

fn u16s(values: &[u16]) -> Vec<u8> {
    values.iter().flat_map(|v| v.to_le_bytes()).collect()
}

fn bytes8(text: &[u8]) -> Vec<u8> {
    [&[text.len() as u8][..], text].concat()
}

fn bytes16(text: &[u8]) -> Vec<u8> {
    [&(text.len() as u16).to_le_bytes()[..], text].concat()
}

fn parse(data: Vec<u8>) -> crate::Result<crate::Document> {
    XlsParser::from_bytes(data)?.parse()
}

fn in_book_stream(stream: &[u8]) -> Vec<u8> {
    let mut container = cfb::CompoundFile::create(std::io::Cursor::new(Vec::new())).unwrap();
    let mut s = container.create_stream("/Book").unwrap();
    std::io::Write::write_all(&mut s, stream).unwrap();
    drop(s);
    container.flush().unwrap();
    container.into_inner().into_inner()
}

// ---------------------------------------------------------------------------------------------
// BIFF5 (Excel 5.0/95)

/// A BIFF5 cell record.
enum C5 {
    /// `LABEL`, its text bytes in the workbook's code page.
    Label(u16, u16, &'static [u8]),
    Rich(u16, u16, &'static [u8]),
    Number(u16, u16, u16, f64),
    Rk(u16, u16, u16, u32),
    MulRk(u16, u16, Vec<(u16, u32)>),
    FormulaString(u16, u16, &'static [u8]),
    FormulaNumber(u16, u16, f64),
    Bool(u16, u16, bool),
    /// A comment in its `NOTE`, split into `NOTE` continuations every `chunk` bytes.
    Note(u16, u16, &'static [u8], usize),
}

struct Book5 {
    codepage: Option<u16>,
    /// The `FONT` records' character sets.
    charsets: Vec<u8>,
    sheets: Vec<(&'static [u8], Vec<C5>)>,
    encrypted: bool,
}

impl Default for Book5 {
    fn default() -> Self {
        Self {
            codepage: Some(1252),
            charsets: vec![0],
            sheets: Vec::new(),
            encrypted: false,
        }
    }
}

fn bof5(kind: u16) -> Vec<u8> {
    u16s(&[0x0500, kind, 0x096C, 0x07C9])
}

fn head(row: u16, col: u16, xf: u16) -> Vec<u8> {
    u16s(&[row, col, xf])
}

fn build5(book: Book5) -> Vec<u8> {
    let mut globals = Vec::new();
    rec(&mut globals, 0x0809, &bof5(0x0005));
    if book.encrypted {
        rec(&mut globals, 0x002F, &[0, 0, 0, 0]);
    }
    if let Some(cp) = book.codepage {
        rec(&mut globals, 0x0042, &cp.to_le_bytes());
    }
    for charset in &book.charsets {
        let mut font = u16s(&[200, 0, 0x7FFF, 400, 0]);
        font.extend([0, 0, *charset, 0]);
        font.extend(bytes8(b"Arial"));
        rec(&mut globals, 0x0031, &font);
    }
    let mut format = 164u16.to_le_bytes().to_vec();
    format.extend(bytes8(b"yyyy\\-mm\\-dd"));
    rec(&mut globals, 0x041E, &format);
    // XF 0 general, XF 1 built-in date (14), XF 2 the custom date format.
    for ifmt in [0u16, 14, 164] {
        let mut xf = u16s(&[0, ifmt]);
        xf.extend([0u8; 12]);
        rec(&mut globals, 0x00E0, &xf);
    }
    let mut patch_at = Vec::new();
    for (name, _) in &book.sheets {
        let mut d = vec![0u8; 4];
        d.extend([0, 0]);
        d.extend(bytes8(name));
        patch_at.push(globals.len() + 4);
        rec(&mut globals, 0x0085, &d);
    }
    rec(&mut globals, EOF, &[]);

    let mut stream = globals;
    for (i, (_, cells)) in book.sheets.iter().enumerate() {
        let offset = stream.len() as u32;
        stream[patch_at[i]..patch_at[i] + 4].copy_from_slice(&offset.to_le_bytes());
        rec(&mut stream, 0x0809, &bof5(0x0010));
        let mut notes = Vec::new();
        for c in cells {
            match c {
                C5::Label(r, col, text) => {
                    rec(
                        &mut stream,
                        0x0204,
                        &[head(*r, *col, 0), bytes16(text)].concat(),
                    );
                }
                C5::Rich(r, col, text) => {
                    let d = [head(*r, *col, 0), bytes16(text), vec![1, 0, 0]].concat();
                    rec(&mut stream, 0x00D6, &d);
                }
                C5::Number(r, col, xf, v) => {
                    let d = [head(*r, *col, *xf), v.to_le_bytes().to_vec()].concat();
                    rec(&mut stream, 0x0203, &d);
                }
                C5::Rk(r, col, xf, rk) => {
                    let d = [head(*r, *col, *xf), rk.to_le_bytes().to_vec()].concat();
                    rec(&mut stream, 0x027E, &d);
                }
                C5::MulRk(r, first, values) => {
                    let mut d = u16s(&[*r, *first]);
                    for (xf, rk) in values {
                        d.extend(xf.to_le_bytes());
                        d.extend(rk.to_le_bytes());
                    }
                    d.extend((first + values.len() as u16 - 1).to_le_bytes());
                    rec(&mut stream, 0x00BD, &d);
                }
                C5::FormulaString(r, col, text) => {
                    let mut d = head(*r, *col, 0);
                    d.extend([0, 0, 0, 0, 0, 0, 0xFF, 0xFF]);
                    d.extend([0u8; 6]);
                    rec(&mut stream, 0x0006, &d);
                    rec(&mut stream, 0x0207, &bytes16(text));
                }
                C5::FormulaNumber(r, col, v) => {
                    let mut d = head(*r, *col, 0);
                    d.extend(v.to_le_bytes());
                    d.extend([0u8; 6]);
                    rec(&mut stream, 0x0006, &d);
                }
                C5::Bool(r, col, v) => {
                    rec(
                        &mut stream,
                        0x0205,
                        &[head(*r, *col, 0), vec![*v as u8, 0]].concat(),
                    );
                }
                C5::Note(r, col, text, chunk) => notes.push((*r, *col, *text, *chunk)),
            }
        }
        // Old-style comments: the first NOTE carries the whole text's length, then as much of
        // it as fits; NOTEs for row 0xFFFF carry the rest.
        for (r, col, text, chunk) in notes {
            let mut pieces = text.chunks(chunk);
            let first = pieces.next().unwrap_or(&[]);
            let d = [u16s(&[r, col, text.len() as u16]), first.to_vec()].concat();
            rec(&mut stream, 0x001C, &d);
            for more in pieces {
                let d = [u16s(&[0xFFFF, 0, more.len() as u16]), more.to_vec()].concat();
                rec(&mut stream, 0x001C, &d);
            }
        }
        rec(&mut stream, EOF, &[]);
    }
    in_book_stream(&stream)
}

fn rk_int(n: i32) -> u32 {
    ((n << 2) as u32) | 0x02
}

#[test]
fn an_excel_95_workbook_reads_like_a_97_one() {
    let doc = parse(build5(Book5 {
        sheets: vec![
            (
                b"Caf\xE9",
                vec![
                    C5::Label(0, 0, b"Item"),
                    C5::Rich(0, 1, b"Qty"),
                    C5::Label(1, 0, b"\x93Bolt\x94"),
                    C5::Rk(1, 1, 0, rk_int(12)),
                    C5::MulRk(2, 1, vec![(0, rk_int(3)), (0, (125 << 2) | 0x03)]),
                    C5::FormulaString(3, 0, b"computed"),
                    C5::Bool(3, 1, true),
                    C5::FormulaNumber(3, 2, 2.5),
                    C5::Number(4, 0, 1, 45292.0),
                    C5::Rk(4, 1, 2, rk_int(45293)),
                    C5::Number(4, 2, 0, 45292.0),
                ],
            ),
            (b"Second", vec![C5::Label(0, 0, b"x")]),
        ],
        ..Default::default()
    }))
    .unwrap();
    assert_eq!(doc.format, crate::FormatType::Xls);
    let names: Vec<_> = doc.sections.iter().map(|s| s.name.as_deref()).collect();
    assert_eq!(names, [Some("Café"), Some("Second")]);
    assert_eq!(
        grid(&doc, 0),
        [
            vec!["Item", "Qty"],
            vec!["\u{201C}Bolt\u{201D}", "12"],
            vec!["", "3", "1.25"],
            vec!["computed", "TRUE", "2.5"],
            vec!["2024-01-01", "2024-01-02", "45292"],
        ]
    );
}

#[test]
fn an_old_style_comment_is_joined_across_its_note_records() {
    let doc = parse(build5(Book5 {
        sheets: vec![(
            b"S",
            vec![
                C5::Label(0, 0, b"value"),
                C5::Note(0, 0, b"split across\nthree records", 10),
                C5::Note(1, 1, b"on an empty cell", 100),
            ],
        )],
        ..Default::default()
    }))
    .unwrap();
    let Some(Block::Table(table)) = doc.sections[0].content.first() else {
        panic!()
    };
    let runs = |r: usize, c: usize| table.rows[r].cells[c].content[0].runs.clone();
    assert_eq!(
        runs(0, 0)[1].text,
        " [Comment: split across\nthree records]"
    );
    assert!(runs(0, 0)[1].style.italic);
    assert_eq!(runs(1, 1)[1].text, " [Comment: on an empty cell]");
}

#[cfg(feature = "codepages")]
#[test]
fn text_is_decoded_in_the_declared_code_page() {
    let doc = parse(build5(Book5 {
        codepage: Some(949),
        sheets: vec![(
            b"\xBD\xC3\xC6\xAE",                                // 시트
            vec![C5::Label(0, 0, b"\xBA\xB8\xB0\xED\xBC\xAD")], // 보고서
        )],
        ..Default::default()
    }))
    .unwrap();
    assert_eq!(doc.sections[0].name.as_deref(), Some("시트"));
    assert_eq!(grid(&doc, 0), [["보고서"]]);
}

/// A workbook that names no code page is read in its fonts' character set — not as
/// Windows-1252, which would turn every Korean label into Latin-1 mojibake.
#[cfg(feature = "codepages")]
#[test]
fn without_a_code_page_the_fonts_character_set_decides() {
    let doc = parse(build5(Book5 {
        codepage: None,
        charsets: vec![129, 129, 0],
        sheets: vec![(b"S", vec![C5::Label(0, 0, b"\xBA\xB8\xB0\xED\xBC\xAD")])],
        ..Default::default()
    }))
    .unwrap();
    assert_eq!(grid(&doc, 0), [["보고서"]]);
}

#[test]
fn without_a_code_page_or_a_telling_font_text_is_windows_1252() {
    let doc = parse(build5(Book5 {
        codepage: None,
        charsets: vec![0, 2],
        sheets: vec![(b"S", vec![C5::Label(0, 0, b"na\xEFve \x80")])],
        ..Default::default()
    }))
    .unwrap();
    assert_eq!(grid(&doc, 0), [["naïve €"]]);
}

#[test]
fn an_encrypted_excel_95_workbook_is_reported_as_encrypted() {
    let err = parse(build5(Book5 {
        encrypted: true,
        sheets: vec![(b"S", vec![C5::Label(0, 0, b"x")])],
        ..Default::default()
    }))
    .unwrap_err();
    assert!(matches!(err, Error::Encrypted), "{err}");
}

#[test]
fn the_library_entry_points_read_an_excel_95_workbook() {
    let bytes = build5(Book5 {
        sheets: vec![(
            b"People",
            vec![C5::Label(0, 0, b"Name"), C5::Label(1, 0, b"Ada")],
        )],
        ..Default::default()
    });
    assert_eq!(
        crate::detect_format_from_bytes(&bytes).unwrap(),
        crate::FormatType::Xls
    );
    let doc = crate::parse_bytes(&bytes).unwrap();
    let markdown =
        crate::render::to_markdown(&doc, &crate::render::RenderOptions::default()).unwrap();
    assert!(markdown.contains("| Name |"), "{markdown}");
    assert!(markdown.contains("| Ada |"), "{markdown}");
}

// ---------------------------------------------------------------------------------------------
// BIFF2–BIFF4: the record stream is the file

/// A bare BIFF3/BIFF4 worksheet stream. Every number format is written out, indexed by its
/// position — here `General`, `0.00`, `m/d/yy` — and XF 0, 1, 2 point at them in turn.
fn sheet34(biff: u16, cells: &[Vec<u8>]) -> Vec<u8> {
    let (bof, format, xf, version) = match biff {
        3 => (0x0209, 0x001E, 0x0243, 0x0000),
        _ => (0x0409, 0x041E, 0x0443, 0x0000),
    };
    let mut s = Vec::new();
    rec(&mut s, bof, &u16s(&[version, 0x0010, 0x0000]));
    rec(&mut s, 0x0042, &32769u16.to_le_bytes());
    for code in [&b"General"[..], b"0.00", b"m/d/yy"] {
        let d = if biff == 3 {
            bytes8(code)
        } else {
            [vec![0, 0], bytes8(code)].concat()
        };
        rec(&mut s, format, &d);
    }
    for ifmt in [0u8, 1, 2] {
        let mut d = vec![0, ifmt];
        d.extend([0u8; 10]);
        rec(&mut s, xf, &d);
    }
    for cell in cells {
        s.extend(cell);
    }
    rec(&mut s, EOF, &[]);
    s
}

fn record(kind: u16, data: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    rec(&mut out, kind, data);
    out
}

#[test]
fn an_excel_4_worksheet_reads_from_its_bare_stream() {
    let formula4 = [
        head(2, 0, 0),
        vec![0, 0, 0, 0, 0, 0, 0xFF, 0xFF],
        vec![0u8; 6],
    ]
    .concat();
    let data = sheet34(
        4,
        &[
            record(0x0204, &[head(0, 0, 0), bytes16(b"Year")].concat()),
            record(0x0204, &[head(0, 1, 0), bytes16(b"Amount")].concat()),
            record(
                0x0203,
                &[head(1, 0, 2), 45292.0f64.to_le_bytes().to_vec()].concat(),
            ),
            record(
                0x027E,
                &[head(1, 1, 1), rk_int(7).to_le_bytes().to_vec()].concat(),
            ),
            record(0x0406, &formula4),
            record(0x0207, &bytes16(b"from a formula")),
            record(0x0205, &[head(2, 1, 0), vec![0x07, 1]].concat()),
            record(0x001C, &[u16s(&[2, 1, 4]), b"note".to_vec()].concat()),
        ],
    );
    assert_eq!(
        crate::detect_format_from_bytes(&data).unwrap(),
        crate::FormatType::Xls
    );
    let doc = parse(data).unwrap();
    assert_eq!(doc.sections.len(), 1);
    assert_eq!(
        doc.sections[0].name, None,
        "a one-sheet file names no sheet"
    );
    assert_eq!(
        grid(&doc, 0),
        [
            vec!["Year", "Amount"],
            vec!["2024-01-01", "7"],
            vec!["from a formula", "#ERROR:#DIV/0! [Comment: note]"],
        ]
    );
}

/// Before BIFF5 there are no built-in formats: index 14 is whatever the file's fifteenth
/// `FORMAT` says, not BIFF5's built-in date.
#[test]
fn a_format_index_before_biff5_is_only_a_position() {
    let mut data = sheet34(4, &[]);
    data.truncate(data.len() - 4); // reopen before the EOF
    for i in 3..15u8 {
        let code = if i == 14 { &b"0.0"[..] } else { b"@" };
        data.extend(record(0x041E, &[vec![0, 0], bytes8(code)].concat()));
    }
    let mut xf = vec![0, 14];
    xf.extend([0u8; 10]);
    data.extend(record(0x0443, &xf)); // XF 3 → format 14
    data.extend(record(
        0x0203,
        &[head(0, 0, 3), 45292.0f64.to_le_bytes().to_vec()].concat(),
    ));
    data.extend(record(EOF, &[]));
    assert_eq!(grid(&parse(data).unwrap(), 0), [["45292"]]);
}

#[test]
fn an_excel_3_worksheet_reads_from_its_bare_stream() {
    let formula3 = [head(1, 0, 1), 2.5f64.to_le_bytes().to_vec(), vec![0u8; 6]].concat();
    let data = sheet34(
        3,
        &[
            record(0x0204, &[head(0, 0, 0), bytes16(b"Caf\xE9")].concat()),
            record(0x0206, &formula3),
            record(
                0x0203,
                &[head(1, 1, 2), 45293.0f64.to_le_bytes().to_vec()].concat(),
            ),
        ],
    );
    let doc = parse(data).unwrap();
    assert_eq!(grid(&doc, 0), [vec!["Café"], vec!["2.5", "2024-01-02"]]);
}

/// BIFF2 cells carry three attribute bytes instead of an XF index; the second holds the
/// number format's index.
#[test]
fn an_excel_2_worksheet_reads_from_its_bare_stream() {
    let attr = |fmt: u8| vec![0u8, fmt, 0];
    let cell = |r: u16, c: u16, fmt: u8| [u16s(&[r, c]), attr(fmt)].concat();
    let mut s = Vec::new();
    rec(&mut s, 0x0009, &u16s(&[0x0002, 0x0010]));
    for code in [&b"General"[..], b"0", b"d-mmm-yy"] {
        rec(&mut s, 0x001E, &bytes8(code));
    }
    rec(&mut s, 0x0004, &[cell(0, 0, 0), bytes8(b"Label")].concat());
    rec(
        &mut s,
        0x0002,
        &[cell(0, 1, 1), 42u16.to_le_bytes().to_vec()].concat(),
    );
    rec(
        &mut s,
        0x0003,
        &[cell(1, 0, 2), 45292.0f64.to_le_bytes().to_vec()].concat(),
    );
    rec(&mut s, 0x0005, &[cell(1, 1, 0), vec![0, 0]].concat());
    let formula = [
        cell(2, 0, 0),
        vec![0, 0, 0, 0, 0, 0, 0xFF, 0xFF],
        vec![0u8; 3],
    ]
    .concat();
    rec(&mut s, 0x0006, &formula);
    rec(&mut s, 0x0007, &bytes8(b"text result"));
    rec(&mut s, EOF, &[]);
    assert_eq!(
        crate::detect_format_from_bytes(&s).unwrap(),
        crate::FormatType::Xls
    );
    let doc = parse(s).unwrap();
    assert_eq!(
        grid(&doc, 0),
        [
            vec!["Label", "42"],
            vec!["2024-01-01", "FALSE"],
            vec!["text result"],
        ]
    );
}

/// A BIFF4 workbook keeps its worksheets in one stream, each after a `SHEETHDR` naming it and
/// each with formats of its own; a chart among them is not a section.
#[test]
fn an_excel_4_workbook_lists_each_worksheet() {
    let sheet = |label: &[u8], date_format: &[u8]| {
        let mut s = Vec::new();
        rec(&mut s, 0x0409, &u16s(&[0, 0x0010, 0]));
        for code in [&b"General"[..], date_format] {
            rec(&mut s, 0x041E, &[vec![0, 0], bytes8(code)].concat());
        }
        for ifmt in [0u8, 1] {
            rec(&mut s, 0x0443, &[vec![0, ifmt], vec![0u8; 10]].concat());
        }
        rec(&mut s, 0x0204, &[head(0, 0, 0), bytes16(label)].concat());
        rec(
            &mut s,
            0x0203,
            &[head(0, 1, 1), 45292.0f64.to_le_bytes().to_vec()].concat(),
        );
        rec(&mut s, EOF, &[]);
        s
    };
    let chart = {
        let mut s = Vec::new();
        rec(&mut s, 0x0409, &u16s(&[0, 0x0020, 0]));
        rec(&mut s, EOF, &[]);
        s
    };
    let mut stream = Vec::new();
    rec(&mut stream, 0x0409, &u16s(&[0, 0x0100, 0]));
    rec(&mut stream, 0x0042, &1252u16.to_le_bytes());
    for (name, substream) in [
        (&b"First"[..], sheet(b"one", b"m/d/yy")),
        (b"Chart", chart),
        (b"Second", sheet(b"two", b"0.00")),
    ] {
        let d = [
            (substream.len() as u32).to_le_bytes().to_vec(),
            bytes8(name),
        ]
        .concat();
        rec(&mut stream, 0x008F, &d);
        stream.extend(substream);
    }
    rec(&mut stream, EOF, &[]);

    let doc = parse(stream).unwrap();
    let names: Vec<_> = doc.sections.iter().map(|s| s.name.as_deref()).collect();
    assert_eq!(names, [Some("First"), Some("Second")]);
    // The same cell, formatted by each sheet's own second format.
    assert_eq!(grid(&doc, 0), [["one", "2024-01-01"]]);
    assert_eq!(grid(&doc, 1), [["two", "45292"]]);
}

#[test]
fn an_encrypted_bare_stream_is_reported_as_encrypted() {
    let mut s = Vec::new();
    rec(&mut s, 0x0409, &u16s(&[0, 0x0010, 0]));
    rec(&mut s, 0x002F, &[0, 0, 0, 0]);
    rec(&mut s, EOF, &[]);
    assert!(matches!(parse(s).unwrap_err(), Error::Encrypted));
}
