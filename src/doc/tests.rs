//! End-to-end tests over Word binary documents the tests assemble themselves: a File
//! Information Block, a piece table, one page of paragraph properties and a style sheet, in a
//! CFB container — the parts a Word 97 file is read through.

use super::DocParser;
use crate::error::Error;
use crate::model::{Block, HeadingLevel};

/// One paragraph of a test document.
struct Para {
    text: &'static str,
    /// Paragraph mark: '\r', or '\u{07}' for a cell or row-end mark.
    mark: char,
    istd: u16,
    in_table: bool,
    row_end: bool,
    /// List level, in the document's one list.
    list_level: Option<u8>,
}

fn p(text: &'static str) -> Para {
    Para {
        text,
        mark: '\r',
        istd: 0,
        in_table: false,
        row_end: false,
        list_level: None,
    }
}

fn item(text: &'static str, level: u8) -> Para {
    Para {
        list_level: Some(level),
        ..p(text)
    }
}

fn heading(text: &'static str) -> Para {
    Para { istd: 1, ..p(text) }
}

fn cell(text: &'static str) -> Para {
    Para {
        mark: '\u{07}',
        in_table: true,
        ..p(text)
    }
}

fn row_end() -> Para {
    Para {
        mark: '\u{07}',
        in_table: true,
        row_end: true,
        ..p("")
    }
}

#[derive(Default)]
struct Options {
    compressed: bool,
    encrypted: bool,
    n_fib: Option<u16>,
    /// Character formatting: the first occurrence of each text in the main story gets the
    /// `grpprl`.
    formats: Vec<(&'static str, Vec<u8>)>,
    /// Footnote texts, referenced in order by the `\u{02}` marks in the main story.
    footnotes: Vec<&'static str>,
    /// The number format of the document's one list's levels (0 decimal, 0x17 bullet).
    list_nfc: u8,
    /// A `Data` stream, for inline pictures.
    data: Option<Vec<u8>>,
}

const TEXT_AT: usize = 1024;

fn le16(v: &mut Vec<u8>, x: u16) {
    v.extend_from_slice(&x.to_le_bytes());
}

fn le32(v: &mut Vec<u8>, x: u32) {
    v.extend_from_slice(&x.to_le_bytes());
}

/// The style sheet: istd 0 "Normal" (sti 0), istd 1 "heading 1" (sti 1).
fn stsh() -> Vec<u8> {
    let mut out = Vec::new();
    le16(&mut out, 18); // cbStshi
    let mut stshi = vec![0u8; 18];
    stshi[0..2].copy_from_slice(&2u16.to_le_bytes()); // cstd
    stshi[2..4].copy_from_slice(&10u16.to_le_bytes()); // cbSTDBaseInFile
    out.extend_from_slice(&stshi);
    for (sti, name) in [(0u16, "Normal"), (1u16, "heading 1")] {
        let mut std = vec![0u8; 10];
        std[0..2].copy_from_slice(&sti.to_le_bytes());
        le16(&mut std, name.encode_utf16().count() as u16);
        for unit in name.encode_utf16() {
            le16(&mut std, unit);
        }
        le16(&mut std, 0);
        le16(&mut out, std.len() as u16);
        out.extend_from_slice(&std);
    }
    out
}

fn build(paras: &[Para], opts: Options) -> Vec<u8> {
    // The text, with each paragraph's [start, end) in characters.
    let mut text = String::new();
    let mut spans = Vec::new();
    for para in paras {
        let start = text.chars().count();
        text.push_str(para.text);
        text.push(para.mark);
        spans.push((start, text.chars().count()));
    }
    let ccp_text = text.chars().count();
    // The footnote story follows the main text: each note opens with its own mark.
    let mut note_starts = Vec::new();
    for note in &opts.footnotes {
        note_starts.push(text.chars().count() - ccp_text);
        text.push('\u{02}');
        text.push_str(note);
        text.push('\r');
    }
    let ccp_ftn = text.chars().count() - ccp_text;
    let note_refs: Vec<usize> = text
        .chars()
        .take(ccp_text)
        .enumerate()
        .filter(|(_, c)| *c == '\u{02}')
        .map(|(i, _)| i)
        .collect();
    let width = if opts.compressed { 1 } else { 2 };
    let encoded: Vec<u8> = if opts.compressed {
        text.chars().map(|c| c as u32 as u8).collect()
    } else {
        text.encode_utf16().flat_map(|u| u.to_le_bytes()).collect()
    };
    let ccp = text.chars().count() as u32;
    let fc_of = |cp: usize| (TEXT_AT + cp * width) as u32;

    // One page of paragraph properties, right after the text.
    let fkp_at = (TEXT_AT + encoded.len()).div_ceil(512) * 512;
    let mut page = vec![0u8; 512];
    let crun = paras.len();
    assert!(
        4 * (crun + 1) + 13 * crun < 400,
        "too many paragraphs for one test page"
    );
    for (i, &(start, _)) in spans.iter().enumerate() {
        page[i * 4..i * 4 + 4].copy_from_slice(&fc_of(start).to_le_bytes());
    }
    let last = fc_of(spans.last().unwrap().1);
    page[crun * 4..crun * 4 + 4].copy_from_slice(&last.to_le_bytes());
    let mut free = 510;
    for (i, para) in paras.iter().enumerate() {
        let mut body = Vec::new();
        le16(&mut body, para.istd);
        if para.in_table {
            body.extend_from_slice(&[0x16, 0x24, 0x01]); // sprmPFInTable
        }
        if para.row_end {
            body.extend_from_slice(&[0x17, 0x24, 0x01]); // sprmPFTtp
        }
        if let Some(level) = para.list_level {
            body.extend_from_slice(&[0x0B, 0x46, 0x01, 0x00]); // sprmPIlfo 1
            body.extend_from_slice(&[0x0A, 0x26, level]); // sprmPIlvl
        }
        let mut papx = Vec::new();
        if body.len() % 2 == 1 {
            papx.push(body.len().div_ceil(2) as u8);
        } else {
            papx.push(0);
            papx.push((body.len() / 2) as u8);
        }
        papx.extend_from_slice(&body);
        free -= papx.len();
        free -= free % 2;
        page[free..free + papx.len()].copy_from_slice(&papx);
        page[4 * (crun + 1) + i * 13] = (free / 2) as u8;
    }
    page[511] = crun as u8;

    // A page of character properties after it: runs over every boundary a format sets.
    let chpx_at = fkp_at + 512;
    let main: Vec<char> = text.chars().take(ccp_text).collect();
    let mut ranges = Vec::new();
    for (needle, grpprl) in &opts.formats {
        let needle: Vec<char> = needle.chars().collect();
        let from = main
            .windows(needle.len())
            .position(|w| w == needle.as_slice())
            .expect("formatted text is in the main story");
        ranges.push((from, from + needle.len(), grpprl.clone()));
    }
    let mut bounds = vec![0, ccp_text];
    for (from, to, _) in &ranges {
        bounds.extend([*from, *to]);
    }
    bounds.sort_unstable();
    bounds.dedup();
    let mut chpx_page = vec![0u8; 512];
    let crun_c = bounds.len() - 1;
    for (i, cp) in bounds.iter().enumerate() {
        chpx_page[i * 4..i * 4 + 4].copy_from_slice(&fc_of(*cp).to_le_bytes());
    }
    let mut free_c = 510;
    for i in 0..crun_c {
        if let Some((_, _, grpprl)) = ranges
            .iter()
            .find(|(f, t, _)| *f <= bounds[i] && bounds[i] < *t)
        {
            free_c -= 1 + grpprl.len();
            free_c -= free_c % 2;
            chpx_page[free_c] = grpprl.len() as u8;
            chpx_page[free_c + 1..free_c + 1 + grpprl.len()].copy_from_slice(grpprl);
            chpx_page[4 * (crun_c + 1) + i] = (free_c / 2) as u8;
        }
    }
    chpx_page[511] = crun_c as u8;

    // The table stream: piece table, paragraph property index, style sheet.
    let mut table = Vec::new();
    let clx_at = table.len();
    let raw_fc = if opts.compressed {
        (TEXT_AT as u32 * 2) | 0x4000_0000
    } else {
        TEXT_AT as u32
    };
    let mut plc = Vec::new();
    le32(&mut plc, 0);
    le32(&mut plc, ccp);
    le16(&mut plc, 0);
    le32(&mut plc, raw_fc);
    le16(&mut plc, 0);
    table.push(0x02);
    le32(&mut table, plc.len() as u32);
    table.extend_from_slice(&plc);
    let clx_len = table.len() - clx_at;

    let papx_at = table.len();
    le32(&mut table, fc_of(0));
    le32(&mut table, last);
    le32(&mut table, (fkp_at / 512) as u32);
    let papx_len = table.len() - papx_at;

    let stsh_at = table.len();
    table.extend_from_slice(&stsh());
    let stsh_len = table.len() - stsh_at;

    let chpx_plc_at = table.len();
    le32(&mut table, fc_of(0));
    le32(&mut table, fc_of(ccp_text));
    le32(&mut table, (chpx_at / 512) as u32);
    let chpx_plc_len = table.len() - chpx_plc_at;

    // Footnote references (CPs + 2-byte FRDs) and text ranges (CPs, with Word's extra one).
    let fnd_ref_at = table.len();
    if !note_refs.is_empty() {
        for cp in &note_refs {
            le32(&mut table, *cp as u32);
        }
        le32(&mut table, ccp_text as u32);
        for _ in &note_refs {
            le16(&mut table, 1);
        }
    }
    let fnd_ref_len = table.len() - fnd_ref_at;
    let fnd_txt_at = table.len();
    if !note_refs.is_empty() {
        for start in &note_starts {
            le32(&mut table, *start as u32);
        }
        le32(&mut table, ccp_ftn as u32);
        le32(&mut table, ccp_ftn as u32 + 1);
    }
    let fnd_txt_len = table.len() - fnd_txt_at;

    // One nine-level list (PlfLst: count, LSTF; its LVLs follow), and one override.
    let lst_at = table.len();
    le16(&mut table, 1);
    let mut lstf = vec![0u8; 28];
    lstf[0..4].copy_from_slice(&42i32.to_le_bytes());
    table.extend_from_slice(&lstf);
    let lst_len = table.len() - lst_at;
    for _ in 0..9 {
        let mut lvlf = vec![0u8; 28];
        lvlf[0..4].copy_from_slice(&1i32.to_le_bytes()); // iStartAt
        lvlf[4] = opts.list_nfc;
        table.extend_from_slice(&lvlf);
        le16(&mut table, 0); // empty number text
    }
    let lfo_at = table.len();
    table.extend_from_slice(&1i32.to_le_bytes());
    let mut lfo = vec![0u8; 16];
    lfo[0..4].copy_from_slice(&42i32.to_le_bytes());
    table.extend_from_slice(&lfo);
    let lfo_len = table.len() - lfo_at;

    // The File Information Block.
    let mut wd = vec![0u8; chpx_at + 512];
    wd[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
    wd[2..4].copy_from_slice(&opts.n_fib.unwrap_or(0x00C1).to_le_bytes());
    let flags: u16 = 0x0200 | if opts.encrypted { 0x0100 } else { 0 };
    wd[0x0A..0x0C].copy_from_slice(&flags.to_le_bytes());
    let mut at = 32;
    wd[at..at + 2].copy_from_slice(&14u16.to_le_bytes());
    at += 2 + 28;
    wd[at..at + 2].copy_from_slice(&22u16.to_le_bytes());
    let rg_lw = at + 2;
    wd[rg_lw + 12..rg_lw + 16].copy_from_slice(&(ccp_text as u32).to_le_bytes());
    wd[rg_lw + 16..rg_lw + 20].copy_from_slice(&(ccp_ftn as u32).to_le_bytes());
    at = rg_lw + 88;
    wd[at..at + 2].copy_from_slice(&93u16.to_le_bytes());
    let rg_fc_lcb = at + 2;
    let mut pair = |index: usize, fc: usize, lcb: usize| {
        let at = rg_fc_lcb + index * 8;
        wd[at..at + 4].copy_from_slice(&(fc as u32).to_le_bytes());
        wd[at + 4..at + 8].copy_from_slice(&(lcb as u32).to_le_bytes());
    };
    pair(1, stsh_at, stsh_len);
    pair(2, fnd_ref_at, fnd_ref_len);
    pair(3, fnd_txt_at, fnd_txt_len);
    pair(12, chpx_plc_at, chpx_plc_len);
    pair(13, papx_at, papx_len);
    pair(33, clx_at, clx_len);
    pair(73, lst_at, lst_len);
    pair(74, lfo_at, lfo_len);
    assert!(rg_fc_lcb + 93 * 8 <= TEXT_AT);
    wd[TEXT_AT..TEXT_AT + encoded.len()].copy_from_slice(&encoded);
    wd[fkp_at..fkp_at + 512].copy_from_slice(&page);
    wd[chpx_at..chpx_at + 512].copy_from_slice(&chpx_page);

    let mut container = cfb::CompoundFile::create(std::io::Cursor::new(Vec::new())).unwrap();
    let mut streams = vec![("/WordDocument", wd), ("/1Table", table)];
    if let Some(data) = opts.data {
        streams.push(("/Data", data));
    }
    for (name, data) in &streams {
        let mut stream = container.create_stream(name).unwrap();
        std::io::Write::write_all(&mut stream, data).unwrap();
    }
    container.flush().unwrap();
    container.into_inner().into_inner()
}

fn parse(paras: &[Para], opts: Options) -> crate::Result<crate::Document> {
    DocParser::from_bytes(build(paras, opts))?.parse()
}

fn paragraphs(doc: &crate::Document) -> Vec<String> {
    doc.sections[0]
        .content
        .iter()
        .filter_map(|b| match b {
            Block::Paragraph(p) => Some(p.plain_text()),
            _ => None,
        })
        .collect()
}

#[test]
fn paragraphs_and_a_heading_style_are_read() {
    let doc = parse(
        &[heading("Results"), p("The rate held."), p("")],
        Options::default(),
    )
    .unwrap();
    assert_eq!(doc.format, crate::FormatType::Doc);
    assert_eq!(paragraphs(&doc), ["Results", "The rate held."]);
    let Block::Paragraph(first) = &doc.sections[0].content[0] else {
        panic!()
    };
    assert_eq!(first.heading, HeadingLevel::H1);
    assert_eq!(first.style_name.as_deref(), Some("heading 1"));
}

#[test]
fn compressed_text_is_windows_1252() {
    // 0x93 / 0x94 are curly quotes in Windows-1252, not C1 controls.
    let doc = parse(
        &[p("say \u{93}hi\u{94}")],
        Options {
            compressed: true,
            ..Default::default()
        },
    )
    .unwrap();
    assert_eq!(paragraphs(&doc), ["say \u{201C}hi\u{201D}"]);
}

#[test]
fn a_table_keeps_its_rows_and_an_empty_cell() {
    let doc = parse(
        &[
            p("Before"),
            cell("A"),
            cell(""),
            row_end(),
            cell("C"),
            cell("D"),
            row_end(),
            p("After"),
        ],
        Options::default(),
    )
    .unwrap();
    let content = &doc.sections[0].content;
    assert_eq!(content.len(), 3, "{content:?}");
    let Block::Table(table) = &content[1] else {
        panic!("{content:?}")
    };
    let cells: Vec<Vec<String>> = table
        .rows
        .iter()
        .map(|r| r.cells.iter().map(|c| c.plain_text()).collect())
        .collect();
    assert_eq!(cells, [["A", ""], ["C", "D"]]);
    assert_eq!(paragraphs(&doc), ["Before", "After"]);
}

#[test]
fn a_hyperlink_field_keeps_its_result_and_target_and_drops_the_instruction() {
    let doc = parse(
        &[p(
            "See \u{13} HYPERLINK \"https://example.com/\" \u{14}the site\u{15} now.",
        )],
        Options::default(),
    )
    .unwrap();
    assert_eq!(paragraphs(&doc), ["See the site now."]);
    let Block::Paragraph(para) = &doc.sections[0].content[0] else {
        panic!()
    };
    let link = para
        .runs
        .iter()
        .find(|r| r.hyperlink.is_some())
        .expect("a linked run");
    assert_eq!(link.text, "the site");
    assert_eq!(link.hyperlink.as_deref(), Some("https://example.com/"));
}

#[test]
fn a_field_without_a_result_contributes_nothing() {
    let doc = parse(&[p("Page \u{13} PAGE \u{15}end")], Options::default()).unwrap();
    assert_eq!(paragraphs(&doc), ["Page end"]);
}

#[test]
fn a_line_break_and_a_page_break_are_kept_on_their_runs() {
    let doc = parse(&[p("one\u{0B}two\u{0C}")], Options::default()).unwrap();
    let Block::Paragraph(para) = &doc.sections[0].content[0] else {
        panic!()
    };
    assert!(para.runs[0].line_break, "{:?}", para.runs);
    assert!(para.runs.last().unwrap().page_break, "{:?}", para.runs);
}

#[test]
fn an_encrypted_document_is_reported_as_encrypted() {
    let err = parse(
        &[p("secret")],
        Options {
            encrypted: true,
            ..Default::default()
        },
    )
    .unwrap_err();
    assert!(matches!(err, Error::Encrypted), "{err}");
}

/// A Word 6.0/95 document: the `WordDocument` stream alone, its FIB at fixed offsets, its text
/// one byte per character from `fcMin` -- or, fast-saved, in `pieces` (character count, byte
/// offset) through a piece table -- and a font table whose character sets name the code page.
fn build_word6(text: &[u8], n_fib: u16, fonts: &[u8], pieces: Option<&[(u32, u32)]>) -> Vec<u8> {
    build_word6_with(text, n_fib, fonts, pieces, &[])
}

/// One Word 6.0 paragraph: its length in characters, its style, whether it is in a table and
/// whether it ends a table row.
type Word6Para = (u32, u16, bool, bool);

/// [`build_word6`] with paragraph properties, written as one Word 6.0 paragraph page, and a
/// style sheet whose `istd` 1 is "heading 1".
fn build_word6_with(
    text: &[u8],
    n_fib: u16,
    fonts: &[u8],
    pieces: Option<&[(u32, u32)]>,
    paras: &[Word6Para],
) -> Vec<u8> {
    const FC_MIN: usize = 0x300;
    const STSH_AT: usize = 0x500;
    const FONTS_AT: usize = 0x600;
    const CLX_AT: usize = 0x700;
    const PAPX_PLC_AT: usize = 0x780;
    const PAPX_PAGE: usize = 4; // the page at 0x800
    let mut wd = vec![0u8; 0xA00];
    wd[0..2].copy_from_slice(&0xA5DCu16.to_le_bytes());
    wd[2..4].copy_from_slice(&n_fib.to_le_bytes());
    let flags: u16 = if pieces.is_some() { 0x0004 } else { 0 };
    wd[0x0A..0x0C].copy_from_slice(&flags.to_le_bytes());
    wd[0x18..0x1C].copy_from_slice(&(FC_MIN as u32).to_le_bytes());
    wd[0x1C..0x20].copy_from_slice(&((FC_MIN + text.len()) as u32).to_le_bytes());
    let ccp: u32 = match pieces {
        Some(pieces) => pieces.iter().map(|&(n, _)| n).sum(),
        None => text.len() as u32,
    };
    wd[0x34..0x38].copy_from_slice(&ccp.to_le_bytes());
    wd[FC_MIN..FC_MIN + text.len()].copy_from_slice(text);

    // Font table: total length, then per font its length less one, three bytes of family and
    // weight, the character set, the alternate-name index, and the name.
    let mut table = vec![0u8, 0];
    for (i, &chs) in fonts.iter().enumerate() {
        let name = format!("Font{i}\0");
        table.push((5 + name.len()) as u8);
        table.extend_from_slice(&[0, 0x90, 0x01, chs, 0]);
        table.extend_from_slice(name.as_bytes());
    }
    let len = table.len() as u16;
    table[0..2].copy_from_slice(&len.to_le_bytes());
    wd[FONTS_AT..FONTS_AT + table.len()].copy_from_slice(&table);
    wd[0xD0..0xD4].copy_from_slice(&(FONTS_AT as u32).to_le_bytes());
    wd[0xD4..0xD8].copy_from_slice(&(table.len() as u32).to_le_bytes());

    if let Some(pieces) = pieces {
        let mut plc = Vec::new();
        let mut cp = 0u32;
        le32(&mut plc, cp);
        for &(n, _) in pieces {
            cp += n;
            le32(&mut plc, cp);
        }
        for &(_, fc) in pieces {
            le16(&mut plc, 0);
            le32(&mut plc, FC_MIN as u32 + fc);
            le16(&mut plc, 0);
        }
        let mut clx = vec![0x02];
        le32(&mut clx, plc.len() as u32);
        clx.extend_from_slice(&plc);
        wd[CLX_AT..CLX_AT + clx.len()].copy_from_slice(&clx);
        wd[0x160..0x164].copy_from_slice(&(CLX_AT as u32).to_le_bytes());
        wd[0x164..0x168].copy_from_slice(&(clx.len() as u32).to_le_bytes());
    }

    if !paras.is_empty() {
        // The paragraph page: run bounds, 7-byte entries, then each run's properties -- a
        // word count, the style index, and sprmPFInTable / sprmPFTtp.
        let mut page = vec![0u8; 512];
        let crun = paras.len();
        let mut fc = FC_MIN as u32;
        for (i, &(len, ..)) in paras.iter().enumerate() {
            page[i * 4..i * 4 + 4].copy_from_slice(&fc.to_le_bytes());
            fc += len;
        }
        page[crun * 4..crun * 4 + 4].copy_from_slice(&fc.to_le_bytes());
        let mut papx_at = 510usize;
        for (i, &(_, istd, in_table, ttp)) in paras.iter().enumerate() {
            let mut papx = istd.to_le_bytes().to_vec();
            if in_table {
                papx.extend_from_slice(&[24, 1]);
            }
            if ttp {
                papx.extend_from_slice(&[25, 1]);
            }
            let words = papx.len().div_ceil(2);
            papx.resize(words * 2, 0);
            papx_at -= 1 + papx.len();
            papx_at &= !1;
            page[papx_at] = words as u8;
            page[papx_at + 1..papx_at + 1 + papx.len()].copy_from_slice(&papx);
            page[4 * (crun + 1) + i * 7] = (papx_at / 2) as u8;
        }
        page[511] = crun as u8;
        let at = PAPX_PAGE * 512;
        wd[at..at + 512].copy_from_slice(&page);
        let mut plc = Vec::new();
        le32(&mut plc, FC_MIN as u32);
        le32(&mut plc, fc);
        le16(&mut plc, PAPX_PAGE as u16);
        wd[PAPX_PLC_AT..PAPX_PLC_AT + plc.len()].copy_from_slice(&plc);
        wd[0xC0..0xC4].copy_from_slice(&(PAPX_PLC_AT as u32).to_le_bytes());
        wd[0xC4..0xC8].copy_from_slice(&(plc.len() as u32).to_le_bytes());

        // The style sheet: a 14-byte header (cstd 2, an 8-byte style base), "Normal" (sti 0)
        // and "heading 1" (sti 1), each a style base and a length-prefixed name.
        let mut stsh = Vec::new();
        le16(&mut stsh, 14);
        le16(&mut stsh, 2);
        le16(&mut stsh, 8);
        stsh.extend_from_slice(&[0; 10]);
        for (sti, name) in [(0u16, "Normal"), (1, "heading 1")] {
            let mut std = sti.to_le_bytes().to_vec();
            std.extend_from_slice(&[0; 6]);
            std.push(name.len() as u8);
            std.extend_from_slice(name.as_bytes());
            std.push(0);
            le16(&mut stsh, std.len() as u16);
            stsh.extend_from_slice(&std);
        }
        wd[STSH_AT..STSH_AT + stsh.len()].copy_from_slice(&stsh);
        wd[0x60..0x64].copy_from_slice(&(STSH_AT as u32).to_le_bytes());
        wd[0x64..0x68].copy_from_slice(&(stsh.len() as u32).to_le_bytes());
    }

    let mut container = cfb::CompoundFile::create(std::io::Cursor::new(Vec::new())).unwrap();
    let mut stream = container.create_stream("/WordDocument").unwrap();
    std::io::Write::write_all(&mut stream, &wd).unwrap();
    drop(stream);
    container.flush().unwrap();
    container.into_inner().into_inner()
}

fn word6_paragraphs(bytes: Vec<u8>) -> Vec<String> {
    paragraphs(&DocParser::from_bytes(bytes).unwrap().parse().unwrap())
}

#[test]
fn a_word_6_document_is_read_for_its_paragraphs() {
    let text = b"The quick brown fox jumps over the lazy dog\r\rParagraph 2\r";
    assert_eq!(
        word6_paragraphs(build_word6(text, 101, &[0], None)),
        ["The quick brown fox jumps over the lazy dog", "Paragraph 2"]
    );
}

#[test]
fn a_word_95_document_keeps_a_fields_result_and_drops_its_instruction() {
    let text = b"Page \x13 PAGE \x147\x15 of the report\r";
    assert_eq!(
        word6_paragraphs(build_word6(text, 104, &[0], None)),
        ["Page 7 of the report"]
    );
}

#[test]
fn a_fast_saved_word_6_document_is_read_in_its_pieces_order() {
    // Two pieces stored in the opposite order of the text they make.
    let text = b"second half.\rFirst half, ";
    let pieces = [(13u32, 13u32), (13, 0)];
    assert_eq!(
        word6_paragraphs(build_word6(text, 101, &[0], Some(&pieces))),
        ["First half, second half."]
    );
}

/// A paragraph in the heading 1 style is a heading, as in a Word 97 document.
#[test]
fn a_word_6_heading_style_makes_a_heading() {
    let text = b"Overview\rBody text.\r";
    let doc = DocParser::from_bytes(build_word6_with(
        text,
        101,
        &[0],
        None,
        &[(9, 1, false, false), (11, 0, false, false)],
    ))
    .unwrap()
    .parse()
    .unwrap();
    let paras: Vec<(&str, crate::model::HeadingLevel)> = doc.sections[0]
        .content
        .iter()
        .filter_map(|b| match b {
            Block::Paragraph(p) => Some((p.runs[0].text.as_str(), p.heading)),
            _ => None,
        })
        .collect();
    assert_eq!(
        paras,
        [
            ("Overview", crate::model::HeadingLevel::H1),
            ("Body text.", crate::model::HeadingLevel::None)
        ]
    );
}

/// A Word 6.0 table -- cells ended by cell marks, a row ended by a row-end mark -- is a table.
#[test]
fn a_word_6_table_is_a_table() {
    let text = b"Name\x07Qty\x07\x07Pen\x072\x07\x07After.\r";
    let doc = DocParser::from_bytes(build_word6_with(
        text,
        101,
        &[0],
        None,
        &[
            (5, 0, true, false),
            (4, 0, true, false),
            (1, 0, true, true),
            (4, 0, true, false),
            (2, 0, true, false),
            (1, 0, true, true),
            (7, 0, false, false),
        ],
    ))
    .unwrap()
    .parse()
    .unwrap();
    let table = doc.sections[0]
        .content
        .iter()
        .find_map(|b| match b {
            Block::Table(t) => Some(t),
            _ => None,
        })
        .expect("a table");
    let cells: Vec<Vec<String>> = table
        .rows
        .iter()
        .map(|r| r.cells.iter().map(|c| c.plain_text()).collect())
        .collect();
    assert_eq!(cells, [["Name", "Qty"], ["Pen", "2"]]);
    assert_eq!(paragraphs(&doc), ["After."]);
}

/// The text is in the code page the fonts' character set names: here Cyrillic.
#[cfg(feature = "codepages")]
#[test]
fn word_6_text_is_decoded_in_the_code_page_its_fonts_name() {
    // "привет" in Windows-1251; a symbol font first, which names no code page.
    let text = b"\xef\xf0\xe8\xe2\xe5\xf2\r";
    assert_eq!(
        word6_paragraphs(build_word6(text, 101, &[2, 204], None)),
        ["\u{43F}\u{440}\u{438}\u{432}\u{435}\u{442}"]
    );
}

/// Without a font that names one, the text is Windows-1252.
#[test]
fn word_6_text_without_a_named_code_page_is_windows_1252() {
    let text = b"caf\xe9 \x93quoted\x94\r";
    assert_eq!(
        word6_paragraphs(build_word6(text, 101, &[0], None)),
        ["caf\u{e9} \u{201C}quoted\u{201D}"]
    );
}

#[test]
fn the_library_entry_points_read_a_doc() {
    let bytes = build(&[heading("Title"), p("Body text.")], Options::default());
    let doc = crate::parse_bytes(&bytes).unwrap();
    let markdown =
        crate::render::to_markdown(&doc, &crate::render::RenderOptions::default()).unwrap();
    assert!(markdown.contains("# Title"), "{markdown}");
    assert!(markdown.contains("Body text."), "{markdown}");
}

#[test]
fn character_formatting_becomes_run_styles() {
    let doc = parse(
        &[p("plain bold italic x2 done")],
        Options {
            formats: vec![
                ("bold", vec![0x35, 0x08, 0x01]),
                ("italic", vec![0x36, 0x08, 0x01]),
                ("2", vec![0x48, 0x2A, 0x01]),
            ],
            ..Default::default()
        },
    )
    .unwrap();
    let Block::Paragraph(para) = &doc.sections[0].content[0] else {
        panic!()
    };
    let styled = |t: &str| {
        para.runs
            .iter()
            .find(|r| r.text == t)
            .map(|r| r.style.clone())
    };
    assert!(
        styled("bold").is_some_and(|s| s.bold && !s.italic),
        "{:?}",
        para.runs
    );
    assert!(
        styled("italic").is_some_and(|s| s.italic && !s.bold),
        "{:?}",
        para.runs
    );
    assert!(
        styled("2").is_some_and(|s| s.superscript),
        "{:?}",
        para.runs
    );
    assert_eq!(para.plain_text(), "plain bold italic x2 done");
}

#[test]
fn hidden_text_is_not_content() {
    let doc = parse(
        &[p("shown secret shown")],
        Options {
            formats: vec![("secret ", vec![0x3C, 0x08, 0x01])],
            ..Default::default()
        },
    )
    .unwrap();
    assert_eq!(paragraphs(&doc), ["shown shown"]);
}

#[test]
fn a_symbol_character_reads_as_the_symbol() {
    // Word stores a Symbol-font character as a placeholder with sprmCSymbol naming it.
    let doc = parse(
        &[p("a ( b")],
        Options {
            formats: vec![("(", vec![0x09, 0x6A, 0x00, 0x00, 0xB1, 0x00])],
            ..Default::default()
        },
    )
    .unwrap();
    assert_eq!(paragraphs(&doc), ["a \u{B1} b"]);
}

#[test]
fn footnotes_are_referenced_inline_and_written_at_the_end() {
    let doc = parse(
        &[p("Claim\u{02} and another\u{02}."), p("After.")],
        Options {
            footnotes: vec!["First source.", "Second\rsource."],
            ..Default::default()
        },
    )
    .unwrap();
    assert_eq!(paragraphs(&doc), ["Claim[^1] and another[^2].", "After."]);
    // The note bodies close the section as notes, not as paragraphs that spell out `[^n]:`.
    let notes: Vec<_> = doc.sections[0]
        .content
        .iter()
        .filter_map(|block| match block {
            Block::Note { label, content } => Some((
                label.as_str(),
                content.iter().map(|p| p.plain_text()).collect::<Vec<_>>(),
            )),
            _ => None,
        })
        .collect();
    assert_eq!(
        notes,
        [
            ("1", vec!["First source.".to_string()]),
            ("2", vec!["Second source.".to_string()]),
        ]
    );
}

#[test]
fn numbered_list_items_count_per_level() {
    let doc = parse(
        &[
            p("Steps:"),
            item("Mix", 0),
            item("Stir", 1),
            item("Fold", 1),
            item("Bake", 0),
            p("Done."),
        ],
        Options::default(),
    )
    .unwrap();
    /// A paragraph's text, and its list level and number when it is a list item.
    type Item = (String, Option<(u8, Option<u32>)>);
    let lists: Vec<Item> = doc.sections[0]
        .content
        .iter()
        .filter_map(|b| match b {
            Block::Paragraph(p) => Some((
                p.plain_text(),
                p.list_info.as_ref().map(|l| (l.level, l.number)),
            )),
            _ => None,
        })
        .collect();
    assert_eq!(
        lists,
        [
            ("Steps:".into(), None),
            ("Mix".into(), Some((0, Some(1)))),
            ("Stir".into(), Some((1, Some(1)))),
            ("Fold".into(), Some((1, Some(2)))),
            ("Bake".into(), Some((0, Some(2)))),
            ("Done.".into(), None),
        ]
    );
}

#[test]
fn a_bulleted_list_has_no_numbers() {
    let doc = parse(
        &[item("one", 0), item("two", 0)],
        Options {
            list_nfc: 0x17,
            ..Default::default()
        },
    )
    .unwrap();
    let Block::Paragraph(first) = &doc.sections[0].content[0] else {
        panic!()
    };
    let info = first.list_info.as_ref().expect("a list item");
    assert_eq!(info.list_type, crate::model::ListType::Bullet);
    assert_eq!(info.number, None);
}

/// A document whose one picture is a compressed EMF: `b"packed"` standing for the zlib
/// stream, `b"original"` for what it inflates to.
fn compressed_emf_document() -> Vec<u8> {
    let mut data = vec![0u8; 8];
    data.extend(super::pictures::tests::picf_with_emf(
        b"original",
        b"packed",
    ));
    build(
        &[p("Figure: \u{01}")],
        Options {
            formats: vec![("\u{01}", vec![0x03, 0x6A, 8, 0, 0, 0, 0x55, 0x08, 0x01])],
            data: Some(data),
            ..Default::default()
        },
    )
}

/// The two-pass protocol a host with only an asynchronous decompressor relies on: a first
/// parse reports the compressed pictures, a second one is handed them inflated.
#[test]
fn compressed_pictures_can_be_supplied_after_a_first_pass_reports_them() {
    let mut parser = DocParser::from_bytes(compressed_emf_document()).unwrap();
    let seen = std::cell::RefCell::new(Vec::new());
    let first = parser
        .parse_with(&|body: &[u8]| {
            seen.borrow_mut().push(body.to_vec());
            None
        })
        .unwrap();
    assert_eq!(*seen.borrow(), vec![b"packed".to_vec()]);
    assert!(
        first.resources.is_empty(),
        "an uninflated picture is left out"
    );

    let second = parser
        .parse_with(&|body: &[u8]| (body == b"packed").then(|| b"original".to_vec()))
        .unwrap();
    let resource = second.resources.values().next().expect("the picture");
    assert_eq!(resource.data, b"original");
    assert_eq!(resource.filename.as_deref(), Some("image1.emf"));
}

#[test]
fn an_inline_picture_becomes_a_resource_the_paragraph_references() {
    let png = b"\x89PNG\r\n\x1a\npixels".to_vec();
    let mut data = vec![0u8; 8];
    data.extend(super::pictures::tests::picf_with_png(
        &png,
        Some("Sales by region"),
    ));
    let doc = parse(
        &[p("Figure: \u{01}"), p("After.")],
        Options {
            // sprmCPicLocation 8, then sprmCFSpec.
            formats: vec![("\u{01}", vec![0x03, 0x6A, 8, 0, 0, 0, 0x55, 0x08, 0x01])],
            data: Some(data),
            ..Default::default()
        },
    )
    .unwrap();
    let Block::Paragraph(para) = &doc.sections[0].content[0] else {
        panic!()
    };
    assert_eq!(para.plain_text(), "Figure: ");
    assert_eq!(para.images.len(), 1);
    let image = &para.images[0];
    assert_eq!(image.alt_text.as_deref(), Some("Sales by region"));
    let resource = &doc.resources[&image.resource_id];
    assert_eq!(resource.data, png);
    assert_eq!(resource.filename.as_deref(), Some("image1.png"));
    assert_eq!(resource.mime_type.as_deref(), Some("image/png"));
}
