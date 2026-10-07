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
    for (name, data) in [("/WordDocument", &wd), ("/1Table", &table)] {
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

#[test]
fn word_95_is_named_as_unsupported() {
    let err = parse(
        &[p("old")],
        Options {
            n_fib: Some(104),
            ..Default::default()
        },
    )
    .unwrap_err();
    assert!(
        matches!(err, Error::UnsupportedFormat(ref m) if m.contains("Word 6.0/95")),
        "{err}"
    );
}

#[test]
fn a_word_6_ident_is_named_as_unsupported() {
    let mut bytes = build(&[p("old")], Options::default());
    // Rewrite the stream's wIdent in place: find the FIB by its Word 97 ident.
    let at = bytes
        .windows(4)
        .position(|w| w == [0xEC, 0xA5, 0xC1, 0x00])
        .expect("the FIB is in the container");
    bytes[at] = 0xDC;
    let err = DocParser::from_bytes(bytes).unwrap().parse().unwrap_err();
    assert!(
        matches!(err, Error::UnsupportedFormat(ref m) if m.contains("Word 6.0/95")),
        "{err}"
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
    assert_eq!(
        paragraphs(&doc),
        [
            "Claim[^1] and another[^2].",
            "After.",
            "[^1]: First source.",
            "[^2]: Second source.",
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
