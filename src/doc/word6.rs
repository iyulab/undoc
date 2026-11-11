//! What a Word 6.0 or Word 95 document keeps differently from Word 97: text one byte per
//! character in the document's code page, and the font table that names that code page.

use super::chars::{toggle, CharProps, ChpxIndex};
use super::fib::FcLcb;
use super::props::{PapxIndex, ParaProps, StyleInfo};
use super::text::{self, DocChar, Piece};
use crate::error::{Error, Result};

/// Windows `charset` values that name no code page of their own: `ANSI_CHARSET`,
/// `DEFAULT_CHARSET` and `SYMBOL_CHARSET`.
const ANSI_CHARSET: u8 = 0;
const DEFAULT_CHARSET: u8 = 1;
const SYMBOL_CHARSET: u8 = 2;

/// The code page the document's text is in: that of the first font whose character set names
/// one, else Windows-1252. Word 6.0 and Word 95 store no code page for the text; the fonts it
/// is set in say which script it is.
pub(super) fn code_page(word_document: &[u8], sttbf_ffn: FcLcb) -> u16 {
    let start = sttbf_ffn.fc as usize;
    let Some(table) = start
        .checked_add(sttbf_ffn.lcb as usize)
        .and_then(|end| word_document.get(start..end))
    else {
        return 1252;
    };
    font_charsets(table)
        .find(|&chs| !matches!(chs, ANSI_CHARSET | DEFAULT_CHARSET | SYMBOL_CHARSET))
        .and_then(crate::codepage::from_charset)
        .unwrap_or(1252)
}

/// The character set of each font in a Word 6.0 font table: a 16-bit total length, then
/// records each led by its length less one, with the character set at offset 4.
fn font_charsets(table: &[u8]) -> impl Iterator<Item = u8> + '_ {
    let mut at = 2;
    std::iter::from_fn(move || {
        let len = *table.get(at)? as usize + 1;
        let record = table.get(at..at + len)?;
        at += len;
        record.get(4).copied()
    })
}

/// The pieces of a fast-saved document. Their descriptors are laid out as Word 97's, but a
/// piece's offset is a plain byte offset -- all of the text is one byte per character -- and
/// never reaches the bit Word 97 marks one-byte pieces with.
pub(super) fn pieces(word_document: &[u8], clx: FcLcb) -> Result<Vec<Piece>> {
    let mut pieces = text::parse_clx(word_document, clx)?;
    for piece in &mut pieces {
        piece.compressed = true;
    }
    Ok(pieces)
}

/// The main document text, characters `[0, ccp_text)`, decoded from `codepage`.
///
/// A double-byte code page decodes a lead byte and its trail byte into one character, which
/// then stands at the position of the next byte; nothing is looked up by position in this
/// text, so the positions only need to keep their order.
pub(super) fn read_chars(
    word_document: &[u8],
    pieces: &[Piece],
    ccp_text: u32,
    codepage: u16,
) -> Result<Vec<DocChar>> {
    let mut out = Vec::with_capacity((ccp_text as usize).min(word_document.len()));
    for piece in pieces {
        let (start, end) = (piece.cp_start, piece.cp_end.min(ccp_text));
        if start >= end {
            continue;
        }
        let first = piece.fc as usize;
        let bytes = word_document
            .get(first..first + (end - start) as usize)
            .ok_or_else(|| {
                Error::InvalidData(
                    "Word document: a piece of text lies past the end of the WordDocument stream"
                        .into(),
                )
            })?;
        let text = crate::codepage::decode_lossy(bytes, codepage);
        for (i, ch) in text.chars().enumerate() {
            out.push(DocChar {
                ch,
                fc: (first + i) as u32,
                cp: start + i as u32,
            });
        }
    }
    Ok(out)
}

/// Formatted disk pages are 512 bytes, as in Word 97.
const FKP_SIZE: usize = 512;
/// A Word 6.0 paragraph page's entries: a one-byte offset and a 6-byte paragraph height.
const BX_SIZE: usize = 7;

/// Paragraph sprms that decide structure -- one byte each in Word 6.0, where Word 97 has two.
const SPRM_P_ISTD: u8 = 2;
const SPRM_P_F_IN_TABLE: u8 = 24;
const SPRM_P_F_TTP: u8 = 25;
/// `sprmPChgTabs`: its operand's size is not where the other variable ones keep it.
const SPRM_P_CHG_TABS: u8 = 23;

/// How a Word 6.0 sprm's operand is sized: a fixed number of bytes; a length byte after the
/// sprm, plus a fixed count; or a 16-bit length one more than the operand, plus a fixed count.
#[derive(Clone, Copy)]
enum Operand {
    Fix(u8),
    Var(u8),
    Var2(u8),
}

/// Every Word 6.0 sprm and how its operand is sized, as the Word for Windows 6.0 binary file
/// format defines them. A sprm missing here cannot be stepped over.
const SPRMS: &[(u8, Operand)] = &[
    (0, Operand::Fix(0)),    // a no-op
    (2, Operand::Fix(2)),    // sprmPIstd
    (3, Operand::Var(3)),    // sprmPIstdPermute
    (4, Operand::Fix(1)),    // sprmPIncLv1
    (5, Operand::Fix(1)),    // sprmPJc
    (6, Operand::Fix(1)),    // sprmPFSideBySide
    (7, Operand::Fix(1)),    // sprmPFKeep
    (8, Operand::Fix(1)),    // sprmPFKeepFollow
    (9, Operand::Fix(1)),    // sprmPPageBreakBefore
    (10, Operand::Fix(1)),   // sprmPBrcl
    (11, Operand::Fix(1)),   // sprmPBrcp
    (12, Operand::Var(0)),   // sprmPAnld
    (13, Operand::Fix(1)),   // sprmPNLvlAnm
    (14, Operand::Fix(1)),   // sprmPFNoLineNumb
    (15, Operand::Var(0)),   // sprmPChgTabsPapx
    (16, Operand::Fix(2)),   // sprmPDxaRight
    (17, Operand::Fix(2)),   // sprmPDxaLeft
    (18, Operand::Fix(2)),   // sprmPNest
    (19, Operand::Fix(2)),   // sprmPDxaLeft1
    (20, Operand::Fix(4)),   // sprmPDyaLine
    (21, Operand::Fix(2)),   // sprmPDyaBefore
    (22, Operand::Fix(2)),   // sprmPDyaAfter
    (23, Operand::Var(0)),   // sprmPChgTabs
    (24, Operand::Fix(1)),   // sprmPFInTable
    (25, Operand::Fix(1)),   // sprmPTtp
    (26, Operand::Fix(2)),   // sprmPDxaAbs
    (27, Operand::Fix(2)),   // sprmPDyaAbs
    (28, Operand::Fix(2)),   // sprmPDxaWidth
    (29, Operand::Fix(1)),   // sprmPPc
    (30, Operand::Fix(2)),   // sprmPBrcTop10
    (31, Operand::Fix(2)),   // sprmPBrcLeft10
    (32, Operand::Fix(2)),   // sprmPBrcBottom10
    (33, Operand::Fix(2)),   // sprmPBrcRight10
    (34, Operand::Fix(2)),   // sprmPBrcBetween10
    (35, Operand::Fix(2)),   // sprmPBrcBar10
    (36, Operand::Fix(2)),   // sprmPFromText10
    (37, Operand::Fix(1)),   // sprmPWr
    (38, Operand::Fix(2)),   // sprmPBrcTop
    (39, Operand::Fix(2)),   // sprmPBrcLeft
    (40, Operand::Fix(2)),   // sprmPBrcBottom
    (41, Operand::Fix(2)),   // sprmPBrcRight
    (42, Operand::Fix(2)),   // sprmPBrcBetween
    (43, Operand::Fix(2)),   // sprmPBrcBar
    (44, Operand::Fix(1)),   // sprmPFNoAutoHyph
    (45, Operand::Fix(2)),   // sprmPWHeightAbs
    (46, Operand::Fix(2)),   // sprmPDcs
    (47, Operand::Fix(2)),   // sprmPShd
    (48, Operand::Fix(2)),   // sprmPDyaFromText
    (49, Operand::Fix(2)),   // sprmPDxaFromText
    (50, Operand::Fix(1)),   // sprmPFLocked
    (51, Operand::Fix(1)),   // sprmPFWidowControl
    (52, Operand::Fix(0)),   // sprmPRuler
    (65, Operand::Fix(1)),   // sprmCFStrikeRM
    (66, Operand::Fix(1)),   // sprmCFRMark
    (67, Operand::Fix(1)),   // sprmCFFldVanish
    (68, Operand::Var(0)),   // sprmCPicLocation
    (69, Operand::Fix(2)),   // sprmCIbstRMark
    (70, Operand::Fix(4)),   // sprmCDttmRMark
    (71, Operand::Fix(1)),   // sprmCFData
    (72, Operand::Fix(2)),   // sprmCRMReason
    (73, Operand::Fix(3)),   // sprmCChse
    (74, Operand::Var(0)),   // sprmCSymbol
    (75, Operand::Fix(1)),   // sprmCFOle2
    (80, Operand::Fix(2)),   // sprmCIstd
    (81, Operand::Var(0)),   // sprmCIstdPermute
    (82, Operand::Var(0)),   // sprmCDefault
    (83, Operand::Fix(0)),   // sprmCPlain
    (85, Operand::Fix(1)),   // sprmCFBold
    (86, Operand::Fix(1)),   // sprmCFItalic
    (87, Operand::Fix(1)),   // sprmCFStrike
    (88, Operand::Fix(1)),   // sprmCFOutline
    (89, Operand::Fix(1)),   // sprmCFShadow
    (90, Operand::Fix(1)),   // sprmCFSmallCaps
    (91, Operand::Fix(1)),   // sprmCFCaps
    (92, Operand::Fix(1)),   // sprmCFVanish
    (93, Operand::Fix(2)),   // sprmCFtc
    (94, Operand::Fix(1)),   // sprmCKul
    (95, Operand::Fix(3)),   // sprmCSizePos
    (96, Operand::Fix(2)),   // sprmCDxaSpace
    (97, Operand::Fix(2)),   // sprmCLid
    (98, Operand::Fix(1)),   // sprmCIco
    (99, Operand::Fix(2)),   // sprmCHps
    (100, Operand::Fix(1)),  // sprmCHpsInc
    (101, Operand::Fix(2)),  // sprmCHpsPos
    (102, Operand::Fix(1)),  // sprmCHpsPosAdj
    (103, Operand::Var(0)),  // sprmCMajority
    (104, Operand::Fix(1)),  // sprmCIss
    (105, Operand::Var(0)),  // sprmCHpsNew50
    (106, Operand::Var(0)),  // sprmCHpsInc1
    (107, Operand::Fix(2)),  // sprmCHpsKern
    (108, Operand::Var(0)),  // sprmCMajority50
    (109, Operand::Fix(2)),  // sprmCHpsMul
    (110, Operand::Fix(2)),  // sprmCCondHyhen
    (117, Operand::Fix(1)),  // sprmCFSpec
    (118, Operand::Fix(1)),  // sprmCFObj
    (119, Operand::Fix(1)),  // sprmPicBrcl
    (120, Operand::Var(12)), // sprmPicScale
    (121, Operand::Fix(2)),  // sprmPicBrcTop
    (122, Operand::Fix(2)),  // sprmPicBrcLeft
    (123, Operand::Fix(2)),  // sprmPicBrcBottom
    (124, Operand::Fix(2)),  // sprmPicBrcRight
    (131, Operand::Fix(1)),  // sprmSScnsPgn
    (132, Operand::Fix(1)),  // sprmSiHeadingPgn
    (133, Operand::Var(0)),  // sprmSOlstAnm
    (136, Operand::Fix(3)),  // sprmSDxaColWidth
    (137, Operand::Fix(3)),  // sprmSDxaColSpacing
    (138, Operand::Fix(1)),  // sprmSFEvenlySpaced
    (139, Operand::Fix(1)),  // sprmSFProtected
    (140, Operand::Fix(2)),  // sprmSDmBinFirst
    (141, Operand::Fix(2)),  // sprmSDmBinOther
    (142, Operand::Fix(1)),  // sprmSBkc
    (143, Operand::Fix(1)),  // sprmSFTitlePage
    (144, Operand::Fix(2)),  // sprmSCcolumns
    (145, Operand::Fix(2)),  // sprmSDxaColumns
    (146, Operand::Fix(1)),  // sprmSFAutoPgn
    (147, Operand::Fix(1)),  // sprmSNfcPgn
    (148, Operand::Fix(2)),  // sprmSDyaPgn
    (149, Operand::Fix(2)),  // sprmSDxaPgn
    (150, Operand::Fix(1)),  // sprmSFPgnRestart
    (151, Operand::Fix(1)),  // sprmSFEndnote
    (152, Operand::Fix(1)),  // sprmSLnc
    (153, Operand::Fix(1)),  // sprmSGprfIhdt
    (154, Operand::Fix(2)),  // sprmSNLnnMod
    (155, Operand::Fix(2)),  // sprmSDxaLnn
    (156, Operand::Fix(2)),  // sprmSDyaHdrTop
    (157, Operand::Fix(2)),  // sprmSDyaHdrBottom
    (158, Operand::Fix(1)),  // sprmSLBetween
    (159, Operand::Fix(1)),  // sprmSVjc
    (160, Operand::Fix(2)),  // sprmSLnnMin
    (161, Operand::Fix(2)),  // sprmSPgnStart
    (162, Operand::Fix(1)),  // sprmSBOrientation
    (163, Operand::Fix(0)),  // sprmSBCustomize
    (164, Operand::Fix(2)),  // sprmSXaPage
    (165, Operand::Fix(2)),  // sprmSYaPage
    (166, Operand::Fix(2)),  // sprmSDxaLeft
    (167, Operand::Fix(2)),  // sprmSDxaRight
    (168, Operand::Fix(2)),  // sprmSDyaTop
    (169, Operand::Fix(2)),  // sprmSDyaBottom
    (170, Operand::Fix(2)),  // sprmSDzaGutter
    (171, Operand::Fix(2)),  // sprmSDMPaperReq
    (182, Operand::Fix(2)),  // sprmTJc
    (183, Operand::Fix(2)),  // sprmTDxaLeft
    (184, Operand::Fix(2)),  // sprmTDxaGapHalf
    (185, Operand::Fix(1)),  // sprmTFCantSplit
    (186, Operand::Fix(1)),  // sprmTTableHeader
    (187, Operand::Fix(12)), // sprmTTableBorders
    (188, Operand::Var2(0)), // sprmTDefTable10
    (189, Operand::Fix(2)),  // sprmTDyaRowHeight
    (190, Operand::Var2(0)), // sprmTDefTable
    (191, Operand::Var(1)),  // sprmTDefTableShd
    (192, Operand::Fix(4)),  // sprmTTlp
    (193, Operand::Fix(5)),  // sprmTSetBrc
    (194, Operand::Fix(4)),  // sprmTInsert
    (195, Operand::Fix(2)),  // sprmTDelete
    (196, Operand::Fix(4)),  // sprmTDxaCol
    (197, Operand::Fix(2)),  // sprmTMerge
    (198, Operand::Fix(2)),  // sprmTSplit
    (199, Operand::Fix(5)),  // sprmTSetBrc10
    (200, Operand::Fix(4)),  // sprmTSetShd
];

fn operand(sprm: u8) -> Option<Operand> {
    SPRMS
        .iter()
        .find(|(code, _)| *code == sprm)
        .map(|&(_, operand)| operand)
}

/// Walk a Word 6.0 `grpprl`, calling `visit` with each sprm and its operand. An unknown sprm
/// or an operand running past the end stops the walk: what was read stands.
fn for_each_sprm(grpprl: &[u8], mut visit: impl FnMut(u8, &[u8])) {
    let mut at = 0;
    while let Some(&sprm) = grpprl.get(at) {
        at += 1;
        let (operand_at, len) = match (sprm, operand(sprm)) {
            (SPRM_P_CHG_TABS, _) => match grpprl.get(at) {
                Some(&cb) => (at + 1, cb as usize),
                None => return,
            },
            (_, Some(Operand::Fix(n))) => (at, n as usize),
            (_, Some(Operand::Var(n))) => match grpprl.get(at) {
                Some(&cb) => (at + 1, cb as usize + n as usize),
                None => return,
            },
            (_, Some(Operand::Var2(n))) => match grpprl.get(at..at + 2) {
                Some(b) => (
                    at + 2,
                    (u16::from_le_bytes([b[0], b[1]]) as usize).saturating_sub(1) + n as usize,
                ),
                None => return,
            },
            (_, None) => return,
        };
        let Some(operand) = grpprl.get(operand_at..operand_at + len) else {
            return;
        };
        visit(sprm, operand);
        at = operand_at + len;
    }
}

/// The paragraph properties of a Word 6.0 document: its `PlcfBtePapx` (stream offsets, then
/// 16-bit page numbers) and the pages it points at. A damaged index or page gives no
/// properties rather than an error -- they decide structure, never whether text is recovered.
pub(super) fn read_papx(word_document: &[u8], plc: FcLcb) -> PapxIndex {
    let start = plc.fc as usize;
    let Some(data) = start
        .checked_add(plc.lcb as usize)
        .and_then(|end| word_document.get(start..end))
    else {
        return PapxIndex::default();
    };
    if data.len() < 4 || (data.len() - 4) % 6 != 0 {
        return PapxIndex::default();
    }
    let n = (data.len() - 4) / 6;
    let mut runs = Vec::new();
    for i in 0..n {
        let at = 4 * (n + 1) + i * 2;
        let pn = u16::from_le_bytes([data[at], data[at + 1]]) as usize;
        if let Some(page) = word_document.get(pn * FKP_SIZE..(pn + 1) * FKP_SIZE) {
            read_papx_fkp(page, &mut runs);
        }
    }
    runs.sort_by_key(|(start, _, _)| *start);
    PapxIndex::from_runs(runs)
}

/// One paragraph page: `crun` runs bounded by `crun + 1` stream offsets, then 7-byte entries
/// whose first byte, doubled, is where the run's properties are -- a count of 16-bit words,
/// the style index, and the sprms.
fn read_papx_fkp(page: &[u8], runs: &mut Vec<(u32, u32, ParaProps)>) {
    let crun = page[FKP_SIZE - 1] as usize;
    if 4 * (crun + 1) + BX_SIZE * crun > FKP_SIZE - 1 {
        return;
    }
    let fc = |i: usize| u32::from_le_bytes(page[i * 4..i * 4 + 4].try_into().unwrap());
    for i in 0..crun {
        let mut props = ParaProps::default();
        let at = page[4 * (crun + 1) + i * BX_SIZE] as usize * 2;
        if at != 0 {
            let words = page.get(at).copied().unwrap_or(0) as usize;
            if let Some(body) = page
                .get(at + 1..at + 1 + words * 2)
                .filter(|b| b.len() >= 2)
            {
                props.istd = u16::from_le_bytes([body[0], body[1]]);
                for_each_sprm(&body[2..], |sprm, op| match sprm {
                    SPRM_P_ISTD if op.len() == 2 => {
                        props.istd = u16::from_le_bytes([op[0], op[1]]);
                    }
                    SPRM_P_F_IN_TABLE => props.in_table = op.first().is_some_and(|&b| b != 0),
                    SPRM_P_F_TTP => props.row_end = op.first().is_some_and(|&b| b != 0),
                    _ => {}
                });
                if props.in_table {
                    props.depth = 1;
                }
            }
        }
        runs.push((fc(i), fc(i + 1), props));
    }
}

/// The style sheet's identifiers and names, indexed by `istd`. Its layout is Word 97's but
/// for the names: one byte per character after a length byte, in the document's code page.
pub(super) fn read_styles(word_document: &[u8], stsh: FcLcb, codepage: u16) -> Vec<StyleInfo> {
    let start = stsh.fc as usize;
    let Some(data) = start
        .checked_add(stsh.lcb as usize)
        .and_then(|end| word_document.get(start..end))
    else {
        return Vec::new();
    };
    let u16_at = |at: usize| {
        data.get(at..at + 2)
            .map(|b| u16::from_le_bytes([b[0], b[1]]))
    };
    let (Some(cb_stshi), Some(cstd), Some(cb_base)) = (u16_at(0), u16_at(2), u16_at(4)) else {
        return Vec::new();
    };
    let cb_base = cb_base as usize;
    let mut styles = Vec::with_capacity(cstd as usize);
    let mut at = 2 + cb_stshi as usize;
    for _ in 0..cstd {
        let Some(cb_std) = u16_at(at) else {
            break;
        };
        let std = data.get(at + 2..at + 2 + cb_std as usize).unwrap_or(&[]);
        at += 2 + cb_std as usize;
        if std.len() < 2 {
            styles.push(StyleInfo::default());
            continue;
        }
        let sti = u16::from_le_bytes([std[0], std[1]]) & 0x0FFF;
        let name = std
            .get(cb_base)
            .and_then(|&len| std.get(cb_base + 1..cb_base + 1 + len as usize))
            .map(|bytes| crate::codepage::decode_lossy(bytes, codepage))
            .unwrap_or_default();
        styles.push(StyleInfo { sti, name });
    }
    styles
}

/// Character sprms -- one byte each in Word 6.0.
const SPRM_C_F_STRIKE_RM: u8 = 65;
const SPRM_C_F_R_MARK: u8 = 66;
const SPRM_C_F_BOLD: u8 = 85;
const SPRM_C_F_ITALIC: u8 = 86;
const SPRM_C_F_STRIKE: u8 = 87;
const SPRM_C_F_VANISH: u8 = 92;
const SPRM_C_KUL: u8 = 94;
const SPRM_C_ISS: u8 = 104;

/// The character properties of a Word 6.0 document: its `PlcfBteChpx` (stream offsets, then
/// 16-bit page numbers) and the pages it points at, laid out as Word 97's but with one-byte
/// sprms. Like the paragraph properties, damage gives plain text rather than an error.
///
/// Runs are keyed by stream offset; in a double-byte code page the characters after a
/// two-byte one stand a little ahead of their bytes, so their formatting may start late.
pub(super) fn read_chpx(word_document: &[u8], plc: FcLcb) -> ChpxIndex {
    let start = plc.fc as usize;
    let Some(data) = start
        .checked_add(plc.lcb as usize)
        .and_then(|end| word_document.get(start..end))
    else {
        return ChpxIndex::default();
    };
    if data.len() < 4 || (data.len() - 4) % 6 != 0 {
        return ChpxIndex::default();
    }
    let n = (data.len() - 4) / 6;
    let mut runs = Vec::new();
    for i in 0..n {
        let at = 4 * (n + 1) + i * 2;
        let pn = u16::from_le_bytes([data[at], data[at + 1]]) as usize;
        let Some(page) = word_document.get(pn * FKP_SIZE..(pn + 1) * FKP_SIZE) else {
            continue;
        };
        let crun = page[FKP_SIZE - 1] as usize;
        if 4 * (crun + 1) + crun > FKP_SIZE - 1 {
            continue;
        }
        let fc = |k: usize| u32::from_le_bytes(page[k * 4..k * 4 + 4].try_into().unwrap());
        for r in 0..crun {
            let mut props = CharProps::default();
            let at = page[4 * (crun + 1) + r] as usize * 2;
            if at != 0 {
                let cb = page.get(at).copied().unwrap_or(0) as usize;
                if let Some(grpprl) = page.get(at + 1..at + 1 + cb) {
                    for_each_sprm(grpprl, |sprm, op| match sprm {
                        SPRM_C_F_BOLD => props.bold = toggle(op),
                        SPRM_C_F_ITALIC => props.italic = toggle(op),
                        SPRM_C_F_STRIKE => props.strike = toggle(op),
                        SPRM_C_F_VANISH => props.hidden = toggle(op),
                        SPRM_C_F_STRIKE_RM => props.deleted = toggle(op),
                        SPRM_C_F_R_MARK => props.inserted = toggle(op),
                        SPRM_C_KUL => props.underline = op.first().is_some_and(|&b| b != 0),
                        SPRM_C_ISS => {
                            props.superscript = op.first() == Some(&1);
                            props.subscript = op.first() == Some(&2);
                        }
                        _ => {}
                    });
                }
            }
            runs.push((fc(r), fc(r + 1), props));
        }
    }
    runs.sort_by_key(|(start, _, _)| *start);
    ChpxIndex::from_runs(runs)
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A stream holding one Word 6.0 character page (page 1) for `runs` -- (start, end, grpprl)
    /// -- and its index at offset 0.
    fn chpx_stream(runs: &[(u32, u32, &[u8])]) -> (Vec<u8>, FcLcb) {
        let mut stream = vec![0u8; 1024];
        let page = &mut stream[512..1024];
        let crun = runs.len();
        for (i, &(start, end, _)) in runs.iter().enumerate() {
            page[i * 4..i * 4 + 4].copy_from_slice(&start.to_le_bytes());
            page[(i + 1) * 4..(i + 1) * 4 + 4].copy_from_slice(&end.to_le_bytes());
        }
        let mut at = 510;
        for (i, &(_, _, grpprl)) in runs.iter().enumerate() {
            if grpprl.is_empty() {
                continue;
            }
            at = (at - 1 - grpprl.len()) & !1;
            page[at] = grpprl.len() as u8;
            page[at + 1..at + 1 + grpprl.len()].copy_from_slice(grpprl);
            page[4 * (crun + 1) + i] = (at / 2) as u8;
        }
        page[511] = crun as u8;
        let mut index = Vec::new();
        index.extend_from_slice(&runs[0].0.to_le_bytes());
        index.extend_from_slice(&runs[crun - 1].1.to_le_bytes());
        index.extend_from_slice(&1u16.to_le_bytes());
        stream[..index.len()].copy_from_slice(&index);
        let plc = FcLcb {
            fc: 0,
            lcb: index.len() as u32,
        };
        (stream, plc)
    }

    #[test]
    fn character_sprms_become_run_properties() {
        // Plain (a font only), bold and coloured, a superscript, "opposite of the style" bold
        // with an underline, and a run with no properties.
        let (stream, plc) = chpx_stream(&[
            (100, 110, &[93, 0, 0]),
            (110, 120, &[85, 1, 98, 6]),
            (120, 122, &[104, 1]),
            (122, 130, &[85, 0x81, 94, 1]),
            (130, 131, &[]),
        ]);
        let index = read_chpx(&stream, plc);
        assert_eq!(index.at(105), CharProps::default());
        assert!(index.at(115).bold);
        assert!(index.at(121).superscript);
        let toggled = index.at(125);
        assert!(toggled.bold && toggled.underline);
        assert_eq!(index.at(130), CharProps::default());
    }

    #[test]
    fn an_unknown_sprm_stops_the_walk_and_keeps_what_was_read() {
        // Bold, then a code no Word 6.0 sprm has, then italic, which is not reached.
        let (stream, plc) = chpx_stream(&[(100, 110, &[85, 1, 255, 86, 1])]);
        let props = read_chpx(&stream, plc).at(100);
        assert!(props.bold && !props.italic);
    }
}
