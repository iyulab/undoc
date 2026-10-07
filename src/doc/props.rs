//! Paragraph properties ([MS-DOC] 2.9.209 `PlcBtePapx`, 2.9.175 `PapxFkp`) and the style
//! sheet ([MS-DOC] 2.9.271 `STSH`) — just what decides structure: whether a paragraph is in a
//! table, whether it closes a table row, and whether it is a heading.

use super::fib::FcLcb;
use crate::error::{Error, Result};

/// Formatted disk pages are 512 bytes.
const FKP_SIZE: usize = 512;

const SPRM_P_OUT_LVL: u16 = 0x2640;
const SPRM_P_F_IN_TABLE: u16 = 0x2416;
const SPRM_P_F_TTP: u16 = 0x2417;
const SPRM_P_F_INNER_TTP: u16 = 0x244C;
const SPRM_P_ITAP: u16 = 0x6649;
const SPRM_P_ILVL: u16 = 0x260A;
const SPRM_P_ILFO: u16 = 0x460B;
const SPRM_T_DEF_TABLE: u16 = 0xD608;

/// The structural properties of one paragraph.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(super) struct ParaProps {
    /// Index of the paragraph's style in the style sheet.
    pub istd: u16,
    /// The paragraph is inside a table.
    pub in_table: bool,
    /// The paragraph is the row-end mark of a table row.
    pub row_end: bool,
    /// The paragraph is the row-end mark of a table nested in a cell.
    pub inner_row_end: bool,
    /// Table nesting depth (0 outside tables, 1 in a top-level table).
    pub depth: u32,
    /// Outline level set on the paragraph itself (0 = level 1; 9 = body text).
    pub outline_level: Option<u8>,
    /// List format override (1-based; 0 = not in a list) and list level.
    pub ilfo: u16,
    pub ilvl: u8,
}

/// Paragraph properties by stream offset: `[fc_start, fc_end)` runs, sorted.
#[derive(Debug, Default)]
pub(super) struct PapxIndex {
    runs: Vec<(u32, u32, ParaProps)>,
}

impl PapxIndex {
    /// The properties of the paragraph whose mark is at stream offset `fc`.
    pub fn at(&self, fc: u32) -> Option<&ParaProps> {
        let i = self.runs.partition_point(|(start, _, _)| *start <= fc);
        let (start, end, props) = self.runs.get(i.checked_sub(1)?)?;
        (*start <= fc && fc < *end).then_some(props)
    }
}

fn invalid(what: impl Into<String>) -> Error {
    Error::InvalidData(format!("Word document: {}", what.into()))
}

fn u16_at(data: &[u8], at: usize) -> Option<u16> {
    data.get(at..at + 2)
        .map(|b| u16::from_le_bytes([b[0], b[1]]))
}

fn u32_at(data: &[u8], at: usize) -> Option<u32> {
    data.get(at..at + 4)
        .map(|b| u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
}

/// Walk a `grpprl`, calling `visit` with each sprm and its operand.
///
/// The operand size is encoded in the sprm itself (`spra`, [MS-DOC] 2.2.5.1), except for the
/// variable-size kind, whose first byte is the size — and `sprmTDefTable`, whose size is a
/// 16-bit count. An operand running past the end stops the walk: the properties read so far
/// stand, and nothing past the damage is guessed at.
pub(super) fn for_each_sprm(grpprl: &[u8], mut visit: impl FnMut(u16, &[u8])) {
    let mut at = 0;
    while let Some(sprm) = u16_at(grpprl, at) {
        at += 2;
        let (operand_at, len) = match sprm >> 13 {
            0 | 1 => (at, 1),
            2 | 4 | 5 => (at, 2),
            3 => (at, 4),
            7 => (at, 3),
            _ if sprm == SPRM_T_DEF_TABLE => match u16_at(grpprl, at) {
                // The count includes one byte of itself beyond the 16-bit field.
                Some(cb) => (at + 2, (cb as usize).saturating_sub(1)),
                None => return,
            },
            _ => match grpprl.get(at) {
                Some(&cb) => (at + 1, cb as usize),
                None => return,
            },
        };
        let Some(operand) = grpprl.get(operand_at..operand_at + len) else {
            return;
        };
        visit(sprm, operand);
        at = operand_at + len;
    }
}

fn apply_papx(props: &mut ParaProps, grpprl: &[u8]) {
    for_each_sprm(grpprl, |sprm, op| match sprm {
        SPRM_P_F_IN_TABLE => props.in_table = op[0] != 0,
        SPRM_P_F_TTP => props.row_end = op[0] != 0,
        SPRM_P_F_INNER_TTP => props.inner_row_end = op[0] != 0,
        SPRM_P_ITAP => props.depth = u32::from_le_bytes([op[0], op[1], op[2], op[3]]),
        SPRM_P_OUT_LVL => props.outline_level = Some(op[0]),
        SPRM_P_ILVL => props.ilvl = op[0],
        SPRM_P_ILFO => props.ilfo = u16::from_le_bytes([op[0], op[1]]),
        _ => {}
    });
    // A paragraph marked as in a table without a depth is at depth one.
    if props.in_table && props.depth == 0 {
        props.depth = 1;
    }
}

/// Read every paragraph property run from the `PlcBtePapx` and the pages it points at.
pub(super) fn read_papx(word_document: &[u8], table: &[u8], plc: FcLcb) -> Result<PapxIndex> {
    if plc.lcb == 0 {
        return Ok(PapxIndex::default());
    }
    let start = plc.fc as usize;
    let data = table
        .get(start..start + plc.lcb as usize)
        .ok_or_else(|| invalid("the paragraph property index lies outside the table stream"))?;
    if data.len() < 4 || (data.len() - 4) % 8 != 0 {
        return Err(invalid(
            "the paragraph property index has an impossible size",
        ));
    }
    let n = (data.len() - 4) / 8;

    let mut runs = Vec::new();
    for i in 0..n {
        let pn = u32_at(data, 4 * (n + 1) + i * 4).unwrap() & 0x003F_FFFF;
        let page_at = pn as usize * FKP_SIZE;
        let page = word_document
            .get(page_at..page_at + FKP_SIZE)
            .ok_or_else(|| invalid("a paragraph property page lies past the end of the stream"))?;
        read_papx_fkp(page, &mut runs);
    }
    runs.sort_by_key(|(start, _, _)| *start);
    Ok(PapxIndex { runs })
}

/// One `PapxFkp`: `crun` paragraph runs bounded by `crun + 1` stream offsets.
fn read_papx_fkp(page: &[u8], runs: &mut Vec<(u32, u32, ParaProps)>) {
    let crun = page[FKP_SIZE - 1] as usize;
    // rgfc (crun + 1 offsets) and rgbx (crun 13-byte entries) must fit before the count byte.
    if 4 * (crun + 1) + 13 * crun > FKP_SIZE - 1 {
        return;
    }
    for i in 0..crun {
        let (fc_start, fc_end) = (
            u32_at(page, i * 4).unwrap(),
            u32_at(page, i * 4 + 4).unwrap(),
        );
        let b_offset = page[4 * (crun + 1) + i * 13] as usize * 2;
        let mut props = ParaProps::default();
        if b_offset != 0 {
            if let Some((istd, grpprl)) = papx_in_fkp(page, b_offset) {
                props.istd = istd;
                apply_papx(&mut props, grpprl);
            }
        }
        runs.push((fc_start, fc_end, props));
    }
}

/// `PapxInFkp` at `at`: a size, then `istd` and the `grpprl`.
fn papx_in_fkp(page: &[u8], at: usize) -> Option<(u16, &[u8])> {
    let cb = *page.get(at)? as usize;
    let (start, len) = if cb == 0 {
        (at + 2, 2 * *page.get(at + 1)? as usize)
    } else {
        (at + 1, 2 * cb - 1)
    };
    let body = page.get(start..start + len)?;
    let istd = u16_at(body, 0)?;
    Some((istd, &body[2..]))
}

/// A style's built-in identifier and name.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(super) struct StyleInfo {
    pub sti: u16,
    pub name: String,
}

/// Built-in style identifiers that mean "heading": `sti` 1–9 are "heading 1" … "heading 9".
const STI_TITLE: u16 = 62;

impl StyleInfo {
    /// The heading level this style confers, 1-based.
    pub fn heading_level(&self) -> Option<u8> {
        match self.sti {
            1..=9 => Some(self.sti as u8),
            STI_TITLE => Some(1),
            _ => None,
        }
    }
}

/// Read the style sheet's identifiers and names, indexed by `istd`. A style sheet that cannot be
/// read yields an empty table: styles decide headings, never whether text is recovered.
pub(super) fn read_styles(table: &[u8], stshf: FcLcb) -> Vec<StyleInfo> {
    let start = stshf.fc as usize;
    let Some(data) = table.get(start..start + stshf.lcb as usize) else {
        return Vec::new();
    };
    let Some(cb_stshi) = u16_at(data, 0) else {
        return Vec::new();
    };
    let (Some(cstd), Some(cb_base)) = (u16_at(data, 2), u16_at(data, 4)) else {
        return Vec::new();
    };
    let mut styles = Vec::with_capacity(cstd as usize);
    let mut at = 2 + cb_stshi as usize;
    for _ in 0..cstd {
        let Some(cb_std) = u16_at(data, at) else {
            break;
        };
        let std = data.get(at + 2..at + 2 + cb_std as usize).unwrap_or(&[]);
        at += 2 + cb_std as usize;
        if std.is_empty() {
            styles.push(StyleInfo::default());
            continue;
        }
        let sti = u16_at(std, 0).unwrap_or(0) & 0x0FFF;
        // The name follows the fixed base: a 16-bit length, UTF-16 characters, a terminator.
        let name = u16_at(std, cb_base as usize)
            .and_then(|len| {
                let from = cb_base as usize + 2;
                std.get(from..from + len as usize * 2)
            })
            .map(|bytes| {
                let units: Vec<u16> = bytes
                    .as_chunks::<2>()
                    .0
                    .iter()
                    .map(|&c| u16::from_le_bytes(c))
                    .collect();
                String::from_utf16_lossy(&units)
            })
            .unwrap_or_default();
        styles.push(StyleInfo { sti, name });
    }
    styles
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sprm_walk_reads_each_operand_size() {
        // sprmPFInTable (1 byte), sprmPItap (4 bytes), a variable one (size byte 3), sprmPFTtp.
        let grpprl = [
            0x16, 0x24, 0x01, //
            0x49, 0x66, 0x02, 0x00, 0x00, 0x00, //
            0x15, 0xC6, 0x03, 0xAA, 0xBB, 0xCC, //
            0x17, 0x24, 0x01,
        ];
        let mut props = ParaProps::default();
        apply_papx(&mut props, &grpprl);
        assert!(props.in_table);
        assert!(props.row_end);
        assert_eq!(props.depth, 2);
    }

    #[test]
    fn a_truncated_operand_stops_the_walk_without_panicking() {
        let mut seen = Vec::new();
        for_each_sprm(&[0x16, 0x24, 0x01, 0x49, 0x66, 0x02], |sprm, _| {
            seen.push(sprm)
        });
        assert_eq!(seen, [SPRM_P_F_IN_TABLE]);
    }

    #[test]
    fn lookup_finds_the_run_holding_an_offset() {
        let index = PapxIndex {
            runs: vec![
                (
                    100,
                    150,
                    ParaProps {
                        istd: 1,
                        ..Default::default()
                    },
                ),
                (
                    150,
                    300,
                    ParaProps {
                        istd: 2,
                        ..Default::default()
                    },
                ),
            ],
        };
        assert_eq!(index.at(100).map(|p| p.istd), Some(1));
        assert_eq!(index.at(149).map(|p| p.istd), Some(1));
        assert_eq!(index.at(150).map(|p| p.istd), Some(2));
        assert_eq!(index.at(300), None);
        assert_eq!(index.at(99), None);
    }
}
