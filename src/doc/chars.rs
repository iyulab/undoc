//! Character properties ([MS-DOC] 2.9.27 `PlcBteChpx`, 2.9.34 `ChpxFkp`): the formatting the
//! Markdown can carry, and the properties that decide whether text is shown at all.

use super::fib::FcLcb;
use super::props::for_each_sprm;
use crate::error::{Error, Result};

const FKP_SIZE: usize = 512;

const SPRM_C_F_R_MARK_DEL: u16 = 0x0800;
const SPRM_C_F_R_MARK_INS: u16 = 0x0801;
const SPRM_C_F_BOLD: u16 = 0x0835;
const SPRM_C_F_ITALIC: u16 = 0x0836;
const SPRM_C_F_STRIKE: u16 = 0x0837;
const SPRM_C_F_VANISH: u16 = 0x083C;
const SPRM_C_KUL: u16 = 0x2A3E;
const SPRM_C_ISS: u16 = 0x2A48;
const SPRM_C_F_D_STRIKE: u16 = 0x2A53;
const SPRM_C_SYMBOL: u16 = 0x6A09;

/// The character properties of a run of text.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(super) struct CharProps {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
    pub superscript: bool,
    pub subscript: bool,
    /// Hidden text — not shown, not document content.
    pub hidden: bool,
    /// Tracked deletion and insertion.
    pub deleted: bool,
    pub inserted: bool,
    /// A symbol-font character: the font index and the code Word stores for it.
    pub symbol: Option<(u16, u16)>,
}

/// A toggle operand ([MS-DOC] 2.9.311 `ToggleOperand`): 0 off, 1 on, and 0x80 / 0x81 "as the
/// style" / "opposite of the style". Styles' character properties are not resolved here, so
/// the style is taken as off: 0x80 reads off and 0x81 on.
fn toggle(op: &[u8]) -> bool {
    matches!(op[0], 1 | 0x81)
}

fn apply_chpx(props: &mut CharProps, grpprl: &[u8]) {
    for_each_sprm(grpprl, |sprm, op| match sprm {
        SPRM_C_F_BOLD => props.bold = toggle(op),
        SPRM_C_F_ITALIC => props.italic = toggle(op),
        SPRM_C_F_STRIKE | SPRM_C_F_D_STRIKE => props.strike = toggle(op),
        SPRM_C_F_VANISH => props.hidden = toggle(op),
        SPRM_C_F_R_MARK_DEL => props.deleted = toggle(op),
        SPRM_C_F_R_MARK_INS => props.inserted = toggle(op),
        SPRM_C_KUL => props.underline = op[0] != 0,
        SPRM_C_ISS => {
            props.superscript = op[0] == 1;
            props.subscript = op[0] == 2;
        }
        SPRM_C_SYMBOL => {
            let font = u16::from_le_bytes([op[0], op[1]]);
            let xchar = u16::from_le_bytes([op[2], op[3]]);
            props.symbol = Some((font, xchar));
        }
        _ => {}
    });
}

/// Character properties by stream offset: `[fc_start, fc_end)` runs, sorted.
#[derive(Debug, Default)]
pub(super) struct ChpxIndex {
    runs: Vec<(u32, u32, CharProps)>,
}

impl ChpxIndex {
    /// The properties of the character at stream offset `fc`; plain text where none are set.
    pub fn at(&self, fc: u32) -> CharProps {
        let i = self.runs.partition_point(|(start, _, _)| *start <= fc);
        match i.checked_sub(1).and_then(|i| self.runs.get(i)) {
            Some((start, end, props)) if *start <= fc && fc < *end => *props,
            _ => CharProps::default(),
        }
    }
}

fn invalid(what: &str) -> Error {
    Error::InvalidData(format!("Word document: {what}"))
}

/// Read every character property run from the `PlcBteChpx` and the pages it points at.
pub(super) fn read_chpx(word_document: &[u8], table: &[u8], plc: FcLcb) -> Result<ChpxIndex> {
    if plc.lcb == 0 {
        return Ok(ChpxIndex::default());
    }
    let start = plc.fc as usize;
    let data = table
        .get(start..start + plc.lcb as usize)
        .ok_or_else(|| invalid("the character property index lies outside the table stream"))?;
    if data.len() < 4 || (data.len() - 4) % 8 != 0 {
        return Err(invalid(
            "the character property index has an impossible size",
        ));
    }
    let n = (data.len() - 4) / 8;
    let mut runs = Vec::new();
    for i in 0..n {
        let at = 4 * (n + 1) + i * 4;
        let pn = u32::from_le_bytes(data[at..at + 4].try_into().unwrap()) & 0x003F_FFFF;
        let page_at = pn as usize * FKP_SIZE;
        let page = word_document
            .get(page_at..page_at + FKP_SIZE)
            .ok_or_else(|| invalid("a character property page lies past the end of the stream"))?;
        let crun = page[FKP_SIZE - 1] as usize;
        if 4 * (crun + 1) + crun > FKP_SIZE - 1 {
            continue;
        }
        for r in 0..crun {
            let fc = |k: usize| u32::from_le_bytes(page[k * 4..k * 4 + 4].try_into().unwrap());
            let offset = page[4 * (crun + 1) + r] as usize * 2;
            let mut props = CharProps::default();
            if offset != 0 {
                if let Some(&cb) = page.get(offset) {
                    if let Some(grpprl) = page.get(offset + 1..offset + 1 + cb as usize) {
                        apply_chpx(&mut props, grpprl);
                    }
                }
            }
            runs.push((fc(r), fc(r + 1), props));
        }
    }
    runs.sort_by_key(|(start, _, _)| *start);
    Ok(ChpxIndex { runs })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn toggles_and_values_are_read() {
        let grpprl = [
            0x35, 0x08, 0x01, // bold on
            0x36, 0x08, 0x81, // italic: opposite of the (unresolved) style
            0x48, 0x2A, 0x01, // superscript
            0x3C, 0x08, 0x00, // not hidden
            0x09, 0x6A, 0x02, 0x00, 0xB1, 0x00, // symbol: U+00B1
        ];
        let mut props = CharProps::default();
        apply_chpx(&mut props, &grpprl);
        assert!(props.bold && props.italic && props.superscript);
        assert!(!props.hidden && !props.subscript);
        assert_eq!(props.symbol, Some((2, 0xB1)));
    }

    #[test]
    fn an_unlisted_offset_is_plain_text() {
        let index = ChpxIndex {
            runs: vec![(
                10,
                20,
                CharProps {
                    bold: true,
                    ..Default::default()
                },
            )],
        };
        assert!(index.at(15).bold);
        assert_eq!(index.at(20), CharProps::default());
    }
}
