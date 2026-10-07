//! Lists ([MS-DOC] 2.9.147 `PlfLst`, 2.9.146 `PlfLfo`): which list a paragraph belongs to, and
//! whether that list's level is numbered or bulleted.

use std::collections::HashMap;

use super::fib::FcLcb;
use crate::model::ListType;

/// `nfc` of a bulleted level, and of a level that shows no number at all.
const NFC_BULLET: u8 = 0x17;
const NFC_NONE: u8 = 0xFF;

/// One level of a list definition.
#[derive(Debug, Clone, Copy)]
struct Level {
    start: u32,
    nfc: u8,
}

/// The document's list definitions, reachable through its list format overrides.
#[derive(Debug, Default)]
pub(super) struct ListTable {
    /// List definitions by `lsid`.
    lists: HashMap<i32, Vec<Level>>,
    /// `lsid` of each list format override, by 1-based `ilfo`.
    overrides: Vec<i32>,
}

/// How a paragraph's list level is shown.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct ListLevel {
    pub lsid: i32,
    pub kind: ListType,
    pub start: u32,
}

impl ListTable {
    /// The list level a paragraph with `ilfo` / `ilvl` is in, if it is in one that shows a
    /// number or a bullet.
    pub fn level(&self, ilfo: u16, ilvl: u8) -> Option<ListLevel> {
        let lsid = *self.overrides.get((ilfo as usize).checked_sub(1)?)?;
        let levels = self.lists.get(&lsid)?;
        // A one-level list applies its only level at every depth.
        let level = levels.get(ilvl as usize).or_else(|| levels.first())?;
        let kind = match level.nfc {
            NFC_NONE => return None,
            NFC_BULLET => ListType::Bullet,
            _ => ListType::Numbered,
        };
        Some(ListLevel {
            lsid,
            kind,
            start: level.start,
        })
    }
}

fn i32_at(data: &[u8], at: usize) -> Option<i32> {
    data.get(at..at + 4)
        .map(|b| i32::from_le_bytes([b[0], b[1], b[2], b[3]]))
}

/// Read the list definitions and overrides. What cannot be read is left out: lists decide
/// markers, never whether text is recovered.
pub(super) fn read_lists(table: &[u8], plf_lst: FcLcb, plf_lfo: FcLcb) -> ListTable {
    let mut out = ListTable::default();
    if plf_lst.lcb >= 2 {
        let start = plf_lst.fc as usize;
        if let Some(count) = table
            .get(start..start + 2)
            .map(|b| i16::from_le_bytes([b[0], b[1]]))
        {
            // The LSTFs, then — past the PlfLst's own extent — each list's levels in order.
            let mut lsts = Vec::new();
            for i in 0..count.max(0) as usize {
                let at = start + 2 + i * 28;
                let (Some(lsid), Some(&flags)) = (i32_at(table, at), table.get(at + 26)) else {
                    break;
                };
                lsts.push((lsid, if flags & 0x01 != 0 { 1 } else { 9 }));
            }
            let mut at = start + 2 + lsts.len() * 28;
            'lists: for (lsid, count) in lsts {
                let mut levels = Vec::with_capacity(count);
                for _ in 0..count {
                    let (Some(start_at), Some(&nfc), Some(&cb_chpx), Some(&cb_papx)) = (
                        i32_at(table, at),
                        table.get(at + 4),
                        table.get(at + 24),
                        table.get(at + 25),
                    ) else {
                        break 'lists;
                    };
                    levels.push(Level {
                        start: start_at.max(0) as u32,
                        nfc,
                    });
                    at += 28 + cb_papx as usize + cb_chpx as usize;
                    // The number text: a 16-bit count of UTF-16 units.
                    let Some(cch) = table
                        .get(at..at + 2)
                        .map(|b| u16::from_le_bytes([b[0], b[1]]))
                    else {
                        break 'lists;
                    };
                    at += 2 + cch as usize * 2;
                }
                out.lists.insert(lsid, levels);
            }
        }
    }
    if plf_lfo.lcb >= 4 {
        let start = plf_lfo.fc as usize;
        if let Some(count) = i32_at(table, start) {
            for i in 0..count.max(0) as usize {
                match i32_at(table, start + 4 + i * 16) {
                    Some(lsid) => out.overrides.push(lsid),
                    None => break,
                }
            }
        }
    }
    out
}

/// Numbers list items as Word counts them: per list and level, restarting a level's count
/// when an item at a shallower level of the same list appears.
#[derive(Debug, Default)]
pub(super) struct ListCounter {
    counts: HashMap<(i32, u8), u32>,
}

impl ListCounter {
    pub fn next(&mut self, level: &ListLevel, ilvl: u8) -> u32 {
        self.counts
            .retain(|&(lsid, depth), _| lsid != level.lsid || depth <= ilvl);
        let count = self
            .counts
            .entry((level.lsid, ilvl))
            .or_insert(level.start.saturating_sub(1));
        *count += 1;
        *count
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn deeper_levels_restart_under_a_new_parent_item() {
        let level = ListLevel {
            lsid: 7,
            kind: ListType::Numbered,
            start: 1,
        };
        let mut counter = ListCounter::default();
        assert_eq!(counter.next(&level, 0), 1);
        assert_eq!(counter.next(&level, 1), 1);
        assert_eq!(counter.next(&level, 1), 2);
        assert_eq!(counter.next(&level, 0), 2);
        assert_eq!(counter.next(&level, 1), 1);
    }

    #[test]
    fn a_list_starting_elsewhere_counts_from_its_start() {
        let level = ListLevel {
            lsid: 1,
            kind: ListType::Numbered,
            start: 5,
        };
        let mut counter = ListCounter::default();
        assert_eq!(counter.next(&level, 0), 5);
        assert_eq!(counter.next(&level, 0), 6);
    }
}
