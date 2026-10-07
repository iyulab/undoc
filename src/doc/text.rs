//! The piece table ([MS-DOC] 2.9.38 `Clx`, 2.8.35 `PlcPcd`): where each run of document
//! characters sits in the `WordDocument` stream, and how it is encoded.

use super::fib::FcLcb;
use crate::error::{Error, Result};

/// One piece: character positions `[cp_start, cp_end)` stored from byte `fc` on.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct Piece {
    pub cp_start: u32,
    pub cp_end: u32,
    /// Byte offset of the first character in the `WordDocument` stream.
    pub fc: u32,
    /// One byte per character (cp1252) rather than UTF-16LE.
    pub compressed: bool,
}

/// A character of the document together with the stream offset it was read from — the key
/// paragraph and character properties are looked up by.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct DocChar {
    pub ch: char,
    pub fc: u32,
    /// Character position in the document's text.
    pub cp: u32,
}

fn invalid(what: impl Into<String>) -> Error {
    Error::InvalidData(format!("Word document: {}", what.into()))
}

/// Read the piece table out of the `Clx` in the table stream.
pub(super) fn parse_clx(table: &[u8], clx: FcLcb) -> Result<Vec<Piece>> {
    let start = clx.fc as usize;
    let end = start
        .checked_add(clx.lcb as usize)
        .filter(|&end| end <= table.len())
        .ok_or_else(|| invalid("the piece table lies outside the table stream"))?;
    let clx = &table[start..end];

    // Zero or more Prc (0x01, i16 size, grpprl), then exactly one Pcdt (0x02).
    let mut at = 0;
    while clx.get(at) == Some(&0x01) {
        let size = clx
            .get(at + 1..at + 3)
            .map(|b| i16::from_le_bytes([b[0], b[1]]))
            .filter(|&n| n >= 0)
            .ok_or_else(|| invalid("a property modifier in the piece table is truncated"))?;
        at += 3 + size as usize;
    }
    if clx.get(at) != Some(&0x02) {
        return Err(invalid(
            "the piece table has no piece descriptor table (Pcdt)",
        ));
    }
    let lcb = clx
        .get(at + 1..at + 5)
        .map(|b| u32::from_le_bytes([b[0], b[1], b[2], b[3]]) as usize)
        .ok_or_else(|| invalid("the piece descriptor table is truncated"))?;
    let plc = clx
        .get(at + 5..at + 5 + lcb)
        .ok_or_else(|| invalid("the piece descriptor table is truncated"))?;

    // PlcPcd: (n + 1) character positions, then n 8-byte piece descriptors.
    if plc.len() < 4 || (plc.len() - 4) % 12 != 0 {
        return Err(invalid("the piece descriptor table has an impossible size"));
    }
    let n = (plc.len() - 4) / 12;
    let cp = |i: usize| u32::from_le_bytes(plc[i * 4..i * 4 + 4].try_into().unwrap());
    let mut pieces = Vec::with_capacity(n);
    for i in 0..n {
        let pcd = 4 * (n + 1) + i * 8;
        let fc_compressed = u32::from_le_bytes(plc[pcd + 2..pcd + 6].try_into().unwrap());
        let compressed = fc_compressed & 0x4000_0000 != 0;
        let fc = fc_compressed & 0x3FFF_FFFF;
        let (cp_start, cp_end) = (cp(i), cp(i + 1));
        if cp_end < cp_start {
            return Err(invalid(
                "the piece table's character positions run backwards",
            ));
        }
        pieces.push(Piece {
            cp_start,
            cp_end,
            // A compressed piece stores its offset doubled.
            fc: if compressed { fc / 2 } else { fc },
            compressed,
        });
    }
    Ok(pieces)
}

/// Decode a byte of compressed text: Windows-1252, as [MS-DOC] 2.4.1 specifies.
pub(super) fn cp1252(byte: u8) -> char {
    const HIGH: [u16; 32] = [
        0x20AC, 0x0081, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, 0x02C6, 0x2030, 0x0160,
        0x2039, 0x0152, 0x008D, 0x017D, 0x008F, 0x0090, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022,
        0x2013, 0x2014, 0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x009D, 0x017E, 0x0178,
    ];
    match byte {
        0x80..=0x9F => char::from_u32(HIGH[(byte - 0x80) as usize] as u32).unwrap_or('\u{FFFD}'),
        _ => byte as char,
    }
}

/// The characters at positions `[from, to)`, in order, each with its stream offset.
///
/// Positions no piece covers are absent from the result rather than invented; a piece that
/// points past the end of the stream is an error, since its text cannot be recovered.
pub(super) fn read_chars(
    word_document: &[u8],
    pieces: &[Piece],
    from: u32,
    to: u32,
) -> Result<Vec<DocChar>> {
    let mut out = Vec::with_capacity(to.saturating_sub(from) as usize);
    for piece in pieces {
        let start = piece.cp_start.max(from);
        let end = piece.cp_end.min(to);
        if start >= end {
            continue;
        }
        let width = if piece.compressed { 1 } else { 2 };
        let first = piece.fc as usize + (start - piece.cp_start) as usize * width;
        let len = (end - start) as usize * width;
        let bytes = word_document.get(first..first + len).ok_or_else(|| {
            invalid("a piece of text lies past the end of the WordDocument stream")
        })?;
        if piece.compressed {
            for (i, &b) in bytes.iter().enumerate() {
                out.push(DocChar {
                    ch: cp1252(b),
                    fc: (first + i) as u32,
                    cp: start + i as u32,
                });
            }
        } else {
            let units: Vec<u16> = bytes
                .as_chunks::<2>()
                .0
                .iter()
                .map(|&c| u16::from_le_bytes(c))
                .collect();
            let mut i = 0;
            for decoded in char::decode_utf16(units.iter().copied()) {
                let unit_len = match decoded {
                    Ok(c) if (c as u32) > 0xFFFF => 2,
                    _ => 1,
                };
                out.push(DocChar {
                    ch: decoded.unwrap_or('\u{FFFD}'),
                    fc: (first + i * 2) as u32,
                    cp: start + i as u32,
                });
                i += unit_len;
            }
        }
    }
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn cp1252_maps_the_windows_block_and_passes_latin1_through() {
        assert_eq!(cp1252(b'A'), 'A');
        assert_eq!(cp1252(0x93), '\u{201C}');
        assert_eq!(cp1252(0x80), '\u{20AC}');
        assert_eq!(cp1252(0xE9), 'é');
    }

    fn clx(pieces: &[(u32, u32, u32, bool)]) -> Vec<u8> {
        let mut plc = Vec::new();
        for (cp, ..) in pieces {
            plc.extend_from_slice(&cp.to_le_bytes());
        }
        plc.extend_from_slice(&pieces.last().unwrap().1.to_le_bytes());
        for &(_, _, fc, compressed) in pieces {
            plc.extend_from_slice(&0u16.to_le_bytes());
            let raw = if compressed {
                (fc * 2) | 0x4000_0000
            } else {
                fc
            };
            plc.extend_from_slice(&raw.to_le_bytes());
            plc.extend_from_slice(&0u16.to_le_bytes());
        }
        // A Prc in front, which the reader must step over.
        let mut out = vec![0x01, 0x02, 0x00, 0xAA, 0xBB, 0x02];
        out.extend_from_slice(&(plc.len() as u32).to_le_bytes());
        out.extend_from_slice(&plc);
        out
    }

    #[test]
    fn mixed_pieces_read_in_character_order() {
        // "Hi" compressed at byte 100, then "é!" as UTF-16 at byte 200.
        let table = clx(&[(0, 2, 100, true), (2, 4, 200, false)]);
        let pieces = parse_clx(
            &table,
            FcLcb {
                fc: 0,
                lcb: table.len() as u32,
            },
        )
        .unwrap();
        let mut wd = vec![0u8; 300];
        wd[100..102].copy_from_slice(b"Hi");
        wd[200..204].copy_from_slice(&[0xE9, 0x00, b'!', 0x00]);

        let chars = read_chars(&wd, &pieces, 0, 4).unwrap();
        let text: String = chars.iter().map(|c| c.ch).collect();
        assert_eq!(text, "Hié!");
        assert_eq!(
            chars.iter().map(|c| c.fc).collect::<Vec<_>>(),
            [100, 101, 200, 202]
        );

        let tail: String = read_chars(&wd, &pieces, 1, 3)
            .unwrap()
            .iter()
            .map(|c| c.ch)
            .collect();
        assert_eq!(tail, "ié");
    }

    #[test]
    fn a_piece_past_the_stream_is_an_error() {
        let table = clx(&[(0, 10, 100, true)]);
        let pieces = parse_clx(
            &table,
            FcLcb {
                fc: 0,
                lcb: table.len() as u32,
            },
        )
        .unwrap();
        let err = read_chars(&[0u8; 50], &pieces, 0, 10).unwrap_err();
        assert!(matches!(err, Error::InvalidData(_)), "{err}");
    }

    #[test]
    fn a_clx_without_a_pcdt_is_an_error() {
        let err = parse_clx(&[0x05, 0, 0], FcLcb { fc: 0, lcb: 3 }).unwrap_err();
        assert!(matches!(err, Error::InvalidData(_)), "{err}");
    }
}
