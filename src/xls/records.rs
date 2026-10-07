//! BIFF8 record stream ([MS-XLS] 2.1.4): a sequence of `(type, size, data)` records, where a
//! record too long for one is carried on in `CONTINUE` records.

use crate::error::{Error, Result};

pub(super) const BOF: u16 = 0x0809;
pub(super) const EOF: u16 = 0x000A;
pub(super) const CONTINUE: u16 = 0x003C;
pub(super) const FILEPASS: u16 = 0x002F;
pub(super) const DATEMODE: u16 = 0x0022;
pub(super) const BOUNDSHEET8: u16 = 0x0085;
pub(super) const SST: u16 = 0x00FC;
pub(super) const FORMAT: u16 = 0x041E;
pub(super) const XF: u16 = 0x00E0;
pub(super) const LABELSST: u16 = 0x00FD;
pub(super) const LABEL: u16 = 0x0204;
pub(super) const NUMBER: u16 = 0x0203;
pub(super) const RK: u16 = 0x027E;
pub(super) const MULRK: u16 = 0x00BD;
pub(super) const FORMULA: u16 = 0x0006;
pub(super) const STRING: u16 = 0x0207;
pub(super) const BOOLERR: u16 = 0x0205;
pub(super) const MERGEDCELLS: u16 = 0x00E5;
pub(super) const HLINK: u16 = 0x01B8;
pub(super) const OBJ: u16 = 0x005D;
pub(super) const TXO: u16 = 0x01B6;
pub(super) const NOTE: u16 = 0x001C;

/// The BIFF version in a BIFF8 `BOF`.
pub(super) const BIFF8: u16 = 0x0600;

/// One record, with the data of the `CONTINUE` records that follow it kept as separate
/// segments: a string split across them restarts with a flags byte, so the boundaries matter.
#[derive(Debug)]
pub(super) struct Record<'a> {
    pub kind: u16,
    pub segments: Vec<&'a [u8]>,
}

impl<'a> Record<'a> {
    /// The record's own data, without its continuations.
    pub fn data(&self) -> &'a [u8] {
        self.segments[0]
    }
}

/// Reads records from `offset` on, each with its continuations attached.
pub(super) struct Records<'a> {
    stream: &'a [u8],
    at: usize,
}

impl<'a> Records<'a> {
    pub fn new(stream: &'a [u8], offset: usize) -> Self {
        Self { stream, at: offset }
    }

    fn header(&self, at: usize) -> Option<(u16, usize)> {
        let h = self.stream.get(at..at + 4)?;
        Some((
            u16::from_le_bytes([h[0], h[1]]),
            u16::from_le_bytes([h[2], h[3]]) as usize,
        ))
    }
}

impl<'a> Iterator for Records<'a> {
    type Item = Result<Record<'a>>;

    fn next(&mut self) -> Option<Self::Item> {
        let (kind, size) = self.header(self.at)?;
        let offset = self.at;
        let Some(data) = self.stream.get(offset + 4..offset + 4 + size) else {
            self.at = self.stream.len();
            return Some(Err(Error::InvalidData(format!(
                "Excel workbook: record 0x{kind:04X} at offset {offset} runs past the end of \
                 the stream"
            ))));
        };
        let mut segments = vec![data];
        self.at = offset + 4 + size;
        while let Some((CONTINUE, size)) = self.header(self.at) {
            match self.stream.get(self.at + 4..self.at + 4 + size) {
                Some(more) => segments.push(more),
                None => break,
            }
            self.at += 4 + size;
        }
        Some(Ok(Record { kind, segments }))
    }
}

/// A cursor over a record's segments.
pub(super) struct Cursor<'a> {
    segments: &'a [&'a [u8]],
    segment: usize,
    at: usize,
}

fn truncated(what: &str) -> Error {
    Error::InvalidData(format!("Excel workbook: {what} is truncated"))
}

impl<'a> Cursor<'a> {
    pub fn new(segments: &'a [&'a [u8]]) -> Self {
        Self {
            segments,
            segment: 0,
            at: 0,
        }
    }

    /// Bytes left in the current segment.
    fn left(&self) -> usize {
        self.segments
            .get(self.segment)
            .map_or(0, |s| s.len() - self.at)
    }

    /// Move to the next segment when the current one is used up.
    fn advance_if_done(&mut self) {
        while self.segment < self.segments.len() && self.left() == 0 {
            self.segment += 1;
            self.at = 0;
        }
    }

    pub fn u8(&mut self) -> Result<u8> {
        self.advance_if_done();
        let b = *self
            .segments
            .get(self.segment)
            .and_then(|s| s.get(self.at))
            .ok_or_else(|| truncated("a record"))?;
        self.at += 1;
        Ok(b)
    }

    pub fn u16(&mut self) -> Result<u16> {
        Ok(u16::from_le_bytes([self.u8()?, self.u8()?]))
    }

    pub fn u32(&mut self) -> Result<u32> {
        Ok(u32::from_le_bytes([
            self.u8()?,
            self.u8()?,
            self.u8()?,
            self.u8()?,
        ]))
    }

    pub fn skip(&mut self, mut n: usize) -> Result<()> {
        while n > 0 {
            self.advance_if_done();
            let step = n.min(self.left());
            if step == 0 {
                return Err(truncated("a record"));
            }
            self.at += step;
            n -= step;
        }
        Ok(())
    }

    /// `cch` characters, one or two bytes each as `high_byte` says. Where the characters run
    /// into the next segment, that segment restarts with a flags byte giving their width
    /// anew ([MS-XLS] 2.5.293).
    fn chars(&mut self, cch: usize, mut high_byte: bool) -> Result<String> {
        let mut units = Vec::with_capacity(cch);
        while units.len() < cch {
            if self.left() == 0 {
                self.segment += 1;
                self.at = 0;
                if self.segment >= self.segments.len() {
                    return Err(truncated("a string"));
                }
                high_byte = self.u8()? & 0x01 != 0;
                continue;
            }
            units.push(if high_byte {
                self.u16()?
            } else {
                self.u8()? as u16
            });
        }
        Ok(String::from_utf16_lossy(&units))
    }

    /// `cch` characters preceded by their flags byte — the text a `TXO` carries in the
    /// `CONTINUE` that follows it.
    pub fn flagged_chars(&mut self, cch: usize) -> Result<String> {
        let flags = self.u8()?;
        self.chars(cch, flags & 0x01 != 0)
    }

    /// A NUL-terminated UTF-16 string of at most `max_bytes` bytes, the rest of which is
    /// skipped.
    pub fn utf16_within(&mut self, max_bytes: usize) -> Result<String> {
        let mut units = Vec::new();
        let mut used = 0;
        while used + 2 <= max_bytes {
            let unit = self.u16()?;
            used += 2;
            if unit == 0 {
                break;
            }
            units.push(unit);
        }
        self.skip(max_bytes - used)?;
        Ok(String::from_utf16_lossy(&units))
    }

    /// `HyperlinkString` ([MS-OSHARED] 2.3.7.9): a count of UTF-16 units, NUL included.
    pub fn hyperlink_string(&mut self) -> Result<String> {
        let count = self.u32()? as usize;
        self.utf16_within(count * 2)
    }

    /// `XLUnicodeRichExtendedString` ([MS-XLS] 2.5.293) — the shared string table's entries.
    pub fn rich_extended_string(&mut self) -> Result<String> {
        let cch = self.u16()? as usize;
        let flags = self.u8()?;
        let runs = if flags & 0x08 != 0 {
            self.u16()? as usize
        } else {
            0
        };
        let ext = if flags & 0x04 != 0 {
            self.u32()? as usize
        } else {
            0
        };
        let text = self.chars(cch, flags & 0x01 != 0)?;
        self.skip(runs * 4 + ext)?;
        Ok(text)
    }

    /// `XLUnicodeString` ([MS-XLS] 2.5.294): a 16-bit count, a flags byte, the characters.
    pub fn unicode_string(&mut self) -> Result<String> {
        let cch = self.u16()? as usize;
        let flags = self.u8()?;
        self.chars(cch, flags & 0x01 != 0)
    }

    /// `ShortXLUnicodeString` ([MS-XLS] 2.5.240): an 8-bit count, a flags byte, the characters.
    pub fn short_unicode_string(&mut self) -> Result<String> {
        let cch = self.u8()? as usize;
        let flags = self.u8()?;
        self.chars(cch, flags & 0x01 != 0)
    }
}

/// Decode an `RkNumber` ([MS-XLS] 2.5.217).
pub(super) fn rk_value(rk: u32) -> f64 {
    let value = if rk & 0x02 != 0 {
        ((rk as i32) >> 2) as f64
    } else {
        f64::from_bits(((rk & 0xFFFF_FFFC) as u64) << 32)
    };
    if rk & 0x01 != 0 {
        value / 100.0
    } else {
        value
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rk_numbers_decode_in_all_four_encodings() {
        assert_eq!(rk_value((42 << 2) | 0x02), 42.0);
        assert_eq!(rk_value(((-7i32 << 2) as u32) | 0x02), -7.0);
        assert_eq!(rk_value((1234 << 2) | 0x03), 12.34);
        let bits = (1.5f64.to_bits() >> 32) as u32;
        assert_eq!(rk_value(bits), 1.5);
    }

    #[test]
    fn a_string_split_across_continue_changes_width_at_the_boundary() {
        // "ab" compressed, then a CONTINUE restarting with fHighByte for "한".
        let first: &[u8] = &[3, 0, 0x00, b'a', b'b'];
        let second: &[u8] = &[0x01, 0x5C, 0xD5];
        let segments = [first, second];
        let mut cursor = Cursor::new(&segments);
        assert_eq!(cursor.rich_extended_string().unwrap(), "ab한");
    }

    #[test]
    fn records_carry_their_continuations() {
        let mut stream = Vec::new();
        for (kind, data) in [(SST, &b"xy"[..]), (CONTINUE, b"z"), (EOF, b"")] {
            stream.extend_from_slice(&kind.to_le_bytes());
            stream.extend_from_slice(&(data.len() as u16).to_le_bytes());
            stream.extend_from_slice(data);
        }
        let records: Vec<_> = Records::new(&stream, 0).map(|r| r.unwrap()).collect();
        assert_eq!(records.len(), 2);
        assert_eq!(records[0].segments, [&b"xy"[..], b"z"]);
        assert_eq!(records[1].kind, EOF);
    }
}
