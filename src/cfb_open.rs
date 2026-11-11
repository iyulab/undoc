//! Opening a compound file the way readers in the wild do: with a file-allocation table that lists
//! more sectors than the file holds.
//!
//! The last sector of the FAT is always whole, so for a file whose sector count is not a multiple
//! of the FAT's capacity the table has entries for sectors past the end of the file. Nothing
//! points at them, so they do not matter ([MS-CFB] 2.3 calls them unused) -- though writers fill
//! them with `FREESECT`, with `ENDOFCHAIN` (the Word document that is the regression file of
//! Apache POI bug 47950 does), or leave the sectors off the end of the file. POI, `olefile` and
//! Microsoft's own readers open such files. The `cfb` crate refuses the ones whose unused
//! entries are not `FREESECT` (`Malformed FAT (FAT has 128 entries, but file has only 52
//! sectors)`), so those are padded with the missing sectors first -- only when every entry past the
//! end of the file is that `ENDOFCHAIN` filler and no entry inside the file, and none of the
//! header's start sectors, points past the end of it; a file that does is truncated and stays an
//! error.
//!
//! TODO: remove this module once the `cfb` crate opens such a file itself.

use std::borrow::Cow;
use std::io::{self, Cursor};

/// A compound file over bytes that may have been padded to the FAT's length.
pub(crate) type Container<'a> = cfb::CompoundFile<Cursor<Cow<'a, [u8]>>>;

/// The largest id of a sector that holds data; above it are the special values (`DIFSECT`,
/// `FATSECT`, `ENDOFCHAIN`, `FREESECT`).
const MAX_REGULAR_SECTOR: u32 = 0xFFFF_FFFA;
/// The value a FAT entry has for the last sector of a chain.
const ENDOFCHAIN: u32 = 0xFFFF_FFFE;
#[cfg(test)]
const FREESECT: u32 = 0xFFFF_FFFF;
/// The header's DIFAT holds the first 109 FAT sector ids.
const HEADER_DIFAT: usize = 109;

/// Open the compound file in `data`.
pub(crate) fn open(data: &[u8]) -> io::Result<Container<'_>> {
    match cfb::CompoundFile::open(Cursor::new(Cow::Borrowed(data))) {
        Ok(container) => Ok(container),
        Err(error) => match padded(data) {
            Some(bytes) => cfb::CompoundFile::open(Cursor::new(Cow::Owned(bytes))),
            None => Err(error),
        },
    }
}

fn u32_at(data: &[u8], at: usize) -> Option<u32> {
    Some(u32::from_le_bytes(data.get(at..at + 4)?.try_into().ok()?))
}

/// `data` with free sectors added up to the length of its FAT, when that is what is missing.
fn padded(data: &[u8]) -> Option<Vec<u8>> {
    let sector = match u16::from_le_bytes(data.get(30..32)?.try_into().ok()?) {
        9 => 512,
        12 => 4096,
        _ => return None,
    };
    let fat_sectors = u32_at(data, 44)? as usize;
    if fat_sectors == 0 || fat_sectors > HEADER_DIFAT {
        return None;
    }
    // Sector 0 follows the header, which takes a sector's length (512 or, for 4 KiB sectors, 4096).
    let present = data.len().checked_sub(sector)? / sector;
    let entries = sector / 4;
    let listed = fat_sectors * entries;
    if listed <= present {
        return None;
    }
    // Nothing in the file may point past its end: not the header's start sectors (directory,
    // mini FAT, DIFAT) and not a FAT entry of a sector that is there.
    let past = |id: u32| id <= MAX_REGULAR_SECTOR && id as usize >= present;
    if [48, 60, 68]
        .into_iter()
        .any(|at| u32_at(data, at).is_none_or(past))
    {
        return None;
    }
    for index in 0..fat_sectors {
        let id = u32_at(data, 76 + index * 4)? as usize;
        if id >= present {
            return None;
        }
        let base = (id + 1).checked_mul(sector)?;
        for entry in 0..entries {
            let value = u32_at(data, base + entry * 4)?;
            if index * entries + entry < present {
                if past(value) {
                    return None;
                }
            } else if value != ENDOFCHAIN {
                // Entries past the end of the file are the filler this module is about; any other
                // value is a sector the file lacks (the cfb crate already takes a FREESECT tail).
                return None;
            }
        }
    }
    let mut bytes = data.to_vec();
    bytes.resize((listed + 1) * sector, 0);
    Some(bytes)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Read, Write};

    /// A compound file (512-byte sectors) holding one small stream, in bytes, and how many of its
    /// sectors, header excluded, hold something.
    fn small_file() -> (Vec<u8>, usize) {
        let mut file =
            cfb::CompoundFile::create_with_version(cfb::Version::V3, Cursor::new(Vec::new()))
                .unwrap();
        file.create_stream("/Data")
            .unwrap()
            .write_all(b"hello compound file")
            .unwrap();
        file.flush().unwrap();
        let data = file.into_inner().into_inner();
        // The first FAT sector's id is in the header's DIFAT; entries from the last used sector
        // on are free.
        let fat_at = (u32_at(&data, 76).unwrap() as usize + 1) * 512;
        let used = (0..128)
            .filter(|&i| u32_at(&data, fat_at + i * 4).unwrap() != FREESECT)
            .max()
            .unwrap()
            + 1;
        (data, used)
    }

    fn read_data(container: &mut Container<'_>) -> Vec<u8> {
        let mut out = Vec::new();
        container
            .open_stream("/Data")
            .unwrap()
            .read_to_end(&mut out)
            .unwrap();
        out
    }

    #[test]
    fn a_fat_whose_unused_entries_are_end_of_chain_is_opened() {
        let (mut whole, used) = small_file();
        // The unused entries of the FAT, as POI's bug 47950 file has them: ENDOFCHAIN.
        let fat_at = (u32_at(&whole, 76).unwrap() as usize + 1) * 512;
        for entry in used..128 {
            let at = fat_at + entry * 4;
            whole[at..at + 4].copy_from_slice(&ENDOFCHAIN.to_le_bytes());
        }
        assert!(
            cfb::CompoundFile::open(Cursor::new(&whole[..])).is_err(),
            "the cfb crate reads it, so this proves nothing"
        );
        let mut container = open(&whole).expect("padded with the missing sectors, it opens");
        assert_eq!(read_data(&mut container), b"hello compound file");
    }

    #[test]
    fn a_file_that_lacks_a_used_sector_stays_an_error() {
        let (whole, used) = small_file();
        // Cut into the sectors the FAT marks as used: a truncated file, not a short one.
        let cut = &whole[..512 * used];
        assert!(open(cut).is_err());
    }

    #[test]
    fn an_intact_file_is_opened_untouched() {
        let (whole, _) = small_file();
        let mut container = open(&whole).unwrap();
        assert_eq!(read_data(&mut container), b"hello compound file");
    }
}
