//! Inflating with the host's `DecompressionStream`.
//!
//! `undoc` is built here without its `deflate` feature, so it carries no decompressor.
//! Its ZIP reader only understands stored entries; [`inflate_package`] bridges that: it reads the archive's central directory,
//! inflates each DEFLATE entry through the browser/Node stream API, and writes the entries
//! back out as a stored-only archive that the core crate can open.
//!
//! Only what Office packages use is handled: no ZIP64, no encryption, no data descriptors
//! (the repacked archive carries the original CRCs and sizes up front). Anything this module
//! does not recognise as a plain ZIP is returned unchanged, so the core crate stays the one
//! place that classifies bad input.

use wasm_bindgen::prelude::*;

#[wasm_bindgen(inline_js = r#"
export async function inflate(bytes, format) {
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream(format));
    return new Uint8Array(await new Response(stream).arrayBuffer());
}
"#)]
extern "C" {
    #[wasm_bindgen(catch)]
    async fn inflate(bytes: &[u8], format: &str) -> Result<JsValue, JsValue>;
}

/// Inflate `bytes` with a `DecompressionStream` of `format`: `deflate-raw` for ZIP entries,
/// `deflate` (zlib framing) for the compressed pictures inside .doc files.
async fn inflate_as(bytes: &[u8], format: &str) -> Result<Vec<u8>, JsValue> {
    let inflated = inflate(bytes, format).await?;
    Ok(js_sys::Uint8Array::new(&inflated).to_vec())
}

/// Inflate a zlib stream (RFC 1950).
#[cfg(feature = "doc")]
pub(crate) async fn inflate_zlib(bytes: &[u8]) -> Result<Vec<u8>, JsValue> {
    inflate_as(bytes, "deflate").await
}

const LOCAL_SIG: u32 = 0x0403_4b50;
const CENTRAL_SIG: u32 = 0x0201_4b50;
const EOCD_SIG: u32 = 0x0605_4b50;
const METHOD_STORED: u16 = 0;
const METHOD_DEFLATE: u16 = 8;
/// Bit 11: the entry name is UTF-8. The one flag that has to survive repacking.
const FLAG_UTF8: u16 = 1 << 11;

struct Entry<'a> {
    name: &'a [u8],
    flags: u16,
    method: u16,
    time: u16,
    date: u16,
    crc: u32,
    data: &'a [u8],
}

/// Return `data` as an archive whose entries are all stored, or `data` itself when it is not
/// a ZIP this module can read.
pub(crate) async fn inflate_package(data: &[u8]) -> Result<Vec<u8>, JsValue> {
    let Some(entries) = read_entries(data) else {
        return Ok(data.to_vec());
    };

    let mut inflated: Vec<Option<Vec<u8>>> = Vec::with_capacity(entries.len());
    for entry in &entries {
        inflated.push(if entry.method == METHOD_DEFLATE {
            Some(inflate_as(entry.data, "deflate-raw").await?)
        } else {
            None
        });
    }
    Ok(write_stored(&entries, &inflated))
}

fn u16_at(b: &[u8], at: usize) -> Option<u16> {
    Some(u16::from_le_bytes(b.get(at..at + 2)?.try_into().ok()?))
}

fn u32_at(b: &[u8], at: usize) -> Option<u32> {
    Some(u32::from_le_bytes(b.get(at..at + 4)?.try_into().ok()?))
}

/// Walk the central directory. `None` for anything unreadable or outside the supported subset.
fn read_entries(data: &[u8]) -> Option<Vec<Entry<'_>>> {
    if u32_at(data, 0)? != LOCAL_SIG {
        return None;
    }
    // The end-of-central-directory record sits at the tail, after a comment of up to 64 KiB.
    let eocd = (0..=data.len().checked_sub(22)?)
        .rev()
        .take(22 + usize::from(u16::MAX))
        .find(|&at| u32_at(data, at) == Some(EOCD_SIG))?;
    let count = u16_at(data, eocd + 10)?;
    let mut at = u32_at(data, eocd + 16)? as usize;
    if count == u16::MAX || at == u32::MAX as usize {
        return None; // ZIP64
    }

    let mut entries = Vec::with_capacity(count.into());
    for _ in 0..count {
        if u32_at(data, at)? != CENTRAL_SIG {
            return None;
        }
        let flags = u16_at(data, at + 8)?;
        let compressed_size = u32_at(data, at + 20)? as usize;
        let name_len = u16_at(data, at + 28)? as usize;
        let extra_len = u16_at(data, at + 30)? as usize;
        let comment_len = u16_at(data, at + 32)? as usize;
        let local = u32_at(data, at + 42)? as usize;
        if flags & 1 != 0 || compressed_size == u32::MAX as usize || local == u32::MAX as usize {
            return None; // encrypted or ZIP64
        }
        let method = u16_at(data, at + 10)?;
        if method != METHOD_STORED && method != METHOD_DEFLATE {
            return None;
        }

        // The local header repeats the name and extra lengths, and its extra field can differ
        // from the central one, so the payload offset has to come from the local copy.
        if u32_at(data, local)? != LOCAL_SIG {
            return None;
        }
        let payload =
            local + 30 + u16_at(data, local + 26)? as usize + u16_at(data, local + 28)? as usize;

        entries.push(Entry {
            name: data.get(at + 46..at + 46 + name_len)?,
            flags,
            method,
            time: u16_at(data, at + 12)?,
            date: u16_at(data, at + 14)?,
            crc: u32_at(data, at + 16)?,
            data: data.get(payload..payload.checked_add(compressed_size)?)?,
        });
        at += 46 + name_len + extra_len + comment_len;
    }
    Some(entries)
}

/// Write `entries` as a stored archive; `inflated[i]` replaces entry `i`'s payload when set.
fn write_stored(entries: &[Entry<'_>], inflated: &[Option<Vec<u8>>]) -> Vec<u8> {
    let mut out = Vec::new();
    let mut central = Vec::new();
    for (entry, inflated) in entries.iter().zip(inflated) {
        let payload = inflated.as_deref().unwrap_or(entry.data);
        let flags = entry.flags & FLAG_UTF8;
        let (name_len, size) = (entry.name.len() as u16, payload.len() as u32);
        let offset = out.len() as u32;

        let header = |out: &mut Vec<u8>, central: bool| {
            out.extend(if central { CENTRAL_SIG } else { LOCAL_SIG }.to_le_bytes());
            if central {
                out.extend(20u16.to_le_bytes()); // version made by
            }
            out.extend(20u16.to_le_bytes()); // version needed
            out.extend(flags.to_le_bytes());
            out.extend(METHOD_STORED.to_le_bytes());
            out.extend(entry.time.to_le_bytes());
            out.extend(entry.date.to_le_bytes());
            out.extend(entry.crc.to_le_bytes());
            out.extend(size.to_le_bytes());
            out.extend(size.to_le_bytes());
            out.extend(name_len.to_le_bytes());
            out.extend(0u16.to_le_bytes()); // extra length
            if central {
                out.extend([0; 6]); // comment length, disk, internal attrs
                out.extend(0u32.to_le_bytes()); // external attrs
                out.extend(offset.to_le_bytes());
            }
            out.extend(entry.name);
        };
        header(&mut out, false);
        out.extend(payload);
        header(&mut central, true);
    }

    let (count, cd_offset, cd_size) =
        (entries.len() as u16, out.len() as u32, central.len() as u32);
    out.extend(central);
    out.extend(EOCD_SIG.to_le_bytes());
    out.extend([0; 4]); // disk numbers
    out.extend(count.to_le_bytes());
    out.extend(count.to_le_bytes());
    out.extend(cd_size.to_le_bytes());
    out.extend(cd_offset.to_le_bytes());
    out.extend(0u16.to_le_bytes()); // comment length
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    fn stored(files: &[(&str, &[u8])]) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        for (name, body) in files {
            zip.start_file(*name, options).unwrap();
            std::io::Write::write_all(&mut zip, body).unwrap();
        }
        zip.finish().unwrap().into_inner()
    }

    #[test]
    fn repacked_archive_reads_back_with_the_same_entries() {
        let original = stored(&[("a.xml", b"<a/>"), ("word/b.xml", b"<b>text</b>")]);
        let entries = read_entries(&original).expect("a stored archive is readable");
        let repacked = write_stored(&entries, &[None, None]);

        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(repacked)).unwrap();
        for (name, body) in [("a.xml", &b"<a/>"[..]), ("word/b.xml", b"<b>text</b>")] {
            let mut got = Vec::new();
            std::io::Read::read_to_end(&mut archive.by_name(name).unwrap(), &mut got).unwrap();
            assert_eq!(got, body, "{name}");
        }
    }

    #[test]
    fn inflated_payload_replaces_the_compressed_one() {
        let original = stored(&[("a.xml", b"stand-in")]);
        let entries = read_entries(&original).unwrap();
        let repacked = write_stored(&entries, &[Some(b"inflated".to_vec())]);
        let again = read_entries(&repacked).unwrap();
        assert_eq!(again[0].data, b"inflated");
    }

    #[test]
    fn non_zip_input_is_not_an_archive() {
        assert!(read_entries(b"garbage data").is_none());
        assert!(read_entries(b"PK\x03\x04 truncated").is_none());
        assert!(read_entries(&[]).is_none());
    }
}
