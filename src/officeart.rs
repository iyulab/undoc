//! OfficeArt ([MS-ODRAW]) records shared by the 97-2003 readers: Word keeps its pictures in
//! the `Data` stream and PowerPoint in the `Pictures` stream, both as blip store entries
//! (`OfficeArtFBSE`) around the bytes of a picture (`OfficeArtBlip`).

pub(crate) const RT_SP_CONTAINER: u16 = 0xF004;
pub(crate) const RT_FSP: u16 = 0xF00A;
pub(crate) const RT_FOPT: u16 = 0xF00B;
pub(crate) const RT_TERTIARY_FOPT: u16 = 0xF122;
pub(crate) const RT_FBSE: u16 = 0xF007;

/// `wzDescription`: the shape's alternative text.
const PROP_DESCRIPTION: u16 = 0x0381;
/// `pib`: the 1-based index, in the blip store, of the picture a shape shows.
const PROP_PIB: u16 = 0x0104;

#[derive(Debug, Clone, Copy)]
pub(crate) struct Record {
    pub instance: u16,
    pub kind: u16,
    pub start: usize,
    pub end: usize,
}

pub(crate) fn record(data: &[u8], at: usize, limit: usize) -> Option<Record> {
    let h = data.get(at..at + 8)?;
    let ver_inst = u16::from_le_bytes([h[0], h[1]]);
    let len = u32::from_le_bytes([h[4], h[5], h[6], h[7]]) as usize;
    let start = at + 8;
    let end = start.checked_add(len).filter(|&e| e <= limit)?;
    Some(Record {
        instance: ver_inst >> 4,
        kind: u16::from_le_bytes([h[2], h[3]]),
        start,
        end,
    })
}

pub(crate) fn u16_at(d: &[u8], at: usize) -> Option<u16> {
    d.get(at..at + 2).map(|b| u16::from_le_bytes([b[0], b[1]]))
}

pub(crate) fn u32_at(d: &[u8], at: usize) -> Option<u32> {
    d.get(at..at + 4)
        .map(|b| u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
}

/// What a shape's property table (`OfficeArtFOPT`) says about its picture.
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub(crate) struct ShapePicture {
    /// The 1-based blip store index of the picture the shape shows.
    pub blip: Option<u32>,
    /// The shape's alternative text.
    pub alt_text: Option<String>,
}

/// Read a property table — `fopt` is a `FOPT` or tertiary `FOPT` record — for the picture it
/// names and its alternative text. Properties of 6 bytes: `instance` of them, then the data of
/// the complex ones in order.
pub(crate) fn picture_properties(data: &[u8], fopt: &Record) -> ShapePicture {
    let mut out = ShapePicture::default();
    let count = fopt.instance as usize;
    let mut complex = fopt.start + count * 6;
    for i in 0..count {
        let at = fopt.start + i * 6;
        let (Some(id), Some(value)) = (u16_at(data, at), u32_at(data, at + 2)) else {
            break;
        };
        let is_complex = id & 0x8000 != 0;
        match id & 0x3FFF {
            PROP_PIB if !is_complex => out.blip = Some(value).filter(|&b| b != 0),
            PROP_DESCRIPTION if is_complex => {
                let units: Vec<u16> = data
                    .get(complex..(complex + value as usize).min(fopt.end))
                    .unwrap_or_default()
                    .as_chunks::<2>()
                    .0
                    .iter()
                    .map(|&c| u16::from_le_bytes(c))
                    .take_while(|&u| u != 0)
                    .collect();
                let text = String::from_utf16_lossy(&units).trim().to_string();
                out.alt_text = (!text.is_empty()).then_some(text);
            }
            _ => {}
        }
        if is_complex {
            complex += value as usize;
        }
    }
    out
}

/// A shape property: its value, or — for a complex property — the bytes the value counts.
#[cfg(feature = "ppt")]
pub(crate) struct Property<'a> {
    pub value: u32,
    pub complex: Option<&'a [u8]>,
}

/// The property `id` of a property table, if it has one.
#[cfg(feature = "ppt")]
pub(crate) fn property<'a>(data: &'a [u8], fopt: &Record, id: u16) -> Option<Property<'a>> {
    let count = fopt.instance as usize;
    let mut complex = fopt.start + count * 6;
    for i in 0..count {
        let at = fopt.start + i * 6;
        let (pid, value) = (u16_at(data, at)?, u32_at(data, at + 2)?);
        let is_complex = pid & 0x8000 != 0;
        let bytes = is_complex
            .then(|| data.get(complex..(complex.checked_add(value as usize)?).min(fopt.end)))
            .flatten();
        if pid & 0x3FFF == id {
            return Some(Property {
                value,
                complex: bytes,
            });
        }
        if is_complex {
            complex = complex.checked_add(value as usize)?;
        }
    }
    None
}

/// How deep shape groups are followed — real documents nest a handful.
#[cfg(any(feature = "doc", feature = "xls"))]
const MAX_GROUP_DEPTH: u8 = 16;

/// The shapes within `container` (a drawing, or a group) that show a picture, in drawing order,
/// as (shape id, what their property tables say about the picture): a shape is an
/// `OfficeArtSpContainer`, possibly inside groups.
#[cfg(any(feature = "doc", feature = "xls"))]
pub(crate) fn picture_shapes(
    data: &[u8],
    container: &Record,
    depth: u8,
    out: &mut Vec<(u32, ShapePicture)>,
) {
    const RT_DG_CONTAINER: u16 = 0xF002;
    const RT_SPGR_CONTAINER: u16 = 0xF003;
    let mut at = container.start;
    while let Some(r) = record(data, at, container.end) {
        at = r.end;
        match r.kind {
            RT_DG_CONTAINER | RT_SPGR_CONTAINER if depth < MAX_GROUP_DEPTH => {
                picture_shapes(data, &r, depth + 1, out)
            }
            RT_SP_CONTAINER => {
                if let Some(shape) = picture_shape(data, &r) {
                    out.push(shape);
                }
            }
            _ => {}
        }
    }
}

/// A shape's id and what its property tables say about its picture — only a shape that names
/// a picture.
#[cfg(any(feature = "doc", feature = "xls"))]
fn picture_shape(data: &[u8], container: &Record) -> Option<(u32, ShapePicture)> {
    let mut id = None;
    let mut picture = ShapePicture::default();
    let mut at = container.start;
    while let Some(r) = record(data, at, container.end) {
        at = r.end;
        match r.kind {
            RT_FSP => id = u32_at(data, r.start),
            RT_FOPT | RT_TERTIARY_FOPT => {
                let found = picture_properties(data, &r);
                picture.blip = picture.blip.or(found.blip);
                picture.alt_text = picture.alt_text.take().or(found.alt_text);
            }
            _ => {}
        }
    }
    picture.blip?;
    Some((id?, picture))
}

/// The shape container's alternative text, from its property tables.
#[cfg(feature = "doc")]
pub(crate) fn description(data: &[u8], container: &Record) -> Option<String> {
    let mut p = container.start;
    while let Some(r) = record(data, p, container.end) {
        p = r.end;
        if r.kind != RT_FOPT && r.kind != RT_TERTIARY_FOPT {
            continue;
        }
        if let Some(text) = picture_properties(data, &r).alt_text {
            return Some(text);
        }
    }
    None
}

/// The picture a blip store entry holds, as file bytes and an extension. The entry carries the
/// blip itself after its name, or — when it has none — says where the blip is in a stream kept
/// apart (`delay`: PowerPoint's `Pictures` stream, Word's main stream).
pub(crate) fn blip_in_fbse(
    data: &[u8],
    fbse: &Record,
    delay: Option<&[u8]>,
) -> Option<(Vec<u8>, &'static str)> {
    // 36 fixed bytes (`foDelay` at 28, `cbName` at 33), then `cbName` bytes of name, then the
    // embedded blip.
    let name_len = *data.get(fbse.start + 33)? as usize;
    let inline = fbse.start + 36 + name_len;
    if inline < fbse.end {
        return blip_at(data, inline, fbse.end);
    }
    // A slot the presentation emptied has size 0 and `foDelay` 0 — not the first blip.
    if u32_at(data, fbse.start + 20)? == 0 {
        return None;
    }
    let delay = delay?;
    let at = u32_at(data, fbse.start + 28)? as usize;
    blip_at(delay, at, delay.len())
}

/// The picture of the blip record at `at`.
pub(crate) fn blip_at(data: &[u8], at: usize, limit: usize) -> Option<(Vec<u8>, &'static str)> {
    let blip = record(data, at, limit)?;
    // Bitmap blips: one or two 16-byte ids (the odd instance carries two), a tag byte, the file.
    let extension = match (blip.kind, blip.instance & !1) {
        (0xF01D | 0xF02A, 0x46A | 0x6E2) => "jpg",
        (0xF01E, 0x6E0) => "png",
        (0xF01F, 0x7A8) => "bmp",
        (0xF029, 0x6E4) => "tif",
        (0xF01A, 0x3D4) => return metafile(data, &blip, "emf"),
        (0xF01B, 0x216) => return metafile(data, &blip, "wmf"),
        // PICT (Macintosh) and anything else: not a file this reader produces.
        _ => return None,
    };
    let skip = 16 + if blip.instance & 1 == 1 { 16 } else { 0 } + 1;
    let bytes = data.get(blip.start + skip..blip.end)?;
    if extension == "bmp" {
        return Some((bmp_from_dib(bytes)?, extension));
    }
    // Office keeps a GIF in a PNG blip: the file is what its first bytes say, not the blip type.
    let extension = if extension == "png" && bytes.starts_with(b"GIF8") {
        "gif"
    } else {
        extension
    };
    Some((bytes.to_vec(), extension))
}

/// A metafile blip's file: one or two 16-byte ids, a 34-byte `OfficeArtMetafileHeader`, then
/// the metafile — DEFLATE-compressed (zlib framing) unless the header says it is stored.
fn metafile(
    data: &[u8],
    blip: &Record,
    extension: &'static str,
) -> Option<(Vec<u8>, &'static str)> {
    use std::io::Read;
    let header = blip.start + 16 + if blip.instance & 1 == 1 { 16 } else { 0 };
    let size = u32_at(data, header)? as usize;
    let saved = u32_at(data, header + 28)? as usize;
    let compression = *data.get(header + 32)?;
    let body = data.get(header + 34..(header + 34 + saved).min(blip.end))?;
    let bytes = match compression {
        0xFE => body.to_vec(),
        0x00 => {
            let mut out = Vec::with_capacity(size.min(64 << 20));
            flate2::read::ZlibDecoder::new(body)
                .take(64 << 20)
                .read_to_end(&mut out)
                .ok()?;
            out
        }
        _ => return None,
    };
    (!bytes.is_empty()).then_some((bytes, extension))
}

/// A BMP file from a device-independent bitmap: the 14-byte file header the DIB lacks.
fn bmp_from_dib(dib: &[u8]) -> Option<Vec<u8>> {
    let header_size = u32_at(dib, 0)? as usize;
    let bit_count = u16_at(dib, 14)?;
    let compression = u32_at(dib, 16)?;
    let colors_used = u32_at(dib, 32).unwrap_or(0) as usize;
    let palette = match (colors_used, bit_count) {
        (0, 1..=8) => 1usize << bit_count,
        (n, _) => n,
    };
    // BI_BITFIELDS with a plain info header keeps its three masks after it.
    let masks = if compression == 3 && header_size == 40 {
        12
    } else {
        0
    };
    let pixels_at = 14 + header_size + masks + palette * 4;
    let mut out = Vec::with_capacity(14 + dib.len());
    out.extend_from_slice(b"BM");
    out.extend_from_slice(&((14 + dib.len()) as u32).to_le_bytes());
    out.extend_from_slice(&[0; 4]);
    out.extend_from_slice(&(pixels_at as u32).to_le_bytes());
    out.extend_from_slice(dib);
    Some(out)
}

#[cfg(test)]
pub(crate) mod tests {
    use super::*;

    pub(crate) const PNG: &[u8] = b"\x89PNG\r\n\x1a\nfake";

    /// A PNG blip record.
    pub(crate) fn png_blip(png: &[u8]) -> Vec<u8> {
        let mut blip = (0x6E0u16 << 4).to_le_bytes().to_vec();
        blip.extend(0xF01Eu16.to_le_bytes());
        blip.extend(((16 + 1 + png.len()) as u32).to_le_bytes());
        blip.extend([0xAB; 16]);
        blip.push(0xFF);
        blip.extend(png);
        blip
    }

    /// A blip store entry: 36 fixed bytes (`size` of the blip it holds, `foDelay`), then `blip`
    /// when it is embedded.
    pub(crate) fn fbse(size: u32, fo_delay: u32, blip: &[u8]) -> Vec<u8> {
        let mut fixed = [0u8; 36];
        fixed[20..24].copy_from_slice(&size.to_le_bytes());
        fixed[28..32].copy_from_slice(&fo_delay.to_le_bytes());
        let mut out = (((6u16) << 4) | 0x2).to_le_bytes().to_vec();
        out.extend(RT_FBSE.to_le_bytes());
        out.extend(((36 + blip.len()) as u32).to_le_bytes());
        out.extend(fixed);
        out.extend(blip);
        out
    }

    /// A property table of `pib` (when given) and a description (when given).
    pub(crate) fn fopt(pib: Option<u32>, alt: Option<&str>) -> Vec<u8> {
        let mut table = Vec::new();
        let mut complex = Vec::new();
        let mut count = 0u16;
        if let Some(pib) = pib {
            table.extend((PROP_PIB | 0x4000).to_le_bytes());
            table.extend(pib.to_le_bytes());
            count += 1;
        }
        if let Some(alt) = alt {
            let text: Vec<u8> = alt
                .encode_utf16()
                .chain([0])
                .flat_map(|u| u.to_le_bytes())
                .collect();
            table.extend((PROP_DESCRIPTION | 0x8000).to_le_bytes());
            table.extend((text.len() as u32).to_le_bytes());
            complex.extend(text);
            count += 1;
        }
        table.extend(complex);
        let mut out = ((count << 4) | 0x3).to_le_bytes().to_vec();
        out.extend(RT_FOPT.to_le_bytes());
        out.extend((table.len() as u32).to_le_bytes());
        out.extend(table);
        out
    }

    #[test]
    fn a_png_blip_is_read_from_an_embedded_entry() {
        let blip = png_blip(PNG);
        let data = fbse(blip.len() as u32, 0, &blip);
        let entry = record(&data, 0, data.len()).unwrap();
        assert_eq!(
            blip_in_fbse(&data, &entry, None),
            Some((PNG.to_vec(), "png"))
        );
    }

    #[test]
    fn an_entry_without_a_blip_reads_it_from_the_delay_stream_at_fo_delay() {
        let blip = png_blip(PNG);
        let mut delay = vec![0u8; 7];
        delay.extend(&blip);
        let data = fbse(blip.len() as u32, 7, &[]);
        let entry = record(&data, 0, data.len()).unwrap();
        assert_eq!(
            blip_in_fbse(&data, &entry, Some(&delay)),
            Some((PNG.to_vec(), "png"))
        );
        // Without the stream, or past its end: no picture, not a failure.
        assert_eq!(blip_in_fbse(&data, &entry, None), None);
        assert_eq!(blip_in_fbse(&data, &entry, Some(&delay[..10])), None);
    }

    #[test]
    fn an_emptied_slot_is_no_picture_even_though_fo_delay_names_the_first_blip() {
        let mut delay = Vec::new();
        delay.extend(png_blip(PNG));
        // Size 0, `foDelay` 0: PowerPoint's mark of a deleted picture.
        let data = fbse(0, 0, &[]);
        let entry = record(&data, 0, data.len()).unwrap();
        assert_eq!(blip_in_fbse(&data, &entry, Some(&delay)), None);
    }

    #[test]
    fn a_gif_kept_in_a_png_blip_is_named_for_what_it_is() {
        let gif = b"GIF89a\x01\x00\x01\x00";
        let blip = png_blip(gif);
        let data = fbse(blip.len() as u32, 0, &blip);
        let entry = record(&data, 0, data.len()).unwrap();
        assert_eq!(
            blip_in_fbse(&data, &entry, None),
            Some((gif.to_vec(), "gif"))
        );
    }

    #[test]
    fn a_property_table_names_the_picture_and_its_alternative_text() {
        let data = fopt(Some(3), Some("A chart"));
        let table = record(&data, 0, data.len()).unwrap();
        assert_eq!(
            picture_properties(&data, &table),
            ShapePicture {
                blip: Some(3),
                alt_text: Some("A chart".into())
            }
        );
        let none = fopt(None, None);
        let table = record(&none, 0, none.len()).unwrap();
        assert_eq!(picture_properties(&none, &table), ShapePicture::default());
    }

    #[test]
    fn a_compressed_emf_blip_is_inflated() {
        use std::io::Write;
        let emf = b" EMF metafile bytes".repeat(4);
        let mut z = flate2::write::ZlibEncoder::new(Vec::new(), flate2::Compression::default());
        z.write_all(&emf).unwrap();
        let packed = z.finish().unwrap();

        let mut body = vec![0xCD; 16];
        let mut header = vec![0u8; 34];
        header[0..4].copy_from_slice(&(emf.len() as u32).to_le_bytes());
        header[28..32].copy_from_slice(&(packed.len() as u32).to_le_bytes());
        header[33] = 0xFE;
        body.extend(header);
        body.extend(&packed);
        let mut data = (0x3D4u16 << 4).to_le_bytes().to_vec();
        data.extend(0xF01Au16.to_le_bytes());
        data.extend((body.len() as u32).to_le_bytes());
        data.extend(body);
        let blip = record(&data, 0, data.len()).unwrap();
        assert_eq!(metafile(&data, &blip, "emf"), Some((emf, "emf")));
    }

    #[test]
    fn a_dib_becomes_a_bmp_file() {
        // 40-byte header, 1x1, 24 bits, no palette.
        let mut dib = vec![0u8; 40];
        dib[0] = 40;
        dib[14] = 24;
        dib.extend([1, 2, 3, 0]);
        let bmp = bmp_from_dib(&dib).unwrap();
        assert_eq!(&bmp[..2], b"BM");
        assert_eq!(u32_at(&bmp, 10), Some(54));
        assert_eq!(bmp.len(), 14 + dib.len());
    }
}
