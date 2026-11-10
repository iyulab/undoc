//! Inline pictures ([MS-DOC] 2.9.192 `PICFAndOfficeArtData`): a picture character points, by
//! `sprmCPicLocation`, at a `PICF` in the `Data` stream, followed by the shape that draws it
//! and the blip store entry holding its bytes ([MS-ODRAW] `OfficeArtFBSE`, `OfficeArtBlip`).

use crate::officeart::{self, record, u16_at, u32_at, RT_FBSE, RT_SP_CONTAINER};

/// `mfpf.mm` of a picture whose file name follows the header.
const MM_SHAPE_FILE: u16 = 0x0066;

/// A picture's bytes as a file, with what describes it.
#[derive(Debug)]
pub(super) struct Picture {
    pub data: Vec<u8>,
    pub extension: &'static str,
    pub alt_text: Option<String>,
}

/// The picture whose `PICF` is at `offset` in the `Data` stream. None for a picture this
/// reader cannot turn into a file — a compressed metafile, or a structure it cannot follow.
pub(super) fn read(data: &[u8], offset: u32) -> Option<Picture> {
    let at = offset as usize;
    let lcb = u32_at(data, at)? as usize;
    let header = u16_at(data, at + 4)? as usize;
    let mm = u16_at(data, at + 6)?;
    let limit = at.checked_add(lcb).filter(|&l| l <= data.len())?;
    let mut p = at + header;
    if mm == MM_SHAPE_FILE {
        p += 1 + *data.get(p)? as usize;
    }

    let mut alt_text = None;
    let mut picture = None;
    while let Some(r) = record(data, p, limit) {
        match r.kind {
            RT_SP_CONTAINER => alt_text = alt_text.or_else(|| officeart::description(data, &r)),
            RT_FBSE => picture = picture.or_else(|| officeart::blip_in_fbse(data, &r, None)),
            _ => {}
        }
        p = r.end;
    }
    let (data, extension) = picture?;
    Some(Picture {
        data,
        extension,
        alt_text,
    })
}

#[cfg(test)]
pub(super) mod tests {
    use super::*;
    use crate::officeart::tests::{fbse, fopt, png_blip};

    /// A `PICF` + shape container (with alternative text) + blip store entry holding `png`.
    pub(crate) fn picf_with_png(png: &[u8], alt: Option<&str>) -> Vec<u8> {
        let mut shape = Vec::new();
        if alt.is_some() {
            shape.extend(fopt(None, alt));
        }
        let mut sp = (0x000Fu16).to_le_bytes().to_vec();
        sp.extend(RT_SP_CONTAINER.to_le_bytes());
        sp.extend((shape.len() as u32).to_le_bytes());
        sp.extend(shape);

        let blip = png_blip(png);
        let fbse = fbse(blip.len() as u32, 0, &blip);

        let mut out = Vec::new();
        let total = 68 + sp.len() + fbse.len();
        out.extend((total as u32).to_le_bytes());
        out.extend(68u16.to_le_bytes());
        out.extend(0x64u16.to_le_bytes());
        out.extend([0u8; 60]);
        out.extend(sp);
        out.extend(fbse);
        out
    }

    #[test]
    fn a_png_blip_and_its_alternative_text_are_read() {
        let png = b"\x89PNG\r\n\x1a\nfake";
        let mut data = vec![0u8; 10];
        data.extend(picf_with_png(png, Some("A chart of sales")));
        let picture = read(&data, 10).expect("a picture");
        assert_eq!(picture.extension, "png");
        assert_eq!(picture.data, png);
        assert_eq!(picture.alt_text.as_deref(), Some("A chart of sales"));
    }

    #[test]
    fn a_picf_past_the_stream_is_none() {
        let data = picf_with_png(b"x", None);
        assert!(read(&data[..data.len() - 1], 0).is_none());
    }
}
