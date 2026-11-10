//! Floating pictures: the shapes the main document's drawing anchors in its text. A drawn-object
//! character (`0x08`) at a character position is tied by the position table of anchors
//! (`PlcfSpaMom`, [MS-DOC] 2.9.253 `Spa`) to a shape id; the shape is in the main drawing of the
//! `OfficeArtContent` ([MS-DOC] 2.9.172) and names its picture by an index into the drawing
//! group's blip store, whose entries point at the blips in the `WordDocument` stream.

use std::collections::HashMap;

use super::fib::FcLcb;
use crate::officeart::{self, picture_shapes, record, u32_at, Record, ShapePicture, RT_FBSE};

const RT_DGG_CONTAINER: u16 = 0xF000;
const RT_BSTORE_CONTAINER: u16 = 0xF001;
const RT_DG_CONTAINER: u16 = 0xF002;

/// `OfficeArtWordDrawing.dgglbl` of the main document's drawing; the headers' is 1.
const DGGLBL_MAIN: u8 = 0;
/// Size of a `Spa`: the shape id, the anchor's rectangle, flags and the text box count.
const SPA_SIZE: usize = 26;

/// A picture a drawn-object character shows.
#[derive(Debug, PartialEq, Eq)]
pub(super) struct Floating {
    /// The character position of the anchor.
    pub cp: u32,
    /// The index of its picture in [`Drawing::pictures`].
    pub picture: usize,
    pub alt_text: Option<String>,
}

/// The pictures the main document's shapes show and where they are anchored.
#[derive(Debug, Default)]
pub(super) struct Drawing {
    /// In character position order.
    pub anchors: Vec<Floating>,
    /// Each picture once, however many shapes show it: the file bytes and the extension.
    pub pictures: Vec<(Vec<u8>, &'static str)>,
}

impl Drawing {
    /// The floating picture anchored at `cp`.
    pub fn at(&self, cp: u32) -> Option<&Floating> {
        self.anchors
            .binary_search_by_key(&cp, |a| a.cp)
            .ok()
            .map(|i| &self.anchors[i])
    }
}

/// A range of a stream, if it lies inside it.
fn range(data: &[u8], at: FcLcb) -> Option<&[u8]> {
    let start = at.fc as usize;
    data.get(start..start.checked_add(at.lcb as usize)?)
}

/// The floating pictures of the main document: `spa` is the anchor table and `dgg` the
/// `OfficeArtContent`, both ranges of the table stream. A structure that cannot be followed
/// gives no pictures — the body text is never lost to it.
pub(super) fn read(word_document: &[u8], table: &[u8], spa: FcLcb, dgg: FcLcb) -> Drawing {
    read_in(word_document, table, spa, dgg).unwrap_or_default()
}

fn read_in(word_document: &[u8], table: &[u8], spa: FcLcb, dgg: FcLcb) -> Option<Drawing> {
    let spas = range(table, spa)?;
    let content = range(table, dgg)?;

    // The anchors: `n + 1` character positions, then `n` of the 26-byte `Spa`, whose first four
    // bytes are the shape id.
    if spas.len() < 4 || (spas.len() - 4) % (4 + SPA_SIZE) != 0 {
        return None;
    }
    let n = (spas.len() - 4) / (4 + SPA_SIZE);
    let anchors: Vec<(u32, u32)> = (0..n)
        .map(|i| {
            let cp = u32_at(spas, i * 4)?;
            let shape = u32_at(spas, (n + 1) * 4 + i * SPA_SIZE)?;
            Some((cp, shape))
        })
        .collect::<Option<_>>()?;

    // The drawing group data, then the drawings, each led by a byte that says whose it is.
    let group = record(content, 0, content.len()).filter(|r| r.kind == RT_DGG_CONTAINER)?;
    let store = blip_store(content, &group);
    let mut shapes: HashMap<u32, ShapePicture> = HashMap::new();
    let mut at = group.end;
    while let Some(label) = content.get(at) {
        let Some(drawing) =
            record(content, at + 1, content.len()).filter(|r| r.kind == RT_DG_CONTAINER)
        else {
            break;
        };
        if *label == DGGLBL_MAIN {
            let mut found = Vec::new();
            picture_shapes(content, &drawing, 0, &mut found);
            shapes.extend(found);
        }
        at = drawing.end;
    }

    let mut out = Drawing::default();
    let mut index_of: HashMap<u32, Option<usize>> = HashMap::new();
    for (cp, shape) in anchors {
        let Some(ShapePicture {
            blip: Some(blip),
            alt_text,
        }) = shapes.get(&shape)
        else {
            continue;
        };
        let picture = *index_of.entry(*blip).or_insert_with(|| {
            let entry = store.get(blip.checked_sub(1)? as usize)?.as_ref()?;
            let found = officeart::blip_in_fbse(content, entry, Some(word_document))?;
            out.pictures.push(found);
            Some(out.pictures.len() - 1)
        });
        if let Some(picture) = picture {
            out.anchors.push(Floating {
                cp,
                picture,
                alt_text: alt_text.clone(),
            });
        }
    }
    out.anchors.sort_by_key(|a| a.cp);
    Some(out)
}

/// The blip store entries of the drawing group, one per picture in order — a shape's picture
/// index counts from 1. An entry that is not a blip store entry stays as a gap.
fn blip_store(content: &[u8], group: &Record) -> Vec<Option<Record>> {
    let mut at = group.start;
    while let Some(r) = record(content, at, group.end) {
        at = r.end;
        if r.kind != RT_BSTORE_CONTAINER {
            continue;
        }
        let mut entries = Vec::new();
        let mut p = r.start;
        while let Some(e) = record(content, p, r.end) {
            p = e.end;
            entries.push((e.kind == RT_FBSE).then_some(e));
        }
        return entries;
    }
    Vec::new()
}
