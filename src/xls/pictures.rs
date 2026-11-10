//! Pictures of a BIFF8 workbook: the pictures are kept once in the workbook's drawing group
//! (`MsoDrawingGroup`, [MS-XLS] 2.4.155) — a blip store whose entries hold the picture files —
//! and each sheet's drawing (`MsoDrawing`, 2.4.156, in as many records as it takes) is a list
//! of shapes, those that show a picture naming one by its index in the store.

use super::records::{self as rec, Records};
use crate::officeart::{self, picture_shapes, record, Record, ShapePicture, RT_FBSE};

/// `OfficeArtDggContainer` and its `OfficeArtBStoreContainer`.
const RT_DGG_CONTAINER: u16 = 0xF000;
const RT_BSTORE_CONTAINER: u16 = 0xF001;
/// `MsoDrawingGroup` and `MsoDrawing`: records carrying OfficeArt data.
pub(super) const MSO_DRAWING_GROUP: u16 = 0x00EB;
const MSO_DRAWING: u16 = 0x00EC;

/// The workbook's blip store: the drawing group's bytes and where its entries are.
#[derive(Debug)]
pub(super) struct DrawingGroup {
    bytes: Vec<u8>,
    /// One per picture in order — a shape's picture index counts from 1. An entry that is not a
    /// blip store entry stays as a gap.
    store: Vec<Option<Record>>,
}

impl DrawingGroup {
    /// The blip store of the `OfficeArtDggContainer` the `MsoDrawingGroup` records hold (all of
    /// them, and their continuations, joined). None for a drawing group without one.
    pub(super) fn new(bytes: Vec<u8>) -> Option<Self> {
        let group = record(&bytes, 0, bytes.len()).filter(|r| r.kind == RT_DGG_CONTAINER)?;
        let mut store = Vec::new();
        let mut at = group.start;
        while let Some(r) = record(&bytes, at, group.end) {
            at = r.end;
            if r.kind != RT_BSTORE_CONTAINER {
                continue;
            }
            let mut p = r.start;
            while let Some(entry) = record(&bytes, p, r.end) {
                p = entry.end;
                store.push((entry.kind == RT_FBSE).then_some(entry));
            }
            break;
        }
        Some(Self { bytes, store })
    }

    /// The picture a shape's picture index names, as a file and its extension. The entry holds
    /// the blip itself; an emptied slot, or a picture format that is not a file this reader
    /// produces (Macintosh PICT), gives none.
    pub(super) fn picture(&self, blip: u32) -> Option<(Vec<u8>, &'static str)> {
        let entry = self.store.get(blip.checked_sub(1)? as usize)?.as_ref()?;
        officeart::blip_in_fbse(&self.bytes, entry, None)
    }
}

/// How many of the `CONTINUE` segments after a `TXO` are its own: the text (`cchText`
/// characters, each segment led by a flags byte that says one or two bytes a character) and then
/// the formatting runs (`cbRuns` bytes). Any that follow carry on the sheet's drawing.
fn txo_own_segments(segments: &[&[u8]]) -> usize {
    let head = segments[0];
    let field = |at: usize| {
        head.get(at..at + 2)
            .map(|b| usize::from(u16::from_le_bytes([b[0], b[1]])))
    };
    let (Some(chars), Some(run_bytes)) = (field(10), field(12)) else {
        return 0;
    };
    let mut used = 0;
    let mut remaining = chars;
    while remaining > 0 {
        let Some(segment) = segments.get(1 + used) else {
            return used;
        };
        let width = if segment.first().is_some_and(|f| f & 1 != 0) {
            2
        } else {
            1
        };
        remaining = remaining.saturating_sub(segment.len().saturating_sub(1) / width);
        used += 1;
    }
    let mut remaining = run_bytes;
    while remaining > 0 {
        let Some(segment) = segments.get(1 + used) else {
            return used;
        };
        remaining = remaining.saturating_sub(segment.len());
        used += 1;
    }
    used
}

/// The shapes of the sheet whose substream begins at `offset` that show a picture, in drawing
/// order. Substreams nested in it (an embedded chart) are passed over.
///
/// The sheet's drawing is written as `MsoDrawing` records, and carries on in `CONTINUE` records
/// wherever it runs on: after an `MsoDrawing`, after an `Obj`, and after a text box's own text
/// and formatting. Its bytes, in stream order, are one `OfficeArtDgContainer`.
pub(super) fn sheet_pictures(stream: &[u8], offset: usize) -> Vec<ShapePicture> {
    let mut drawing = Vec::new();
    let mut records = Records::new(stream, offset);
    let _bof = records.next();
    let mut depth = 0usize;
    while let Some(Ok(record)) = records.next() {
        // The segments of the record that are drawing: all of an `MsoDrawing`, the continuations
        // of an `Obj`, those of a `TXO` that are not its own.
        let skip = match record.kind {
            rec::BOF | rec::BOF_BIFF2 | rec::BOF_BIFF3 | rec::BOF_BIFF4 => {
                depth += 1;
                continue;
            }
            rec::EOF if depth == 0 => break,
            rec::EOF => {
                depth -= 1;
                continue;
            }
            _ if depth > 0 => continue,
            MSO_DRAWING => 0,
            rec::OBJ => 1,
            rec::TXO => 1 + txo_own_segments(&record.segments),
            _ => continue,
        };
        for segment in record.segments.iter().skip(skip) {
            drawing.extend_from_slice(segment);
        }
        // An `Obj` or `TXO` is not itself drawing data.
        if skip > 0 && record.segments.len() <= skip {
            continue;
        }
    }
    // The sheet's drawing is one container written over the records of the sheet: its declared
    // length is that of all of them together.
    let Some(container) = record(&drawing, 0, drawing.len()) else {
        return Vec::new();
    };
    let mut shapes = Vec::new();
    picture_shapes(&drawing, &container, 0, &mut shapes);
    shapes.into_iter().map(|(_, picture)| picture).collect()
}
