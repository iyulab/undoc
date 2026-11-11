//! Reads a PowerPoint 97-2003 presentation's slides and notes into the document model.

use std::collections::HashMap;
use std::io::Read;

use crate::detect::FormatType;
use crate::error::{Error, Result};
use crate::model::{
    Block, Cell, Document, HeadingLevel, Paragraph, Resource, Row, Section, Table, TextRun,
};
use crate::officeart::{self, ShapePicture};

/// Record types ([MS-PPT] 2.13.24 `RecordType`).
const RT_DOCUMENT: u16 = 0x03E8;
const RT_SLIDE: u16 = 0x03EE;
const RT_SLIDE_ATOM: u16 = 0x03EF;
const RT_NOTES: u16 = 0x03F0;
const RT_SLIDE_PERSIST_ATOM: u16 = 0x03F3;
const RT_SLIDE_LIST_WITH_TEXT: u16 = 0x0FF0;
const RT_USER_EDIT_ATOM: u16 = 0x0FF5;
const RT_PERSIST_DIRECTORY_ATOM: u16 = 0x1772;
const RT_OUTLINE_TEXT_REF_ATOM: u16 = 0x0F9E;
const RT_TEXT_HEADER_ATOM: u16 = 0x0F9F;
const RT_TEXT_CHARS_ATOM: u16 = 0x0FA0;
const RT_TEXT_BYTES_ATOM: u16 = 0x0FA8;
const RT_EX_OBJ_LIST: u16 = 0x0409;
/// `PPDrawingGroup`: the drawing group, whose blip store lists the presentation's pictures.
const RT_DRAWING_GROUP: u16 = 0x040B;
/// OfficeArt containers on the way to the blip store ([MS-ODRAW] `OfficeArtDggContainer`,
/// `OfficeArtBStoreContainer`).
const RT_DGG_CONTAINER: u16 = 0xF000;
const RT_BSTORE_CONTAINER: u16 = 0xF001;
const RT_SPGR_CONTAINER: u16 = 0xF003;
/// `OfficeArtChildAnchor`: a shape inside a group, in the group's coordinate space.
const RT_CHILD_ANCHOR: u16 = 0xF00F;
/// Group shape properties that make a group a table: `tableProperties` (bit 0 set; bit 1: the
/// first row is a header) and `tableRowProperties` (an array with one element per row).
const PROP_TABLE_PROPERTIES: u16 = 0x039F;
const PROP_TABLE_ROW_PROPERTIES: u16 = 0x03A0;
const TABLE_FIRST_ROW: u32 = 0x2;
/// Shape type of a table cell (`msosptRectangle`).
const SHAPE_RECTANGLE: u16 = 1;
/// How many cells, and how many grid positions, a table may have: a group with more is not
/// read as a table.
const MAX_TABLE_CELLS: usize = 10_000;
const MAX_GRID_POSITIONS: usize = 1_000_000;
/// How deep shape groups are followed.
const MAX_GROUP_DEPTH: u8 = 16;
const RT_EX_HYPERLINK: u16 = 0x0FD7;
const RT_EX_HYPERLINK_ATOM: u16 = 0x0FD3;
const RT_CSTRING: u16 = 0x0FBA;
const RT_INTERACTIVE_INFO_ATOM: u16 = 0x0FF3;
const RT_TX_INTERACTIVE_INFO_ATOM: u16 = 0x0FDF;
/// `CString` instance holding an `ExHyperlink`'s target.
const HYPERLINK_TARGET: u16 = 1;

/// `CurrentUserAtom.headerToken` of an encrypted presentation.
const ENCRYPTED_TOKEN: u32 = 0xF3D1_C4DF;
/// `SlideListWithText` instances: slides and notes.
const LIST_SLIDES: u16 = 0;
const LIST_NOTES: u16 = 2;

/// `TextHeaderAtom.textType` values ([MS-PPT] 2.13.33 `TextTypeEnum`).
const TEXT_TITLE: u32 = 0;
const TEXT_NOTES: u32 = 2;
const TEXT_CENTER_BODY: u32 = 5;
const TEXT_CENTER_TITLE: u32 = 6;

/// A record header: version, instance, type, and the data's extent in the stream.
#[derive(Debug, Clone, Copy)]
struct Header {
    version: u8,
    instance: u16,
    kind: u16,
    start: usize,
    end: usize,
}

fn header(stream: &[u8], at: usize) -> Option<Header> {
    let h = stream.get(at..at + 8)?;
    let ver_inst = u16::from_le_bytes([h[0], h[1]]);
    let len = u32::from_le_bytes([h[4], h[5], h[6], h[7]]) as usize;
    let start = at + 8;
    let end = start.checked_add(len).filter(|&e| e <= stream.len())?;
    Some(Header {
        version: (ver_inst & 0x0F) as u8,
        instance: ver_inst >> 4,
        kind: u16::from_le_bytes([h[2], h[3]]),
        start,
        end,
    })
}

/// The records directly inside `[start, end)`.
fn children(stream: &[u8], start: usize, end: usize) -> Vec<Header> {
    let mut out = Vec::new();
    let mut at = start;
    while at + 8 <= end {
        let Some(h) = header(stream, at).filter(|h| h.end <= end) else {
            break;
        };
        at = h.end;
        out.push(h);
    }
    out
}

fn u32_at(data: &[u8], at: usize) -> Option<u32> {
    data.get(at..at + 4)
        .map(|b| u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
}

fn invalid(what: impl Into<String>) -> Error {
    Error::InvalidData(format!("PowerPoint presentation: {}", what.into()))
}

/// A run of text from a text atom, with the kind of placeholder it fills and the hyperlinks
/// over it: `[begin, end)` in UTF-16 units of the text, and the hyperlink's id.
#[derive(Debug, Clone)]
struct Text {
    kind: u32,
    text: String,
    links: Vec<(u32, u32, u32)>,
}

/// Text atoms in order: each `TextHeaderAtom` names the kind of the characters after it, and
/// an `InteractiveInfoAtom` / `TxInteractiveInfoAtom` pair after them links a range of them.
/// Containers among `records` are looked into.
fn texts_in(stream: &[u8], records: &[Header]) -> Vec<Text> {
    let mut atoms = Vec::new();
    for &r in records {
        if r.version == 0x0F {
            atoms.extend(atoms_in(stream, r));
        } else {
            atoms.push(r);
        }
    }
    let mut out: Vec<Text> = Vec::new();
    let mut kind = None;
    let mut link = None;
    for r in &atoms {
        match r.kind {
            RT_INTERACTIVE_INFO_ATOM => link = u32_at(stream, r.start + 4).filter(|&id| id != 0),
            RT_TX_INTERACTIVE_INFO_ATOM => {
                if let (Some(id), Some(begin), Some(end), Some(text)) = (
                    link.take(),
                    u32_at(stream, r.start),
                    u32_at(stream, r.start + 4),
                    out.last_mut(),
                ) {
                    text.links.push((begin, end, id));
                }
            }
            RT_TEXT_HEADER_ATOM => kind = u32_at(stream, r.start),
            RT_TEXT_CHARS_ATOM | RT_TEXT_BYTES_ATOM => {
                let data = &stream[r.start..r.end];
                let text = if r.kind == RT_TEXT_CHARS_ATOM {
                    let units: Vec<u16> = data
                        .as_chunks::<2>()
                        .0
                        .iter()
                        .map(|&c| u16::from_le_bytes(c))
                        .collect();
                    String::from_utf16_lossy(&units)
                } else {
                    // The low bytes of UTF-16 code units: Latin-1.
                    data.iter().map(|&b| b as char).collect()
                };
                out.push(Text {
                    kind: kind.take().unwrap_or(u32::MAX),
                    text,
                    links: Vec::new(),
                });
            }
            _ => {}
        }
    }
    out
}

/// Reader for PowerPoint 97-2003 presentations (.ppt).
pub struct PptParser {
    data: Vec<u8>,
}

impl PptParser {
    /// Open a .ppt file.
    #[cfg(not(target_arch = "wasm32"))]
    pub fn open(path: impl AsRef<std::path::Path>) -> Result<Self> {
        let path = path.as_ref();
        Ok(Self {
            data: std::fs::read(path).map_err(|e| Error::read_file(path, e))?,
        })
    }

    /// Read a .ppt presentation held in memory.
    pub fn from_bytes(data: Vec<u8>) -> Result<Self> {
        Ok(Self { data })
    }

    /// Parse the presentation: one section per slide, in presentation order, with the
    /// slide's text as paragraphs and its notes as the section's notes.
    pub fn parse(&mut self) -> Result<Document> {
        let mut container = crate::cfb_open::open(&self.data)
            .map_err(|e| invalid(format!("unreadable container: {e}")))?;
        let read = |container: &mut crate::cfb_open::Container<'_>, name: &str| {
            let mut out = Vec::new();
            container
                .open_stream(name)
                .map_err(|_| {
                    invalid(format!(
                        "the {} stream is missing",
                        name.trim_start_matches('/')
                    ))
                })?
                .read_to_end(&mut out)?;
            Ok::<_, Error>(out)
        };
        let current_user = read(&mut container, "/Current User")?;
        let stream = read(&mut container, "/PowerPoint Document")?;

        // CurrentUserAtom: header, size, headerToken, offsetToCurrentEdit.
        let token =
            u32_at(&current_user, 12).ok_or_else(|| invalid("Current User is truncated"))?;
        if token == ENCRYPTED_TOKEN {
            return Err(Error::Encrypted);
        }
        let current_edit =
            u32_at(&current_user, 16).ok_or_else(|| invalid("Current User is truncated"))? as usize;

        let (persist, document_ref) = read_persist_directory(&stream, current_edit)?;
        let document_at = *persist
            .get(&document_ref)
            .ok_or_else(|| invalid("the document's persist object is missing"))?;
        let document = header(&stream, document_at)
            .filter(|h| h.kind == RT_DOCUMENT)
            .ok_or_else(|| invalid("the document container is missing"))?;

        let urls = hyperlink_targets(&stream, document);
        // The `Pictures` stream is absent from a presentation without delayed pictures.
        let pictures_stream = read(&mut container, "/Pictures").ok();
        let store = blip_store(&stream, document);
        let mut pictures = PictureIds::default();

        // The slide and notes lists: each SlidePersistAtom, then that slide's outline texts.
        let mut slides: Vec<(u32, Vec<Text>)> = Vec::new();
        let mut notes_by_id: HashMap<u32, u32> = HashMap::new();
        for list in children(&stream, document.start, document.end)
            .into_iter()
            .filter(|h| h.kind == RT_SLIDE_LIST_WITH_TEXT)
        {
            let records = children(&stream, list.start, list.end);
            let mut i = 0;
            while i < records.len() {
                let r = records[i];
                i += 1;
                if r.kind != RT_SLIDE_PERSIST_ATOM {
                    continue;
                }
                let (Some(persist_ref), Some(slide_id)) =
                    (u32_at(&stream, r.start), u32_at(&stream, r.start + 12))
                else {
                    continue;
                };
                let next = records[i..]
                    .iter()
                    .position(|h| h.kind == RT_SLIDE_PERSIST_ATOM)
                    .map_or(records.len(), |p| i + p);
                match list.instance {
                    LIST_SLIDES => slides.push((persist_ref, texts_in(&stream, &records[i..next]))),
                    LIST_NOTES => {
                        notes_by_id.insert(slide_id, persist_ref);
                    }
                    _ => {}
                }
                i = next;
            }
        }

        let mut doc = Document::new();
        doc.format = FormatType::Ppt;
        doc.metadata = crate::summary::read(&mut container);
        for (index, (persist_ref, outline)) in slides.iter().enumerate() {
            let mut section = Section::new(index);
            section.name = Some(format!("Slide {}", index + 1));
            let Some(slide) = persist
                .get(persist_ref)
                .and_then(|&at| header(&stream, at))
                .filter(|h| h.kind == RT_SLIDE)
            else {
                // Without its container the slide's outline text is all there is.
                for text in outline {
                    add_text(&mut section, text, &urls);
                }
                doc.add_section(section);
                continue;
            };
            for item in slide_items(&stream, slide, outline, &urls) {
                match item {
                    Item::Text(text) => add_text(&mut section, &text, &urls),
                    Item::Table(table) => section.add_table(table),
                    Item::Picture(shape) => {
                        let blip = shape.blip.and_then(|blip| {
                            pictures.resolve(&mut doc, blip, |blip| {
                                let entry = blip
                                    .checked_sub(1)
                                    .and_then(|i| store.get(i as usize))
                                    .copied()
                                    .flatten()?;
                                officeart::blip_in_fbse(&stream, &entry, pictures_stream.as_deref())
                            })
                        });
                        if let Some(id) = blip {
                            section.add_block(Block::Image {
                                resource_id: id,
                                alt_text: shape.alt_text,
                                width: None,
                                height: None,
                            });
                        }
                    }
                }
            }

            // The notes page this slide names, through the notes list.
            let notes_id = children(&stream, slide.start, slide.end)
                .into_iter()
                .find(|h| h.kind == RT_SLIDE_ATOM)
                .and_then(|atom| u32_at(&stream, atom.start + 16))
                .filter(|&id| id != 0);
            if let Some(notes) = notes_id
                .and_then(|id| notes_by_id.get(&id))
                .and_then(|r| persist.get(r))
                .and_then(|&at| header(&stream, at))
                .filter(|h| h.kind == RT_NOTES)
            {
                let mut paragraphs = Vec::new();
                for text in drawing_texts(&stream, notes)
                    .into_iter()
                    .filter(|t| t.kind == TEXT_NOTES)
                {
                    paragraphs.extend(paragraphs_of(&text, &urls));
                }
                if !paragraphs.is_empty() {
                    section.notes = Some(paragraphs);
                }
            }
            doc.add_section(section);
        }
        Ok(doc)
    }
}

/// Hyperlink targets by id, from the document's external object list.
fn hyperlink_targets(stream: &[u8], document: Header) -> HashMap<u32, String> {
    let mut urls = HashMap::new();
    for list in children(stream, document.start, document.end)
        .into_iter()
        .filter(|h| h.kind == RT_EX_OBJ_LIST)
    {
        for link in children(stream, list.start, list.end)
            .into_iter()
            .filter(|h| h.kind == RT_EX_HYPERLINK)
        {
            let parts = children(stream, link.start, link.end);
            let id = parts
                .iter()
                .find(|h| h.kind == RT_EX_HYPERLINK_ATOM)
                .and_then(|h| u32_at(stream, h.start));
            let target = parts
                .iter()
                .find(|h| h.kind == RT_CSTRING && h.instance == HYPERLINK_TARGET)
                .map(|h| {
                    let units: Vec<u16> = stream[h.start..h.end]
                        .as_chunks::<2>()
                        .0
                        .iter()
                        .map(|&c| u16::from_le_bytes(c))
                        .collect();
                    String::from_utf16_lossy(&units)
                });
            if let (Some(id), Some(target)) = (id, target.filter(|t| !t.is_empty())) {
                urls.insert(id, target);
            }
        }
    }
    urls
}

/// Follow the user edits from the current one back to the first, and build the persist
/// object directory — later edits override earlier ones.
fn read_persist_directory(
    stream: &[u8],
    current_edit: usize,
) -> Result<(HashMap<u32, usize>, u32)> {
    let mut edits = Vec::new();
    let mut at = current_edit;
    loop {
        let edit = header(stream, at)
            .filter(|h| h.kind == RT_USER_EDIT_ATOM)
            .ok_or_else(|| invalid(format!("no user edit at offset {at}")))?;
        let field = |offset: usize| {
            u32_at(stream, edit.start + offset).ok_or_else(|| invalid("a user edit is truncated"))
        };
        let (last_edit, directory, document_ref) = (field(8)?, field(12)?, field(16)?);
        edits.push((directory as usize, document_ref));
        if last_edit == 0 || edits.len() > 4096 {
            break;
        }
        // A chain must move backwards; anything else is a loop.
        if last_edit as usize >= at {
            return Err(invalid("the user edit chain does not move backwards"));
        }
        at = last_edit as usize;
    }

    let mut persist = HashMap::new();
    for &(directory, _) in edits.iter().rev() {
        let Some(dir) = header(stream, directory).filter(|h| h.kind == RT_PERSIST_DIRECTORY_ATOM)
        else {
            return Err(invalid("a persist directory is missing"));
        };
        let mut p = dir.start;
        while p + 4 <= dir.end {
            let entry = u32_at(stream, p).unwrap();
            p += 4;
            let (first, count) = (entry & 0x000F_FFFF, entry >> 20);
            for k in 0..count {
                let Some(offset) = u32_at(stream, p).filter(|_| p + 4 <= dir.end) else {
                    break;
                };
                persist.insert(first + k, offset as usize);
                p += 4;
            }
        }
    }
    Ok((persist, edits[0].1))
}

/// Every atom inside a container, depth first — drawing order for a slide's shapes.
fn atoms_in(stream: &[u8], container: Header) -> Vec<Header> {
    fn walk(stream: &[u8], start: usize, end: usize, out: &mut Vec<Header>) {
        for h in children(stream, start, end) {
            if h.version == 0x0F {
                walk(stream, h.start, h.end, out);
            } else {
                out.push(h);
            }
        }
    }
    let mut atoms = Vec::new();
    walk(stream, container.start, container.end, &mut atoms);
    atoms
}

/// Every text inside a container, in drawing order.
fn drawing_texts(stream: &[u8], container: Header) -> Vec<Text> {
    texts_in(stream, &atoms_in(stream, container))
}

/// What a slide shows, in reading order.
enum Item {
    Text(Text),
    Picture(ShapePicture),
    Table(Table),
}

/// A shape's property table, as the atom it is read from.
fn art_record(atom: Header) -> officeart::Record {
    officeart::Record {
        instance: atom.instance,
        kind: atom.kind,
        start: atom.start,
        end: atom.end,
    }
}

/// A slide's content in reading order: its shapes as drawn — text, the pictures that name a
/// blip, and tables — each outline reference resolved to the outline text it points at, then
/// any outline text no shape referenced.
fn slide_items(
    stream: &[u8],
    slide: Header,
    outline: &[Text],
    urls: &HashMap<u32, String>,
) -> Vec<Item> {
    /// Close the shape being read: a picture goes after the text drawn before it.
    fn end_shape(
        out: &mut Vec<Item>,
        pending: &mut Vec<Header>,
        stream: &[u8],
        shape: &mut ShapePicture,
    ) {
        let done = std::mem::take(shape);
        if done.blip.is_some() {
            out.extend(texts_in(stream, pending).into_iter().map(Item::Text));
            pending.clear();
            out.push(Item::Picture(done));
        }
    }

    let mut tables = Vec::new();
    find_tables(stream, slide.start, slide.end, urls, 0, &mut tables);
    let mut emitted = vec![false; tables.len()];

    let mut out = Vec::new();
    let mut used = vec![false; outline.len()];
    let mut pending: Vec<Header> = Vec::new();
    let mut shape = ShapePicture::default();
    for atom in atoms_in(stream, slide) {
        // Everything inside a table group belongs to the table, which appears where the group
        // is drawn.
        if let Some(i) = tables
            .iter()
            .position(|t| t.start <= atom.start && atom.start < t.end)
        {
            if !emitted[i] {
                emitted[i] = true;
                end_shape(&mut out, &mut pending, stream, &mut shape);
                out.extend(texts_in(stream, &pending).into_iter().map(Item::Text));
                pending.clear();
                out.push(Item::Table(tables[i].table.clone()));
            }
            continue;
        }
        match atom.kind {
            RT_OUTLINE_TEXT_REF_ATOM => {
                out.extend(texts_in(stream, &pending).into_iter().map(Item::Text));
                pending.clear();
                if let Some(index) = u32_at(stream, atom.start).map(|i| i as usize) {
                    if let Some(text) = outline.get(index) {
                        used[index] = true;
                        out.push(Item::Text(text.clone()));
                    }
                }
            }
            // A shape starts with its `OfficeArtFSP`.
            officeart::RT_FSP => end_shape(&mut out, &mut pending, stream, &mut shape),
            officeart::RT_FOPT | officeart::RT_TERTIARY_FOPT => {
                let found = officeart::picture_properties(stream, &art_record(atom));
                shape.blip = shape.blip.or(found.blip);
                shape.alt_text = shape.alt_text.take().or(found.alt_text);
            }
            _ => pending.push(atom),
        }
    }
    end_shape(&mut out, &mut pending, stream, &mut shape);
    out.extend(texts_in(stream, &pending).into_iter().map(Item::Text));
    out.extend(
        outline
            .iter()
            .zip(&used)
            .filter(|(_, used)| !**used)
            .map(|(t, _)| Item::Text(t.clone())),
    );
    out
}

/// A table drawn as a group of shapes: the byte range of the group in the stream, and the table.
struct TableGroup {
    start: usize,
    end: usize,
    table: Table,
}

/// The groups of shapes within `[start, end)` that are tables ([MS-ODRAW] group shape
/// properties `tableProperties` and `tableRowProperties`; PowerPoint 97-2003 has no table
/// object). A group that sets them but whose cells do not tile a grid is not read as a table:
/// its texts are read as the texts they are.
fn find_tables(
    stream: &[u8],
    start: usize,
    end: usize,
    urls: &HashMap<u32, String>,
    depth: u8,
    out: &mut Vec<TableGroup>,
) {
    for h in children(stream, start, end) {
        if h.version != 0x0F || depth >= MAX_GROUP_DEPTH {
            continue;
        }
        if h.kind == RT_SPGR_CONTAINER {
            if let Some(table) = table_of(stream, h, urls) {
                out.push(TableGroup {
                    // The group's own header is not an atom; its first byte is.
                    start: h.start,
                    end: h.end,
                    table,
                });
                continue;
            }
        }
        find_tables(stream, h.start, h.end, urls, depth + 1, out);
    }
}

/// One cell shape of a table group: its place in the group's coordinate space and its text.
struct CellShape {
    left: i32,
    top: i32,
    right: i32,
    bottom: i32,
    content: Vec<Paragraph>,
}

/// The table a group of shapes is, if it is one that can be read: the group's own shape names
/// the table properties and its rows; the cells are its rectangles, which together tile a grid
/// whose rows are those the group states. Cells are laid out by position — the order they are
/// stored in is not the reading order — and a cell that spans grid rows or columns is one
/// shape, as in the `.pptx` and `.docx` models.
fn table_of(stream: &[u8], group: Header, urls: &HashMap<u32, String>) -> Option<Table> {
    let shapes = children(stream, group.start, group.end);
    // The first shape is the group's own: its property tables make the group a table.
    let own = shapes
        .first()
        .filter(|h| h.kind == officeart::RT_SP_CONTAINER)?;
    let mut flags = None;
    let mut rows = None;
    for property_table in children(stream, own.start, own.end)
        .into_iter()
        .filter(|h| h.kind == officeart::RT_FOPT || h.kind == officeart::RT_TERTIARY_FOPT)
    {
        let record = art_record(property_table);
        if let Some(p) = officeart::property(stream, &record, PROP_TABLE_PROPERTIES) {
            flags = Some(p.value);
        }
        if let Some(p) = officeart::property(stream, &record, PROP_TABLE_ROW_PROPERTIES) {
            // An array property: element count (`u16`), allocated, element size — one element
            // per row.
            rows = p
                .complex
                .and_then(|c| c.get(..2))
                .map(|b| usize::from(u16::from_le_bytes([b[0], b[1]])));
        }
    }
    let (flags, rows) = (flags.filter(|f| f & 1 != 0)?, rows.filter(|&r| r > 0)?);

    let mut cells = Vec::new();
    for shape in &shapes[1..] {
        if shape.kind != officeart::RT_SP_CONTAINER {
            continue;
        }
        let parts = children(stream, shape.start, shape.end);
        let is_cell = parts
            .iter()
            .any(|h| h.kind == officeart::RT_FSP && h.instance == SHAPE_RECTANGLE);
        let Some(anchor) = parts
            .iter()
            .find(|h| h.kind == RT_CHILD_ANCHOR && h.end - h.start == 16)
        else {
            continue;
        };
        if !is_cell {
            continue;
        }
        let at = |i: usize| u32_at(stream, anchor.start + i * 4).map(|v| v as i32);
        let (left, top, right, bottom) = (at(0)?, at(1)?, at(2)?, at(3)?);
        let mut content = Vec::new();
        for text in texts_in(stream, &atoms_in(stream, *shape)) {
            content.extend(paragraphs_of(&text, urls));
        }
        cells.push(CellShape {
            left,
            top,
            right,
            bottom,
            content,
        });
        if cells.len() > MAX_TABLE_CELLS {
            return None;
        }
    }
    if cells.is_empty() {
        return None;
    }

    // The grid lines: every edge any cell has.
    let lines = |edges: Vec<i32>| {
        let mut edges = edges;
        edges.sort_unstable();
        edges.dedup();
        edges
    };
    let xs = lines(cells.iter().flat_map(|c| [c.left, c.right]).collect());
    let ys = lines(cells.iter().flat_map(|c| [c.top, c.bottom]).collect());
    let (n_cols, n_rows) = (xs.len() - 1, ys.len() - 1);
    if n_rows != rows || n_cols == 0 || n_rows.saturating_mul(n_cols) > MAX_GRID_POSITIONS {
        return None;
    }

    // Every position of the grid is covered by exactly one cell.
    let index = |lines: &[i32], v: i32| lines.binary_search(&v).ok();
    let mut covered = vec![false; n_rows * n_cols];
    let mut placed = Vec::new();
    for cell in cells {
        let (c0, c1) = (index(&xs, cell.left)?, index(&xs, cell.right)?);
        let (r0, r1) = (index(&ys, cell.top)?, index(&ys, cell.bottom)?);
        if c1 <= c0 || r1 <= r0 {
            return None;
        }
        for r in r0..r1 {
            for c in c0..c1 {
                if std::mem::replace(&mut covered[r * n_cols + c], true) {
                    return None;
                }
            }
        }
        placed.push((r0, c0, r1 - r0, c1 - c0, cell.content));
    }
    if covered.iter().any(|c| !c) {
        return None;
    }

    placed.sort_by_key(|&(row, col, ..)| (row, col));
    let mut table = Table::new();
    let mut placed = placed.into_iter().peekable();
    for row_index in 0..n_rows {
        let mut row = Row::new();
        row.is_header = row_index == 0 && flags & TABLE_FIRST_ROW != 0;
        while let Some((_, _, row_span, col_span, content)) =
            placed.next_if(|&(r, ..)| r == row_index)
        {
            let mut cell = Cell::new();
            cell.content = content;
            cell.row_span = u32::try_from(row_span).unwrap_or(1);
            cell.col_span = u32::try_from(col_span).unwrap_or(1);
            row.cells.push(cell);
        }
        table.add_row(row);
    }
    Some(table)
}

/// The blip store of the document's drawing group: one entry per picture, in order — a shape's
/// `pib` is a 1-based index into it. An entry that is not a blip store entry stays as a gap so
/// the indices hold.
fn blip_store(stream: &[u8], document: Header) -> Vec<Option<officeart::Record>> {
    let child = |parent: Header, kind: u16| {
        children(stream, parent.start, parent.end)
            .into_iter()
            .find(|h| h.kind == kind)
    };
    let Some(store) = children(stream, document.start, document.end)
        .into_iter()
        .filter(|h| h.kind == RT_DRAWING_GROUP)
        .filter_map(|group| child(group, RT_DGG_CONTAINER))
        .find_map(|dgg| child(dgg, RT_BSTORE_CONTAINER))
    else {
        return Vec::new();
    };
    children(stream, store.start, store.end)
        .into_iter()
        .map(|h| (h.kind == officeart::RT_FBSE).then(|| art_record(h)))
        .collect()
}

/// The resources the slides' pictures became, by blip store index: a picture shown on several
/// slides is one resource.
#[derive(Default)]
struct PictureIds {
    ids: HashMap<u32, Option<String>>,
}

impl PictureIds {
    /// The resource id of picture `blip`, adding it to `doc` on first use. None when the entry
    /// holds no picture this reader turns into a file.
    fn resolve(
        &mut self,
        doc: &mut Document,
        blip: u32,
        read: impl FnOnce(u32) -> Option<(Vec<u8>, &'static str)>,
    ) -> Option<String> {
        if let Some(known) = self.ids.get(&blip) {
            return known.clone();
        }
        let id = read(blip).map(|(data, extension)| {
            let id = format!("image{}", doc.resources.len() + 1);
            doc.add_resource(
                id.clone(),
                Resource::image(data, Some(format!("{id}.{extension}"))),
            );
            id
        });
        self.ids.insert(blip, id.clone());
        id
    }
}

/// A text atom's paragraphs: `\r` separates paragraphs, `\v` breaks a line within one, and
/// each linked range becomes a run carrying its target.
fn paragraphs_of(text: &Text, urls: &HashMap<u32, String>) -> Vec<Paragraph> {
    let link_at = |unit: u32| {
        text.links
            .iter()
            .find(|(begin, end, _)| *begin <= unit && unit < *end)
            .and_then(|(_, _, id)| urls.get(id))
    };
    let mut paragraphs = Vec::new();
    let mut paragraph = Paragraph::new();
    let mut run = String::new();
    let mut run_link: Option<&String> = None;
    let mut unit = 0u32;
    let flush = |paragraph: &mut Paragraph, run: &mut String, link: Option<&String>| {
        if !run.is_empty() {
            let mut r = TextRun::plain(std::mem::take(run));
            r.hyperlink = link.cloned();
            paragraph.runs.push(r);
        }
    };
    for ch in text.text.chars() {
        match ch {
            '\r' => {
                flush(&mut paragraph, &mut run, run_link);
                let done = std::mem::take(&mut paragraph);
                if !done.plain_text().trim().is_empty() {
                    paragraphs.push(done);
                }
            }
            '\u{0B}' => {
                flush(&mut paragraph, &mut run, run_link);
                if let Some(last) = paragraph.runs.last_mut() {
                    last.line_break = true;
                }
            }
            ch => {
                let link = link_at(unit);
                if link != run_link {
                    flush(&mut paragraph, &mut run, run_link);
                    run_link = link;
                }
                run.push(ch);
            }
        }
        unit += ch.len_utf16() as u32;
    }
    flush(&mut paragraph, &mut run, run_link);
    if !paragraph.plain_text().trim().is_empty() {
        paragraphs.push(paragraph);
    }
    paragraphs
}

/// Add a slide text's paragraphs; a title is a heading, as in the .pptx reader.
fn add_text(section: &mut Section, text: &Text, urls: &HashMap<u32, String>) {
    let heading = match text.kind {
        TEXT_TITLE | TEXT_CENTER_TITLE => HeadingLevel::H1,
        TEXT_CENTER_BODY => HeadingLevel::H2,
        _ => HeadingLevel::None,
    };
    for mut paragraph in paragraphs_of(text, urls) {
        paragraph.heading = heading;
        section.add_paragraph(paragraph);
    }
}
