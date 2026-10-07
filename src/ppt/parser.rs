//! Reads a PowerPoint 97-2003 presentation's slides and notes into the document model.

use std::collections::HashMap;
use std::io::{Cursor as IoCursor, Read};

use crate::detect::FormatType;
use crate::error::{Error, Result};
use crate::model::{Document, HeadingLevel, Paragraph, Section, TextRun};

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
        Ok(Self {
            data: std::fs::read(path)?,
        })
    }

    /// Read a .ppt presentation held in memory.
    pub fn from_bytes(data: Vec<u8>) -> Result<Self> {
        Ok(Self { data })
    }

    /// Parse the presentation: one section per slide, in presentation order, with the
    /// slide's text as paragraphs and its notes as the section's notes.
    pub fn parse(&mut self) -> Result<Document> {
        let mut container = cfb::CompoundFile::open(IoCursor::new(&self.data[..]))
            .map_err(|e| invalid(format!("unreadable container: {e}")))?;
        let read = |container: &mut cfb::CompoundFile<IoCursor<&[u8]>>, name: &str| {
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
            for text in slide_texts(&stream, slide, outline) {
                add_text(&mut section, &text, &urls);
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

/// A slide's text in reading order: its shapes as drawn, each outline reference resolved to
/// the outline text it points at, then any outline text no shape referenced.
fn slide_texts(stream: &[u8], slide: Header, outline: &[Text]) -> Vec<Text> {
    let mut out = Vec::new();
    let mut used = vec![false; outline.len()];
    let mut pending: Vec<Header> = Vec::new();
    for atom in atoms_in(stream, slide) {
        if atom.kind == RT_OUTLINE_TEXT_REF_ATOM {
            out.extend(texts_in(stream, &pending));
            pending.clear();
            if let Some(index) = u32_at(stream, atom.start).map(|i| i as usize) {
                if let Some(text) = outline.get(index) {
                    used[index] = true;
                    out.push(text.clone());
                }
            }
        } else {
            pending.push(atom);
        }
    }
    out.extend(texts_in(stream, &pending));
    out.extend(
        outline
            .iter()
            .zip(&used)
            .filter(|(_, used)| !**used)
            .map(|(t, _)| t.clone()),
    );
    out
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
