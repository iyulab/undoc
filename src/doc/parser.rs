//! Assembles the main document text of a Word 97-2003 file into the document model.

use std::io::{Cursor, Read};

use super::chars::{self, CharProps, ChpxIndex};
use super::fib;
use super::fib::FcLcb;
use super::lists::{self, ListCounter, ListTable};
use super::props::{self, PapxIndex, ParaProps, StyleInfo};
use super::symbol;
use super::text::{self, DocChar};
use crate::detect::FormatType;
use crate::error::{Error, Result};
use crate::model::{
    Block, Cell, Document, HeadingLevel, InlineImage, ListInfo, ListType, Paragraph, Resource,
    RevisionType, Row, Section, Table, TextRun, TextStyle,
};

/// Paragraph mark.
const PARAGRAPH_END: char = '\r';
/// Cell mark, and the row-end mark of a table row.
const CELL_END: char = '\u{07}';
/// Manual line break.
const LINE_BREAK: char = '\u{0B}';
/// Page break or section mark.
const PAGE_BREAK: char = '\u{0C}';
/// Field begin, separator and end.
const FIELD_BEGIN: char = '\u{13}';
const FIELD_SEPARATOR: char = '\u{14}';
const FIELD_END: char = '\u{15}';
/// Picture anchor: an inline picture whose bytes are in the `Data` stream.
const PICTURE: char = '\u{01}';
/// Auto-numbered note reference.
const NOTE_REFERENCE: char = '\u{02}';
/// Non-breaking and optional hyphens.
const NON_BREAKING_HYPHEN: char = '\u{1E}';
const OPTIONAL_HYPHEN: char = '\u{1F}';

/// Reader for Word 97-2003 binary documents (.doc).
pub struct DocParser {
    data: Vec<u8>,
}

impl DocParser {
    /// Open a .doc file.
    #[cfg(not(target_arch = "wasm32"))]
    pub fn open(path: impl AsRef<std::path::Path>) -> Result<Self> {
        Ok(Self {
            data: std::fs::read(path)?,
        })
    }

    /// Read a .doc document held in memory.
    pub fn from_bytes(data: Vec<u8>) -> Result<Self> {
        Ok(Self { data })
    }

    /// Parse the document.
    pub fn parse(&mut self) -> Result<Document> {
        let mut container = cfb::CompoundFile::open(Cursor::new(&self.data[..]))
            .map_err(|e| Error::InvalidData(format!("Word document: unreadable container: {e}")))?;
        let word_document = read_stream(&mut container, "/WordDocument")?;
        let fib = fib::parse(&word_document)?;
        let table = read_stream(&mut container, &format!("/{}", fib.table_stream))?;
        // Only documents with inline pictures have one.
        let data = read_stream(&mut container, "/Data").unwrap_or_default();

        let pieces = text::parse_clx(&table, fib.clx)?;
        let chars = text::read_chars(&word_document, &pieces, 0, fib.ccp_text)?;
        let papx = props::read_papx(&word_document, &table, fib.plcf_bte_papx)?;
        let chpx = chars::read_chpx(&word_document, &table, fib.plcf_bte_chpx)?;
        let styles = props::read_styles(&table, fib.stshf);
        let fonts = symbol::read_font_names(&table, fib.sttbf_ffn);
        let list_table = lists::read_lists(&table, fib.plf_lst, fib.plf_lfo);

        // Footnotes follow the main text; endnotes follow the footnote, header and comment
        // stories. Each is listed by the position of its reference and the range of its text.
        let read_notes = |refs: FcLcb, txt: FcLcb, story: u32, marker: &str| -> Result<Vec<Note>> {
            let refs = plc_positions(&table, refs, 2);
            let txt = plc_positions(&table, txt, 0);
            let mut notes = Vec::with_capacity(refs.len());
            for (i, &cp) in refs.iter().enumerate() {
                let (Some(&from), Some(&to)) = (txt.get(i), txt.get(i + 1)) else {
                    break;
                };
                let text = note_text(&text::read_chars(
                    &word_document,
                    &pieces,
                    story + from,
                    story + to.max(from),
                )?);
                notes.push(Note {
                    cp,
                    label: format!("{marker}{}", i + 1),
                    text,
                });
            }
            Ok(notes)
        };
        let footnotes = read_notes(fib.plcffnd_ref, fib.plcffnd_txt, fib.ccp_text, "")?;
        let endnote_story = fib.ccp_text + fib.ccp_ftn + fib.ccp_hdd + fib.ccp_atn;
        let endnotes = read_notes(fib.plcfend_ref, fib.plcfend_txt, endnote_story, "e")?;

        let mut doc = Document::new();
        doc.format = FormatType::Doc;
        doc.metadata = crate::summary::read(&mut container);
        let mut section = Section::new(0);
        let notes: Vec<&Note> = footnotes.iter().chain(&endnotes).collect();
        let mut assembler =
            Assembler::new(&papx, &chpx, &styles, &fonts, &list_table, &notes, &data);
        for block in assembler.run(&chars) {
            section.add_block(block);
        }
        for (id, resource) in assembler.resources {
            doc.add_resource(id, resource);
        }
        // Note bodies close the section, as the .docx reader writes them.
        for note in notes.iter().filter(|n| !n.text.is_empty()) {
            section.add_paragraph(Paragraph::with_text(format!(
                "[^{}]: {}",
                note.label, note.text
            )));
        }
        // The first section's headers and footers — the stories Word shows on odd pages,
        // which are every page's unless the document asks for different ones.
        let hdd = plc_positions(&table, fib.plcf_hdd, 0);
        let story = fib.ccp_text + fib.ccp_ftn;
        let read_story = |index: usize| -> Result<Option<Vec<Paragraph>>> {
            let (Some(&from), Some(&to)) = (hdd.get(index), hdd.get(index + 1)) else {
                return Ok(None);
            };
            let chars =
                text::read_chars(&word_document, &pieces, story + from, story + to.max(from))?;
            let paragraphs = story_paragraphs(&chars);
            Ok((!paragraphs.is_empty()).then_some(paragraphs))
        };
        section.header = read_story(HDD_ODD_HEADER)?;
        section.footer = read_story(HDD_ODD_FOOTER)?;
        doc.add_section(section);
        Ok(doc)
    }
}

fn read_stream(container: &mut cfb::CompoundFile<Cursor<&[u8]>>, name: &str) -> Result<Vec<u8>> {
    let mut stream = container.open_stream(name).map_err(|_| {
        Error::InvalidData(format!(
            "Word document: the {} stream is missing",
            name.trim_start_matches('/')
        ))
    })?;
    let mut data = Vec::new();
    stream.read_to_end(&mut data)?;
    Ok(data)
}

/// Positions in `PlcfHdd`: six note separator stories, then six stories per section — even
/// header, odd header, even footer, odd footer, first-page header, first-page footer.
const HDD_ODD_HEADER: usize = 7;
const HDD_ODD_FOOTER: usize = 9;

/// A header or footer story as paragraphs: field instructions dropped, results kept, control
/// characters left out.
fn story_paragraphs(chars: &[DocChar]) -> Vec<Paragraph> {
    let mut paragraphs = Vec::new();
    let mut line = String::new();
    // Each open field: whether its result has started.
    let mut fields: Vec<bool> = Vec::new();
    for c in chars {
        match c.ch {
            FIELD_BEGIN => fields.push(false),
            FIELD_SEPARATOR => {
                if let Some(in_result) = fields.last_mut() {
                    *in_result = true;
                }
            }
            FIELD_END => {
                fields.pop();
            }
            _ if fields.iter().any(|r| !r) => {}
            PARAGRAPH_END | CELL_END => {
                let text = line.split_whitespace().collect::<Vec<_>>().join(" ");
                if !text.is_empty() {
                    paragraphs.push(Paragraph::with_text(text));
                }
                line.clear();
            }
            '\t' | LINE_BREAK => line.push(' '),
            ch if (ch as u32) < 0x20 => {}
            ch => line.push(ch),
        }
    }
    let text = line.split_whitespace().collect::<Vec<_>>().join(" ");
    if !text.is_empty() {
        paragraphs.push(Paragraph::with_text(text));
    }
    paragraphs
}

/// A footnote or endnote: where its reference is, its label, and its text.
struct Note {
    cp: u32,
    label: String,
    text: String,
}

/// The character positions of a `PLC` whose data elements are `data_size` bytes each — its
/// `n + 1` positions, the data being what this reader does not need. An unreadable one is
/// empty: notes are read where they can be, never at the cost of the body text.
fn plc_positions(table: &[u8], plc: FcLcb, data_size: usize) -> Vec<u32> {
    let start = plc.fc as usize;
    let Some(data) = table.get(start..start + plc.lcb as usize) else {
        return Vec::new();
    };
    if data.len() < 4 || (data.len() - 4) % (4 + data_size) != 0 {
        return Vec::new();
    }
    let n = (data.len() - 4) / (4 + data_size);
    (0..=n)
        .map(|i| u32::from_le_bytes(data[i * 4..i * 4 + 4].try_into().unwrap()))
        .collect()
}

/// A note's text on one line: its own reference mark and control characters dropped,
/// paragraph marks read as spaces.
fn note_text(chars: &[DocChar]) -> String {
    let mut out = String::new();
    for c in chars {
        match c.ch {
            PARAGRAPH_END | CELL_END | LINE_BREAK => out.push(' '),
            '\t' => out.push(' '),
            ch if (ch as u32) < 0x20 => {}
            ch => out.push(ch),
        }
    }
    out.split_whitespace().collect::<Vec<_>>().join(" ")
}

/// A field being read: its instruction text, and whether its result has started.
struct OpenField {
    instruction: String,
    in_result: bool,
    hyperlink: bool,
}

/// Turns the character stream into blocks, one paragraph mark at a time.
struct Assembler<'a> {
    papx: &'a PapxIndex,
    chpx: &'a ChpxIndex,
    styles: &'a [StyleInfo],
    fonts: &'a [String],
    lists: &'a ListTable,
    counter: ListCounter,
    notes: &'a [&'a Note],
    /// The `Data` stream, and the pictures read from it so far.
    data: &'a [u8],
    resources: Vec<(String, Resource)>,
    images: Vec<InlineImage>,
    /// The character properties of the run being read.
    props: CharProps,
    blocks: Vec<Block>,
    /// The paragraph being read.
    runs: Vec<TextRun>,
    text: String,
    hyperlink: Option<String>,
    fields: Vec<OpenField>,
    /// The table being read: finished rows, the row being read, the cell being read.
    rows: Vec<Row>,
    cells: Vec<Cell>,
    cell: Vec<Paragraph>,
}

impl<'a> Assembler<'a> {
    fn new(
        papx: &'a PapxIndex,
        chpx: &'a ChpxIndex,
        styles: &'a [StyleInfo],
        fonts: &'a [String],
        lists: &'a ListTable,
        notes: &'a [&'a Note],
        data: &'a [u8],
    ) -> Self {
        Self {
            papx,
            chpx,
            styles,
            fonts,
            lists,
            counter: ListCounter::default(),
            notes,
            data,
            resources: Vec::new(),
            images: Vec::new(),
            props: CharProps::default(),
            blocks: Vec::new(),
            runs: Vec::new(),
            text: String::new(),
            hyperlink: None,
            fields: Vec::new(),
            rows: Vec::new(),
            cells: Vec::new(),
            cell: Vec::new(),
        }
    }

    fn run(&mut self, chars: &[DocChar]) -> Vec<Block> {
        for c in chars {
            match c.ch {
                FIELD_BEGIN => {
                    self.fields.push(OpenField {
                        instruction: String::new(),
                        in_result: false,
                        hyperlink: false,
                    });
                }
                FIELD_SEPARATOR => self.separate_field(),
                FIELD_END => self.end_field(),
                // Instruction text is not document text.
                _ if self.in_instruction() => {
                    if let Some(field) = self.fields.last_mut() {
                        field.instruction.push(c.ch);
                    }
                }
                PARAGRAPH_END | CELL_END => self.end_paragraph(c.ch, c.fc),
                PAGE_BREAK => {
                    self.flush_run();
                    match self.runs.last_mut() {
                        Some(run) => run.page_break = true,
                        None => {
                            self.close_table();
                            self.blocks.push(Block::PageBreak);
                        }
                    }
                }
                LINE_BREAK => {
                    self.flush_run();
                    if let Some(run) = self.runs.last_mut() {
                        run.line_break = true;
                    }
                }
                PICTURE => self.picture(c.fc),
                NOTE_REFERENCE => {
                    if let Some(note) = self.notes.iter().find(|n| n.cp == c.cp) {
                        self.flush_run();
                        self.runs.push(TextRun::plain(format!("[^{}]", note.label)));
                    }
                }
                OPTIONAL_HYPHEN => {}
                // Anchors of pictures, drawn objects, comment references, and other control
                // characters carry no text.
                ch if ch != '\t' && ch != NON_BREAKING_HYPHEN && (ch as u32) < 0x20 => {}
                ch => self.push_char(ch, c.fc),
            }
        }
        // Text after the last paragraph mark still belongs to the document.
        if !self.runs.is_empty() || !self.text.is_empty() {
            self.end_paragraph(PARAGRAPH_END, u32::MAX);
        }
        self.close_table();
        std::mem::take(&mut self.blocks)
    }

    /// An inline picture: its bytes become a resource, and the paragraph references it.
    fn picture(&mut self, fc: u32) {
        let props = self.chpx.at(fc);
        if props.hidden {
            return;
        }
        let Some(picture) = props
            .picture
            .and_then(|offset| super::pictures::read(self.data, offset))
        else {
            return;
        };
        let id = format!("image{}", self.resources.len() + 1);
        let mut resource =
            Resource::image(picture.data, Some(format!("{id}.{}", picture.extension)));
        resource.alt_text = picture.alt_text.clone();
        self.images.push(InlineImage {
            resource_id: id.clone(),
            alt_text: picture.alt_text,
            width: None,
            height: None,
        });
        self.resources.push((id, resource));
    }

    /// Append a character of document text, starting a new run where its formatting changes.
    /// Hidden text is not document content.
    fn push_char(&mut self, ch: char, fc: u32) {
        let props = self.chpx.at(fc);
        if props.hidden {
            return;
        }
        if props != self.props {
            self.flush_run();
            self.props = props;
        }
        self.text.push(match ch {
            NON_BREAKING_HYPHEN => '-',
            _ => props
                .symbol
                .and_then(|(font, code)| {
                    symbol::resolve(self.fonts.get(font as usize).map(String::as_str), code)
                })
                .unwrap_or(ch),
        });
    }

    fn in_instruction(&self) -> bool {
        self.fields.iter().any(|f| !f.in_result)
    }

    fn separate_field(&mut self) {
        let Some(field) = self.fields.last_mut() else {
            return;
        };
        field.in_result = true;
        if let Some(url) = hyperlink_target(&field.instruction) {
            field.hyperlink = true;
            self.flush_run();
            self.hyperlink = Some(url);
        }
    }

    fn end_field(&mut self) {
        if let Some(field) = self.fields.pop() {
            if field.hyperlink {
                self.flush_run();
                self.hyperlink = None;
            }
        }
    }

    fn flush_run(&mut self) {
        if self.text.is_empty() {
            return;
        }
        let text = std::mem::take(&mut self.text);
        let p = self.props;
        let mut run = TextRun::styled(
            text,
            TextStyle {
                bold: p.bold,
                italic: p.italic,
                underline: p.underline,
                strikethrough: p.strike,
                superscript: p.superscript,
                subscript: p.subscript,
                ..TextStyle::default()
            },
        );
        run.hyperlink = self.hyperlink.clone();
        run.revision = if p.deleted {
            RevisionType::Deleted
        } else if p.inserted {
            RevisionType::Inserted
        } else {
            RevisionType::None
        };
        self.runs.push(run);
    }

    fn take_paragraph(&mut self, props: Option<&ParaProps>) -> Paragraph {
        self.flush_run();
        let mut paragraph = Paragraph::new();
        paragraph.runs = std::mem::take(&mut self.runs);
        paragraph.images = std::mem::take(&mut self.images);
        if let Some(props) = props {
            if let Some(level) = (props.ilfo > 0)
                .then(|| self.lists.level(props.ilfo, props.ilvl))
                .flatten()
            {
                let number = self.counter.next(&level, props.ilvl);
                paragraph.list_info = Some(ListInfo {
                    list_type: level.kind,
                    level: props.ilvl,
                    number: (level.kind == ListType::Numbered).then_some(number),
                    marker_image: None,
                });
            }
            paragraph.heading = self.heading_of(props);
            if let Some(style) = self.styles.get(props.istd as usize) {
                if !style.name.is_empty() {
                    paragraph.style_name = Some(style.name.clone());
                }
            }
        }
        paragraph
    }

    fn heading_of(&self, props: &ParaProps) -> HeadingLevel {
        if props.in_table {
            return HeadingLevel::None;
        }
        let level = match props.outline_level {
            Some(level) if level < 9 => Some(level + 1),
            _ => self
                .styles
                .get(props.istd as usize)
                .and_then(StyleInfo::heading_level),
        };
        level.map_or(HeadingLevel::None, HeadingLevel::from_number)
    }

    fn end_paragraph(&mut self, mark: char, fc: u32) {
        let props = self.papx.at(fc).cloned();
        let in_table = props.as_ref().is_some_and(|p| p.in_table);
        let row_end = props.as_ref().is_some_and(|p| p.row_end);
        let inner_row_end = props.as_ref().is_some_and(|p| p.inner_row_end);
        let paragraph = self.take_paragraph(props.as_ref());

        if !in_table {
            self.close_table();
            if !paragraph.is_empty() {
                self.blocks.push(Block::Paragraph(paragraph));
            }
            return;
        }
        if props.as_ref().is_some_and(|p| p.depth > 1) {
            // A nested table is read as paragraphs of the cell that holds it; its row-end
            // marks carry no text.
            if !inner_row_end && !paragraph.is_empty() {
                self.cell.push(paragraph);
            }
            return;
        }
        if row_end {
            // The row-end mark holds no text; it closes the row. Its cells were closed by
            // their own marks -- anything left over had none.
            if !self.cell.is_empty() {
                self.close_cell();
            }
            let cells = std::mem::take(&mut self.cells);
            if !cells.is_empty() {
                let mut row = Row::new();
                row.cells = cells;
                self.rows.push(row);
            }
            return;
        }
        if !paragraph.is_empty() {
            self.cell.push(paragraph);
        }
        if mark == CELL_END {
            self.close_cell();
        }
    }

    fn close_cell(&mut self) {
        let mut cell = Cell::new();
        cell.content = std::mem::take(&mut self.cell);
        self.cells.push(cell);
    }

    fn close_table(&mut self) {
        if !self.cell.is_empty() || !self.cells.is_empty() {
            self.close_cell();
            let mut row = Row::new();
            row.cells = std::mem::take(&mut self.cells);
            self.rows.push(row);
        }
        if self.rows.is_empty() {
            return;
        }
        let mut table = Table::new();
        for row in std::mem::take(&mut self.rows) {
            table.add_row(row);
        }
        self.blocks.push(Block::Table(table));
    }
}

/// The target of a `HYPERLINK` field instruction: `HYPERLINK "url"`, or `HYPERLINK \l "name"`
/// for a bookmark in the same document.
fn hyperlink_target(instruction: &str) -> Option<String> {
    let rest = instruction.trim_start().strip_prefix("HYPERLINK")?;
    let mut url = None;
    let mut bookmark = false;
    let mut chars = rest.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            '\\' => {
                if chars.next() == Some('l') {
                    bookmark = true;
                }
            }
            '"' => {
                let quoted: String = chars.by_ref().take_while(|&c| c != '"').collect();
                if url.is_none() {
                    url = Some(quoted);
                }
            }
            c if !c.is_whitespace() && url.is_none() => {
                let mut bare = String::from(c);
                while let Some(&next) = chars.peek() {
                    if next.is_whitespace() {
                        break;
                    }
                    bare.push(next);
                    chars.next();
                }
                url = Some(bare);
            }
            _ => {}
        }
    }
    let url = url.filter(|u| !u.is_empty())?;
    Some(if bookmark { format!("#{url}") } else { url })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn hyperlink_targets_are_read_from_the_instruction() {
        assert_eq!(
            hyperlink_target(" HYPERLINK \"https://example.com/a b\" ").as_deref(),
            Some("https://example.com/a b")
        );
        assert_eq!(
            hyperlink_target("HYPERLINK \\l \"_Toc1\"").as_deref(),
            Some("#_Toc1")
        );
        assert_eq!(
            hyperlink_target("HYPERLINK https://example.com").as_deref(),
            Some("https://example.com")
        );
        assert_eq!(hyperlink_target(" PAGEREF _Toc1 \\h "), None);
    }
}
