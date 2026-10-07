//! Reads an Excel 97-2003 workbook's worksheets into the document model.

use std::collections::{BTreeMap, HashMap};
use std::io::{Cursor as IoCursor, Read};

use super::records::{self as rec, rk_value, Cursor, Record, Records};
use crate::detect::FormatType;
use crate::error::{Error, Result};
use crate::model::{Block, Cell, Document, Paragraph, Section, TextRun, TextStyle};
use crate::sheet::{self, MergeRange, SheetTableBuilder};

/// A worksheet listed in the workbook globals.
struct SheetEntry {
    name: String,
    /// Stream offset of the sheet's `BOF`.
    offset: usize,
}

/// What the workbook globals say about every sheet's cells.
#[derive(Default)]
struct Globals {
    sheets: Vec<SheetEntry>,
    strings: Vec<String>,
    /// Number format of each cell format (XF), by XF index.
    xf_formats: Vec<u16>,
    /// Number format strings declared by the workbook, by format index.
    formats: HashMap<u16, String>,
    /// The 1904 date system.
    date1904: bool,
}

impl Globals {
    fn is_date(&self, xf: u16) -> bool {
        let Some(&ifmt) = self.xf_formats.get(xf as usize) else {
            return false;
        };
        sheet::is_builtin_date_format(ifmt as u32)
            || self
                .formats
                .get(&ifmt)
                .is_some_and(|code| sheet::is_date_format_code(code))
    }

    /// A number as the cell shows it: a date for a date-formatted cell, otherwise the
    /// shortest decimal that reads back as the same value.
    fn number(&self, value: f64, xf: u16) -> String {
        if self.is_date(xf) {
            let serial = if self.date1904 { value + 1462.0 } else { value };
            if let Some(date) = sheet::serial_to_date(serial) {
                return date;
            }
        }
        format!("{value}")
    }
}

/// Reader for Excel 97-2003 workbooks (.xls, BIFF8).
pub struct XlsParser {
    data: Vec<u8>,
}

impl XlsParser {
    /// Open a .xls file.
    #[cfg(not(target_arch = "wasm32"))]
    pub fn open(path: impl AsRef<std::path::Path>) -> Result<Self> {
        Ok(Self {
            data: std::fs::read(path)?,
        })
    }

    /// Read a .xls workbook held in memory.
    pub fn from_bytes(data: Vec<u8>) -> Result<Self> {
        Ok(Self { data })
    }

    /// Parse the workbook: one section per worksheet, its cells as a table.
    pub fn parse(&mut self) -> Result<Document> {
        let mut container =
            cfb::CompoundFile::open(IoCursor::new(&self.data[..])).map_err(|e| {
                Error::InvalidData(format!("Excel workbook: unreadable container: {e}"))
            })?;
        if !container.exists("/Workbook") {
            return Err(if container.exists("/Book") {
                Error::UnsupportedFormat(
                    "Excel 5.0/95 workbook (.xls, BIFF5) — only Excel 97 and later are supported"
                        .into(),
                )
            } else {
                Error::InvalidData("Excel workbook: the Workbook stream is missing".into())
            });
        }
        let mut stream = Vec::new();
        container
            .open_stream("/Workbook")
            .map_err(|e| Error::InvalidData(format!("Excel workbook: {e}")))?
            .read_to_end(&mut stream)?;

        let globals = read_globals(&stream)?;
        let mut doc = Document::new();
        doc.format = FormatType::Xls;
        doc.metadata = crate::summary::read(&mut container);
        for (index, entry) in globals.sheets.iter().enumerate() {
            let mut section = Section::new(index);
            section.name = Some(entry.name.clone());
            let table = read_sheet(&stream, entry.offset, &globals)?;
            if !table.rows.is_empty() {
                section.add_block(Block::Table(table));
            }
            doc.add_section(section);
        }
        Ok(doc)
    }
}

fn read_globals(stream: &[u8]) -> Result<Globals> {
    let mut globals = Globals::default();
    let mut records = Records::new(stream, 0);
    match records.next().transpose()? {
        Some(bof) if bof.kind == rec::BOF => check_biff8(&bof)?,
        _ => {
            return Err(Error::InvalidData(
                "Excel workbook: the Workbook stream does not begin with a BOF record".into(),
            ))
        }
    }
    for record in records {
        let record = record?;
        let mut c = Cursor::new(&record.segments);
        match record.kind {
            rec::EOF => break,
            rec::FILEPASS => return Err(Error::Encrypted),
            rec::DATEMODE => globals.date1904 = c.u16()? == 1,
            rec::BOUNDSHEET8 => {
                let offset = c.u32()? as usize;
                let _visibility = c.u8()?;
                let kind = c.u8()?;
                let name = c.short_unicode_string()?;
                // 0 is a worksheet; charts, macro sheets and VBA modules hold no cell grid.
                if kind == 0 {
                    globals.sheets.push(SheetEntry { name, offset });
                }
            }
            rec::SST => {
                let _total = c.u32()?;
                let unique = c.u32()? as usize;
                globals.strings.reserve(unique.min(1 << 20));
                for _ in 0..unique {
                    globals.strings.push(c.rich_extended_string()?);
                }
            }
            rec::FORMAT => {
                let index = c.u16()?;
                globals.formats.insert(index, c.unicode_string()?);
            }
            rec::XF => {
                let _font = c.u16()?;
                globals.xf_formats.push(c.u16()?);
            }
            _ => {}
        }
    }
    Ok(globals)
}

fn check_biff8(bof: &Record<'_>) -> Result<()> {
    let version = Cursor::new(&bof.segments).u16()?;
    if version == rec::BIFF8 {
        Ok(())
    } else {
        Err(Error::UnsupportedFormat(format!(
            "Excel workbook in BIFF version 0x{version:04X} (.xls) — only Excel 97 and later \
             (BIFF8) are supported"
        )))
    }
}

/// The text of a BIFF error code ([MS-XLS] 2.5.30 `BErr`), as the xlsx reader renders errors.
fn error_text(code: u8) -> String {
    let name = match code {
        0x00 => "#NULL!",
        0x07 => "#DIV/0!",
        0x0F => "#VALUE!",
        0x17 => "#REF!",
        0x1D => "#NAME?",
        0x24 => "#NUM!",
        0x2A => "#N/A",
        0x2B => "#GETTING_DATA",
        _ => "#ERROR",
    };
    format!("#ERROR:{name}")
}

fn read_sheet(stream: &[u8], offset: usize, globals: &Globals) -> Result<crate::model::Table> {
    let mut cells: BTreeMap<(u16, u16), String> = BTreeMap::new();
    let mut merges = Vec::new();
    let mut records = Records::new(stream, offset);
    match records.next().transpose()? {
        Some(bof) if bof.kind == rec::BOF => {}
        _ => {
            return Err(Error::InvalidData(format!(
                "Excel workbook: no worksheet begins at offset {offset}"
            )))
        }
    }
    // A FORMULA whose cached result is a string leaves it to the STRING record that follows.
    let mut pending_string: Option<(u16, u16)> = None;
    let mut links: HashMap<(u16, u16), String> = HashMap::new();
    // A comment's text sits in the TXO after the OBJ that names it; its NOTE, which says
    // which cell it belongs to, comes at the end of the sheet.
    let mut comment_obj: Option<u16> = None;
    let mut comment_texts: HashMap<u16, String> = HashMap::new();
    let mut comments: HashMap<(u16, u16), String> = HashMap::new();
    for record in records {
        let record = record?;
        let mut c = Cursor::new(&record.segments);
        match record.kind {
            rec::EOF => break,
            rec::LABELSST => {
                let (row, col, _xf) = (c.u16()?, c.u16()?, c.u16()?);
                let index = c.u32()? as usize;
                let text = globals.strings.get(index).cloned().ok_or_else(|| {
                    Error::InvalidData(format!(
                        "Excel workbook: shared string index out of range: {index}"
                    ))
                })?;
                cells.insert((row, col), text);
            }
            rec::LABEL => {
                let (row, col, _xf) = (c.u16()?, c.u16()?, c.u16()?);
                cells.insert((row, col), c.unicode_string()?);
            }
            rec::NUMBER => {
                let (row, col, xf) = (c.u16()?, c.u16()?, c.u16()?);
                let value = f64::from_bits(c.u32()? as u64 | (c.u32()? as u64) << 32);
                cells.insert((row, col), globals.number(value, xf));
            }
            rec::RK => {
                let (row, col, xf) = (c.u16()?, c.u16()?, c.u16()?);
                cells.insert((row, col), globals.number(rk_value(c.u32()?), xf));
            }
            rec::MULRK => {
                let data = record.data();
                let (row, first) = (c.u16()?, c.u16()?);
                // Six bytes per cell between the 4-byte head and the 2-byte last column.
                let count = data.len().saturating_sub(6) / 6;
                for i in 0..count {
                    let xf = c.u16()?;
                    let value = rk_value(c.u32()?);
                    cells.insert((row, first + i as u16), globals.number(value, xf));
                }
            }
            rec::FORMULA => {
                let (row, col, xf) = (c.u16()?, c.u16()?, c.u16()?);
                let mut value = [0u8; 8];
                for b in &mut value {
                    *b = c.u8()?;
                }
                if value[6] == 0xFF && value[7] == 0xFF {
                    match value[0] {
                        0 => pending_string = Some((row, col)),
                        1 => {
                            cells.insert((row, col), bool_text(value[2] != 0));
                        }
                        2 => {
                            cells.insert((row, col), error_text(value[2]));
                        }
                        _ => {}
                    }
                } else {
                    cells.insert((row, col), globals.number(f64::from_le_bytes(value), xf));
                }
            }
            rec::STRING => {
                if let Some(at) = pending_string.take() {
                    cells.insert(at, c.unicode_string()?);
                }
            }
            rec::BOOLERR => {
                let (row, col, _xf) = (c.u16()?, c.u16()?, c.u16()?);
                let (value, is_error) = (c.u8()?, c.u8()?);
                cells.insert(
                    (row, col),
                    if is_error != 0 {
                        error_text(value)
                    } else {
                        bool_text(value != 0)
                    },
                );
            }
            rec::HLINK => {
                let (top, bottom, left, right) = (c.u16()?, c.u16()?, c.u16()?, c.u16()?);
                // A link that cannot be read is left out, not the sheet.
                if let Ok(Some(url)) = hyperlink_target(&mut c) {
                    for row in top..=bottom.min(top.saturating_add(4096)) {
                        for col in left..=right.min(left.saturating_add(256)) {
                            links.insert((row, col), url.clone());
                        }
                    }
                }
            }
            rec::OBJ => {
                // The first sub-record is the common object data: ft, cb, ot, id.
                let (ft, _cb, ot, id) = (c.u16()?, c.u16()?, c.u16()?, c.u16()?);
                comment_obj = (ft == 0x15 && ot == 0x19).then_some(id);
            }
            rec::TXO => {
                if let Some(id) = comment_obj.take() {
                    c.skip(10)?;
                    let cch = c.u16()? as usize;
                    if cch > 0 && record.segments.len() > 1 {
                        let mut text = Cursor::new(&record.segments[1..]);
                        if let Ok(t) = text.flagged_chars(cch) {
                            comment_texts.insert(id, t);
                        }
                    }
                }
            }
            rec::NOTE => {
                let (row, col, _flags, id) = (c.u16()?, c.u16()?, c.u16()?, c.u16()?);
                if let Some(text) = comment_texts.get(&id) {
                    comments.insert((row, col), text.clone());
                }
            }
            rec::MERGEDCELLS => {
                let count = c.u16()?;
                for _ in 0..count {
                    let (top, bottom, left, right) = (c.u16()?, c.u16()?, c.u16()?, c.u16()?);
                    merges.push(MergeRange {
                        left: left.min(right) as u32,
                        right: left.max(right) as u32,
                        top: top.min(bottom) as u32 + 1,
                        bottom: top.max(bottom) as u32 + 1,
                    });
                }
            }
            _ => {}
        }
    }

    // A comment on a cell with no value still belongs to that cell.
    for &at in comments.keys() {
        cells.entry(at).or_default();
    }

    let mut builder = SheetTableBuilder::new(merges);
    let mut current_row = None;
    for ((row, col), text) in cells {
        if current_row != Some(row) {
            builder.start_row(Some(row as u32 + 1));
            current_row = Some(row);
        }
        // Same shape as the .xlsx reader: the value's run carries the link, and a comment
        // follows as an italic run.
        let mut runs = vec![TextRun::plain(normalize_breaks(&text))];
        runs[0].hyperlink = links.get(&(row, col)).cloned();
        if let Some(comment) = comments.get(&(row, col)) {
            runs.push(TextRun::styled(
                format!(" [Comment: {}]", normalize_breaks(comment)),
                TextStyle::italic(),
            ));
        }
        let cell = Cell {
            content: vec![Paragraph {
                runs,
                ..Default::default()
            }],
            is_header: builder.in_first_row(),
            ..Cell::new()
        };
        builder.place(cell, Some((col as u32, row as u32 + 1)));
    }
    Ok(builder.finish())
}

/// `URLMoniker` and `FileMoniker` class identifiers ([MS-OSHARED] 2.3.7.6, 2.3.7.8).
const URL_MONIKER: [u8; 16] = [
    0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B,
];
const FILE_MONIKER: [u8; 16] = [
    0x03, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46,
];

/// The external target of a `Hyperlink` object ([MS-OSHARED] 2.3.7.1), past its 16-byte class
/// id: a URL or a file path. A link to a place in the workbook only — no moniker — has none,
/// as in the .xlsx reader.
fn hyperlink_target(c: &mut Cursor<'_>) -> Result<Option<String>> {
    c.skip(16)?;
    let _version = c.u32()?;
    let flags = c.u32()?;
    let (has_moniker, saved_as_string) = (flags & 0x01 != 0, flags & 0x100 != 0);
    if flags & 0x10 != 0 {
        c.hyperlink_string()?; // display name
    }
    if flags & 0x80 != 0 {
        c.hyperlink_string()?; // target frame
    }
    if !has_moniker {
        return Ok(None);
    }
    if saved_as_string {
        return Ok(Some(c.hyperlink_string()?).filter(|s| !s.is_empty()));
    }
    let mut class = [0u8; 16];
    for b in &mut class {
        *b = c.u8()?;
    }
    if class == URL_MONIKER {
        let length = c.u32()? as usize;
        return Ok(Some(c.utf16_within(length)?).filter(|s| !s.is_empty()));
    }
    if class == FILE_MONIKER {
        let _anti = c.u16()?;
        let ansi_len = c.u32()? as usize;
        let mut ansi = Vec::with_capacity(ansi_len);
        for _ in 0..ansi_len {
            ansi.push(c.u8()?);
        }
        let ansi: String = ansi
            .iter()
            .take_while(|&&b| b != 0)
            // The ANSI path is in the writer's code page, which is not recorded; read as
            // Latin-1. The Unicode path, when present, wins.
            .map(|&b| b as char)
            .collect();
        // endServer, versionNumber, 20 reserved bytes, then an optional Unicode path.
        c.skip(2 + 2 + 20)?;
        let unicode_size = c.u32()? as usize;
        if unicode_size >= 6 {
            let bytes = c.u32()? as usize;
            let _key = c.u16()?;
            let path = c.utf16_within(bytes)?;
            if !path.is_empty() {
                return Ok(Some(path));
            }
        }
        return Ok(Some(ansi).filter(|s| !s.is_empty()));
    }
    Ok(None)
}

fn bool_text(value: bool) -> String {
    if value { "TRUE" } else { "FALSE" }.to_string()
}

/// In-cell line breaks are stored as CR LF or LF; the model uses LF.
fn normalize_breaks(text: &str) -> String {
    text.replace("\r\n", "\n").replace('\r', "\n")
}
