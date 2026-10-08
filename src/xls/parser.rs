//! Reads an Excel workbook's worksheets — Excel 2.x through 2003, BIFF2 to BIFF8 — into the
//! document model.

use std::collections::{BTreeMap, HashMap};
use std::io::{Cursor as IoCursor, Read};

use super::records::{self as rec, rk_value, Cursor, Record, Records};
use crate::codepage;
use crate::detect::FormatType;
use crate::error::{Error, Result};
use crate::model::{Block, Cell, Document, Metadata, Paragraph, Section, TextRun, TextStyle};
use crate::sheet::{self, MergeRange, SheetTableBuilder};

/// The BIFF generation a stream is written in. Each one changed the record layouts the reader
/// relies on: BIFF2 heads its cells with attribute bytes rather than an XF index, BIFF3–BIFF4
/// keep one worksheet's formats in its own stream, BIFF5 (and Excel 95's BIFF7, which declares
/// itself BIFF5) stores text in the workbook's code page, BIFF8 in UTF-16.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
enum Biff {
    V2,
    V3,
    V4,
    V5,
    V8,
}

/// Substream types a `BOF` declares.
const WORKBOOK_GLOBALS: u16 = 0x0005;
const WORKSHEET: u16 = 0x0010;
const CHART: u16 = 0x0020;
const MACRO_SHEET: u16 = 0x0040;
/// A BIFF4 workbook: worksheets follow in the same stream, each after a `SHEETHDR`.
const WORKBOOK_BIFF4: u16 = 0x0100;

/// The version and substream type a `BOF` record declares, or `None` for a record that is not
/// a `BOF`.
fn bof_info(record: &Record<'_>) -> Result<Option<(Biff, u16)>> {
    let biff = match record.kind {
        rec::BOF_BIFF2 => Biff::V2,
        rec::BOF_BIFF3 => Biff::V3,
        rec::BOF_BIFF4 => Biff::V4,
        rec::BOF => Biff::V8, // refined below
        _ => return Ok(None),
    };
    let mut c = Cursor::new(&record.segments);
    let version = c.u16()?;
    let kind = c.u16()?;
    let biff = match (biff, version) {
        (Biff::V8, rec::BIFF8) => Biff::V8,
        (Biff::V8, rec::BIFF5) => Biff::V5,
        (Biff::V8, _) => {
            return Err(Error::UnsupportedFormat(format!(
                "Excel workbook in BIFF version 0x{version:04X} (.xls)"
            )))
        }
        (biff, _) => biff,
    };
    Ok(Some((biff, kind)))
}

/// Text as a record stores it: decoded already (BIFF8), or bytes in a code page the workbook
/// may only declare later in the stream — or never, leaving its fonts to say.
enum Text {
    Decoded(String),
    Bytes(Vec<u8>),
}

impl Text {
    fn resolve(self, codepage: u16) -> String {
        match self {
            Text::Decoded(text) => text,
            Text::Bytes(bytes) => codepage::decode_lossy(&bytes, codepage),
        }
    }
}

/// What a substream says about every cell it governs: the workbook globals for BIFF5 and
/// later, or for BIFF2–BIFF4 the worksheet's own stream, which carries its formats.
#[derive(Clone)]
struct Globals {
    biff: Biff,
    sheets: Vec<(String, usize)>,
    strings: Vec<String>,
    /// Number format of each cell format (XF), by XF index.
    xf_formats: Vec<u16>,
    /// Number format strings declared by the stream, by format index.
    formats: HashMap<u16, String>,
    /// The 1904 date system.
    date1904: bool,
    /// The code page of byte strings (BIFF2–BIFF7).
    codepage: u16,
}

/// Where a cell's number format comes from: its XF record (BIFF3 and later), or for BIFF2 the
/// format index its attribute bytes carry.
#[derive(Clone, Copy)]
enum Fmt {
    Xf(u16),
    Format(u16),
}

impl Globals {
    fn is_date(&self, fmt: Fmt) -> bool {
        let ifmt = match fmt {
            Fmt::Format(ifmt) => ifmt,
            Fmt::Xf(xf) => match self.xf_formats.get(xf as usize) {
                Some(&ifmt) => ifmt,
                None => return false,
            },
        };
        // Built-in format numbers exist from BIFF5 on; before that every format a workbook
        // uses is written out in it, and the index is only its position.
        (self.biff >= Biff::V5 && sheet::is_builtin_date_format(ifmt as u32))
            || self
                .formats
                .get(&ifmt)
                .is_some_and(|code| sheet::is_date_format_code(code))
    }

    /// A number as the cell shows it: a date for a date-formatted cell, otherwise the
    /// shortest decimal that reads back as the same value.
    fn number(&self, value: f64, fmt: Fmt) -> String {
        if self.is_date(fmt) {
            let serial = if self.date1904 { value + 1462.0 } else { value };
            if let Some(date) = sheet::serial_to_date(serial) {
                return date;
            }
        }
        format!("{value}")
    }

    fn text(&self, bytes: &[u8]) -> String {
        codepage::decode_lossy(bytes, self.codepage)
    }
}

/// Reader for Excel workbooks in the binary .xls format: Excel 97-2003 (BIFF8), Excel 5.0/95
/// (BIFF5/BIFF7) in a compound file, and the bare record streams of Excel 2.x–4.0
/// (BIFF2–BIFF4).
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
        let mut doc = Document::new();
        doc.format = FormatType::Xls;
        let owned;
        let stream: &[u8] = if self.data.starts_with(&crate::detect::CFB_MAGIC) {
            let (workbook, metadata) = open_container(&self.data)?;
            doc.metadata = metadata;
            owned = workbook;
            &owned
        } else {
            // Excel 2.x–4.0 wrote the record stream as the file itself.
            &self.data
        };

        let first = Records::new(stream, 0).next().transpose()?;
        let Some((biff, kind)) = first.as_ref().map(bof_info).transpose()?.flatten() else {
            return Err(Error::InvalidData(
                "Excel workbook: the stream does not begin with a BOF record".into(),
            ));
        };

        let sheets: Vec<(Option<String>, Globals, usize)> = match kind {
            WORKBOOK_GLOBALS | WORKBOOK_BIFF4 => {
                let globals = scan(stream, 0, biff, None)?;
                let mut sheets = Vec::new();
                for (name, offset) in &globals.sheets {
                    if biff >= Biff::V5 {
                        sheets.push((Some(name.clone()), globals.clone(), *offset));
                        continue;
                    }
                    // A BIFF4 workbook's sheets each carry their own formats; charts and
                    // macro sheets among them are told apart by their own BOF.
                    let sheet_kind = Records::new(stream, *offset)
                        .next()
                        .transpose()?
                        .as_ref()
                        .map(bof_info)
                        .transpose()?
                        .flatten()
                        .map(|(_, kind)| kind);
                    if matches!(sheet_kind, Some(WORKSHEET | MACRO_SHEET)) {
                        let own = scan(stream, *offset, biff, Some(globals.codepage))?;
                        sheets.push((Some(name.clone()), own, *offset));
                    }
                }
                sheets
            }
            // A workbook of one worksheet (Excel 2.x–4.0): no globals, no sheet name.
            WORKSHEET | MACRO_SHEET => vec![(None, scan(stream, 0, biff, None)?, 0)],
            CHART => Vec::new(),
            other => {
                return Err(Error::UnsupportedFormat(format!(
                    "Excel file whose stream is a substream of type 0x{other:04X}, not a \
                     workbook or worksheet"
                )))
            }
        };

        for (index, (name, globals, offset)) in sheets.into_iter().enumerate() {
            let mut section = Section::new(index);
            section.name = name;
            let table = read_sheet(stream, offset, &globals)?;
            if !table.rows.is_empty() {
                section.add_block(Block::Table(table));
            }
            doc.add_section(section);
        }
        Ok(doc)
    }
}

/// The workbook stream of a compound file — `Workbook` from Excel 97 on, `Book` before — and
/// the document properties beside it.
fn open_container(data: &[u8]) -> Result<(Vec<u8>, Metadata)> {
    let mut container = cfb::CompoundFile::open(IoCursor::new(data))
        .map_err(|e| Error::InvalidData(format!("Excel workbook: unreadable container: {e}")))?;
    let name = ["/Workbook", "/Book"]
        .into_iter()
        .find(|name| container.exists(name))
        .ok_or_else(|| {
            Error::InvalidData("Excel workbook: the Workbook stream is missing".into())
        })?;
    let mut stream = Vec::new();
    container
        .open_stream(name)
        .map_err(|e| Error::InvalidData(format!("Excel workbook: {e}")))?
        .read_to_end(&mut stream)?;
    Ok((stream, crate::summary::read(&mut container)))
}

/// Read one substream — from its `BOF` at `offset` to its `EOF` — for what it says about the
/// cells it governs. Substreams nested in it (a BIFF4 workbook's worksheets, embedded charts)
/// are passed over; a BIFF4 workbook's worksheets are listed by the `SHEETHDR` before each.
fn scan(stream: &[u8], offset: usize, biff: Biff, inherited: Option<u16>) -> Result<Globals> {
    let mut codepage = None;
    let mut charsets: Vec<u8> = Vec::new();
    let mut sheets: Vec<(Text, usize)> = Vec::new();
    let mut formats: Vec<(u16, Text)> = Vec::new();
    let mut globals = Globals {
        biff,
        sheets: Vec::new(),
        strings: Vec::new(),
        xf_formats: Vec::new(),
        formats: HashMap::new(),
        date1904: false,
        codepage: 0,
    };
    let mut records = Records::new(stream, offset);
    let _bof = records.next(); // this substream's BOF
    let mut depth = 0usize;
    while let Some(record) = records.next() {
        let record = record?;
        if bof_info(&record)?.is_some() {
            depth += 1;
            continue;
        }
        if record.kind == rec::EOF {
            if depth == 0 {
                break;
            }
            depth -= 1;
            continue;
        }
        if depth > 0 {
            continue;
        }
        let mut c = Cursor::new(&record.segments);
        match record.kind {
            rec::FILEPASS => return Err(Error::Encrypted),
            rec::DATEMODE => globals.date1904 = c.u16()? == 1,
            rec::CODEPAGE => codepage = Some(c.u16()?),
            rec::FONT if biff >= Biff::V5 => {
                c.skip(12)?;
                charsets.push(c.u8()?);
            }
            rec::BOUNDSHEET8 if biff >= Biff::V5 => {
                let offset = c.u32()? as usize;
                let _visibility = c.u8()?;
                let kind = c.u8()?;
                let name = if biff == Biff::V8 {
                    Text::Decoded(c.short_unicode_string()?)
                } else {
                    Text::Bytes(c.byte_string8()?)
                };
                // 0 is a worksheet; charts, macro sheets and VBA modules hold no cell grid.
                if kind == 0 {
                    sheets.push((name, offset));
                }
            }
            rec::SHEETHDR if biff == Biff::V4 => {
                let _length = c.u32()?;
                let name = Text::Bytes(c.byte_string8()?);
                sheets.push((name, records.position()));
            }
            rec::SST if biff == Biff::V8 => {
                let _total = c.u32()?;
                let unique = c.u32()? as usize;
                globals.strings.reserve(unique.min(1 << 20));
                for _ in 0..unique {
                    globals.strings.push(c.rich_extended_string()?);
                }
            }
            rec::FORMAT => match biff {
                Biff::V8 => {
                    let index = c.u16()?;
                    formats.push((index, Text::Decoded(c.unicode_string()?)));
                }
                Biff::V5 => {
                    let index = c.u16()?;
                    formats.push((index, Text::Bytes(c.byte_string8()?)));
                }
                // BIFF4's leading field is unused: a format is known by its position.
                _ => {
                    c.skip(2)?;
                    formats.push((formats.len() as u16, Text::Bytes(c.byte_string8()?)));
                }
            },
            rec::FORMAT_BIFF2 if biff <= Biff::V3 => {
                formats.push((formats.len() as u16, Text::Bytes(c.byte_string8()?)));
            }
            rec::XF if biff >= Biff::V5 => {
                let _font = c.u16()?;
                globals.xf_formats.push(c.u16()?);
            }
            rec::XF_BIFF3 | rec::XF_BIFF4 if matches!(biff, Biff::V3 | Biff::V4) => {
                let _font = c.u8()?;
                globals.xf_formats.push(c.u8()? as u16);
            }
            _ => {}
        }
    }

    // A workbook that names no code page writes in its fonts' character set; failing that,
    // in Windows-1252, the code page of the Excel versions that left it out.
    globals.codepage = codepage
        .or(inherited)
        .or_else(|| most_common(charsets.iter().filter_map(|&c| codepage::from_charset(c))))
        .unwrap_or(1252);
    globals.sheets = sheets
        .into_iter()
        .map(|(name, offset)| (name.resolve(globals.codepage), offset))
        .collect();
    globals.formats = formats
        .into_iter()
        .map(|(index, code)| (index, code.resolve(globals.codepage)))
        .collect();
    Ok(globals)
}

fn most_common(values: impl Iterator<Item = u16>) -> Option<u16> {
    let mut counts: Vec<(u16, usize)> = Vec::new();
    for value in values {
        match counts.iter_mut().find(|(v, _)| *v == value) {
            Some((_, n)) => *n += 1,
            None => counts.push((value, 1)),
        }
    }
    // The first seen wins a tie: the workbook's default font comes first.
    counts
        .iter()
        .fold(None, |best: Option<(u16, usize)>, &(v, n)| match best {
            Some((_, m)) if m >= n => best,
            _ => Some((v, n)),
        })
        .map(|(v, _)| v)
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

/// A cell record's head: row, column and where its number format comes from — an XF index,
/// or in BIFF2 three attribute bytes whose second carries the format index in its low 6 bits.
fn cell_head(c: &mut Cursor<'_>, biff: Biff) -> Result<(u16, u16, Fmt)> {
    let (row, col) = (c.u16()?, c.u16()?);
    let fmt = if biff == Biff::V2 {
        let attributes = [c.u8()?, c.u8()?, c.u8()?];
        Fmt::Format((attributes[1] & 0x3F) as u16)
    } else {
        Fmt::Xf(c.u16()?)
    };
    Ok((row, col, fmt))
}

fn f64_le(c: &mut Cursor<'_>) -> Result<f64> {
    Ok(f64::from_bits(c.u32()? as u64 | (c.u32()? as u64) << 32))
}

fn read_sheet(stream: &[u8], offset: usize, globals: &Globals) -> Result<crate::model::Table> {
    let biff = globals.biff;
    let mut cells: BTreeMap<(u16, u16), String> = BTreeMap::new();
    let mut merges = Vec::new();
    let mut records = Records::new(stream, offset);
    match records.next().transpose()? {
        Some(ref bof) if bof_info(bof)?.is_some() => {}
        _ => {
            return Err(Error::InvalidData(format!(
                "Excel workbook: no worksheet begins at offset {offset}"
            )))
        }
    }
    // A FORMULA whose cached result is a string leaves it to the STRING record that follows.
    let mut pending_string: Option<(u16, u16)> = None;
    let mut links: HashMap<(u16, u16), String> = HashMap::new();
    // BIFF8: a comment's text sits in the TXO after the OBJ that names it; its NOTE, which
    // says which cell it belongs to, comes at the end of the sheet.
    let mut comment_obj: Option<u16> = None;
    let mut comment_texts: HashMap<u16, String> = HashMap::new();
    let mut comments: HashMap<(u16, u16), String> = HashMap::new();
    // BIFF2–BIFF7: a comment's text is in its NOTE, carried on by NOTEs for row 0xFFFF.
    let mut notes: Vec<((u16, u16), Vec<u8>)> = Vec::new();
    // Substreams nested in the sheet — an embedded chart — end with an EOF of their own.
    let mut depth = 0usize;
    for record in records {
        let record = record?;
        if bof_info(&record)?.is_some() {
            depth += 1;
            continue;
        }
        if record.kind == rec::EOF {
            if depth == 0 {
                break;
            }
            depth -= 1;
            continue;
        }
        if depth > 0 {
            continue;
        }
        let mut c = Cursor::new(&record.segments);
        match record.kind {
            rec::LABELSST if biff == Biff::V8 => {
                let (row, col, _) = cell_head(&mut c, biff)?;
                let index = c.u32()? as usize;
                let text = globals.strings.get(index).cloned().ok_or_else(|| {
                    Error::InvalidData(format!(
                        "Excel workbook: shared string index out of range: {index}"
                    ))
                })?;
                cells.insert((row, col), text);
            }
            rec::LABEL | rec::RSTRING if biff >= Biff::V3 => {
                let (row, col, _) = cell_head(&mut c, biff)?;
                let text = if biff == Biff::V8 {
                    c.unicode_string()?
                } else {
                    globals.text(&c.byte_string16()?)
                };
                cells.insert((row, col), text);
            }
            rec::LABEL_BIFF2 if biff == Biff::V2 => {
                let (row, col, _) = cell_head(&mut c, biff)?;
                cells.insert((row, col), globals.text(&c.byte_string8()?));
            }
            rec::NUMBER if biff >= Biff::V3 => {
                let (row, col, fmt) = cell_head(&mut c, biff)?;
                cells.insert((row, col), globals.number(f64_le(&mut c)?, fmt));
            }
            rec::NUMBER_BIFF2 if biff == Biff::V2 => {
                let (row, col, fmt) = cell_head(&mut c, biff)?;
                cells.insert((row, col), globals.number(f64_le(&mut c)?, fmt));
            }
            rec::INTEGER_BIFF2 if biff == Biff::V2 => {
                let (row, col, fmt) = cell_head(&mut c, biff)?;
                cells.insert((row, col), globals.number(c.u16()? as f64, fmt));
            }
            rec::RK if biff >= Biff::V3 => {
                let (row, col, fmt) = cell_head(&mut c, biff)?;
                cells.insert((row, col), globals.number(rk_value(c.u32()?), fmt));
            }
            rec::MULRK if biff >= Biff::V5 => {
                let data = record.data();
                let (row, first) = (c.u16()?, c.u16()?);
                // Six bytes per cell between the 4-byte head and the 2-byte last column.
                let count = data.len().saturating_sub(6) / 6;
                for i in 0..count {
                    let xf = c.u16()?;
                    let value = rk_value(c.u32()?);
                    cells.insert((row, first + i as u16), globals.number(value, Fmt::Xf(xf)));
                }
            }
            rec::FORMULA | rec::FORMULA_BIFF3 | rec::FORMULA_BIFF4
                if formula_record(biff) == record.kind =>
            {
                let (row, col, fmt) = cell_head(&mut c, biff)?;
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
                    cells.insert((row, col), globals.number(f64::from_le_bytes(value), fmt));
                }
            }
            rec::STRING if biff >= Biff::V3 => {
                if let Some(at) = pending_string.take() {
                    let text = if biff == Biff::V8 {
                        c.unicode_string()?
                    } else {
                        globals.text(&c.byte_string16()?)
                    };
                    cells.insert(at, text);
                }
            }
            rec::STRING_BIFF2 if biff == Biff::V2 => {
                if let Some(at) = pending_string.take() {
                    cells.insert(at, globals.text(&c.byte_string8()?));
                }
            }
            rec::BOOLERR | rec::BOOLERR_BIFF2
                if (record.kind == rec::BOOLERR_BIFF2) == (biff == Biff::V2) =>
            {
                let (row, col, _) = cell_head(&mut c, biff)?;
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
            rec::HLINK if biff == Biff::V8 => {
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
            rec::OBJ if biff == Biff::V8 => {
                // The first sub-record is the common object data: ft, cb, ot, id.
                let (ft, _cb, ot, id) = (c.u16()?, c.u16()?, c.u16()?, c.u16()?);
                comment_obj = (ft == 0x15 && ot == 0x19).then_some(id);
            }
            rec::TXO if biff == Biff::V8 => {
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
            rec::NOTE if biff == Biff::V8 => {
                let (row, col, _flags, id) = (c.u16()?, c.u16()?, c.u16()?, c.u16()?);
                if let Some(text) = comment_texts.get(&id) {
                    comments.insert((row, col), text.clone());
                }
            }
            rec::NOTE => {
                let (row, col, _length) = (c.u16()?, c.u16()?, c.u16()?);
                let text = c.rest();
                if row == 0xFFFF {
                    if let Some((_, last)) = notes.last_mut() {
                        last.extend_from_slice(&text);
                    }
                } else {
                    notes.push(((row, col), text));
                }
            }
            rec::MERGEDCELLS if biff == Biff::V8 => {
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
    for (at, text) in notes {
        comments.insert(at, globals.text(&text));
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

/// The `FORMULA` record type of a BIFF version.
fn formula_record(biff: Biff) -> u16 {
    match biff {
        Biff::V3 => rec::FORMULA_BIFF3,
        Biff::V4 => rec::FORMULA_BIFF4,
        _ => rec::FORMULA,
    }
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
