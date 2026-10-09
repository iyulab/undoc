//! Text for the slide rasterizer: which font draws a character, and how a shape's paragraphs
//! are laid out in its text box.
//!
//! Fonts are looked up in the faces the caller passes, then the directories it names, then the
//! system's font directories; a character no requested face covers is drawn in a face that
//! does, and the run counts as substituted. Glyphs are taken one per character from the cmap —
//! there is no shaping, so a script that needs it (Arabic, the Indic scripts) is not drawn and
//! counts as a gap. Lines break at spaces, and between characters of the CJK scripts.

use std::collections::HashMap;
use std::fs::File;
use std::io::{Read, Seek, SeekFrom};
use std::path::{Path, PathBuf};
use std::sync::{Arc, OnceLock};

use skrifa::instance::{LocationRef, Size};
use skrifa::outline::{DrawSettings, OutlinePen};
use skrifa::raw::FileRef;
use skrifa::{FontRef, MetadataProvider};
use tiny_skia::{PathBuilder, Transform};

/// Where a face's bytes are.
#[derive(Debug, Clone)]
enum Source {
    Bytes(Arc<Vec<u8>>),
    File(PathBuf),
}

/// A face that may be used, by family name (lower case, every name its `name` table gives).
#[derive(Debug, Clone)]
struct Face {
    source: Source,
    index: u32,
    families: Vec<String>,
    bold: bool,
    italic: bool,
}

/// Families tried, in order, for a character no requested face covers.
const FALLBACK_FAMILIES: &[&str] = &[
    "malgun gothic",
    "noto sans cjk kr",
    "noto sans kr",
    "nanumgothic",
    "apple sd gothic neo",
    "arial",
    "helvetica",
    "segoe ui",
    "liberation sans",
    "dejavu sans",
    "noto sans",
];

/// The faces a render may use.
pub(super) struct FontBook {
    faces: Vec<Face>,
    loaded: HashMap<usize, Arc<Vec<u8>>>,
}

impl FontBook {
    pub(super) fn new(fonts: &[Vec<u8>], dirs: &[PathBuf], system: bool) -> Self {
        let mut faces = Vec::new();
        for data in fonts {
            let data = Arc::new(data.clone());
            faces.extend(faces_in(&data, |index, families, bold, italic| Face {
                source: Source::Bytes(data.clone()),
                index,
                families,
                bold,
                italic,
            }));
        }
        for dir in dirs {
            faces.extend(scan_dir(dir));
        }
        if system {
            faces.extend(system_faces().iter().cloned());
        }
        Self {
            faces,
            loaded: HashMap::new(),
        }
    }

    fn data(&mut self, face: usize) -> Option<Arc<Vec<u8>>> {
        if let Some(d) = self.loaded.get(&face) {
            return Some(d.clone());
        }
        let data = match &self.faces[face].source {
            Source::Bytes(d) => d.clone(),
            Source::File(path) => Arc::new(std::fs::read(path).ok()?),
        };
        self.loaded.insert(face, data.clone());
        Some(data)
    }

    fn covers(&mut self, face: usize, ch: char) -> bool {
        let index = self.faces[face].index;
        self.data(face).is_some_and(|d| {
            FontRef::from_index(&d, index)
                .ok()
                .and_then(|f| f.charmap().map(ch))
                .is_some()
        })
    }

    /// The face to draw `ch` in, and whether it is a stand-in for the requested families.
    pub(super) fn pick(
        &mut self,
        families: &[&str],
        bold: bool,
        italic: bool,
        ch: char,
    ) -> Option<(usize, bool)> {
        let style_rank = |f: &Face| (f.bold != bold) as u8 + (f.italic != italic) as u8;
        for (n, family) in families.iter().enumerate() {
            let wanted = family.to_lowercase();
            let mut candidates: Vec<usize> = (0..self.faces.len())
                .filter(|&i| self.faces[i].families.contains(&wanted))
                .collect();
            candidates.sort_by_key(|&i| style_rank(&self.faces[i]));
            for i in candidates {
                if self.covers(i, ch) {
                    return Some((i, n > 0 && families[0] != *family));
                }
            }
        }
        // A face the caller gave covers it before any the system has.
        let given: Vec<usize> = (0..self.faces.len())
            .filter(|&i| matches!(self.faces[i].source, Source::Bytes(_)))
            .collect();
        for i in given {
            if self.covers(i, ch) {
                return Some((i, true));
            }
        }
        for family in FALLBACK_FAMILIES {
            let mut candidates: Vec<usize> = (0..self.faces.len())
                .filter(|&i| self.faces[i].families.iter().any(|f| f == family))
                .collect();
            candidates.sort_by_key(|&i| style_rank(&self.faces[i]));
            for i in candidates {
                if self.covers(i, ch) {
                    return Some((i, true));
                }
            }
        }
        None
    }

    // Sizes here are in EMU (254,000 for 20 pt), past what skrifa's 16.16 fixed-point scaling
    // holds -- so metrics and outlines are read in font units and scaled here.

    /// The advance of `ch` in `face` at `size` (EMU).
    pub(super) fn advance(&mut self, face: usize, ch: char, size: f32) -> f32 {
        let index = self.faces[face].index;
        self.data(face)
            .and_then(|d| {
                let font = FontRef::from_index(&d, index).ok()?;
                let gid = font.charmap().map(ch)?;
                let units = font
                    .glyph_metrics(Size::unscaled(), LocationRef::default())
                    .advance_width(gid)?;
                Some(units * size / units_per_em(&font))
            })
            .unwrap_or(size * 0.5)
    }

    /// The ascent of `face` at `size` (EMU).
    pub(super) fn ascent(&mut self, face: usize, size: f32) -> f32 {
        let index = self.faces[face].index;
        self.data(face)
            .and_then(|d| {
                let font = FontRef::from_index(&d, index).ok()?;
                let ascent = font
                    .metrics(Size::unscaled(), LocationRef::default())
                    .ascent;
                Some(ascent * size / units_per_em(&font))
            })
            .unwrap_or(size * 0.88)
    }

    /// The outline of `ch` in `face` at `size` (EMU), with its origin at `(x, baseline)`.
    pub(super) fn glyph(
        &mut self,
        face: usize,
        ch: char,
        size: f32,
        x: f32,
        baseline: f32,
    ) -> Option<tiny_skia::Path> {
        let index = self.faces[face].index;
        let data = self.data(face)?;
        let font = FontRef::from_index(&data, index).ok()?;
        let gid = font.charmap().map(ch)?;
        let glyph = font.outline_glyphs().get(gid)?;
        let mut pen = Pen(PathBuilder::new());
        glyph
            .draw(
                DrawSettings::unhinted(Size::unscaled(), LocationRef::default()),
                &mut pen,
            )
            .ok()?;
        // Font outlines are y-up, in font units.
        let k = size / units_per_em(&font);
        pen.0
            .finish()?
            .transform(Transform::from_row(k, 0.0, 0.0, -k, x, baseline))
    }
}

fn units_per_em(font: &FontRef) -> f32 {
    font.metrics(Size::unscaled(), LocationRef::default())
        .units_per_em
        .max(1) as f32
}

struct Pen(PathBuilder);

impl OutlinePen for Pen {
    fn move_to(&mut self, x: f32, y: f32) {
        self.0.move_to(x, y);
    }
    fn line_to(&mut self, x: f32, y: f32) {
        self.0.line_to(x, y);
    }
    fn quad_to(&mut self, x1: f32, y1: f32, x: f32, y: f32) {
        self.0.quad_to(x1, y1, x, y);
    }
    fn curve_to(&mut self, x1: f32, y1: f32, x2: f32, y2: f32, x: f32, y: f32) {
        self.0.cubic_to(x1, y1, x2, y2, x, y);
    }
    fn close(&mut self) {
        self.0.close();
    }
}

/// The faces in one font file or collection, built by `make` from each one's names and style.
fn faces_in<T>(data: &[u8], make: impl Fn(u32, Vec<String>, bool, bool) -> T) -> Vec<T> {
    let Ok(file) = FileRef::new(data) else {
        return Vec::new();
    };
    file.fonts()
        .enumerate()
        .filter_map(|(i, font)| {
            let font = font.ok()?;
            let mut families: Vec<String> = [
                skrifa::string::StringId::FAMILY_NAME,
                skrifa::string::StringId::TYPOGRAPHIC_FAMILY_NAME,
            ]
            .into_iter()
            .flat_map(|id| font.localized_strings(id))
            .map(|s| s.to_string().to_lowercase())
            .collect();
            families.dedup();
            let attrs = font.attributes();
            Some(make(
                i as u32,
                families,
                attrs.weight.value() >= 600.0,
                attrs.style != skrifa::attribute::Style::Normal,
            ))
        })
        .collect()
}

fn scan_dir(dir: &Path) -> Vec<Face> {
    let mut faces = Vec::new();
    let mut stack = vec![dir.to_path_buf()];
    while let Some(d) = stack.pop() {
        let Ok(entries) = std::fs::read_dir(&d) else {
            continue;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                stack.push(path);
                continue;
            }
            let ext = path
                .extension()
                .and_then(|e| e.to_str())
                .map(str::to_ascii_lowercase);
            if !matches!(ext.as_deref(), Some("ttf" | "otf" | "ttc" | "otc")) {
                continue;
            }
            // Only the tables naming the faces are read here; the rest of the file is read
            // when a face is used.
            if let Some(head) = read_naming_tables(&path) {
                faces.extend(faces_in(&head, |index, families, bold, italic| Face {
                    source: Source::File(path.clone()),
                    index,
                    families,
                    bold,
                    italic,
                }));
            }
        }
    }
    faces
}

/// The system's fonts, scanned once per process.
fn system_faces() -> &'static [Face] {
    static FACES: OnceLock<Vec<Face>> = OnceLock::new();
    FACES.get_or_init(|| {
        let mut dirs: Vec<PathBuf> = Vec::new();
        if let Some(windir) = std::env::var_os("WINDIR") {
            dirs.push(PathBuf::from(windir).join("Fonts"));
        }
        if let Some(local) = std::env::var_os("LOCALAPPDATA") {
            dirs.push(PathBuf::from(local).join("Microsoft/Windows/Fonts"));
        }
        dirs.extend(
            [
                "/Library/Fonts",
                "/System/Library/Fonts",
                "/usr/share/fonts",
                "/usr/local/share/fonts",
            ]
            .iter()
            .map(PathBuf::from),
        );
        if let Some(home) = std::env::var_os("HOME") {
            let home = PathBuf::from(home);
            dirs.push(home.join(".fonts"));
            dirs.push(home.join(".local/share/fonts"));
            dirs.push(home.join("Library/Fonts"));
        }
        dirs.iter().flat_map(|d| scan_dir(d)).collect()
    })
}

/// A copy of a font file holding only what naming its faces needs — the headers, the table
/// directories and the `name` and `OS/2` tables, each at its own offset with zeros between —
/// so a font directory can be indexed without reading every glyph in it.
fn read_naming_tables(path: &Path) -> Option<Vec<u8>> {
    let mut file = File::open(path).ok()?;
    let len = file.metadata().ok()?.len();
    let mut read_at = |offset: u64, n: usize| -> Option<Vec<u8>> {
        if offset + n as u64 > len {
            return None;
        }
        file.seek(SeekFrom::Start(offset)).ok()?;
        let mut buf = vec![0u8; n];
        file.read_exact(&mut buf).ok()?;
        Some(buf)
    };
    let be32 = |b: &[u8], at: usize| u32::from_be_bytes([b[at], b[at + 1], b[at + 2], b[at + 3]]);
    let be16 = |b: &[u8], at: usize| u16::from_be_bytes([b[at], b[at + 1]]);

    let head = read_at(0, 12)?;
    let mut copy: Vec<(u64, Vec<u8>)> = Vec::new();
    let mut offsets = vec![0u64];
    if &head[0..4] == b"ttcf" {
        let count = be32(&head, 8).min(64) as usize;
        let table = read_at(12, count * 4)?;
        offsets = (0..count).map(|i| be32(&table, i * 4) as u64).collect();
        copy.push((0, [head.clone(), table].concat()));
    }
    for offset in offsets {
        let dir_head = read_at(offset, 12)?;
        let tables = be16(&dir_head, 4).min(256) as usize;
        let records = read_at(offset + 12, tables * 16)?;
        for r in 0..tables {
            let rec = &records[r * 16..r * 16 + 16];
            if &rec[0..4] == b"name" || &rec[0..4] == b"OS/2" {
                let (at, size) = (be32(rec, 8) as u64, be32(rec, 12) as usize);
                if size <= 1 << 20 {
                    copy.push((at, read_at(at, size)?));
                }
            }
        }
        copy.push((offset, [dir_head, records].concat()));
    }
    let end = copy.iter().map(|(at, b)| at + b.len() as u64).max()?;
    if end > 64 << 20 {
        return None;
    }
    let mut out = vec![0u8; end as usize];
    for (at, bytes) in copy {
        out[at as usize..at as usize + bytes.len()].copy_from_slice(&bytes);
    }
    Some(out)
}

/// Whether `ch` belongs to a script whose lines may break between any two characters.
pub(super) fn breaks_anywhere(ch: char) -> bool {
    matches!(ch as u32,
        0x1100..=0x11FF | 0x3040..=0x30FF | 0x3130..=0x318F | 0x3400..=0x4DBF |
        0x4E00..=0x9FFF | 0xAC00..=0xD7AF | 0xF900..=0xFAFF | 0xFF00..=0xFFEF)
}

/// Whether `ch` is East Asian, drawn with a run's `ea` font rather than its `latin` one.
pub(super) fn is_east_asian(ch: char) -> bool {
    breaks_anywhere(ch) || matches!(ch as u32, 0x3000..=0x303F)
}

/// Whether `ch` needs shaping this does not do.
pub(super) fn needs_shaping(ch: char) -> bool {
    matches!(ch as u32,
        0x0600..=0x06FF | 0x0750..=0x077F | 0x0900..=0x0DFF | 0x0E00..=0x0E7F | 0x1000..=0x109F)
}
