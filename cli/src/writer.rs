//! Multi-format fan-out writer for the convert command.
//!
//! Supports two usage modes:
//!
//! - **Batch** (current default): `write_document(&Document)` renders the full
//!   document in one shot.
//! - **Streaming** (PPTX/XLSX): `write_document_start`, `write_section`,
//!   `finish` — processes one section at a time with bounded memory.

use std::collections::HashMap;
use std::fs::{self, File};
use std::io::{self, BufWriter, Write};
use std::path::{Path, PathBuf};

use undoc::detect::FormatType;
use undoc::model::{Document, Section};
use undoc::render::{
    render_section_to_string, to_json, to_markdown, to_text, JsonFormat, RenderOptions,
};

/// Output formats supported by the convert command.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OutputFormat {
    Markdown,
    Text,
    Json,
}

impl OutputFormat {
    pub fn parse(s: &str) -> Option<Self> {
        match s.trim().to_ascii_lowercase().as_str() {
            "md" | "markdown" => Some(Self::Markdown),
            "txt" | "text" => Some(Self::Text),
            "json" => Some(Self::Json),
            _ => None,
        }
    }
}

/// Summary of files written by the writer.
#[derive(Default)]
pub struct WriteSummary {
    pub md_path: Option<PathBuf>,
    pub txt_path: Option<PathBuf>,
    pub json_path: Option<PathBuf>,
    /// Word count from Markdown output (0 if Markdown not requested).
    pub word_count: usize,
    /// Number of sections written.
    pub section_count: usize,
}

/// Fan-out writer that produces Markdown, text, and/or JSON for a document.
pub struct MultiFormatWriter<'a> {
    out_dir: &'a Path,
    formats: &'a [OutputFormat],
    render_opts: &'a RenderOptions,
}

impl<'a> MultiFormatWriter<'a> {
    pub fn new(
        out_dir: &'a Path,
        formats: &'a [OutputFormat],
        render_opts: &'a RenderOptions,
    ) -> Self {
        Self {
            out_dir,
            formats,
            render_opts,
        }
    }

    /// Batch render: write all formats for a complete Document.
    pub fn write_document(&self, doc: &Document) -> io::Result<WriteSummary> {
        let want_md = self.formats.contains(&OutputFormat::Markdown);
        let want_txt = self.formats.contains(&OutputFormat::Text);
        let want_json = self.formats.contains(&OutputFormat::Json);

        let mut summary = WriteSummary::default();

        summary.section_count = doc.sections.len();

        if want_md {
            let markdown = to_markdown(doc, self.render_opts)
                .map_err(|e| io::Error::new(io::ErrorKind::InvalidData, e.to_string()))?;
            let path = self.out_dir.join("extract.md");
            fs::write(&path, &markdown)?;
            summary.word_count = markdown.split_whitespace().count();
            summary.md_path = Some(path);
        }

        if want_txt {
            let text = to_text(doc, self.render_opts)
                .map_err(|e| io::Error::new(io::ErrorKind::InvalidData, e.to_string()))?;
            let path = self.out_dir.join("extract.txt");
            fs::write(&path, text)?;
            summary.txt_path = Some(path);
        }

        if want_json {
            let json = to_json(doc, JsonFormat::Pretty)
                .map_err(|e| io::Error::new(io::ErrorKind::InvalidData, e.to_string()))?;
            let path = self.out_dir.join("content.json");
            fs::write(&path, json)?;
            summary.json_path = Some(path);
        }

        Ok(summary)
    }

    /// Open streaming writers. Call once before any `write_section` calls.
    pub fn open_streaming(
        &self,
        doc_format: FormatType,
        image_map: HashMap<String, String>,
    ) -> io::Result<StreamingWriter> {
        StreamingWriter::new(
            self.out_dir,
            self.formats,
            self.render_opts,
            doc_format,
            image_map,
        )
    }
}

/// Stateful writer for the streaming path.
///
/// Created via [`MultiFormatWriter::open_streaming`]. Call
/// [`write_section`](StreamingWriter::write_section) for each section, then
/// [`finish`](StreamingWriter::finish) when done.
pub struct StreamingWriter {
    md: Option<BufWriter<File>>,
    md_path: Option<PathBuf>,
    txt: Option<BufWriter<File>>,
    txt_path: Option<PathBuf>,
    json: Option<BufWriter<File>>,
    json_path: Option<PathBuf>,
    json_first_section: bool,
    render_opts: RenderOptions,
    doc_format: FormatType,
    image_map: HashMap<String, String>,
    word_count: usize,
    section_index: usize,
}

impl StreamingWriter {
    fn new(
        out_dir: &Path,
        formats: &[OutputFormat],
        render_opts: &RenderOptions,
        doc_format: FormatType,
        image_map: HashMap<String, String>,
    ) -> io::Result<Self> {
        let want_md = formats.contains(&OutputFormat::Markdown);
        let want_txt = formats.contains(&OutputFormat::Text);
        let want_json = formats.contains(&OutputFormat::Json);

        let (md, md_path) = if want_md {
            let p = out_dir.join("extract.md");
            let f = File::create(&p)?;
            (Some(BufWriter::new(f)), Some(p))
        } else {
            (None, None)
        };

        let (txt, txt_path) = if want_txt {
            let p = out_dir.join("extract.txt");
            let f = File::create(&p)?;
            (Some(BufWriter::new(f)), Some(p))
        } else {
            (None, None)
        };

        let (json, json_path) = if want_json {
            let p = out_dir.join("content.json");
            let f = File::create(&p)?;
            (Some(BufWriter::new(f)), Some(p))
        } else {
            (None, None)
        };

        Ok(Self {
            md,
            md_path,
            txt,
            txt_path,
            json,
            json_path,
            json_first_section: true,
            render_opts: render_opts.clone(),
            doc_format,
            image_map,
            word_count: 0,
            section_index: 0,
        })
    }

    /// Write one section to all open format writers.
    pub fn write_section(&mut self, section: &Section) -> io::Result<()> {
        let idx = self.section_index;
        self.section_index += 1;

        // Markdown
        if let Some(ref mut md) = self.md {
            let rendered = render_section_to_string(
                section,
                idx,
                self.doc_format,
                &self.render_opts,
                &self.image_map,
            );
            self.word_count += rendered.split_whitespace().count();
            md.write_all(rendered.as_bytes())?;
            md.write_all(b"\n")?;
        }

        // Plain text
        if let Some(ref mut txt) = self.txt {
            for block in &section.content {
                use undoc::model::Block;
                match block {
                    Block::Paragraph(p) => {
                        let line = p.plain_text();
                        if !line.is_empty() {
                            writeln!(txt, "{}", line)?;
                        }
                    }
                    Block::Table(t) => {
                        for row in &t.rows {
                            for cell in &row.cells {
                                let text = cell.plain_text();
                                if !text.is_empty() {
                                    writeln!(txt, "{}", text)?;
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
        }

        // JSON
        if let Some(ref mut json) = self.json {
            if !self.json_first_section {
                write!(json, ",")?;
            }
            let section_json = serde_json::to_string(section)
                .map_err(|e| io::Error::new(io::ErrorKind::InvalidData, e))?;
            write!(json, "{}", section_json)?;
            self.json_first_section = false;
        }

        Ok(())
    }

    /// Write JSON envelope header. Must be called before `write_section` for JSON output.
    pub fn write_json_start(&mut self) -> io::Result<()> {
        if let Some(ref mut json) = self.json {
            write!(json, "{{\"sections\":[")?;
        }
        Ok(())
    }

    /// Flush all writers and return a summary of written paths.
    pub fn finish(mut self) -> io::Result<WriteSummary> {
        let mut summary = WriteSummary::default();

        if let (Some(mut md), Some(md_path)) = (self.md.take(), self.md_path.take()) {
            md.flush()?;
            summary.md_path = Some(md_path);
        }

        if let (Some(mut txt), Some(txt_path)) = (self.txt.take(), self.txt_path.take()) {
            txt.flush()?;
            summary.txt_path = Some(txt_path);
        }

        if let (Some(mut json), Some(json_path)) = (self.json.take(), self.json_path.take()) {
            write!(json, "]}}")?;
            json.flush()?;
            summary.json_path = Some(json_path);
        }

        summary.word_count = self.word_count;
        summary.section_count = self.section_index;
        Ok(summary)
    }
}
