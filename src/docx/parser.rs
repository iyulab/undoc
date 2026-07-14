//! DOCX parser implementation.

use std::collections::HashMap;

use crate::charts;
use crate::container::OoxmlContainer;
use crate::error::{Error, Result};
use crate::model::{
    Block, Cell, CellAlignment, Document, ListInfo, ListType, Metadata, Paragraph, Resource,
    ResourceType, RevisionType, Row, Section, Table, TextAlignment, TextRun, TextStyle,
    VerticalAlignment,
};

use super::numbering::NumberingMap;
use super::styles::StyleMap;

/// Parser for DOCX (Word) documents.
pub struct DocxParser {
    container: OoxmlContainer,
    styles: StyleMap,
    numbering: NumberingMap,
    relationships: crate::container::Relationships,
    /// Footnote id → plain text content
    footnotes: HashMap<String, String>,
    /// Endnote id → plain text content
    endnotes: HashMap<String, String>,
}

impl DocxParser {
    /// Open a DOCX file for parsing.
    #[cfg(not(target_arch = "wasm32"))]
    pub fn open(path: impl AsRef<std::path::Path>) -> Result<Self> {
        let container = OoxmlContainer::open(path)?;
        Self::from_container(container)
    }

    /// Create a parser from bytes.
    pub fn from_bytes(data: Vec<u8>) -> Result<Self> {
        let container = OoxmlContainer::from_bytes(data)?;
        Self::from_container(container)
    }

    /// Create a parser from a container.
    fn from_container(container: OoxmlContainer) -> Result<Self> {
        // Parse styles — absent is OK, malformed bytes must surface.
        let styles = match container.read_xml_optional("word/styles.xml")? {
            Some(xml) => StyleMap::parse(&xml)?,
            None => StyleMap::default(),
        };

        // Parse numbering — absent is OK, malformed bytes must surface.
        let numbering = match container.read_xml_optional("word/numbering.xml")? {
            Some(xml) => NumberingMap::parse(&xml)?,
            None => NumberingMap::default(),
        };

        // Parse document relationships when present.
        //
        // Text-only DOCX files can omit document.xml.rels entirely, so keep
        // extraction best-effort and only use relationships for linked assets
        // such as images, charts, headers/footers, and hyperlinks.
        let relationships = container.read_optional_relationships_for_part("word/document.xml")?;

        // Parse footnotes — absent is OK, malformed bytes must surface.
        let footnotes = match container.read_xml_optional("word/footnotes.xml")? {
            Some(xml) => parse_notes_xml(&xml, b"w:footnote"),
            None => HashMap::new(),
        };

        // Parse endnotes — absent is OK, malformed bytes must surface.
        let endnotes = match container.read_xml_optional("word/endnotes.xml")? {
            Some(xml) => parse_notes_xml(&xml, b"w:endnote"),
            None => HashMap::new(),
        };

        Ok(Self {
            container,
            styles,
            numbering,
            relationships,
            footnotes,
            endnotes,
        })
    }

    /// Parse the document and return a Document model.
    pub fn parse(&mut self) -> Result<Document> {
        let mut doc = Document::new();
        doc.format = crate::detect::FormatType::Docx;

        // Parse metadata
        doc.metadata = self.parse_metadata()?;

        // Parse main document content
        let mut main_section = self.parse_document_xml()?;

        // Parse charts and add as tables for RAG-ready output
        let chart_tables = self.parse_charts()?;
        for table in chart_tables {
            main_section.add_block(Block::Table(table));
        }

        // Append footnote definitions at end of section
        if !self.footnotes.is_empty() {
            let mut ids: Vec<&String> = self.footnotes.keys().collect();
            ids.sort_by(|a, b| {
                a.parse::<u64>()
                    .unwrap_or(u64::MAX)
                    .cmp(&b.parse::<u64>().unwrap_or(u64::MAX))
            });
            for id in ids {
                if let Some(text) = self.footnotes.get(id) {
                    let para = Paragraph::with_text(format!("[^{}]: {}", id, text));
                    main_section.add_block(Block::Paragraph(para));
                }
            }
        }

        // Append endnote definitions at end of section
        if !self.endnotes.is_empty() {
            let mut ids: Vec<&String> = self.endnotes.keys().collect();
            ids.sort_by(|a, b| {
                a.parse::<u64>()
                    .unwrap_or(u64::MAX)
                    .cmp(&b.parse::<u64>().unwrap_or(u64::MAX))
            });
            for id in ids {
                if let Some(text) = self.endnotes.get(id) {
                    let para = Paragraph::with_text(format!("[^e{}]: {}", id, text));
                    main_section.add_block(Block::Paragraph(para));
                }
            }
        }

        doc.add_section(main_section);

        // Extract resources (images)
        self.extract_resources(&mut doc)?;

        Ok(doc)
    }

    /// Stream document sections one at a time via a callback.
    ///
    /// DOCX has a single logical document but may contain multiple sections via
    /// page-section breaks.  The entire document is parsed upfront and each
    /// section is delivered as a [`SectionParsed`](crate::streaming::ParseEvent) event.
    #[cfg(not(target_arch = "wasm32"))]
    pub fn for_each_section<F>(
        &mut self,
        opts: crate::streaming::SectionStreamOptions,
        mut f: F,
    ) -> crate::error::Result<()>
    where
        F: FnMut(crate::streaming::ParseEvent<'_>) -> std::ops::ControlFlow<()>,
    {
        let metadata = self.parse_metadata()?;

        // Parse the complete document first so we can report section_count and
        // build the image_map before emitting DocumentStart.
        let doc = match self.parse() {
            Ok(d) => d,
            Err(e) if opts.lenient => {
                // Emit a degenerate stream with a single failure
                let _ = f(crate::streaming::ParseEvent::DocumentStart {
                    metadata: &metadata,
                    section_count: 0,
                    image_map: HashMap::new(),
                });
                let _ = f(crate::streaming::ParseEvent::SectionFailed { index: 0, error: e });
                let _ = f(crate::streaming::ParseEvent::DocumentEnd);
                return Ok(());
            }
            Err(e) => return Err(e),
        };

        let image_map: HashMap<String, String> = doc
            .resources
            .iter()
            .filter_map(|(id, r)| r.filename.as_ref().map(|name| (id.clone(), name.clone())))
            .collect();

        if f(crate::streaming::ParseEvent::DocumentStart {
            metadata: &metadata,
            section_count: doc.sections.len(),
            image_map,
        })
        .is_break()
        {
            return Ok(());
        }

        for section in &doc.sections {
            if f(crate::streaming::ParseEvent::SectionParsed(section)).is_break() {
                return Ok(());
            }
        }

        if f(crate::streaming::ParseEvent::DocumentEnd).is_break() {
            return Ok(());
        }

        if opts.extract_resources {
            for (id, resource) in doc.resources {
                let name = resource.filename.clone().unwrap_or(id);
                if f(crate::streaming::ParseEvent::ResourceExtracted {
                    name,
                    data: resource.data,
                })
                .is_break()
                {
                    return Ok(());
                }
            }
        }

        Ok(())
    }

    /// Parse document metadata from docProps/core.xml.
    fn parse_metadata(&self) -> Result<Metadata> {
        // Use shared metadata parsing from container
        self.container.parse_core_metadata()
    }

    /// Parse charts from word/charts/ and convert to tables for RAG-ready output.
    fn parse_charts(&self) -> Result<Vec<Table>> {
        let mut tables = Vec::new();

        // Find chart relationships in document.xml.rels
        for (rel_type, rels) in &self.relationships.by_type {
            if !rel_type.contains("chart") {
                continue;
            }

            for rel in rels {
                // Resolve chart path relative to document.xml
                // Target is like "charts/chart1.xml"
                let chart_path = if rel.target.starts_with('/') {
                    rel.target[1..].to_string()
                } else {
                    format!("word/{}", rel.target)
                };

                // Read and parse chart XML
                let chart_xml = self.container.read_xml(&chart_path)?;
                match charts::parse_chart_xml(&chart_xml) {
                    Ok(chart_data) => {
                        if !chart_data.is_empty() {
                            let mut table = chart_data.to_table();
                            // Add chart title if available
                            if let Some(ref title) = chart_data.title {
                                if !title.is_empty() {
                                    if let Some(first_row) = table.rows.first_mut() {
                                        if let Some(first_cell) = first_row.cells.first_mut() {
                                            let original = first_cell.plain_text();
                                            first_cell.content.clear();
                                            first_cell.content.push(Paragraph::with_text(format!(
                                                "{} ({})",
                                                original, title
                                            )));
                                        }
                                    }
                                }
                            }
                            tables.push(table);
                        }
                    }
                    Err(e) => return Err(e),
                }
            }
        }

        Ok(tables)
    }

    /// Parse the main document.xml content.
    fn parse_document_xml(&mut self) -> Result<Section> {
        let xml = self.container.read_xml("word/document.xml")?;
        let mut section = Section::new(0);

        let mut reader = crate::decode::reader_for(&xml);
        // IMPORTANT: Don't trim text - preserve whitespace from xml:space="preserve" elements
        // This fixes the "DATE OF BIRTH" -> "DATEOFBIRTH" bug (GitHub Issue #2)
        reader.config_mut().trim_text(false);

        let mut buf = Vec::new();
        let mut in_body = false;
        let mut paragraph_xml = String::new();
        let mut table_xml = String::new();
        let mut in_paragraph = false;
        let mut para_depth: u32 = 0; // Track nested w:p depth (for text boxes)
        let mut table_depth: u32 = 0; // Track nested table depth
                                      // Every header/footer reference (default / first / even), in document order.
        let mut header_rids: Vec<String> = Vec::new();
        let mut footer_rids: Vec<String> = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    let name = e.name();
                    match name.as_ref() {
                        b"w:body" => {
                            in_body = true;
                        }
                        b"w:p" if in_body && table_depth == 0 && !in_paragraph => {
                            in_paragraph = true;
                            paragraph_xml.clear();
                            paragraph_xml.push_str("<w:p");
                            for attr in e.attributes().flatten() {
                                paragraph_xml.push_str(&format!(
                                    " {}=\"{}\"",
                                    String::from_utf8_lossy(attr.key.as_ref()),
                                    String::from_utf8_lossy(&attr.value)
                                ));
                            }
                            paragraph_xml.push('>');
                        }
                        // Only body-level tables enter table mode. A table nested
                        // inside a text box lives within a paragraph; it must stay
                        // in paragraph_xml so the text-box path extracts its text,
                        // otherwise table mode collects an empty shell (cell text is
                        // captured separately) and emits a spurious empty table.
                        b"w:tbl" if in_body && !in_paragraph => {
                            if table_depth == 0 {
                                // Start collecting table XML
                                table_xml.clear();
                            }
                            table_depth += 1;
                            table_xml.push_str("<w:tbl>");
                        }
                        _ => {
                            if in_paragraph {
                                // Track nested w:p depth for text boxes
                                if name.as_ref() == b"w:p" {
                                    para_depth += 1;
                                }
                                paragraph_xml.push('<');
                                paragraph_xml.push_str(&String::from_utf8_lossy(name.as_ref()));
                                for attr in e.attributes().flatten() {
                                    paragraph_xml.push_str(&format!(
                                        " {}=\"{}\"",
                                        String::from_utf8_lossy(attr.key.as_ref()),
                                        String::from_utf8_lossy(&attr.value)
                                    ));
                                }
                                paragraph_xml.push('>');
                            } else if table_depth > 0 {
                                table_xml.push('<');
                                table_xml.push_str(&String::from_utf8_lossy(name.as_ref()));
                                for attr in e.attributes().flatten() {
                                    table_xml.push_str(&format!(
                                        " {}=\"{}\"",
                                        String::from_utf8_lossy(attr.key.as_ref()),
                                        String::from_utf8_lossy(&attr.value)
                                    ));
                                }
                                table_xml.push('>');
                            }
                        }
                    }
                }
                Ok(quick_xml::events::Event::Empty(ref e)) => {
                    let name = e.name();
                    // Header/footer references live inside a w:sectPr, which may be
                    // either the final body-level sectPr OR a section-break sectPr
                    // nested in a paragraph's w:pPr. Collect them wherever they
                    // appear in the body (every type: default/first/even) so
                    // multi-section documents don't lose every section's
                    // header/footer text but the last.
                    if in_body
                        && table_depth == 0
                        && matches!(name.as_ref(), b"w:headerReference" | b"w:footerReference")
                    {
                        let mut r_id = String::new();
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"r:id" {
                                r_id = String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }
                        if !r_id.is_empty() {
                            if name.as_ref() == b"w:headerReference" {
                                header_rids.push(r_id);
                            } else {
                                footer_rids.push(r_id);
                            }
                        }
                    } else if in_paragraph {
                        paragraph_xml.push('<');
                        paragraph_xml.push_str(&String::from_utf8_lossy(name.as_ref()));
                        for attr in e.attributes().flatten() {
                            paragraph_xml.push_str(&format!(
                                " {}=\"{}\"",
                                String::from_utf8_lossy(attr.key.as_ref()),
                                String::from_utf8_lossy(&attr.value)
                            ));
                        }
                        paragraph_xml.push_str("/>");
                    } else if table_depth > 0 {
                        table_xml.push('<');
                        table_xml.push_str(&String::from_utf8_lossy(name.as_ref()));
                        for attr in e.attributes().flatten() {
                            table_xml.push_str(&format!(
                                " {}=\"{}\"",
                                String::from_utf8_lossy(attr.key.as_ref()),
                                String::from_utf8_lossy(&attr.value)
                            ));
                        }
                        table_xml.push_str("/>");
                    }
                }
                Ok(quick_xml::events::Event::Text(ref e)) => {
                    if in_paragraph {
                        let text = crate::decode::decode_text_lossy(e);
                        paragraph_xml.push_str(&escape_xml(&text));
                    } else if table_depth > 0 {
                        let text = crate::decode::decode_text_lossy(e);
                        table_xml.push_str(&escape_xml(&text));
                    }
                }
                // quick-xml 0.40+ emits entity refs as separate events. This buffer
                // is re-serialized XML that gets parsed a second time, so re-emit
                // the resolved entity through escape_xml — the second pass then sees
                // "&amp;" again and decodes it via its own GeneralRef arm.
                Ok(quick_xml::events::Event::GeneralRef(ref e)) => {
                    let decoded = crate::decode::resolve_general_ref(e);
                    if in_paragraph {
                        paragraph_xml.push_str(&escape_xml(&decoded));
                    } else if table_depth > 0 {
                        table_xml.push_str(&escape_xml(&decoded));
                    }
                }
                Ok(quick_xml::events::Event::End(ref e)) => {
                    let name = e.name();
                    match name.as_ref() {
                        b"w:body" => {
                            in_body = false;
                        }
                        b"w:p" if in_paragraph && table_depth == 0 && para_depth == 0 => {
                            paragraph_xml.push_str("</w:p>");
                            // Extract text box paragraphs before parsing the main paragraph
                            let textbox_paras = self.extract_textbox_paragraphs(&paragraph_xml);
                            if let Ok(para) = self.parse_paragraph(&paragraph_xml) {
                                section.add_block(Block::Paragraph(para));
                            }
                            // Add text box paragraphs as separate blocks
                            for tb_para in textbox_paras {
                                section.add_block(Block::Paragraph(tb_para));
                            }
                            in_paragraph = false;
                        }
                        b"w:tbl" if table_depth > 0 => {
                            table_xml.push_str("</w:tbl>");
                            table_depth -= 1;
                            if table_depth == 0 {
                                // Finished collecting outermost table - now parse it
                                if let Ok(table) = self.parse_table(&table_xml) {
                                    section.add_block(Block::Table(table));
                                }
                            }
                        }
                        _ => {
                            if in_paragraph {
                                // Track nested w:p depth for text boxes
                                if name.as_ref() == b"w:p" {
                                    para_depth = para_depth.saturating_sub(1);
                                }
                                paragraph_xml.push_str("</");
                                paragraph_xml.push_str(&String::from_utf8_lossy(name.as_ref()));
                                paragraph_xml.push('>');
                            } else if table_depth > 0 {
                                table_xml.push_str("</");
                                table_xml.push_str(&String::from_utf8_lossy(name.as_ref()));
                                table_xml.push('>');
                            }
                        }
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => {
                    return Err(Error::xml_parse_with_context(
                        e.to_string(),
                        "word/document.xml",
                    ))
                }
                _ => {}
            }
            buf.clear();
        }

        // Resolve and parse header/footer from sectPr references. A single part
        // can be referenced by multiple section types, so parse each part once
        // (preserving document order) and merge the paragraphs into one list.
        section.header = self.resolve_header_footer_parts(&header_rids)?;
        section.footer = self.resolve_header_footer_parts(&footer_rids)?;

        Ok(section)
    }

    /// Resolve a relationship ID to a header/footer XML path and parse its paragraphs.
    ///
    /// Returns `Ok(None)` when the relationship is absent or the target part is
    /// missing, but surfaces malformed content (e.g. `Error::Encoding`) so that
    /// corrupted header/footer parts are not silently dropped.
    /// Resolve a set of header/footer relationship IDs (in document order) into a
    /// single merged paragraph list. Duplicate IDs (the same part referenced by
    /// several section types) are parsed only once. Returns `None` when nothing
    /// resolves to non-empty content, keeping `section.header/footer` absent.
    fn resolve_header_footer_parts(&mut self, rids: &[String]) -> Result<Option<Vec<Paragraph>>> {
        let mut seen = std::collections::HashSet::new();
        let mut merged: Vec<Paragraph> = Vec::new();
        for rid in rids {
            if !seen.insert(rid.as_str()) {
                continue;
            }
            if let Some(paragraphs) = self.parse_header_footer_by_rid(rid)? {
                merged.extend(paragraphs);
            }
        }
        Ok((!merged.is_empty()).then_some(merged))
    }

    fn parse_header_footer_by_rid(&mut self, rid: &str) -> Result<Option<Vec<Paragraph>>> {
        let Some(rel) = self.relationships.get(rid) else {
            return Ok(None);
        };
        let path = OoxmlContainer::resolve_path("word/document.xml", &rel.target);
        let Some(xml) = self.container.read_xml_optional(&path)? else {
            return Ok(None);
        };

        // A header/footer part has its OWN relationships namespace
        // (e.g. word/_rels/header1.xml.rels), independently numbered from the
        // body's. Swap them in while parsing so hyperlink/image r:ids resolve
        // against the correct part and can't collide with document.xml's rIds.
        let part_rels = self.container.read_optional_relationships_for_part(&path)?;
        let saved = std::mem::replace(&mut self.relationships, part_rels);
        let paragraphs = self.parse_header_footer_xml(&xml);
        self.relationships = saved;

        Ok(Some(paragraphs))
    }

    /// Parse a header or footer XML part (`w:hdr` / `w:ftr`) into a flat list of
    /// paragraphs.
    ///
    /// Header/footer XML shares the document body's structure, so this reuses the
    /// same building blocks as the body parser rather than a divergent text-only
    /// scan: [`Self::parse_paragraph`] for run text and formatting and
    /// [`Self::extract_textbox_paragraphs`] for text boxes (kept as separate
    /// paragraphs, matching body behaviour).
    ///
    /// The flat header/footer model (`Vec<Paragraph>`) cannot hold table
    /// structure, so tables are treated transparently: their structural tags
    /// (`w:tbl`/`w:tr`/`w:tc` and props) are skipped and each cell's paragraphs
    /// become loose paragraphs in document order. Because tables are transparent
    /// rather than parsed as a unit, *nested* tables recurse naturally and their
    /// text is never dropped.
    fn parse_header_footer_xml(&mut self, xml: &str) -> Vec<Paragraph> {
        let mut paragraphs = Vec::new();
        let mut reader = crate::decode::reader_for(xml);
        // Preserve whitespace (xml:space="preserve"), consistent with the body parser.
        reader.config_mut().trim_text(false);

        let mut buf = Vec::new();
        let mut paragraph_xml = String::new();
        let mut in_paragraph = false;
        let mut para_depth: u32 = 0; // Nested w:p depth (text boxes)

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    let name = e.name();
                    match name.as_ref() {
                        // A paragraph at any nesting depth (top level or inside a
                        // table cell). Table wrappers are transparent, so cell
                        // paragraphs start here just like top-level ones.
                        b"w:p" if !in_paragraph => {
                            in_paragraph = true;
                            para_depth = 0;
                            paragraph_xml.clear();
                            push_start_tag(&mut paragraph_xml, e);
                        }
                        _ => {
                            if in_paragraph {
                                if name.as_ref() == b"w:p" {
                                    para_depth += 1;
                                }
                                push_start_tag(&mut paragraph_xml, e);
                            }
                            // Outside a paragraph, table structural tags are
                            // ignored (descended into transparently).
                        }
                    }
                }
                // Content events only matter inside a paragraph; outside one (e.g.
                // between table cells) they are ignored.
                Ok(quick_xml::events::Event::Empty(ref e)) if in_paragraph => {
                    push_empty_tag(&mut paragraph_xml, e);
                }
                Ok(quick_xml::events::Event::Text(ref e)) if in_paragraph => {
                    let text = crate::decode::decode_text_lossy(e);
                    paragraph_xml.push_str(&escape_xml(&text));
                }
                // Re-emit resolved entities through escape_xml; the re-parse of this
                // buffer decodes them again (mirrors the body parser).
                Ok(quick_xml::events::Event::GeneralRef(ref e)) if in_paragraph => {
                    let decoded = crate::decode::resolve_general_ref(e);
                    paragraph_xml.push_str(&escape_xml(&decoded));
                }
                Ok(quick_xml::events::Event::End(ref e)) => {
                    let name = e.name();
                    match name.as_ref() {
                        b"w:p" if in_paragraph && para_depth == 0 => {
                            paragraph_xml.push_str("</w:p>");
                            // Text boxes are pulled out first, as separate paragraphs.
                            let textbox_paras = self.extract_textbox_paragraphs(&paragraph_xml);
                            if let Ok(para) = self.parse_paragraph(&paragraph_xml) {
                                // Skip blank lines (e.g. spacer paragraphs), but keep
                                // internal whitespace intact — consistent with the body parser.
                                if !para.plain_text().trim().is_empty() {
                                    paragraphs.push(para);
                                }
                            }
                            paragraphs.extend(textbox_paras);
                            in_paragraph = false;
                        }
                        _ => {
                            if in_paragraph {
                                if name.as_ref() == b"w:p" {
                                    para_depth = para_depth.saturating_sub(1);
                                }
                                push_end_tag(&mut paragraph_xml, name.as_ref());
                            }
                        }
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }

        // Header/footer text is auxiliary: it must not enter the document heading
        // outline (a Heading paragraph style here would otherwise emit a stray `#`
        // in the rendered header/footer and mislead structure-aware consumers).
        // Run-level formatting (bold/italic/hyperlinks) is preserved.
        for para in &mut paragraphs {
            para.heading = crate::model::HeadingLevel::None;
        }

        paragraphs
    }

    /// Parse a single paragraph element.
    fn parse_paragraph(&mut self, xml: &str) -> Result<Paragraph> {
        use crate::model::InlineImage;

        let mut para = Paragraph::new();
        let mut reader = crate::decode::reader_for(xml);
        // Don't trim text - preserve whitespace from xml:space="preserve" elements
        reader.config_mut().trim_text(false);

        let mut buf = Vec::new();
        let mut in_ppr = false;
        let mut in_rpr = false;
        let mut in_run = false;
        let mut in_text = false; // Track w:t elements (regular text)
        let mut in_instr_text = false; // Track w:instrText elements (field codes to skip)
        let mut in_drawing = false; // Track w:drawing elements for images
        let mut in_pict = false; // Track w:pict/w:object elements for VML images
        let mut in_ins = false; // Track w:ins elements (tracked changes - insertions)
        let mut in_del = false; // Track w:del elements (tracked changes - deletions)
        let mut txbx_content_depth: u32 = 0; // Track w:txbxContent nesting (suppress text capture)
        let mut mc_fallback_depth: u32 = 0; // Track mc:Fallback nesting (skip entirely)
        let mut current_style = TextStyle::default();
        let mut current_hyperlink: Option<String> = None;
        let mut current_image_alt: Option<String> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => match e.name().as_ref() {
                    // Skip mc:Fallback branches to avoid duplicating text box content
                    b"mc:Fallback" => {
                        mc_fallback_depth += 1;
                    }
                    // Track w:txbxContent to suppress text capture (extracted separately)
                    b"w:txbxContent" if mc_fallback_depth == 0 => {
                        txbx_content_depth += 1;
                    }
                    _ if mc_fallback_depth > 0 => {} // Skip everything inside mc:Fallback
                    _ if txbx_content_depth > 0 => {} // Skip everything inside w:txbxContent
                    b"w:pPr" => in_ppr = true,
                    b"w:rPr" => in_rpr = true,
                    b"w:r" => {
                        in_run = true;
                        current_style = TextStyle::default();
                    }
                    b"w:t" => in_text = true,
                    b"w:instrText" => in_instr_text = true,
                    b"w:drawing" => {
                        in_drawing = true;
                        current_image_alt = None;
                    }
                    b"w:pict" | b"w:object" => in_pict = true,
                    // Tracked changes - insertions
                    b"w:ins" => in_ins = true,
                    // Tracked changes - deletions
                    b"w:del" => in_del = true,
                    b"w:hyperlink" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"r:id" {
                                let rel_id = String::from_utf8_lossy(&attr.value);
                                if let Some(rel) = self.relationships.get(&rel_id) {
                                    current_hyperlink = Some(rel.target.clone());
                                }
                            }
                        }
                    }
                    _ => {}
                },
                Ok(quick_xml::events::Event::Empty(ref e)) => match e.name().as_ref() {
                    _ if mc_fallback_depth > 0 || txbx_content_depth > 0 => {} // Skip
                    b"w:pStyle" if in_ppr => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"w:val" {
                                let style_id = String::from_utf8_lossy(&attr.value);
                                para.style_id = Some(style_id.to_string());
                                para.heading = self.styles.get_heading_level(&style_id);
                                // Also get style name from StyleMap
                                if let Some(style) = self.styles.styles.get(style_id.as_ref()) {
                                    if !style.name.is_empty() {
                                        para.style_name = Some(style.name.clone());
                                    }
                                }
                            }
                        }
                    }
                    b"w:jc" if in_ppr => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"w:val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                para.alignment = match val.as_ref() {
                                    "center" => TextAlignment::Center,
                                    "right" => TextAlignment::Right,
                                    "both" | "distribute" => TextAlignment::Justify,
                                    _ => TextAlignment::Left,
                                };
                            }
                        }
                    }
                    b"w:b" if in_rpr => {
                        let val = get_bool_attr(e, b"w:val");
                        current_style.bold = val.unwrap_or(true);
                    }
                    b"w:i" if in_rpr => {
                        let val = get_bool_attr(e, b"w:val");
                        current_style.italic = val.unwrap_or(true);
                    }
                    b"w:u" if in_rpr => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"w:val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                current_style.underline = val != "none";
                            }
                        }
                    }
                    b"w:strike" if in_rpr => {
                        let val = get_bool_attr(e, b"w:val");
                        current_style.strikethrough = val.unwrap_or(true);
                    }
                    b"w:vertAlign" if in_rpr => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"w:val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                match val.as_ref() {
                                    "superscript" => current_style.superscript = true,
                                    "subscript" => current_style.subscript = true,
                                    _ => {}
                                }
                            }
                        }
                    }
                    b"w:sz" if in_rpr => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"w:val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                current_style.size = val.parse().ok();
                            }
                        }
                    }
                    b"w:color" if in_rpr => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"w:val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                if val != "auto" {
                                    current_style.color = Some(val.to_string());
                                }
                            }
                        }
                    }
                    b"w:highlight" if in_rpr => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"w:val" {
                                current_style.highlight =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    b"w:rFonts" if in_rpr => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"w:ascii" {
                                current_style.font =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                                break;
                            }
                        }
                    }
                    // Image handling: wp:docPr contains alt text
                    b"wp:docPr" if in_drawing => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"descr" {
                                current_image_alt =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    // Image handling: a:blip contains the image reference
                    b"a:blip" if in_drawing => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"r:embed" {
                                let rel_id = String::from_utf8_lossy(&attr.value).to_string();
                                // Create inline image with the relationship ID
                                let image = InlineImage {
                                    resource_id: rel_id,
                                    alt_text: current_image_alt.clone(),
                                    width: None,
                                    height: None,
                                };
                                para.images.push(image);
                            }
                        }
                    }
                    // VML image handling: v:imagedata references the image part
                    b"v:imagedata" if in_pict => {
                        if let Some(image) = vml_inline_image(e) {
                            para.images.push(image);
                        }
                    }
                    // Break handling - line break or page break
                    b"w:br" if in_run => {
                        // Check for break type: page, column, or text wrapping (default)
                        let mut is_page_break = false;
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"w:type" {
                                let break_type = String::from_utf8_lossy(&attr.value);
                                is_page_break = break_type == "page";
                            }
                        }

                        // Compute current revision type
                        let current_revision = if in_del {
                            RevisionType::Deleted
                        } else if in_ins {
                            RevisionType::Inserted
                        } else {
                            RevisionType::None
                        };

                        if is_page_break {
                            // Page break - mark run with page_break flag
                            if let Some(last_run) = para.runs.last_mut() {
                                last_run.page_break = true;
                            } else {
                                para.runs.push(TextRun {
                                    text: String::new(),
                                    style: current_style.clone(),
                                    hyperlink: None,
                                    line_break: false,
                                    page_break: true,
                                    revision: current_revision,
                                });
                            }
                        } else {
                            // Line break (text wrapping or column break treated as line break)
                            if let Some(last_run) = para.runs.last_mut() {
                                last_run.line_break = true;
                            } else {
                                para.runs.push(TextRun {
                                    text: String::new(),
                                    style: current_style.clone(),
                                    hyperlink: None,
                                    line_break: true,
                                    page_break: false,
                                    revision: current_revision,
                                });
                            }
                        }
                    }
                    // Tab character handling - convert <w:tab/> to tab character
                    b"w:tab" if in_run => {
                        let current_revision = if in_del {
                            RevisionType::Deleted
                        } else if in_ins {
                            RevisionType::Inserted
                        } else {
                            RevisionType::None
                        };
                        para.runs.push(TextRun {
                            text: "\t".to_string(),
                            style: current_style.clone(),
                            hyperlink: current_hyperlink.clone(),
                            line_break: false,
                            page_break: false,
                            revision: current_revision,
                        });
                    }
                    // Carriage return handling - convert <w:cr/> to newline
                    b"w:cr" if in_run => {
                        if let Some(last_run) = para.runs.last_mut() {
                            last_run.line_break = true;
                        } else {
                            let current_revision = if in_del {
                                RevisionType::Deleted
                            } else if in_ins {
                                RevisionType::Inserted
                            } else {
                                RevisionType::None
                            };
                            para.runs.push(TextRun {
                                text: String::new(),
                                style: current_style.clone(),
                                hyperlink: None,
                                line_break: true,
                                page_break: false,
                                revision: current_revision,
                            });
                        }
                    }
                    // Non-breaking hyphen handling
                    b"w:noBreakHyphen" if in_run => {
                        let current_revision = if in_del {
                            RevisionType::Deleted
                        } else if in_ins {
                            RevisionType::Inserted
                        } else {
                            RevisionType::None
                        };
                        para.runs.push(TextRun {
                            text: "\u{2011}".to_string(), // Non-breaking hyphen Unicode
                            style: current_style.clone(),
                            hyperlink: current_hyperlink.clone(),
                            line_break: false,
                            page_break: false,
                            revision: current_revision,
                        });
                    }
                    // Soft hyphen handling (optional hyphen, usually invisible)
                    b"w:softHyphen" if in_run => {
                        let current_revision = if in_del {
                            RevisionType::Deleted
                        } else if in_ins {
                            RevisionType::Inserted
                        } else {
                            RevisionType::None
                        };
                        para.runs.push(TextRun {
                            text: "\u{00AD}".to_string(), // Soft hyphen Unicode
                            style: current_style.clone(),
                            hyperlink: current_hyperlink.clone(),
                            line_break: false,
                            page_break: false,
                            revision: current_revision,
                        });
                    }
                    // Non-breaking space handling
                    b"w:noBreakSpace" if in_run => {
                        let current_revision = if in_del {
                            RevisionType::Deleted
                        } else if in_ins {
                            RevisionType::Inserted
                        } else {
                            RevisionType::None
                        };
                        para.runs.push(TextRun {
                            text: "\u{00A0}".to_string(), // Non-breaking space Unicode
                            style: current_style.clone(),
                            hyperlink: current_hyperlink.clone(),
                            line_break: false,
                            page_break: false,
                            revision: current_revision,
                        });
                    }
                    // Footnote reference handling
                    b"w:footnoteReference" if in_run => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"w:id" {
                                let id = String::from_utf8_lossy(&attr.value).to_string();
                                // Only insert marker if this footnote has content
                                if self.footnotes.contains_key(&id) {
                                    para.runs.push(TextRun::plain(format!("[^{}]", id)));
                                }
                            }
                        }
                    }
                    // Endnote reference handling
                    b"w:endnoteReference" if in_run => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"w:id" {
                                let id = String::from_utf8_lossy(&attr.value).to_string();
                                // Only insert marker if this endnote has content
                                if self.endnotes.contains_key(&id) {
                                    para.runs.push(TextRun::plain(format!("[^e{}]", id)));
                                }
                            }
                        }
                    }
                    _ => {}
                },
                Ok(quick_xml::events::Event::Text(ref e))
                    if in_run
                        && in_text
                        && !in_instr_text
                        && mc_fallback_depth == 0
                        && txbx_content_depth == 0 =>
                {
                    // Only extract text from w:t elements, skip w:instrText (field codes)
                    // Also skip text inside mc:Fallback and w:txbxContent (extracted separately)
                    let text = crate::decode::decode_text_lossy(e);
                    if !text.is_empty() {
                        let current_revision = if in_del {
                            RevisionType::Deleted
                        } else if in_ins {
                            RevisionType::Inserted
                        } else {
                            RevisionType::None
                        };
                        let run = TextRun {
                            text,
                            style: current_style.clone(),
                            hyperlink: current_hyperlink.clone(),
                            line_break: false,
                            page_break: false,
                            revision: current_revision,
                        };
                        para.runs.push(run);
                    }
                }
                // quick-xml 0.40+ delivers entity refs separately; mirror the Text
                // arm so a run like "AT&amp;T" keeps its ampersand. The entity
                // becomes its own run, which concatenates identically on render.
                Ok(quick_xml::events::Event::GeneralRef(ref e))
                    if in_run
                        && in_text
                        && !in_instr_text
                        && mc_fallback_depth == 0
                        && txbx_content_depth == 0 =>
                {
                    let text = crate::decode::resolve_general_ref(e);
                    if !text.is_empty() {
                        let current_revision = if in_del {
                            RevisionType::Deleted
                        } else if in_ins {
                            RevisionType::Inserted
                        } else {
                            RevisionType::None
                        };
                        let run = TextRun {
                            text,
                            style: current_style.clone(),
                            hyperlink: current_hyperlink.clone(),
                            line_break: false,
                            page_break: false,
                            revision: current_revision,
                        };
                        para.runs.push(run);
                    }
                }
                Ok(quick_xml::events::Event::End(ref e)) => match e.name().as_ref() {
                    b"mc:Fallback" if mc_fallback_depth > 0 => {
                        mc_fallback_depth -= 1;
                    }
                    b"w:txbxContent" if txbx_content_depth > 0 => {
                        txbx_content_depth -= 1;
                    }
                    _ if mc_fallback_depth > 0 || txbx_content_depth > 0 => {} // Skip
                    b"w:pPr" => in_ppr = false,
                    b"w:rPr" => in_rpr = false,
                    b"w:r" => in_run = false,
                    b"w:t" => in_text = false,
                    b"w:instrText" => in_instr_text = false,
                    b"w:hyperlink" => current_hyperlink = None,
                    b"w:drawing" => {
                        in_drawing = false;
                        current_image_alt = None;
                    }
                    b"w:pict" | b"w:object" => in_pict = false,
                    b"w:ins" => in_ins = false,
                    b"w:del" => in_del = false,
                    _ => {}
                },
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => return Err(Error::xml_parse_with_context(e.to_string(), "paragraph")),
                _ => {}
            }
            buf.clear();
        }

        // Parse numbering (list info)
        para.list_info = self.parse_list_info(xml);

        Ok(para)
    }

    /// Extract paragraphs from `w:txbxContent` elements (text boxes/shapes).
    ///
    /// Text boxes in DOCX appear inside `w:drawing` or `mc:AlternateContent` elements.
    /// This method finds all `w:txbxContent` sections (skipping those inside `mc:Fallback`
    /// to avoid duplication) and parses the inner `<w:p>` elements as regular paragraphs.
    fn extract_textbox_paragraphs(&mut self, xml: &str) -> Vec<Paragraph> {
        let mut paragraphs = Vec::new();
        let mut reader = crate::decode::reader_for(xml);
        reader.config_mut().trim_text(false);

        let mut buf = Vec::new();
        let mut mc_fallback_depth: u32 = 0;
        let mut txbx_content_depth: u32 = 0;
        let mut in_txbx_para = false;
        let mut txbx_para_xml = String::new();
        let mut txbx_para_depth: u32 = 0; // Track nested elements inside the text box <w:p>

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    let name = e.name();
                    match name.as_ref() {
                        b"mc:Fallback" => {
                            mc_fallback_depth += 1;
                        }
                        b"w:txbxContent" if mc_fallback_depth == 0 => {
                            txbx_content_depth += 1;
                        }
                        b"w:p" if txbx_content_depth > 0 && !in_txbx_para => {
                            in_txbx_para = true;
                            txbx_para_depth = 0;
                            txbx_para_xml.clear();
                            txbx_para_xml.push_str("<w:p");
                            for attr in e.attributes().flatten() {
                                txbx_para_xml.push_str(&format!(
                                    " {}=\"{}\"",
                                    String::from_utf8_lossy(attr.key.as_ref()),
                                    String::from_utf8_lossy(&attr.value)
                                ));
                            }
                            txbx_para_xml.push('>');
                        }
                        _ if in_txbx_para => {
                            txbx_para_depth += 1;
                            txbx_para_xml.push('<');
                            txbx_para_xml.push_str(&String::from_utf8_lossy(name.as_ref()));
                            for attr in e.attributes().flatten() {
                                txbx_para_xml.push_str(&format!(
                                    " {}=\"{}\"",
                                    String::from_utf8_lossy(attr.key.as_ref()),
                                    String::from_utf8_lossy(&attr.value)
                                ));
                            }
                            txbx_para_xml.push('>');
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Empty(ref e)) if in_txbx_para => {
                    let name = e.name();
                    txbx_para_xml.push('<');
                    txbx_para_xml.push_str(&String::from_utf8_lossy(name.as_ref()));
                    for attr in e.attributes().flatten() {
                        txbx_para_xml.push_str(&format!(
                            " {}=\"{}\"",
                            String::from_utf8_lossy(attr.key.as_ref()),
                            String::from_utf8_lossy(&attr.value)
                        ));
                    }
                    txbx_para_xml.push_str("/>");
                }
                Ok(quick_xml::events::Event::Text(ref e)) if in_txbx_para => {
                    let text = crate::decode::decode_text_lossy(e);
                    txbx_para_xml.push_str(&escape_xml(&text));
                }
                // Re-serialized buffer (parsed a second time): re-emit the resolved
                // entity through escape_xml so the second pass decodes it again.
                Ok(quick_xml::events::Event::GeneralRef(ref e)) if in_txbx_para => {
                    let decoded = crate::decode::resolve_general_ref(e);
                    txbx_para_xml.push_str(&escape_xml(&decoded));
                }
                Ok(quick_xml::events::Event::End(ref e)) => {
                    let name = e.name();
                    match name.as_ref() {
                        b"mc:Fallback" if mc_fallback_depth > 0 => {
                            mc_fallback_depth -= 1;
                        }
                        b"w:txbxContent" if txbx_content_depth > 0 => {
                            txbx_content_depth -= 1;
                        }
                        b"w:p" if in_txbx_para && txbx_para_depth == 0 => {
                            txbx_para_xml.push_str("</w:p>");
                            if let Ok(para) = self.parse_paragraph(&txbx_para_xml) {
                                if !para.plain_text().is_empty() {
                                    paragraphs.push(para);
                                }
                            }
                            in_txbx_para = false;
                        }
                        _ if in_txbx_para => {
                            txbx_para_depth = txbx_para_depth.saturating_sub(1);
                            txbx_para_xml.push_str("</");
                            txbx_para_xml.push_str(&String::from_utf8_lossy(name.as_ref()));
                            txbx_para_xml.push('>');
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }

        paragraphs
    }

    /// Parse list info from paragraph XML.
    fn parse_list_info(&mut self, xml: &str) -> Option<ListInfo> {
        let mut reader = crate::decode::reader_for(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut num_id: Option<String> = None;
        let mut level: u8 = 0;
        let mut in_num_pr = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) if e.name().as_ref() == b"w:numPr" => {
                    in_num_pr = true;
                }
                Ok(quick_xml::events::Event::Empty(ref e)) if in_num_pr => {
                    match e.name().as_ref() {
                        b"w:numId" => {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"w:val" {
                                    num_id = Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }
                        b"w:ilvl" => {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"w:val" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    level = val.parse().unwrap_or(0);
                                }
                            }
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::End(ref e)) if e.name().as_ref() == b"w:numPr" => {
                    in_num_pr = false;
                }
                Ok(quick_xml::events::Event::Eof) => break,
                _ => {}
            }
            buf.clear();
        }

        if let Some(ref nid) = num_id {
            if let Some((list_type, number)) = self.numbering.get_list_info(nid, level) {
                return Some(ListInfo {
                    list_type,
                    level,
                    number: if list_type == ListType::Numbered {
                        Some(number)
                    } else {
                        None
                    },
                });
            }
        }

        None
    }

    /// Parse a table element.
    #[allow(clippy::only_used_in_recursion)] // &self needed for recursive nested table parsing
    fn parse_table(&self, xml: &str) -> Result<Table> {
        use crate::model::InlineImage;

        let mut table = Table::new();
        let mut reader = crate::decode::reader_for(xml);
        // Don't trim text - preserve whitespace from xml:space="preserve" elements
        reader.config_mut().trim_text(false);

        let mut buf = Vec::new();
        let mut in_row = false;
        let mut in_cell = false;
        let mut in_paragraph = false;
        let mut in_run = false;
        let mut in_rpr = false; // Track w:rPr (run properties for formatting)
        let mut in_text = false; // Track w:t elements (regular text)
        let mut in_instr_text = false; // Track w:instrText elements (field codes to skip)
        let mut in_drawing = false; // Track w:drawing elements for images
        let mut in_pict = false; // Track w:pict/w:object elements for VML images
        let mut mc_fallback_depth: u32 = 0; // Track mc:Fallback nesting (skip VML duplicates)
        let mut current_image_alt: Option<String> = None;
        let mut current_row: Option<Row> = None;
        let mut cell_paragraphs: Vec<Paragraph> = Vec::new();
        let mut cell_nested_tables: Vec<Table> = Vec::new();
        let mut current_paragraph: Option<Paragraph> = None;
        let mut current_style = TextStyle::default();
        let mut is_header_row = false;
        let mut col_span = 1u32;
        let mut row_span = 1u32;
        let mut cell_alignment = CellAlignment::Left;
        let mut in_tc_pr = false; // Track w:tcPr (table cell properties)

        // vMerge rowspan tracking: col_cursor tracks logical column position within current row;
        // vmerge_origins maps logical_col -> (row_idx, cell_idx) of the origin cell.
        let mut col_cursor = 0usize;
        let mut vmerge_origins: HashMap<usize, (usize, usize)> = HashMap::new();

        // Track nested table depth (0 = we're at the main table level)
        // 1+ = we're inside a nested table and should collect its XML
        let mut nested_table_depth: u32 = 0;
        let mut nested_table_xml = String::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    let name = e.name();

                    // If we're inside a nested table, just collect XML
                    if nested_table_depth > 0 {
                        nested_table_xml.push('<');
                        nested_table_xml.push_str(&String::from_utf8_lossy(name.as_ref()));
                        for attr in e.attributes().flatten() {
                            nested_table_xml.push_str(&format!(
                                " {}=\"{}\"",
                                String::from_utf8_lossy(attr.key.as_ref()),
                                String::from_utf8_lossy(&attr.value)
                            ));
                        }
                        nested_table_xml.push('>');
                        if name.as_ref() == b"w:tbl" {
                            nested_table_depth += 1;
                        }
                        continue;
                    }

                    match name.as_ref() {
                        b"w:tbl" if in_cell => {
                            // Start collecting nested table
                            nested_table_depth = 1;
                            nested_table_xml.clear();
                            nested_table_xml.push_str("<w:tbl>");
                        }
                        b"w:tr" => {
                            in_row = true;
                            current_row = Some(Row {
                                cells: Vec::new(),
                                is_header: false,
                                height: None,
                            });
                            is_header_row = false;
                        }
                        b"w:tc" => {
                            in_cell = true;
                            cell_paragraphs.clear();
                            cell_nested_tables.clear();
                            col_span = 1;
                            row_span = 1;
                            cell_alignment = CellAlignment::Left;
                        }
                        b"w:tcPr" if in_cell => {
                            in_tc_pr = true;
                        }
                        b"w:p" if in_cell => {
                            in_paragraph = true;
                            current_paragraph = Some(Paragraph::new());
                        }
                        b"w:r" if in_paragraph => {
                            in_run = true;
                            current_style = TextStyle::default();
                        }
                        b"w:rPr" if in_run => in_rpr = true,
                        b"w:t" => in_text = true,
                        b"w:instrText" => in_instr_text = true,
                        b"w:drawing" => {
                            in_drawing = true;
                            current_image_alt = None;
                        }
                        b"w:pict" | b"w:object" => in_pict = true,
                        b"mc:Fallback" => mc_fallback_depth += 1,
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Empty(ref e)) => {
                    let name = e.name();

                    // If we're inside a nested table, just collect XML
                    if nested_table_depth > 0 {
                        nested_table_xml.push('<');
                        nested_table_xml.push_str(&String::from_utf8_lossy(name.as_ref()));
                        for attr in e.attributes().flatten() {
                            nested_table_xml.push_str(&format!(
                                " {}=\"{}\"",
                                String::from_utf8_lossy(attr.key.as_ref()),
                                String::from_utf8_lossy(&attr.value)
                            ));
                        }
                        nested_table_xml.push_str("/>");
                        continue;
                    }

                    match name.as_ref() {
                        b"w:tblHeader" if in_row => {
                            is_header_row = true;
                        }
                        b"w:gridSpan" if in_cell => {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"w:val" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    col_span = val.parse().unwrap_or(1);
                                }
                            }
                        }
                        b"w:vMerge" if in_cell => {
                            let mut has_val = false;
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"w:val" {
                                    has_val = true;
                                }
                            }
                            if !has_val {
                                row_span = 0;
                            }
                        }
                        b"w:jc" if in_tc_pr => {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"w:val" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    cell_alignment = match val.as_ref() {
                                        "center" => CellAlignment::Center,
                                        "right" | "end" => CellAlignment::Right,
                                        _ => CellAlignment::Left,
                                    };
                                }
                            }
                        }
                        // Handle formatting in run properties
                        b"w:b" if in_rpr => {
                            let val = get_bool_attr(e, b"w:val");
                            current_style.bold = val.unwrap_or(true);
                        }
                        b"w:i" if in_rpr => {
                            let val = get_bool_attr(e, b"w:val");
                            current_style.italic = val.unwrap_or(true);
                        }
                        b"w:u" if in_rpr => {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"w:val" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    current_style.underline = val != "none";
                                }
                            }
                        }
                        b"w:strike" if in_rpr => {
                            let val = get_bool_attr(e, b"w:val");
                            current_style.strikethrough = val.unwrap_or(true);
                        }
                        // Image handling: wp:docPr contains alt text
                        b"wp:docPr" if in_drawing => {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"descr" {
                                    current_image_alt =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }
                        // Image handling: a:blip contains the image reference
                        b"a:blip" if in_drawing => {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"r:embed" {
                                    let rel_id = String::from_utf8_lossy(&attr.value).to_string();
                                    // Create inline image with the relationship ID
                                    let image = InlineImage {
                                        resource_id: rel_id,
                                        alt_text: current_image_alt.clone(),
                                        width: None,
                                        height: None,
                                    };
                                    if let Some(ref mut para) = current_paragraph {
                                        para.images.push(image);
                                    }
                                }
                            }
                        }
                        // VML image handling: skip mc:Fallback copies, which
                        // duplicate the DrawingML mc:Choice branch
                        b"v:imagedata" if in_pict && mc_fallback_depth == 0 => {
                            if let Some(image) = vml_inline_image(e) {
                                if let Some(ref mut para) = current_paragraph {
                                    para.images.push(image);
                                }
                            }
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Text(ref e)) => {
                    // If we're inside a nested table, just collect XML
                    if nested_table_depth > 0 {
                        let text = crate::decode::decode_text_lossy(e);
                        nested_table_xml.push_str(&escape_xml(&text));
                        continue;
                    }

                    // Only extract text from w:t elements, skip w:instrText (field codes)
                    if in_run && in_text && !in_instr_text {
                        let text = crate::decode::decode_text_lossy(e);
                        if !text.is_empty() {
                            if let Some(ref mut para) = current_paragraph {
                                let run = TextRun {
                                    text,
                                    style: current_style.clone(),
                                    hyperlink: None,
                                    line_break: false,
                                    page_break: false,
                                    revision: RevisionType::None,
                                };
                                para.runs.push(run);
                            }
                        }
                    }
                }
                // quick-xml 0.40+ delivers entity refs separately; mirror both
                // branches of the Text arm (nested-table re-serialization and the
                // cell's own run text).
                Ok(quick_xml::events::Event::GeneralRef(ref e)) => {
                    if nested_table_depth > 0 {
                        let decoded = crate::decode::resolve_general_ref(e);
                        nested_table_xml.push_str(&escape_xml(&decoded));
                        continue;
                    }

                    if in_run && in_text && !in_instr_text {
                        let text = crate::decode::resolve_general_ref(e);
                        if !text.is_empty() {
                            if let Some(ref mut para) = current_paragraph {
                                let run = TextRun {
                                    text,
                                    style: current_style.clone(),
                                    hyperlink: None,
                                    line_break: false,
                                    page_break: false,
                                    revision: RevisionType::None,
                                };
                                para.runs.push(run);
                            }
                        }
                    }
                }
                Ok(quick_xml::events::Event::End(ref e)) => {
                    let name = e.name();

                    // If we're inside a nested table, collect XML and check for end
                    if nested_table_depth > 0 {
                        if name.as_ref() == b"w:tbl" {
                            nested_table_xml.push_str("</w:tbl>");
                            nested_table_depth -= 1;
                            if nested_table_depth == 0 {
                                // Finished collecting nested table - parse recursively
                                if let Ok(nested_table) = self.parse_table(&nested_table_xml) {
                                    cell_nested_tables.push(nested_table);
                                }
                            }
                        } else {
                            nested_table_xml.push_str("</");
                            nested_table_xml.push_str(&String::from_utf8_lossy(name.as_ref()));
                            nested_table_xml.push('>');
                        }
                        continue;
                    }

                    match name.as_ref() {
                        b"w:tr" => {
                            if let Some(mut row) = current_row.take() {
                                row.is_header = is_header_row;
                                table.add_row(row);
                            }
                            in_row = false;
                            col_cursor = 0;
                        }
                        b"w:tcPr" => {
                            in_tc_pr = false;
                        }
                        b"w:tc" => {
                            if row_span > 0 {
                                // Use collected paragraphs, or empty paragraph if none
                                // Deduplicate repeated paragraph blocks within cell
                                // Word may store the same paragraph block twice but only displays once
                                let content = if cell_paragraphs.is_empty() {
                                    vec![Paragraph::new()]
                                } else {
                                    let paragraphs = std::mem::take(&mut cell_paragraphs);
                                    deduplicate_paragraph_block(paragraphs)
                                };
                                let cell = Cell {
                                    content,
                                    nested_tables: std::mem::take(&mut cell_nested_tables),
                                    col_span,
                                    row_span,
                                    alignment: cell_alignment,
                                    vertical_alignment: VerticalAlignment::default(),
                                    is_header: is_header_row,
                                    background: None,
                                };
                                // Track as vMerge origin: row_idx = table.rows.len() (index
                                // the current row will have once pushed in </w:tr> handler)
                                if let Some(ref row) = current_row {
                                    vmerge_origins
                                        .insert(col_cursor, (table.rows.len(), row.cells.len()));
                                }
                                if let Some(ref mut row) = current_row {
                                    row.cells.push(cell);
                                }
                            } else {
                                // Continuation cell: find origin and increment its row_span
                                if let Some(&(origin_row_idx, origin_cell_idx)) =
                                    vmerge_origins.get(&col_cursor)
                                {
                                    if let Some(row) = table.rows.get_mut(origin_row_idx) {
                                        if let Some(cell) = row.cells.get_mut(origin_cell_idx) {
                                            cell.row_span += 1;
                                        }
                                    }
                                }
                                cell_paragraphs.clear();
                                cell_nested_tables.clear();
                            }
                            col_cursor += col_span as usize;
                            in_cell = false;
                        }
                        b"w:p" if in_cell => {
                            // Save the completed paragraph
                            if let Some(para) = current_paragraph.take() {
                                // Only add non-empty paragraphs
                                if !para.is_empty() {
                                    // Skip duplicate paragraphs (same text content as previous)
                                    // Word may store duplicate paragraphs in same cell but only displays one
                                    let is_duplicate = cell_paragraphs
                                        .last()
                                        .map(|last| last.plain_text() == para.plain_text())
                                        .unwrap_or(false);

                                    if !is_duplicate {
                                        cell_paragraphs.push(para);
                                    }
                                }
                            }
                            in_paragraph = false;
                        }
                        b"w:r" => {
                            in_run = false;
                        }
                        b"w:rPr" => in_rpr = false,
                        b"w:t" => in_text = false,
                        b"w:instrText" => in_instr_text = false,
                        b"w:drawing" => {
                            in_drawing = false;
                            current_image_alt = None;
                        }
                        b"w:pict" | b"w:object" => in_pict = false,
                        b"mc:Fallback" => mc_fallback_depth = mc_fallback_depth.saturating_sub(1),
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => return Err(Error::xml_parse_with_context(e.to_string(), "table")),
                _ => {}
            }
            buf.clear();
        }

        Ok(table)
    }

    /// Extract embedded resources (images, etc.).
    fn extract_resources(&self, doc: &mut Document) -> Result<()> {
        for (id, rel) in &self.relationships.by_id {
            if rel.rel_type.contains("/image") && !rel.external {
                let path = OoxmlContainer::resolve_path("word/document.xml", &rel.target);
                if let Ok(data) = self.container.read_binary(&path) {
                    let size = data.len();
                    let ext = std::path::Path::new(&path)
                        .extension()
                        .and_then(|e| e.to_str())
                        .unwrap_or("");
                    let resource = Resource {
                        resource_type: ResourceType::from_extension(ext),
                        filename: Some(
                            std::path::Path::new(&path)
                                .file_name()
                                .unwrap_or_default()
                                .to_string_lossy()
                                .to_string(),
                        ),
                        mime_type: guess_mime_type(&path),
                        data,
                        size,
                        width: None,
                        height: None,
                        alt_text: None,
                    };
                    doc.resources.insert(id.clone(), resource);
                }
            }
        }

        Ok(())
    }

    /// Get a reference to the container.
    pub fn container(&self) -> &OoxmlContainer {
        &self.container
    }
}

/// Parse footnotes.xml or endnotes.xml into a map of id → plain text.
///
/// `note_tag` should be `b"w:footnote"` or `b"w:endnote"`.
/// Entries with `w:type="separator"` or `w:type="continuationSeparator"` are skipped.
fn parse_notes_xml(xml: &str, note_tag: &[u8]) -> HashMap<String, String> {
    let mut notes = HashMap::new();
    let mut reader = crate::decode::reader_for(xml);
    reader.config_mut().trim_text(false);

    let mut buf = Vec::new();
    let mut current_id: Option<String> = None;
    let mut current_text = String::new();
    let mut in_note = false;
    let mut in_text = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e)) => {
                if e.name().as_ref() == note_tag {
                    let mut id = None;
                    let mut note_type = None;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"w:id" => {
                                id = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                            b"w:type" => {
                                note_type = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                            _ => {}
                        }
                    }
                    // Skip separator and continuationSeparator types
                    if let Some(ref t) = note_type {
                        if t == "separator" || t == "continuationSeparator" {
                            // Don't set in_note, so content is ignored
                            buf.clear();
                            continue;
                        }
                    }
                    if let Some(id_val) = id {
                        in_note = true;
                        current_id = Some(id_val);
                        current_text.clear();
                    }
                } else if in_note && e.name().as_ref() == b"w:t" {
                    in_text = true;
                }
            }
            Ok(quick_xml::events::Event::Text(ref e)) if in_note && in_text => {
                current_text.push_str(&crate::decode::decode_text_lossy(e));
            }
            Ok(quick_xml::events::Event::GeneralRef(ref e)) if in_note && in_text => {
                current_text.push_str(&crate::decode::resolve_general_ref(e));
            }
            Ok(quick_xml::events::Event::End(ref e)) => {
                if e.name().as_ref() == note_tag {
                    if in_note {
                        if let Some(id) = current_id.take() {
                            let trimmed = current_text.trim().to_string();
                            if !trimmed.is_empty() {
                                notes.insert(id, trimmed);
                            }
                        }
                        in_note = false;
                    }
                } else if e.name().as_ref() == b"w:t" {
                    in_text = false;
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    notes
}

/// Build an `InlineImage` from a VML `v:imagedata` element, if it carries an
/// `r:id` relationship reference. `o:title` supplies the alt text.
///
/// VML appears standalone in legacy documents (typically .doc → .docx
/// conversions) via `w:pict`, and as the visual representation of embedded
/// OLE objects via `w:object`. VML inside `mc:Fallback` duplicates the
/// DrawingML `mc:Choice` and is skipped by the callers' fallback guards.
fn vml_inline_image(e: &quick_xml::events::BytesStart) -> Option<crate::model::InlineImage> {
    let mut rel_id = None;
    let mut title = None;
    for attr in e.attributes().flatten() {
        match attr.key.as_ref() {
            b"r:id" => rel_id = Some(String::from_utf8_lossy(&attr.value).to_string()),
            b"o:title" => title = Some(String::from_utf8_lossy(&attr.value).to_string()),
            _ => {}
        }
    }
    Some(crate::model::InlineImage {
        resource_id: rel_id?,
        alt_text: title,
        width: None,
        height: None,
    })
}

/// Helper to get a boolean attribute value.
fn get_bool_attr(e: &quick_xml::events::BytesStart, key: &[u8]) -> Option<bool> {
    for attr in e.attributes().flatten() {
        if attr.key.as_ref() == key {
            let val = String::from_utf8_lossy(&attr.value);
            return Some(val != "0" && val != "false");
        }
    }
    None
}

/// Append a re-serialized start tag (`<name attr="val">`) to a buffer that will
/// be parsed a second time. Text/attribute values are already-decoded, so they
/// are re-escaped here for the re-parse.
fn push_start_tag(buf: &mut String, e: &quick_xml::events::BytesStart) {
    buf.push('<');
    buf.push_str(&String::from_utf8_lossy(e.name().as_ref()));
    for attr in e.attributes().flatten() {
        buf.push_str(&format!(
            " {}=\"{}\"",
            String::from_utf8_lossy(attr.key.as_ref()),
            String::from_utf8_lossy(&attr.value)
        ));
    }
    buf.push('>');
}

/// Append a re-serialized empty-element tag (`<name attr="val"/>`) to a buffer.
fn push_empty_tag(buf: &mut String, e: &quick_xml::events::BytesStart) {
    buf.push('<');
    buf.push_str(&String::from_utf8_lossy(e.name().as_ref()));
    for attr in e.attributes().flatten() {
        buf.push_str(&format!(
            " {}=\"{}\"",
            String::from_utf8_lossy(attr.key.as_ref()),
            String::from_utf8_lossy(&attr.value)
        ));
    }
    buf.push_str("/>");
}

/// Append a re-serialized end tag (`</name>`) to a buffer.
fn push_end_tag(buf: &mut String, name: &[u8]) {
    buf.push_str("</");
    buf.push_str(&String::from_utf8_lossy(name));
    buf.push('>');
}

/// Escape XML special characters.
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

/// Guess MIME type from file extension.
fn guess_mime_type(path: &str) -> Option<String> {
    let ext = std::path::Path::new(path)
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_lowercase())?;

    Some(
        match ext.as_str() {
            "png" => "image/png",
            "jpg" | "jpeg" => "image/jpeg",
            "gif" => "image/gif",
            "bmp" => "image/bmp",
            "tiff" | "tif" => "image/tiff",
            "svg" => "image/svg+xml",
            "emf" => "image/x-emf",
            "wmf" => "image/x-wmf",
            _ => return None,
        }
        .to_string(),
    )
}

/// Deduplicate repeated paragraph blocks within a table cell.
/// Word may store the same paragraph block twice but only displays one.
/// This function checks if the first half and second half are identical
/// and returns only the first half if so.
fn deduplicate_paragraph_block(paragraphs: Vec<Paragraph>) -> Vec<Paragraph> {
    let len = paragraphs.len();
    if len < 2 {
        return paragraphs;
    }

    // Check if paragraphs form a duplicated block (first half == second half)
    if len.is_multiple_of(2) {
        let half = len / 2;
        let first_half = &paragraphs[..half];
        let second_half = &paragraphs[half..];

        let is_duplicate = first_half
            .iter()
            .zip(second_half.iter())
            .all(|(a, b)| a.plain_text() == b.plain_text());

        if is_duplicate {
            return paragraphs.into_iter().take(half).collect();
        }
    }

    paragraphs
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_open_docx() {
        let path = "test-files/file-sample_1MB.docx";
        if std::path::Path::new(path).exists() {
            let parser = DocxParser::open(path);
            assert!(parser.is_ok());
        }
    }

    #[test]
    fn test_parse_docx() {
        let path = "test-files/file-sample_1MB.docx";
        if std::path::Path::new(path).exists() {
            let mut parser = DocxParser::open(path).unwrap();
            let doc = parser.parse().unwrap();

            assert!(!doc.sections.is_empty());

            let text = doc.plain_text();
            assert!(!text.is_empty());
            assert!(text.contains("Lorem ipsum"));
        }
    }

    #[test]
    fn test_parse_headings() {
        let path = "test-files/file-sample_1MB.docx";
        if std::path::Path::new(path).exists() {
            let mut parser = DocxParser::open(path).unwrap();
            let doc = parser.parse().unwrap();

            let headings: Vec<_> = doc.sections[0]
                .content
                .iter()
                .filter_map(|block| {
                    if let Block::Paragraph(p) = block {
                        if p.is_heading() {
                            return Some(p);
                        }
                    }
                    None
                })
                .collect();

            assert!(!headings.is_empty());
        }
    }

    #[test]
    fn test_extract_resources() {
        let path = "test-files/file-sample_1MB.docx";
        if std::path::Path::new(path).exists() {
            let mut parser = DocxParser::open(path).unwrap();
            let doc = parser.parse().unwrap();

            if !doc.resources.is_empty() {
                let resource = doc.resources.values().next().unwrap();
                assert!(resource.is_image());
            }
        }
    }

    // =========================================================================
    // Whitespace Preservation Tests (GitHub Issue #2)
    // =========================================================================

    #[test]
    fn test_whitespace_preserved_between_runs() {
        // Test case from GitHub Issue #2: "DATE OF BIRTH" was becoming "DATEOFBIRTH"
        // When xml:space="preserve" is used, spaces must be preserved
        let xml = r#"<w:p>
            <w:r><w:t>DATE</w:t></w:r>
            <w:r><w:t xml:space="preserve"> </w:t></w:r>
            <w:r><w:t>OF</w:t></w:r>
            <w:r><w:t xml:space="preserve"> </w:t></w:r>
            <w:r><w:t>BIRTH</w:t></w:r>
        </w:p>"#;

        // Create a minimal container just for testing paragraph parsing
        let container = crate::container::OoxmlContainer::from_bytes(Vec::new());
        if container.is_err() {
            // Can't create empty container, skip test
            return;
        }
        let container = container.unwrap();
        let mut parser = DocxParser {
            container,
            styles: StyleMap::default(),
            numbering: NumberingMap::default(),
            relationships: crate::container::Relationships::default(),
            footnotes: HashMap::new(),
            endnotes: HashMap::new(),
        };

        let para = parser.parse_paragraph(xml).unwrap();
        let text = para.plain_text();

        // The text should have spaces preserved
        assert!(
            text.contains("DATE") && text.contains("OF") && text.contains("BIRTH"),
            "Expected 'DATE OF BIRTH' with spaces, got: '{}'",
            text
        );
        // Check that spaces are actually there
        assert!(
            text.contains(' '),
            "Expected spaces between words, got: '{}'",
            text
        );
    }

    #[test]
    fn test_whitespace_leading_trailing_preserved() {
        // Test leading/trailing whitespace with xml:space="preserve"
        let xml = r#"<w:p>
            <w:r><w:t xml:space="preserve">  Hello World  </w:t></w:r>
        </w:p>"#;

        let container = crate::container::OoxmlContainer::from_bytes(Vec::new());
        if container.is_err() {
            return;
        }
        let container = container.unwrap();
        let mut parser = DocxParser {
            container,
            styles: StyleMap::default(),
            numbering: NumberingMap::default(),
            relationships: crate::container::Relationships::default(),
            footnotes: HashMap::new(),
            endnotes: HashMap::new(),
        };

        let para = parser.parse_paragraph(xml).unwrap();
        let text = para.plain_text();

        // Leading/trailing spaces should be preserved
        assert!(
            text.starts_with("  ") || text.contains("  Hello"),
            "Expected leading spaces, got: '{}'",
            text
        );
    }

    #[test]
    fn test_tab_character_handling() {
        // Test <w:tab/> element is converted to tab character
        let xml = r#"<w:p>
            <w:r>
                <w:t>Column1</w:t>
            </w:r>
            <w:r>
                <w:tab/>
            </w:r>
            <w:r>
                <w:t>Column2</w:t>
            </w:r>
        </w:p>"#;

        let container = crate::container::OoxmlContainer::from_bytes(Vec::new());
        if container.is_err() {
            return;
        }
        let container = container.unwrap();
        let mut parser = DocxParser {
            container,
            styles: StyleMap::default(),
            numbering: NumberingMap::default(),
            relationships: crate::container::Relationships::default(),
            footnotes: HashMap::new(),
            endnotes: HashMap::new(),
        };

        let para = parser.parse_paragraph(xml).unwrap();
        let text = para.plain_text();

        assert!(
            text.contains('\t'),
            "Expected tab character between columns, got: '{}'",
            text
        );
        assert!(
            text.contains("Column1") && text.contains("Column2"),
            "Expected both column texts, got: '{}'",
            text
        );
    }

    #[test]
    fn test_multiple_spaces_preserved() {
        // Test multiple consecutive spaces are preserved
        let xml = r#"<w:p>
            <w:r><w:t xml:space="preserve">Word1     Word2</w:t></w:r>
        </w:p>"#;

        let container = crate::container::OoxmlContainer::from_bytes(Vec::new());
        if container.is_err() {
            return;
        }
        let container = container.unwrap();
        let mut parser = DocxParser {
            container,
            styles: StyleMap::default(),
            numbering: NumberingMap::default(),
            relationships: crate::container::Relationships::default(),
            footnotes: HashMap::new(),
            endnotes: HashMap::new(),
        };

        let para = parser.parse_paragraph(xml).unwrap();
        let text = para.plain_text();

        // Multiple spaces should be preserved
        assert!(
            text.contains("     "),
            "Expected 5 consecutive spaces, got: '{}'",
            text
        );
    }

    #[test]
    fn test_carriage_return_handling() {
        // Test <w:cr/> element creates line break
        let xml = r#"<w:p>
            <w:r>
                <w:t>Line1</w:t>
                <w:cr/>
                <w:t>Line2</w:t>
            </w:r>
        </w:p>"#;

        let container = crate::container::OoxmlContainer::from_bytes(Vec::new());
        if container.is_err() {
            return;
        }
        let container = container.unwrap();
        let mut parser = DocxParser {
            container,
            styles: StyleMap::default(),
            numbering: NumberingMap::default(),
            relationships: crate::container::Relationships::default(),
            footnotes: HashMap::new(),
            endnotes: HashMap::new(),
        };

        let para = parser.parse_paragraph(xml).unwrap();

        // Check that we have a line break somewhere
        let has_line_break = para.runs.iter().any(|r| r.line_break);
        assert!(has_line_break, "Expected line break from <w:cr/>");
    }

    #[test]
    fn test_non_breaking_hyphen() {
        // Test <w:noBreakHyphen/> is converted to non-breaking hyphen Unicode
        let xml = r#"<w:p>
            <w:r>
                <w:t>non</w:t>
                <w:noBreakHyphen/>
                <w:t>breaking</w:t>
            </w:r>
        </w:p>"#;

        let container = crate::container::OoxmlContainer::from_bytes(Vec::new());
        if container.is_err() {
            return;
        }
        let container = container.unwrap();
        let mut parser = DocxParser {
            container,
            styles: StyleMap::default(),
            numbering: NumberingMap::default(),
            relationships: crate::container::Relationships::default(),
            footnotes: HashMap::new(),
            endnotes: HashMap::new(),
        };

        let para = parser.parse_paragraph(xml).unwrap();
        let text = para.plain_text();

        // Should contain non-breaking hyphen (U+2011) or at least the text parts
        assert!(
            text.contains("non") && text.contains("breaking"),
            "Expected 'non' and 'breaking' text, got: '{}'",
            text
        );
        assert!(
            text.contains('\u{2011}'),
            "Expected non-breaking hyphen U+2011, got: '{}'",
            text
        );
    }

    // =========================================================================
    // Tracked Changes Tests (Revisions)
    // =========================================================================

    #[test]
    fn test_tracked_changes_insertion() {
        // Test <w:ins> element marks text as inserted
        let xml = r#"<w:p>
            <w:r><w:t>Original </w:t></w:r>
            <w:ins>
                <w:r><w:t>inserted </w:t></w:r>
            </w:ins>
            <w:r><w:t>text</w:t></w:r>
        </w:p>"#;

        let container = crate::container::OoxmlContainer::from_bytes(Vec::new());
        if container.is_err() {
            return;
        }
        let container = container.unwrap();
        let mut parser = DocxParser {
            container,
            styles: StyleMap::default(),
            numbering: NumberingMap::default(),
            relationships: crate::container::Relationships::default(),
            footnotes: HashMap::new(),
            endnotes: HashMap::new(),
        };

        let para = parser.parse_paragraph(xml).unwrap();

        // Check that we have an inserted revision
        let has_inserted = para
            .runs
            .iter()
            .any(|r| r.revision == RevisionType::Inserted);
        assert!(has_inserted, "Expected to find inserted revision");

        // The inserted text should be marked
        let inserted_text: String = para
            .runs
            .iter()
            .filter(|r| r.revision == RevisionType::Inserted)
            .map(|r| r.text.as_str())
            .collect();
        assert!(
            inserted_text.contains("inserted"),
            "Expected 'inserted' text in revision, got: '{}'",
            inserted_text
        );
    }

    #[test]
    fn test_tracked_changes_deletion() {
        // Test <w:del> element marks text as deleted
        let xml = r#"<w:p>
            <w:r><w:t>Keep this </w:t></w:r>
            <w:del>
                <w:r><w:t>deleted </w:t></w:r>
            </w:del>
            <w:r><w:t>text</w:t></w:r>
        </w:p>"#;

        let container = crate::container::OoxmlContainer::from_bytes(Vec::new());
        if container.is_err() {
            return;
        }
        let container = container.unwrap();
        let mut parser = DocxParser {
            container,
            styles: StyleMap::default(),
            numbering: NumberingMap::default(),
            relationships: crate::container::Relationships::default(),
            footnotes: HashMap::new(),
            endnotes: HashMap::new(),
        };

        let para = parser.parse_paragraph(xml).unwrap();

        // Check that we have a deleted revision
        let has_deleted = para
            .runs
            .iter()
            .any(|r| r.revision == RevisionType::Deleted);
        assert!(has_deleted, "Expected to find deleted revision");

        // The deleted text should be marked
        let deleted_text: String = para
            .runs
            .iter()
            .filter(|r| r.revision == RevisionType::Deleted)
            .map(|r| r.text.as_str())
            .collect();
        assert!(
            deleted_text.contains("deleted"),
            "Expected 'deleted' text in revision, got: '{}'",
            deleted_text
        );
    }

    // =========================================================================
    // Footnote / Endnote Tests
    // =========================================================================

    #[test]
    fn test_parse_footnotes_xml() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:footnote w:id="0" w:type="separator">
                <w:p><w:r><w:t>___</w:t></w:r></w:p>
            </w:footnote>
            <w:footnote w:id="1" w:type="continuationSeparator">
                <w:p><w:r><w:t>---</w:t></w:r></w:p>
            </w:footnote>
            <w:footnote w:id="2">
                <w:p><w:r><w:t>This is footnote two.</w:t></w:r></w:p>
            </w:footnote>
            <w:footnote w:id="3">
                <w:p><w:r><w:t>Another footnote.</w:t></w:r></w:p>
            </w:footnote>
        </w:footnotes>"#;

        let notes = parse_notes_xml(xml, b"w:footnote");

        // Separator and continuationSeparator should be skipped
        assert!(!notes.contains_key("0"), "separator should be skipped");
        assert!(
            !notes.contains_key("1"),
            "continuationSeparator should be skipped"
        );

        // Content footnotes should be parsed
        assert_eq!(notes.get("2").unwrap(), "This is footnote two.");
        assert_eq!(notes.get("3").unwrap(), "Another footnote.");
        assert_eq!(notes.len(), 2);
    }

    #[test]
    fn test_parse_endnotes_xml() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:endnote w:id="0" w:type="separator">
                <w:p><w:r><w:t>___</w:t></w:r></w:p>
            </w:endnote>
            <w:endnote w:id="1">
                <w:p><w:r><w:t>Endnote content here.</w:t></w:r></w:p>
            </w:endnote>
        </w:endnotes>"#;

        let notes = parse_notes_xml(xml, b"w:endnote");

        assert!(!notes.contains_key("0"), "separator should be skipped");
        assert_eq!(notes.get("1").unwrap(), "Endnote content here.");
        assert_eq!(notes.len(), 1);
    }

    #[test]
    fn test_footnote_multi_run_text() {
        // Footnote with multiple runs should concatenate text
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:footnote w:id="1">
                <w:p>
                    <w:r><w:t>First </w:t></w:r>
                    <w:r><w:t>second </w:t></w:r>
                    <w:r><w:t>third.</w:t></w:r>
                </w:p>
            </w:footnote>
        </w:footnotes>"#;

        let notes = parse_notes_xml(xml, b"w:footnote");
        assert_eq!(notes.get("1").unwrap(), "First second third.");
    }

    #[test]
    fn test_parse_notes_xml_preserves_raw_malformed_entity() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:footnote w:id="2">
                <w:p><w:r><w:t>Footnote &bogus; text</w:t></w:r></w:p>
            </w:footnote>
        </w:footnotes>"#;

        let notes = parse_notes_xml(xml, b"w:footnote");
        assert_eq!(notes.get("2").unwrap(), "Footnote &bogus; text");
    }

    #[test]
    fn test_footnote_reference_in_paragraph() {
        // Test that w:footnoteReference inserts [^N] marker
        let xml = r#"<w:p>
            <w:r><w:t>Some text</w:t></w:r>
            <w:r><w:footnoteReference w:id="2"/></w:r>
            <w:r><w:t> more text</w:t></w:r>
        </w:p>"#;

        let container = crate::container::OoxmlContainer::from_bytes(Vec::new());
        if container.is_err() {
            return;
        }
        let container = container.unwrap();
        let mut footnotes = HashMap::new();
        footnotes.insert("2".to_string(), "Footnote content".to_string());

        let mut parser = DocxParser {
            container,
            styles: StyleMap::default(),
            numbering: NumberingMap::default(),
            relationships: crate::container::Relationships::default(),
            footnotes,
            endnotes: HashMap::new(),
        };

        let para = parser.parse_paragraph(xml).unwrap();
        let text = para.plain_text();

        assert!(
            text.contains("[^2]"),
            "Expected footnote reference [^2], got: '{}'",
            text
        );
        assert!(
            text.contains("Some text"),
            "Expected original text, got: '{}'",
            text
        );
    }

    #[test]
    fn test_endnote_reference_in_paragraph() {
        // Test that w:endnoteReference inserts [^eN] marker
        let xml = r#"<w:p>
            <w:r><w:t>Text with endnote</w:t></w:r>
            <w:r><w:endnoteReference w:id="1"/></w:r>
        </w:p>"#;

        let container = crate::container::OoxmlContainer::from_bytes(Vec::new());
        if container.is_err() {
            return;
        }
        let container = container.unwrap();
        let mut endnotes = HashMap::new();
        endnotes.insert("1".to_string(), "Endnote content".to_string());

        let mut parser = DocxParser {
            container,
            styles: StyleMap::default(),
            numbering: NumberingMap::default(),
            relationships: crate::container::Relationships::default(),
            footnotes: HashMap::new(),
            endnotes,
        };

        let para = parser.parse_paragraph(xml).unwrap();
        let text = para.plain_text();

        assert!(
            text.contains("[^e1]"),
            "Expected endnote reference [^e1], got: '{}'",
            text
        );
    }

    #[test]
    fn test_footnote_reference_skipped_when_no_content() {
        // If the footnote id doesn't exist in the map, no marker should be inserted
        let xml = r#"<w:p>
            <w:r><w:t>Text</w:t></w:r>
            <w:r><w:footnoteReference w:id="99"/></w:r>
        </w:p>"#;

        let container = crate::container::OoxmlContainer::from_bytes(Vec::new());
        if container.is_err() {
            return;
        }
        let container = container.unwrap();
        let mut parser = DocxParser {
            container,
            styles: StyleMap::default(),
            numbering: NumberingMap::default(),
            relationships: crate::container::Relationships::default(),
            footnotes: HashMap::new(), // No footnotes
            endnotes: HashMap::new(),
        };

        let para = parser.parse_paragraph(xml).unwrap();
        let text = para.plain_text();

        assert!(
            !text.contains("[^"),
            "Should not insert marker for unknown footnote, got: '{}'",
            text
        );
    }

    #[test]
    fn test_empty_footnote_skipped() {
        // Footnotes with only whitespace should not be included
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:footnote w:id="1">
                <w:p><w:r><w:t>   </w:t></w:r></w:p>
            </w:footnote>
            <w:footnote w:id="2">
                <w:p><w:r><w:t>Real content.</w:t></w:r></w:p>
            </w:footnote>
        </w:footnotes>"#;

        let notes = parse_notes_xml(xml, b"w:footnote");
        assert!(
            !notes.contains_key("1"),
            "Whitespace-only note should be skipped"
        );
        assert_eq!(notes.get("2").unwrap(), "Real content.");
    }

    #[test]
    fn test_parse_header_footer_xml_basic() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:p>
                <w:r><w:t>Company Name</w:t></w:r>
            </w:p>
            <w:p>
                <w:r><w:t>Confidential</w:t></w:r>
            </w:p>
        </w:hdr>"#;

        let paragraphs = empty_test_parser().parse_header_footer_xml(xml);
        assert_eq!(paragraphs.len(), 2);
        assert_eq!(paragraphs[0].plain_text(), "Company Name");
        assert_eq!(paragraphs[1].plain_text(), "Confidential");
    }

    #[test]
    fn test_parse_header_footer_xml_empty_paragraphs_skipped() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:p></w:p>
            <w:p>
                <w:r><w:t>Page 1</w:t></w:r>
            </w:p>
            <w:p>
                <w:r><w:t>   </w:t></w:r>
            </w:p>
        </w:ftr>"#;

        let paragraphs = empty_test_parser().parse_header_footer_xml(xml);
        assert_eq!(paragraphs.len(), 1);
        assert_eq!(paragraphs[0].plain_text(), "Page 1");
    }

    #[test]
    fn test_parse_header_footer_xml_multiple_runs() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:p>
                <w:r><w:t>Draft - </w:t></w:r>
                <w:r><w:t>Do Not Distribute</w:t></w:r>
            </w:p>
        </w:hdr>"#;

        let paragraphs = empty_test_parser().parse_header_footer_xml(xml);
        assert_eq!(paragraphs.len(), 1);
        assert_eq!(paragraphs[0].plain_text(), "Draft - Do Not Distribute");
    }

    #[test]
    fn test_parse_header_footer_xml_preserves_raw_malformed_entity() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:p><w:r><w:t>Footer &bogus; text</w:t></w:r></w:p>
        </w:ftr>"#;

        let paragraphs = empty_test_parser().parse_header_footer_xml(xml);
        assert_eq!(paragraphs.len(), 1);
        assert_eq!(paragraphs[0].plain_text(), "Footer &bogus; text");
    }

    #[test]
    fn test_parse_header_footer_xml_empty() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        </w:hdr>"#;

        let paragraphs = empty_test_parser().parse_header_footer_xml(xml);
        assert!(paragraphs.is_empty());
    }

    // =========================================================================
    // Text Box Content Extraction Tests (w:txbxContent)
    // =========================================================================

    /// Helper to create a minimal DOCX in memory with given document.xml content.
    fn create_minimal_docx(document_xml: &str) -> Vec<u8> {
        create_minimal_docx_with_document_rels(
            document_xml,
            Some(
                r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#,
            ),
        )
    }

    fn create_minimal_docx_with_document_rels(
        document_xml: &str,
        document_rels_xml: Option<&str>,
    ) -> Vec<u8> {
        use std::io::{Cursor, Write};
        let buf = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(buf);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);

        // [Content_Types].xml
        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#).unwrap();

        // _rels/.rels
        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#).unwrap();

        if let Some(document_rels_xml) = document_rels_xml {
            // word/_rels/document.xml.rels
            zip.start_file("word/_rels/document.xml.rels", options)
                .unwrap();
            zip.write_all(document_rels_xml.as_bytes()).unwrap();
        }

        // word/document.xml
        zip.start_file("word/document.xml", options).unwrap();
        zip.write_all(document_xml.as_bytes()).unwrap();

        zip.finish().unwrap().into_inner()
    }

    /// Build a minimal DOCX with a single extra optional part containing
    /// CP-1252 bytes (e.g. raw `Café`), which is valid XML syntactically but
    /// not valid UTF-8 / UTF-16. Used to verify that `read_xml_optional`
    /// surfaces `Error::Encoding` rather than silently treating the part as
    /// absent.
    fn create_minimal_docx_with_malformed_optional_part(extra_part_path: &str) -> Vec<u8> {
        use std::io::{Cursor, Write};
        let buf = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(buf);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);

        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#).unwrap();

        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#).unwrap();

        zip.start_file("word/document.xml", options).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p/></w:body>
</w:document>"#,
        )
        .unwrap();

        zip.start_file(extra_part_path, options).unwrap();
        zip.write_all(b"<?xml version=\"1.0\"?><root>Caf\xe9</root>")
            .unwrap();

        zip.finish().unwrap().into_inner()
    }

    #[test]
    fn test_docx_non_utf8_optional_parts_surface_encoding_error() {
        // Optional DOCX parts with malformed (non-UTF-8/UTF-16) byte content
        // must surface Error::Encoding instead of being silently dropped as
        // if the part were absent.
        for part_path in &[
            "word/styles.xml",
            "word/numbering.xml",
            "word/footnotes.xml",
            "word/endnotes.xml",
        ] {
            let data = create_minimal_docx_with_malformed_optional_part(part_path);
            let err = match DocxParser::from_bytes(data) {
                Ok(_) => panic!("malformed {part_path} must surface Error::Encoding"),
                Err(err) => err,
            };
            assert!(
                matches!(err, Error::Encoding(_)),
                "expected Error::Encoding for {part_path}, got {err:?}"
            );
        }
    }

    fn empty_test_parser() -> DocxParser {
        DocxParser::from_bytes(create_minimal_docx(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p/></w:body>
</w:document>"#,
        ))
        .unwrap()
    }

    #[test]
    fn test_textbox_content_extracted() {
        // Text box via w:drawing > wps:txbx > w:txbxContent
        let doc_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Normal paragraph</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:drawing>
          <wps:wsp>
            <wps:txbx>
              <w:txbxContent>
                <w:p>
                  <w:r><w:t>Text box content here</w:t></w:r>
                </w:p>
              </w:txbxContent>
            </wps:txbx>
          </wps:wsp>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let data = create_minimal_docx(doc_xml);
        let mut parser = DocxParser::from_bytes(data).unwrap();
        let doc = parser.parse().unwrap();
        let text = doc.plain_text();

        assert!(
            text.contains("Normal paragraph"),
            "Should contain normal paragraph text"
        );
        assert!(
            text.contains("Text box content here"),
            "Should contain text box content, got: {}",
            text
        );
    }

    // A table nested inside a text box (which lives inside a paragraph) must not
    // be treated as a body-level table. Its text is extracted via the text-box
    // path; entering table mode for it left an empty table shell (all cells
    // blank, since the cell text was captured separately) as a spurious block.
    #[test]
    fn test_textbox_with_table_does_not_emit_empty_body_table() {
        let doc_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
  <w:body>
    <w:p>
      <w:r>
        <w:drawing>
          <wps:wsp>
            <wps:txbx>
              <w:txbxContent>
                <w:tbl>
                  <w:tr><w:tc><w:p><w:r><w:t>TextboxTableCell</w:t></w:r></w:p></w:tc></w:tr>
                </w:tbl>
              </w:txbxContent>
            </wps:txbx>
          </wps:wsp>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let data = create_minimal_docx(doc_xml);
        let mut parser = DocxParser::from_bytes(data).unwrap();
        let doc = parser.parse().unwrap();

        assert!(
            doc.plain_text().contains("TextboxTableCell"),
            "text box table text lost, got: {:?}",
            doc.plain_text()
        );
        let table_blocks = doc.sections[0]
            .content
            .iter()
            .filter(|b| matches!(b, crate::model::Block::Table(_)))
            .count();
        assert_eq!(
            table_blocks, 0,
            "text-box table must not surface as a body table block"
        );
    }

    #[test]
    fn test_docx_allows_missing_document_relationships_when_unused() {
        let doc_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body>
</w:document>"#;

        let data = create_minimal_docx_with_document_rels(doc_xml, None);
        let mut parser = DocxParser::from_bytes(data)
            .expect("missing document relationships should be optional");
        let doc = parser.parse().unwrap();

        assert_eq!(doc.plain_text(), "Hello");
    }

    #[test]
    fn test_docx_rejects_malformed_document_relationships() {
        let doc_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body>
</w:document>"#;

        let data = create_minimal_docx_with_document_rels(doc_xml, Some("<Relationships"));
        let err = DocxParser::from_bytes(data)
            .err()
            .expect("malformed document relationships should fail");

        match err {
            Error::XmlParseWithContext { location, .. } => {
                assert_eq!(location, "word/_rels/document.xml.rels")
            }
            other => panic!("expected malformed document rels error, got {other:?}"),
        }
    }

    #[test]
    fn test_docx_body_malformed_entity_preserves_raw_text() {
        let doc_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>Hello &bogus; body</w:t></w:r></w:p></w:body>
</w:document>"#;

        let data = create_minimal_docx(doc_xml);
        let mut parser = DocxParser::from_bytes(data).unwrap();
        let doc = parser.parse().unwrap();

        assert_eq!(doc.plain_text(), "Hello &bogus; body");
    }

    #[test]
    fn test_docx_textbox_malformed_entity_preserves_raw_text() {
        let doc_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
  <w:body>
    <w:p>
      <w:r>
        <w:drawing>
          <wps:wsp>
            <wps:txbx>
              <w:txbxContent>
                <w:p><w:r><w:t>Box &bogus; text</w:t></w:r></w:p>
              </w:txbxContent>
            </wps:txbx>
          </wps:wsp>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let data = create_minimal_docx(doc_xml);
        let mut parser = DocxParser::from_bytes(data).unwrap();
        let doc = parser.parse().unwrap();

        assert!(doc.plain_text().contains("Box &bogus; text"));
    }

    #[test]
    fn test_docx_nested_table_malformed_entity_preserves_raw_text() {
        let parser = empty_test_parser();
        let xml = r#"<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:tr>
    <w:tc>
      <w:tbl>
        <w:tr>
          <w:tc>
            <w:p><w:r><w:t>Inner &bogus; table</w:t></w:r></w:p>
          </w:tc>
        </w:tr>
      </w:tbl>
    </w:tc>
  </w:tr>
</w:tbl>"#;

        let table = parser.parse_table(xml).unwrap();

        assert_eq!(
            table.rows[0].cells[0].nested_tables[0].rows[0].cells[0].plain_text(),
            "Inner &bogus; table"
        );
    }

    #[test]
    fn test_textbox_mc_alternate_content_no_duplication() {
        // mc:AlternateContent with text box in both Choice and Fallback
        // Should only extract once (from Choice branch)
        let doc_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            xmlns:v="urn:schemas-microsoft-com:vml">
  <w:body>
    <w:p>
      <w:r>
        <mc:AlternateContent>
          <mc:Choice>
            <w:drawing>
              <wps:wsp>
                <wps:txbx>
                  <w:txbxContent>
                    <w:p>
                      <w:r><w:t>Unique text box</w:t></w:r>
                    </w:p>
                  </w:txbxContent>
                </wps:txbx>
              </wps:wsp>
            </w:drawing>
          </mc:Choice>
          <mc:Fallback>
            <w:pict>
              <v:shape>
                <v:textbox>
                  <w:txbxContent>
                    <w:p>
                      <w:r><w:t>Unique text box</w:t></w:r>
                    </w:p>
                  </w:txbxContent>
                </v:textbox>
              </v:shape>
            </w:pict>
          </mc:Fallback>
        </mc:AlternateContent>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let data = create_minimal_docx(doc_xml);
        let mut parser = DocxParser::from_bytes(data).unwrap();
        let doc = parser.parse().unwrap();
        let text = doc.plain_text();

        // Count occurrences - should appear exactly once
        let count = text.matches("Unique text box").count();
        assert_eq!(
            count, 1,
            "Text box content should appear exactly once, not duplicated. Full text: {}",
            text
        );
    }

    /// First paragraph's inline images, for VML extraction assertions.
    fn first_paragraph_images(doc: &Document) -> &[crate::model::InlineImage] {
        for block in &doc.sections[0].content {
            if let Block::Paragraph(para) = block {
                return &para.images;
            }
        }
        panic!("no paragraph found in first section");
    }

    #[test]
    fn test_vml_pict_image_extracted() {
        // Legacy documents (typically .doc → .docx conversions) embed images
        // as standalone VML: w:pict > v:shape > v:imagedata[@r:id].
        let doc_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:pict>
          <v:shape style="width:100pt;height:50pt">
            <v:imagedata r:id="rId5" o:title="legacy image"/>
          </v:shape>
        </w:pict>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let data = create_minimal_docx(doc_xml);
        let mut parser = DocxParser::from_bytes(data).unwrap();
        let doc = parser.parse().unwrap();

        let images = first_paragraph_images(&doc);
        assert_eq!(images.len(), 1, "VML image not extracted");
        assert_eq!(images[0].resource_id, "rId5");
        assert_eq!(images[0].alt_text.as_deref(), Some("legacy image"));
    }

    #[test]
    fn test_vml_object_image_extracted() {
        // Embedded OLE objects (w:object) carry their visual as v:imagedata.
        let doc_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:object>
          <v:shape><v:imagedata r:id="rId7"/></v:shape>
        </w:object>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let data = create_minimal_docx(doc_xml);
        let mut parser = DocxParser::from_bytes(data).unwrap();
        let doc = parser.parse().unwrap();

        let images = first_paragraph_images(&doc);
        assert_eq!(images.len(), 1, "w:object VML image not extracted");
        assert_eq!(images[0].resource_id, "rId7");
    }

    #[test]
    fn test_vml_fallback_image_not_duplicated() {
        // mc:AlternateContent pairs a DrawingML Choice with a VML Fallback
        // for the same picture — only the Choice branch must produce an image.
        let doc_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <mc:AlternateContent>
          <mc:Choice>
            <w:drawing>
              <a:blip r:embed="rId3"/>
            </w:drawing>
          </mc:Choice>
          <mc:Fallback>
            <w:pict>
              <v:shape><v:imagedata r:id="rId3"/></v:shape>
            </w:pict>
          </mc:Fallback>
        </mc:AlternateContent>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let data = create_minimal_docx(doc_xml);
        let mut parser = DocxParser::from_bytes(data).unwrap();
        let doc = parser.parse().unwrap();

        let images = first_paragraph_images(&doc);
        assert_eq!(
            images.len(),
            1,
            "fallback VML must not duplicate the Choice image"
        );
        assert_eq!(images[0].resource_id, "rId3");
    }

    #[test]
    fn test_vml_pict_image_in_table_cell() {
        let doc_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:p>
            <w:r>
              <w:pict>
                <v:shape><v:imagedata r:id="rId9"/></v:shape>
              </w:pict>
            </w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"#;

        let data = create_minimal_docx(doc_xml);
        let mut parser = DocxParser::from_bytes(data).unwrap();
        let doc = parser.parse().unwrap();

        let table = doc.sections[0]
            .content
            .iter()
            .find_map(|b| match b {
                Block::Table(t) => Some(t),
                _ => None,
            })
            .expect("table missing");
        let images = &table.rows[0].cells[0].content[0].images;
        assert_eq!(images.len(), 1, "VML image in table cell not extracted");
        assert_eq!(images[0].resource_id, "rId9");
    }

    #[test]
    fn test_vml_fallback_image_not_duplicated_in_table_cell() {
        // Same AlternateContent dedup contract as the body-paragraph path:
        // the table-cell parser must not extract the Fallback VML copy.
        let doc_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:p>
            <w:r>
              <mc:AlternateContent>
                <mc:Choice>
                  <w:drawing>
                    <a:blip r:embed="rId4"/>
                  </w:drawing>
                </mc:Choice>
                <mc:Fallback>
                  <w:pict>
                    <v:shape><v:imagedata r:id="rId4"/></v:shape>
                  </w:pict>
                </mc:Fallback>
              </mc:AlternateContent>
            </w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"#;

        let data = create_minimal_docx(doc_xml);
        let mut parser = DocxParser::from_bytes(data).unwrap();
        let doc = parser.parse().unwrap();

        let table = doc.sections[0]
            .content
            .iter()
            .find_map(|b| match b {
                Block::Table(t) => Some(t),
                _ => None,
            })
            .expect("table missing");
        let images = &table.rows[0].cells[0].content[0].images;
        assert_eq!(
            images.len(),
            1,
            "fallback VML must not duplicate the Choice image in a table cell"
        );
        assert_eq!(images[0].resource_id, "rId4");
    }

    #[test]
    fn test_textbox_multiple_paragraphs() {
        // Text box with multiple paragraphs
        let doc_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
  <w:body>
    <w:p>
      <w:r>
        <w:drawing>
          <wps:wsp>
            <wps:txbx>
              <w:txbxContent>
                <w:p>
                  <w:r><w:t>First text box paragraph</w:t></w:r>
                </w:p>
                <w:p>
                  <w:r><w:t>Second text box paragraph</w:t></w:r>
                </w:p>
              </w:txbxContent>
            </wps:txbx>
          </wps:wsp>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let data = create_minimal_docx(doc_xml);
        let mut parser = DocxParser::from_bytes(data).unwrap();
        let doc = parser.parse().unwrap();
        let text = doc.plain_text();

        assert!(
            text.contains("First text box paragraph"),
            "Should contain first text box paragraph"
        );
        assert!(
            text.contains("Second text box paragraph"),
            "Should contain second text box paragraph"
        );
    }

    #[test]
    fn test_docx_chart_invalid_numeric_value_propagates_error() {
        use std::io::{Cursor, Write};
        use zip::write::SimpleFileOptions;

        let chart_xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart><c:plotArea><c:lineChart>
    <c:ser>
      <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>S</c:v></c:pt></c:strCache></c:strRef></c:tx>
      <c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt></c:strCache></c:strRef></c:cat>
      <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>not-a-number</c:v></c:pt></c:numCache></c:numRef></c:val>
    </c:ser>
  </c:lineChart></c:plotArea></c:chart>
</c:chartSpace>"#;

        let document_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <w:body>
    <w:p><w:r><w:drawing>
      <a:graphic><a:graphicData>
        <c:chart r:id="rIdChart"/>
      </a:graphicData></a:graphic>
    </w:drawing></w:r></w:p>
  </w:body>
</w:document>"#;

        let document_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>
</Relationships>"#;

        let buf = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(buf);
        let options =
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);

        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#).unwrap();

        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#).unwrap();

        zip.start_file("word/document.xml", options).unwrap();
        zip.write_all(document_xml.as_bytes()).unwrap();

        zip.start_file("word/_rels/document.xml.rels", options)
            .unwrap();
        zip.write_all(document_rels.as_bytes()).unwrap();

        zip.start_file("word/charts/chart1.xml", options).unwrap();
        zip.write_all(chart_xml.as_bytes()).unwrap();

        let data = zip.finish().unwrap().into_inner();
        let mut parser = DocxParser::from_bytes(data).unwrap();
        let err = parser
            .parse()
            .expect_err("invalid chart numeric value must surface");

        match err {
            Error::InvalidData(msg) => assert!(
                msg.contains("invalid chart numeric value"),
                "unexpected msg: {msg}"
            ),
            other => panic!("expected InvalidData, got {other:?}"),
        }
    }

    #[test]
    fn test_docx_missing_chart_part_propagates_error() {
        use std::io::{Cursor, Write};
        use zip::write::SimpleFileOptions;

        let document_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <w:body>
    <w:p><w:r><w:drawing>
      <a:graphic><a:graphicData>
        <c:chart r:id="rIdChart"/>
      </a:graphicData></a:graphic>
    </w:drawing></w:r></w:p>
  </w:body>
</w:document>"#;

        let document_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>
</Relationships>"#;

        let buf = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(buf);
        let options =
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);

        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#).unwrap();

        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#).unwrap();

        zip.start_file("word/document.xml", options).unwrap();
        zip.write_all(document_xml.as_bytes()).unwrap();

        zip.start_file("word/_rels/document.xml.rels", options)
            .unwrap();
        zip.write_all(document_rels.as_bytes()).unwrap();

        let data = zip.finish().unwrap().into_inner();
        let mut parser = DocxParser::from_bytes(data).unwrap();
        let err = parser
            .parse()
            .expect_err("missing referenced chart part must surface");

        match err {
            Error::MissingComponent(path) => assert_eq!(path, "word/charts/chart1.xml"),
            other => panic!("expected MissingComponent, got {other:?}"),
        }
    }

    #[test]
    fn test_docx_body_mixed_entities_preserve_legitimate_and_malformed() {
        use std::io::Write;

        let mut buf = Vec::new();
        {
            let cursor = std::io::Cursor::new(&mut buf);
            let mut zip = zip::ZipWriter::new(cursor);
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);

            zip.start_file("[Content_Types].xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#).unwrap();

            zip.start_file("_rels/.rels", options).unwrap();
            zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#).unwrap();

            zip.start_file("word/document.xml", options).unwrap();
            zip.write_all(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>A &amp; B &bogus; C</w:t></w:r></w:p>
  </w:body>
</w:document>"#,
            )
            .unwrap();

            zip.finish().unwrap();
        }

        let mut parser = DocxParser::from_bytes(buf).expect("parser opens");
        let doc = parser.parse().expect("document parses");
        let text = doc.plain_text();
        assert!(
            text.contains("A & B &bogus; C"),
            "expected legitimate entity decoded and malformed preserved; got {text:?}"
        );
        assert!(
            !text.contains("A &amp; B"),
            "legitimate entity must not remain escaped; got {text:?}"
        );
    }

    #[test]
    fn test_docx_ampersand_midword_renders_intact() {
        // Gate-2: under quick-xml 0.40+, "AT&amp;T" splits into runs "AT","&","T".
        // Verify the fragmentation survives full rendering to Markdown — the entity
        // run carries the same style, so it must concatenate with no inserted
        // separator. Assert on rendered Markdown, not just run existence.
        use std::io::Write;

        let mut buf = Vec::new();
        {
            let cursor = std::io::Cursor::new(&mut buf);
            let mut zip = zip::ZipWriter::new(cursor);
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);

            zip.start_file("[Content_Types].xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#).unwrap();

            zip.start_file("_rels/.rels", options).unwrap();
            zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#).unwrap();

            zip.start_file("word/document.xml", options).unwrap();
            zip.write_all(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>AT&amp;T &lt;tag&gt; R &amp; D &#48;&#x30;</w:t></w:r></w:p>
  </w:body>
</w:document>"#,
            )
            .unwrap();

            zip.finish().unwrap();
        }

        let mut parser = DocxParser::from_bytes(buf).expect("parser opens");
        let doc = parser.parse().expect("document parses");
        let md = crate::render::to_markdown(&doc, &crate::render::RenderOptions::new())
            .expect("renders to markdown");
        assert!(
            md.contains("AT&T <tag> R & D 00"),
            "entity runs must concatenate without inserted separators; got {md:?}"
        );
    }

    #[test]
    fn test_docx_core_metadata_preserves_fragmented_ampersand() {
        // The core.xml metadata loop previously assigned text per Text event, so
        // under quick-xml 0.40+ "Tom &amp; Jerry" (Text/GeneralRef/Text) kept only
        // the last fragment. Verify element-level accumulation preserves the whole.
        use std::io::Write;

        let mut buf = Vec::new();
        {
            let cursor = std::io::Cursor::new(&mut buf);
            let mut zip = zip::ZipWriter::new(cursor);
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);

            zip.start_file("[Content_Types].xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#).unwrap();

            zip.start_file("_rels/.rels", options).unwrap();
            zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#).unwrap();

            zip.start_file("docProps/core.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>Tom &amp; Jerry &lt;S1&gt;</dc:title>
  <dc:creator>A &amp; B</dc:creator>
</cp:coreProperties>"#).unwrap();

            zip.start_file("word/document.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>"#).unwrap();

            zip.finish().unwrap();
        }

        let mut parser = DocxParser::from_bytes(buf).expect("parser opens");
        let doc = parser.parse().expect("document parses");
        assert_eq!(doc.metadata.title.as_deref(), Some("Tom & Jerry <S1>"));
        assert_eq!(doc.metadata.author.as_deref(), Some("A & B"));
    }

    #[test]
    fn test_vmerge_origin_gets_correct_rowspan() {
        // BUG-2: parse_table must set row_span=2 on the origin cell of a vMerge pair.
        // The parser fix uses col_cursor + vmerge_origins to track this at parse time.
        let doc_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:tcPr><w:vMerge w:val="restart"/></w:tcPr>
          <w:p><w:r><w:t>A</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:p><w:r><w:t>B</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
      <w:tr>
        <w:tc>
          <w:tcPr><w:vMerge/></w:tcPr>
          <w:p/>
        </w:tc>
        <w:tc>
          <w:p><w:r><w:t>C</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"#;

        let data = create_minimal_docx(doc_xml);
        let mut parser = DocxParser::from_bytes(data).unwrap();
        let doc = parser.parse().unwrap();

        let table = match doc.sections[0].content.first() {
            Some(Block::Table(t)) => t,
            other => panic!("expected Block::Table, got {other:?}"),
        };

        assert_eq!(table.rows.len(), 2, "table should have 2 rows");
        // Row 0: origin cell A (row_span=2) + cell B
        assert_eq!(
            table.rows[0].cells[0].row_span, 2,
            "origin cell must have row_span=2"
        );
        // Row 1: only cell C (continuation cell excluded)
        assert_eq!(
            table.rows[1].cells.len(),
            1,
            "continuation cell must be excluded from row 1"
        );
        assert_eq!(table.rows[1].cells[0].plain_text(), "C");
    }

    #[cfg(not(target_arch = "wasm32"))]
    #[test]
    fn test_docx_streaming_emits_expected_events() {
        // FEAT-2: parse_file_streaming must work for DOCX.
        use crate::streaming::{ParseEvent, SectionStreamOptions};
        use std::ops::ControlFlow;

        let doc_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>Stream me</w:t></w:r></w:p></w:body>
</w:document>"#;

        let data = create_minimal_docx(doc_xml);
        let mut parser = DocxParser::from_bytes(data).unwrap();

        let mut event_log: Vec<&'static str> = Vec::new();
        let mut section_text = String::new();

        parser
            .for_each_section(SectionStreamOptions::default(), |event| {
                match event {
                    ParseEvent::DocumentStart { .. } => event_log.push("start"),
                    ParseEvent::SectionParsed(s) => {
                        event_log.push("section");
                        section_text = s
                            .content
                            .iter()
                            .filter_map(|b| {
                                if let Block::Paragraph(p) = b {
                                    Some(p.plain_text())
                                } else {
                                    None
                                }
                            })
                            .collect::<Vec<_>>()
                            .join(" ");
                    }
                    ParseEvent::DocumentEnd => event_log.push("end"),
                    _ => {}
                }
                ControlFlow::Continue(())
            })
            .unwrap();

        assert_eq!(event_log, ["start", "section", "end"]);
        assert!(
            section_text.contains("Stream me"),
            "section text missing, got: {section_text}"
        );
    }

    fn hf_text(paras: &[Paragraph]) -> String {
        paras
            .iter()
            .map(|p| p.plain_text())
            .collect::<Vec<_>>()
            .join("\n")
    }

    // Bug #3 (verify): text inside a table in a header must be extracted.
    #[test]
    fn header_extracts_table_cell_text() {
        let xml = r#"<?xml version="1.0"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:tbl>
    <w:tr>
      <w:tc><w:p><w:r><w:t>CellLeft</w:t></w:r></w:p></w:tc>
      <w:tc><w:p><w:r><w:t>CellRight</w:t></w:r></w:p></w:tc>
    </w:tr>
  </w:tbl>
</w:hdr>"#;
        let paras = empty_test_parser().parse_header_footer_xml(xml);
        let text = hf_text(&paras);
        assert!(text.contains("CellLeft"), "missing CellLeft, got: {text:?}");
        assert!(
            text.contains("CellRight"),
            "missing CellRight, got: {text:?}"
        );
    }

    // Bug #2 (verify): a text box nested inside a header paragraph must not
    // swallow the surrounding paragraph's own text, and its text must appear.
    #[test]
    fn header_textbox_does_not_swallow_surrounding_text() {
        let xml = r#"<?xml version="1.0"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t>AnchorBefore</w:t></w:r>
    <w:r><w:pict><v:shape><v:textbox><w:txbxContent>
      <w:p><w:r><w:t>BoxInside</w:t></w:r></w:p>
    </w:txbxContent></v:textbox></v:shape></w:pict></w:r>
    <w:r><w:t>AnchorAfter</w:t></w:r>
  </w:p>
</w:hdr>"#;
        let paras = empty_test_parser().parse_header_footer_xml(xml);
        let text = hf_text(&paras);
        assert!(
            text.contains("AnchorBefore") && text.contains("AnchorAfter"),
            "anchor text lost, got: {text:?}"
        );
        assert!(
            text.contains("BoxInside"),
            "textbox text lost, got: {text:?}"
        );
    }

    // Baseline: plain header paragraph text still works.
    #[test]
    fn header_extracts_plain_paragraph_text() {
        let xml = r#"<?xml version="1.0"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>PlainHeader</w:t></w:r></w:p>
</w:hdr>"#;
        let paras = empty_test_parser().parse_header_footer_xml(xml);
        assert_eq!(hf_text(&paras), "PlainHeader");
    }

    /// Build a minimal DOCX with a document.xml.rels and arbitrary extra parts
    /// (e.g. header/footer XML parts referenced from the body's sectPr).
    fn create_docx_with_parts(
        document_xml: &str,
        document_rels_xml: &str,
        extra_parts: &[(&str, &str)],
    ) -> Vec<u8> {
        use std::io::{Cursor, Write};
        let buf = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(buf);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);

        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#).unwrap();

        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#).unwrap();

        zip.start_file("word/_rels/document.xml.rels", options)
            .unwrap();
        zip.write_all(document_rels_xml.as_bytes()).unwrap();

        zip.start_file("word/document.xml", options).unwrap();
        zip.write_all(document_xml.as_bytes()).unwrap();

        for (path, content) in extra_parts {
            zip.start_file(*path, options).unwrap();
            zip.write_all(content.as_bytes()).unwrap();
        }

        zip.finish().unwrap().into_inner()
    }

    fn section_header_text(section: &crate::model::Section) -> String {
        section
            .header
            .as_deref()
            .unwrap_or(&[])
            .iter()
            .map(|p| p.plain_text())
            .collect::<Vec<_>>()
            .join("\n")
    }

    // Bug #1: header references of every type (default / first / even) must be
    // extracted, not just the "default" one. Previously only `w:type="default"`
    // was captured, silently dropping first-page and even-page header text.
    #[test]
    fn extracts_first_and_even_header_references() {
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Body</w:t></w:r></w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rIdHDef"/>
      <w:headerReference w:type="first" r:id="rIdHFirst"/>
      <w:headerReference w:type="even" r:id="rIdHEven"/>
    </w:sectPr>
  </w:body>
</w:document>"#;
        let rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdHDef" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rIdHFirst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>
  <Relationship Id="rIdHEven" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header3.xml"/>
</Relationships>"#;
        let hdr = |text: &str| {
            format!(
                r#"<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:hdr>"#
            )
        };
        let data = create_docx_with_parts(
            document_xml,
            rels,
            &[
                ("word/header1.xml", &hdr("DefaultHeader")),
                ("word/header2.xml", &hdr("FirstPageHeader")),
                ("word/header3.xml", &hdr("EvenPageHeader")),
            ],
        );

        let mut parser = DocxParser::from_bytes(data).unwrap();
        let doc = parser.parse().unwrap();
        let text = section_header_text(&doc.sections[0]);

        assert!(
            text.contains("DefaultHeader"),
            "missing default, got: {text:?}"
        );
        assert!(
            text.contains("FirstPageHeader"),
            "missing first-page, got: {text:?}"
        );
        assert!(
            text.contains("EvenPageHeader"),
            "missing even-page, got: {text:?}"
        );
    }

    // Bug: header/footer references inside a *section-break* sectPr (nested in a
    // paragraph's `w:pPr`) were dropped — only the final body-level `w:sectPr`
    // was read. Multi-section documents place each section's header/footer refs
    // in the paragraph-level sectPr that terminates that section, so all but the
    // last section's header/footer text was silently lost.
    #[test]
    fn extracts_header_footer_from_section_break_sectpr() {
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Section one body</w:t></w:r></w:p>
    <w:p>
      <w:pPr>
        <w:sectPr>
          <w:headerReference w:type="default" r:id="rIdH1"/>
          <w:footerReference w:type="default" r:id="rIdF1"/>
        </w:sectPr>
      </w:pPr>
    </w:p>
    <w:p><w:r><w:t>Section two body</w:t></w:r></w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rIdH2"/>
    </w:sectPr>
  </w:body>
</w:document>"#;
        let rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdH1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rIdF1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
  <Relationship Id="rIdH2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>
</Relationships>"#;
        let hdr = |tag: &str, text: &str| {
            format!(
                r#"<?xml version="1.0"?><w:{tag} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:{tag}>"#
            )
        };
        let data = create_docx_with_parts(
            document_xml,
            rels,
            &[
                ("word/header1.xml", &hdr("hdr", "SectionOneHeader")),
                ("word/footer1.xml", &hdr("ftr", "SectionOneFooter")),
                ("word/header2.xml", &hdr("hdr", "SectionTwoHeader")),
            ],
        );

        let mut parser = DocxParser::from_bytes(data).unwrap();
        let doc = parser.parse().unwrap();
        let section = &doc.sections[0];
        let header = section_header_text(section);
        let footer = section
            .footer
            .as_deref()
            .unwrap_or(&[])
            .iter()
            .map(|p| p.plain_text())
            .collect::<Vec<_>>()
            .join("\n");

        assert!(
            header.contains("SectionOneHeader"),
            "section-break header dropped, got: {header:?}"
        );
        assert!(
            header.contains("SectionTwoHeader"),
            "body-level header dropped, got: {header:?}"
        );
        assert!(
            footer.contains("SectionOneFooter"),
            "section-break footer dropped, got: {footer:?}"
        );
    }

    // A header/footer whose text lives in a *nested* table must not lose the
    // inner table's text. The flat header/footer model discards table structure
    // (by design), but every cell's text — at any nesting depth — must survive.
    #[test]
    fn extracts_nested_table_text_in_header() {
        let inner = r#"<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:tbl><w:tr><w:tc><w:p><w:r><w:t>OuterCellText</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>InnerCellText</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl></w:hdr>"#;
        let paras = empty_test_parser().parse_header_footer_xml(inner);
        let text = hf_text(&paras);
        assert!(
            text.contains("OuterCellText"),
            "outer cell text dropped, got: {text:?}"
        );
        assert!(
            text.contains("InnerCellText"),
            "nested table text dropped, got: {text:?}"
        );
    }

    // A single header referenced by several section types (default + first) must
    // be parsed once, not duplicated.
    #[test]
    fn deduplicates_shared_header_reference() {
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Body</w:t></w:r></w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rIdShared"/>
      <w:headerReference w:type="first" r:id="rIdShared"/>
    </w:sectPr>
  </w:body>
</w:document>"#;
        let rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdShared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
</Relationships>"#;
        let data = create_docx_with_parts(
            document_xml,
            rels,
            &[(
                "word/header1.xml",
                r#"<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>SharedHeader</w:t></w:r></w:p></w:hdr>"#,
            )],
        );

        let mut parser = DocxParser::from_bytes(data).unwrap();
        let doc = parser.parse().unwrap();
        let text = section_header_text(&doc.sections[0]);

        assert_eq!(
            text, "SharedHeader",
            "shared header must appear exactly once"
        );
    }

    // A header's relationship IDs live in its own rels part
    // (word/_rels/header1.xml.rels), independently numbered from the body's.
    // A hyperlink r:id must resolve against the header's rels, not document.xml's
    // — otherwise a colliding rId silently yields the wrong URL.
    #[test]
    fn header_hyperlink_resolves_against_its_own_rels() {
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Body</w:t></w:r></w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rIdHeaderPart"/>
    </w:sectPr>
  </w:body>
</w:document>"#;
        // document.xml.rels: rId1 is a DECOY that points at a body URL.
        let doc_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdHeaderPart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://body-url.example/" TargetMode="External"/>
</Relationships>"#;
        let header_xml = r#"<?xml version="1.0"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:p><w:hyperlink r:id="rId1"><w:r><w:t>HeaderLink</w:t></w:r></w:hyperlink></w:p>
</w:hdr>"#;
        // header1.xml.rels: the SAME rId1 points at the correct header URL.
        let header_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://header-url.example/" TargetMode="External"/>
</Relationships>"#;

        let data = create_docx_with_parts(
            document_xml,
            doc_rels,
            &[
                ("word/header1.xml", header_xml),
                ("word/_rels/header1.xml.rels", header_rels),
            ],
        );

        let mut parser = DocxParser::from_bytes(data).unwrap();
        let doc = parser.parse().unwrap();
        let header = doc.sections[0].header.as_ref().expect("header present");
        let link = header
            .iter()
            .flat_map(|p| p.runs.iter())
            .find_map(|r| r.hyperlink.clone())
            .expect("header hyperlink resolved");

        assert_eq!(
            link, "https://header-url.example/",
            "header hyperlink must resolve against the header's own rels"
        );
    }

    // Header/footer text is auxiliary and must not enter the document heading
    // outline, even when it carries a Heading paragraph style — otherwise it
    // would emit a stray `#` inside the rendered header blockquote and mislead
    // structure-aware consumers.
    #[test]
    fn header_paragraph_is_not_treated_as_a_heading() {
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Body</w:t></w:r></w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rIdHdr"/>
    </w:sectPr>
  </w:body>
</w:document>"#;
        let rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
</Relationships>"#;
        let styles = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="Heading 1"/>
    <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
  </w:style>
</w:styles>"#;
        let header_xml = r#"<?xml version="1.0"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>HeaderTitle</w:t></w:r></w:p>
</w:hdr>"#;
        let data = create_docx_with_parts(
            document_xml,
            rels,
            &[
                ("word/styles.xml", styles),
                ("word/header1.xml", header_xml),
            ],
        );

        let mut parser = DocxParser::from_bytes(data).unwrap();
        let doc = parser.parse().unwrap();
        let header = doc.sections[0].header.as_ref().expect("header present");

        assert_eq!(header[0].plain_text(), "HeaderTitle");
        assert!(
            !header[0].heading.is_heading(),
            "header paragraph must not be a document heading, got {:?}",
            header[0].heading
        );
    }
}
