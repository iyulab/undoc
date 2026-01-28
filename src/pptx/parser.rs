//! PPTX parser implementation.

use crate::charts;
use crate::container::OoxmlContainer;
use crate::error::{Error, Result};
use crate::model::{
    Block, Cell, Document, HeadingLevel, Metadata, Paragraph, Resource, ResourceType, Row, Section,
    Table, TextRun, TextStyle,
};
use std::collections::HashMap;
use std::path::Path;

/// Slide info from presentation.xml.
#[derive(Debug, Clone)]
struct SlideInfo {
    #[allow(dead_code)]
    id: String,
    rel_id: String,
}

/// Parser for PPTX (PowerPoint) presentations.
pub struct PptxParser {
    container: OoxmlContainer,
    slides: Vec<SlideInfo>,
    relationships: HashMap<String, String>,
}

impl PptxParser {
    /// Open a PPTX file for parsing.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
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
        // Parse presentation relationships
        let relationships = Self::parse_presentation_rels(&container)?;

        // Parse presentation for slide info
        let slides = Self::parse_presentation(&container)?;

        Ok(Self {
            container,
            slides,
            relationships,
        })
    }

    /// Parse presentation relationships.
    fn parse_presentation_rels(container: &OoxmlContainer) -> Result<HashMap<String, String>> {
        let mut rels = HashMap::new();

        if let Ok(xml) = container.read_xml("ppt/_rels/presentation.xml.rels") {
            let mut reader = quick_xml::Reader::from_str(&xml);
            reader.config_mut().trim_text(true);

            let mut buf = Vec::new();

            loop {
                match reader.read_event_into(&mut buf) {
                    Ok(quick_xml::events::Event::Empty(e))
                    | Ok(quick_xml::events::Event::Start(e)) => {
                        if e.name().as_ref() == b"Relationship" {
                            let mut id = String::new();
                            let mut target = String::new();

                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"Id" => {
                                        id = String::from_utf8_lossy(&attr.value).to_string();
                                    }
                                    b"Target" => {
                                        target = String::from_utf8_lossy(&attr.value).to_string();
                                    }
                                    _ => {}
                                }
                            }

                            if !id.is_empty() && !target.is_empty() {
                                rels.insert(id, target);
                            }
                        }
                    }
                    Ok(quick_xml::events::Event::Eof) => break,
                    Err(e) => return Err(Error::XmlParse(e.to_string())),
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(rels)
    }

    /// Parse presentation.xml for slide info.
    fn parse_presentation(container: &OoxmlContainer) -> Result<Vec<SlideInfo>> {
        let mut slides = Vec::new();

        if let Ok(xml) = container.read_xml("ppt/presentation.xml") {
            let mut reader = quick_xml::Reader::from_str(&xml);
            reader.config_mut().trim_text(true);

            let mut buf = Vec::new();

            loop {
                match reader.read_event_into(&mut buf) {
                    Ok(quick_xml::events::Event::Empty(e))
                    | Ok(quick_xml::events::Event::Start(e)) => {
                        // Look for p:sldId elements (slide references)
                        let name = e.name();
                        let local_name = name.local_name();
                        if local_name.as_ref() == b"sldId" {
                            let mut id = String::new();
                            let mut rel_id = String::new();

                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"id" => {
                                        id = String::from_utf8_lossy(&attr.value).to_string();
                                    }
                                    // r:id attribute
                                    key if key.ends_with(b"id")
                                        && key != b"id"
                                        && key.len() > 2 =>
                                    {
                                        rel_id = String::from_utf8_lossy(&attr.value).to_string();
                                    }
                                    _ => {}
                                }
                            }

                            if !rel_id.is_empty() {
                                slides.push(SlideInfo { id, rel_id });
                            }
                        }
                    }
                    Ok(quick_xml::events::Event::Eof) => break,
                    Err(e) => return Err(Error::XmlParse(e.to_string())),
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(slides)
    }

    /// Parse the presentation and return a Document model.
    pub fn parse(&mut self) -> Result<Document> {
        let mut doc = Document::new();

        // Parse metadata
        doc.metadata = self.parse_metadata()?;

        // Extract resources (images, media) and add to document
        let resources = self.extract_resources()?;
        for resource in resources {
            if let Some(ref filename) = resource.filename {
                doc.add_resource(filename.clone(), resource);
            }
        }

        // Parse each slide as a section
        for (idx, slide) in self.slides.clone().iter().enumerate() {
            let mut section = Section::new(idx);
            section.name = Some(format!("Slide {}", idx + 1));

            // Get the slide path from relationships
            if let Some(target) = self.relationships.get(&slide.rel_id) {
                let slide_path = if let Some(stripped) = target.strip_prefix('/') {
                    stripped.to_string()
                } else {
                    format!("ppt/{}", target)
                };

                // Parse slide-specific relationships for hyperlinks and images
                let slide_rels = self.parse_slide_relationships(&slide_path)?;

                if let Ok(xml) = self.container.read_xml(&slide_path) {
                    let blocks =
                        self.parse_slide_content_with_rels(&xml, &slide_rels, &slide_path)?;
                    for block in blocks {
                        section.add_block(block);
                    }
                }

                // Try to parse notes for this slide
                let notes_path = slide_path
                    .replace("slides/slide", "notesSlides/notesSlide")
                    .replace("slides\\slide", "notesSlides\\notesSlide");
                if let Ok(xml) = self.container.read_xml(&notes_path) {
                    // Parse notes relationships too
                    let notes_rels = self.parse_slide_relationships(&notes_path)?;
                    let notes = self.parse_notes_with_rels(&xml, &notes_rels)?;
                    if !notes.is_empty() {
                        // Add a separator and notes
                        section.add_block(Block::Paragraph(Paragraph::with_text("--- Notes ---")));
                        for para in notes {
                            section.add_block(Block::Paragraph(para));
                        }
                    }
                }
            }

            doc.add_section(section);
        }

        Ok(doc)
    }

    /// Parse relationships for a specific slide/notes file.
    fn parse_slide_relationships(&self, slide_path: &str) -> Result<HashMap<String, String>> {
        let mut rels = HashMap::new();

        // Convert slide path to its rels path
        // ppt/slides/slide1.xml -> ppt/slides/_rels/slide1.xml.rels
        let rels_path = if let Some(last_slash) = slide_path.rfind('/') {
            let dir = &slide_path[..last_slash];
            let file = &slide_path[last_slash + 1..];
            format!("{}/_rels/{}.rels", dir, file)
        } else {
            format!("_rels/{}.rels", slide_path)
        };

        if let Ok(xml) = self.container.read_xml(&rels_path) {
            let mut reader = quick_xml::Reader::from_str(&xml);
            reader.config_mut().trim_text(true);

            let mut buf = Vec::new();

            loop {
                match reader.read_event_into(&mut buf) {
                    Ok(quick_xml::events::Event::Empty(e))
                    | Ok(quick_xml::events::Event::Start(e)) => {
                        if e.name().as_ref() == b"Relationship" {
                            let mut id = String::new();
                            let mut target = String::new();
                            let mut rel_type = String::new();

                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"Id" => {
                                        id = String::from_utf8_lossy(&attr.value).to_string();
                                    }
                                    b"Target" => {
                                        target = String::from_utf8_lossy(&attr.value).to_string();
                                    }
                                    b"Type" => {
                                        rel_type = String::from_utf8_lossy(&attr.value).to_string();
                                    }
                                    _ => {}
                                }
                            }

                            // Store hyperlinks, image, and chart relationships
                            if !id.is_empty()
                                && !target.is_empty()
                                && (rel_type.contains("hyperlink")
                                    || rel_type.contains("image")
                                    || rel_type.contains("chart"))
                            {
                                rels.insert(id, target);
                            }
                        }
                    }
                    Ok(quick_xml::events::Event::Eof) => break,
                    Err(_) => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(rels)
    }

    /// Parse metadata from docProps/core.xml.
    fn parse_metadata(&self) -> Result<Metadata> {
        // Use shared metadata parsing from container
        let mut meta = self.container.parse_core_metadata()?;
        // Set slide count
        meta.page_count = Some(self.slides.len() as u32);
        Ok(meta)
    }

    /// Parse a slide XML into paragraphs (legacy, kept for compatibility).
    #[allow(dead_code)]
    fn parse_slide(&self, xml: &str) -> Result<Vec<Paragraph>> {
        self.parse_text_content(xml)
    }

    /// Parse slide XML into content blocks (paragraphs and tables).
    #[allow(dead_code)]
    fn parse_slide_content(&self, xml: &str) -> Result<Vec<Block>> {
        self.parse_slide_content_with_rels(xml, &HashMap::new(), "")
    }

    /// Parse slide XML into content blocks with relationship map for hyperlinks, images, and charts.
    fn parse_slide_content_with_rels(
        &self,
        xml: &str,
        rels: &HashMap<String, String>,
        slide_path: &str,
    ) -> Result<Vec<Block>> {
        let mut blocks = Vec::new();

        // Parse text content first (title, headings usually come before tables)
        let paragraphs = self.parse_text_content_excluding_tables_with_rels(xml, rels)?;
        for para in paragraphs {
            blocks.push(Block::Paragraph(para));
        }

        // Parse tables after text content
        let tables = self.parse_tables_with_rels(xml, rels)?;
        for table in tables {
            blocks.push(Block::Table(table));
        }

        // Parse charts and convert to tables for RAG-ready output
        let chart_tables = self.parse_charts(rels, slide_path)?;
        for table in chart_tables {
            blocks.push(Block::Table(table));
        }

        // Parse images (p:pic elements)
        let images = self.parse_images(xml, rels)?;
        for image in images {
            blocks.push(image);
        }

        Ok(blocks)
    }

    /// Parse images from slide XML.
    /// Images are in <p:pic> elements with <a:blip r:embed="rIdN"> referencing relationships.
    fn parse_images(&self, xml: &str, rels: &HashMap<String, String>) -> Result<Vec<Block>> {
        let mut images = Vec::new();
        let mut reader = quick_xml::Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut in_pic = false;
        let mut in_nvpicpr = false;
        let mut in_blipfill = false;
        let mut in_sppr = false;
        let mut current_name: Option<String> = None;
        let mut current_rel_id: Option<String> = None;
        let mut current_width: Option<u32> = None;
        let mut current_height: Option<u32> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    let local_name = e.name().local_name();
                    match local_name.as_ref() {
                        // p:pic - picture element
                        b"pic" => {
                            in_pic = true;
                            current_name = None;
                            current_rel_id = None;
                            current_width = None;
                            current_height = None;
                        }
                        // p:nvPicPr - non-visual picture properties (contains name)
                        b"nvPicPr" if in_pic => {
                            in_nvpicpr = true;
                        }
                        // p:cNvPr - common non-visual properties (has name attribute)
                        b"cNvPr" if in_nvpicpr => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"name" {
                                    current_name =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }
                        // p:blipFill - blip fill (contains the image reference)
                        b"blipFill" if in_pic => {
                            in_blipfill = true;
                        }
                        // a:blip - the actual image reference
                        b"blip" if in_blipfill => {
                            for attr in e.attributes().flatten() {
                                // r:embed attribute contains the relationship ID
                                if attr.key.local_name().as_ref() == b"embed" {
                                    current_rel_id =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }
                        // p:spPr - shape properties (contains size)
                        b"spPr" if in_pic => {
                            in_sppr = true;
                        }
                        // a:ext - extent (size)
                        b"ext" if in_sppr => {
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"cx" => {
                                        if let Ok(cx) =
                                            String::from_utf8_lossy(&attr.value).parse::<u32>()
                                        {
                                            current_width = Some(cx);
                                        }
                                    }
                                    b"cy" => {
                                        if let Ok(cy) =
                                            String::from_utf8_lossy(&attr.value).parse::<u32>()
                                        {
                                            current_height = Some(cy);
                                        }
                                    }
                                    _ => {}
                                }
                            }
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Empty(ref e)) => {
                    let local_name = e.name().local_name();
                    match local_name.as_ref() {
                        // Handle self-closing cNvPr
                        b"cNvPr" if in_nvpicpr => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"name" {
                                    current_name =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }
                        // Handle self-closing blip
                        b"blip" if in_blipfill => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"embed" {
                                    current_rel_id =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }
                        // Handle self-closing ext
                        b"ext" if in_sppr => {
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"cx" => {
                                        if let Ok(cx) =
                                            String::from_utf8_lossy(&attr.value).parse::<u32>()
                                        {
                                            current_width = Some(cx);
                                        }
                                    }
                                    b"cy" => {
                                        if let Ok(cy) =
                                            String::from_utf8_lossy(&attr.value).parse::<u32>()
                                        {
                                            current_height = Some(cy);
                                        }
                                    }
                                    _ => {}
                                }
                            }
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::End(ref e)) => {
                    let local_name = e.name().local_name();
                    match local_name.as_ref() {
                        b"pic" => {
                            // Create image block if we have a valid relationship
                            if let Some(rel_id) = current_rel_id.take() {
                                if let Some(target) = rels.get(&rel_id) {
                                    // Extract filename from target path (e.g., "../media/image1.png" -> "image1.png")
                                    let filename =
                                        target.rsplit('/').next().unwrap_or(target).to_string();

                                    images.push(Block::Image {
                                        resource_id: filename,
                                        alt_text: current_name.take(),
                                        width: current_width.take(),
                                        height: current_height.take(),
                                    });
                                }
                            }
                            in_pic = false;
                        }
                        b"nvPicPr" => {
                            in_nvpicpr = false;
                        }
                        b"blipFill" => {
                            in_blipfill = false;
                        }
                        b"spPr" => {
                            in_sppr = false;
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => return Err(Error::XmlParse(e.to_string())),
                _ => {}
            }
            buf.clear();
        }

        Ok(images)
    }

    /// Parse charts referenced in slide relationships and convert to tables for RAG-ready output.
    /// Chart data is extracted from ppt/charts/chartN.xml files.
    fn parse_charts(&self, rels: &HashMap<String, String>, slide_path: &str) -> Result<Vec<Table>> {
        let mut tables = Vec::new();

        // Find chart relationships (target contains "chart")
        for (_rel_id, target) in rels.iter() {
            if !target.contains("chart") {
                continue;
            }

            // Resolve chart path relative to slide
            // Relationship target is like "../charts/chart1.xml"
            let chart_path = if let Some(stripped) = target.strip_prefix("../") {
                // Relative path from slide directory
                if let Some(last_slash) = slide_path.rfind('/') {
                    let slide_dir = &slide_path[..last_slash];
                    if let Some(parent_slash) = slide_dir.rfind('/') {
                        let parent_dir = &slide_dir[..parent_slash];
                        format!("{}/{}", parent_dir, stripped)
                    } else {
                        stripped.to_string()
                    }
                } else {
                    stripped.to_string()
                }
            } else if let Some(stripped) = target.strip_prefix('/') {
                stripped.to_string()
            } else {
                format!("ppt/{}", target)
            };

            // Read and parse chart XML
            if let Ok(chart_xml) = self.container.read_xml(&chart_path) {
                match charts::parse_chart_xml(&chart_xml) {
                    Ok(chart_data) => {
                        if !chart_data.is_empty() {
                            let mut table = chart_data.to_table();
                            // Add chart title as caption if available
                            if let Some(ref title) = chart_data.title {
                                if !title.is_empty() {
                                    // Update first header cell to include chart title
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
                    Err(_) => {
                        // Chart parsing failed, skip this chart
                        // In Phase 2, we would add a warning here
                    }
                }
            }
        }

        Ok(tables)
    }

    /// Parse notes slide XML into paragraphs.
    #[allow(dead_code)]
    fn parse_notes(&self, xml: &str) -> Result<Vec<Paragraph>> {
        self.parse_notes_with_rels(xml, &HashMap::new())
    }

    /// Parse notes slide XML into paragraphs with relationship map.
    fn parse_notes_with_rels(
        &self,
        xml: &str,
        rels: &HashMap<String, String>,
    ) -> Result<Vec<Paragraph>> {
        self.parse_text_content_with_rels(xml, rels)
    }

    /// Parse all tables from slide XML.
    #[allow(dead_code)]
    fn parse_tables(&self, xml: &str) -> Result<Vec<Table>> {
        self.parse_tables_with_rels(xml, &HashMap::new())
    }

    /// Parse all tables from slide XML with relationship map for hyperlinks.
    fn parse_tables_with_rels(
        &self,
        xml: &str,
        rels: &HashMap<String, String>,
    ) -> Result<Vec<Table>> {
        let mut tables = Vec::new();
        let mut reader = quick_xml::Reader::from_str(xml);
        // Don't trim text - preserve whitespace from xml:space="preserve" elements
        reader.config_mut().trim_text(false);

        let mut buf = Vec::new();
        let mut in_table = false;
        let mut in_row = false;
        let mut in_cell = false;
        let mut in_txbody = false;
        let mut in_paragraph = false;
        let mut in_run = false;
        let mut in_text = false;
        let mut in_rpr = false;

        let mut current_table = Table::new();
        let mut current_row = Row::new();
        let mut current_cell = Cell::new();
        let mut current_paragraphs: Vec<Paragraph> = Vec::new();
        let mut current_runs: Vec<TextRun> = Vec::new();
        let mut current_text = String::new();
        let mut current_style = TextStyle::default();
        let mut current_hyperlink: Option<String> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    let local_name = e.name().local_name();
                    match local_name.as_ref() {
                        // a:tbl - table
                        b"tbl" => {
                            in_table = true;
                            current_table = Table::new();
                        }
                        // a:tr - table row
                        b"tr" if in_table => {
                            in_row = true;
                            current_row = Row::new();
                        }
                        // a:tc - table cell
                        b"tc" if in_row => {
                            in_cell = true;
                            current_cell = Cell::new();
                            current_paragraphs.clear();
                        }
                        // a:txBody - text body in cell
                        b"txBody" if in_cell => {
                            in_txbody = true;
                        }
                        // a:p - paragraph
                        b"p" if in_txbody => {
                            in_paragraph = true;
                            current_runs.clear();
                        }
                        // a:r - text run
                        b"r" if in_paragraph => {
                            in_run = true;
                            current_text.clear();
                            current_style = TextStyle::default();
                            current_hyperlink = None;
                        }
                        // a:t - text element
                        b"t" if in_run => {
                            in_text = true;
                        }
                        // a:rPr - run properties
                        b"rPr" if in_run => {
                            in_rpr = true;
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"b" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        current_style.bold = val != "0" && val != "false";
                                    }
                                    b"i" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        current_style.italic = val != "0" && val != "false";
                                    }
                                    _ => {}
                                }
                            }
                        }
                        // a:hlinkClick - hyperlink (nested in a:rPr)
                        b"hlinkClick" if in_rpr => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"id" {
                                    let rel_id = String::from_utf8_lossy(&attr.value);
                                    if let Some(url) = rels.get(rel_id.as_ref()) {
                                        current_hyperlink = Some(url.clone());
                                    }
                                }
                            }
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Empty(ref e)) => {
                    let local_name = e.name().local_name();
                    match local_name.as_ref() {
                        // Handle self-closing run properties
                        b"rPr" if in_run => {
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"b" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        current_style.bold = val != "0" && val != "false";
                                    }
                                    b"i" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        current_style.italic = val != "0" && val != "false";
                                    }
                                    _ => {}
                                }
                            }
                        }
                        // a:hlinkClick - hyperlink (self-closing)
                        b"hlinkClick" if in_run => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"id" {
                                    let rel_id = String::from_utf8_lossy(&attr.value);
                                    if let Some(url) = rels.get(rel_id.as_ref()) {
                                        current_hyperlink = Some(url.clone());
                                    }
                                }
                            }
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Text(ref e)) => {
                    if in_text {
                        let text = e.unescape().unwrap_or_default();
                        current_text.push_str(&text);
                    }
                }
                Ok(quick_xml::events::Event::End(ref e)) => {
                    let local_name = e.name().local_name();
                    match local_name.as_ref() {
                        b"t" => {
                            in_text = false;
                        }
                        b"rPr" => {
                            in_rpr = false;
                        }
                        b"r" => {
                            if !current_text.is_empty() {
                                current_runs.push(TextRun {
                                    text: current_text.clone(),
                                    style: current_style.clone(),
                                    hyperlink: current_hyperlink.clone(),
                                    line_break: false,
                                });
                            }
                            in_run = false;
                            current_hyperlink = None;
                        }
                        b"p" if in_txbody => {
                            if !current_runs.is_empty() {
                                current_paragraphs.push(Paragraph {
                                    runs: current_runs.clone(),
                                    ..Default::default()
                                });
                            }
                            in_paragraph = false;
                        }
                        b"txBody" => {
                            in_txbody = false;
                        }
                        b"tc" => {
                            current_cell.content = current_paragraphs.clone();
                            current_row.add_cell(current_cell.clone());
                            in_cell = false;
                        }
                        b"tr" => {
                            if !current_row.is_empty() {
                                // Mark first row as header
                                if current_table.is_empty() {
                                    current_row.is_header = true;
                                }
                                current_table.add_row(current_row.clone());
                            }
                            in_row = false;
                        }
                        b"tbl" => {
                            if !current_table.is_empty() {
                                tables.push(current_table.clone());
                            }
                            in_table = false;
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => return Err(Error::XmlParse(e.to_string())),
                _ => {}
            }
            buf.clear();
        }

        Ok(tables)
    }

    /// Parse text content excluding tables (paragraphs from shapes, not table cells).
    #[allow(dead_code)]
    fn parse_text_content_excluding_tables(&self, xml: &str) -> Result<Vec<Paragraph>> {
        self.parse_text_content_excluding_tables_with_rels(xml, &HashMap::new())
    }

    /// Parse text content excluding tables with relationship map for hyperlinks.
    fn parse_text_content_excluding_tables_with_rels(
        &self,
        xml: &str,
        rels: &HashMap<String, String>,
    ) -> Result<Vec<Paragraph>> {
        let mut paragraphs = Vec::new();
        let mut reader = quick_xml::Reader::from_str(xml);
        // Don't trim text - preserve whitespace from xml:space="preserve" elements
        reader.config_mut().trim_text(false);

        let mut buf = Vec::new();
        let mut in_table = false;
        let mut table_depth = 0;
        let mut in_shape = false;
        let mut in_txbody = false;
        let mut in_paragraph = false;
        let mut in_run = false;
        let mut in_text = false;
        let mut in_rpr = false;
        let mut current_runs: Vec<TextRun> = Vec::new();
        let mut current_text = String::new();
        let mut current_style = TextStyle::default();
        let mut current_hyperlink: Option<String> = None;
        let mut current_heading: HeadingLevel = HeadingLevel::None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    let local_name = e.name().local_name();
                    match local_name.as_ref() {
                        // Track table depth to skip table content
                        b"tbl" => {
                            in_table = true;
                            table_depth += 1;
                        }
                        // p:sp - shape
                        b"sp" if !in_table => {
                            in_shape = true;
                            current_heading = HeadingLevel::None;
                        }
                        // p:txBody - text body in shape
                        b"txBody" if in_shape && !in_table => {
                            in_txbody = true;
                        }
                        // a:p - paragraph (only if not in table, but in shape's txBody)
                        b"p" if !in_table && in_txbody => {
                            in_paragraph = true;
                            current_runs.clear();
                        }
                        // a:r - text run
                        b"r" if in_paragraph && !in_table => {
                            in_run = true;
                            current_text.clear();
                            current_style = TextStyle::default();
                            current_hyperlink = None;
                        }
                        // a:t - text element
                        b"t" if in_run && !in_table => {
                            in_text = true;
                        }
                        // a:rPr - run properties
                        b"rPr" if in_run && !in_table => {
                            in_rpr = true;
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"b" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        current_style.bold = val != "0" && val != "false";
                                    }
                                    b"i" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        current_style.italic = val != "0" && val != "false";
                                    }
                                    b"u" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        current_style.underline = val != "none";
                                    }
                                    b"strike" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        current_style.strikethrough =
                                            val != "noStrike" && val != "0" && val != "false";
                                    }
                                    _ => {}
                                }
                            }
                        }
                        // a:hlinkClick - hyperlink (nested in a:rPr)
                        b"hlinkClick" if in_rpr && !in_table => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"id" {
                                    let rel_id = String::from_utf8_lossy(&attr.value);
                                    if let Some(url) = rels.get(rel_id.as_ref()) {
                                        current_hyperlink = Some(url.clone());
                                    }
                                }
                            }
                        }
                        // p:ph - placeholder type (for heading detection)
                        b"ph" if in_shape && !in_table => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"type" {
                                    let ph_type = String::from_utf8_lossy(&attr.value);
                                    current_heading = match ph_type.as_ref() {
                                        "title" | "ctrTitle" => HeadingLevel::H1,
                                        "subTitle" => HeadingLevel::H2,
                                        "body" => HeadingLevel::None,
                                        _ => HeadingLevel::None,
                                    };
                                }
                            }
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Empty(ref e)) => {
                    let local_name = e.name().local_name();
                    match local_name.as_ref() {
                        // p:ph - placeholder type (self-closing)
                        b"ph" if in_shape && !in_table => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"type" {
                                    let ph_type = String::from_utf8_lossy(&attr.value);
                                    current_heading = match ph_type.as_ref() {
                                        "title" | "ctrTitle" => HeadingLevel::H1,
                                        "subTitle" => HeadingLevel::H2,
                                        "body" => HeadingLevel::None,
                                        _ => HeadingLevel::None,
                                    };
                                }
                            }
                        }
                        b"rPr" if in_run && !in_table => {
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"b" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        current_style.bold = val != "0" && val != "false";
                                    }
                                    b"i" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        current_style.italic = val != "0" && val != "false";
                                    }
                                    b"u" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        current_style.underline = val != "none";
                                    }
                                    b"strike" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        current_style.strikethrough =
                                            val != "noStrike" && val != "0" && val != "false";
                                    }
                                    _ => {}
                                }
                            }
                        }
                        // a:hlinkClick - hyperlink (self-closing)
                        b"hlinkClick" if in_run && !in_table => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"id" {
                                    let rel_id = String::from_utf8_lossy(&attr.value);
                                    if let Some(url) = rels.get(rel_id.as_ref()) {
                                        current_hyperlink = Some(url.clone());
                                    }
                                }
                            }
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Text(ref e)) => {
                    if in_text && !in_table {
                        let text = e.unescape().unwrap_or_default();
                        current_text.push_str(&text);
                    }
                }
                Ok(quick_xml::events::Event::End(ref e)) => {
                    let local_name = e.name().local_name();
                    match local_name.as_ref() {
                        b"tbl" => {
                            table_depth -= 1;
                            if table_depth == 0 {
                                in_table = false;
                            }
                        }
                        b"t" if !in_table => {
                            in_text = false;
                        }
                        b"rPr" if !in_table => {
                            in_rpr = false;
                        }
                        b"r" if !in_table => {
                            if !current_text.is_empty() {
                                current_runs.push(TextRun {
                                    text: current_text.clone(),
                                    style: current_style.clone(),
                                    hyperlink: current_hyperlink.clone(),
                                    line_break: false,
                                });
                            }
                            in_run = false;
                            current_hyperlink = None;
                        }
                        b"p" if !in_table => {
                            if !current_runs.is_empty() {
                                paragraphs.push(Paragraph {
                                    runs: current_runs.clone(),
                                    heading: current_heading,
                                    ..Default::default()
                                });
                            }
                            in_paragraph = false;
                        }
                        b"txBody" if !in_table => {
                            in_txbody = false;
                        }
                        b"sp" if !in_table => {
                            in_shape = false;
                            current_heading = HeadingLevel::None;
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => return Err(Error::XmlParse(e.to_string())),
                _ => {}
            }
            buf.clear();
        }

        Ok(paragraphs)
    }

    /// Parse text content from slide or notes XML.
    /// Text is found in: p:sp/p:txBody/a:p/a:r/a:t
    #[allow(dead_code)]
    fn parse_text_content(&self, xml: &str) -> Result<Vec<Paragraph>> {
        self.parse_text_content_with_rels(xml, &HashMap::new())
    }

    /// Parse text content from slide or notes XML with relationship map for hyperlinks.
    fn parse_text_content_with_rels(
        &self,
        xml: &str,
        rels: &HashMap<String, String>,
    ) -> Result<Vec<Paragraph>> {
        let mut paragraphs = Vec::new();
        let mut reader = quick_xml::Reader::from_str(xml);
        // Don't trim text - preserve whitespace from xml:space="preserve" elements
        reader.config_mut().trim_text(false);

        let mut buf = Vec::new();
        let mut in_paragraph = false;
        let mut in_run = false;
        let mut in_text = false;
        let mut in_rpr = false;
        let mut current_runs: Vec<TextRun> = Vec::new();
        let mut current_text = String::new();
        let mut current_style = TextStyle::default();
        let mut current_hyperlink: Option<String> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    let local_name = e.name().local_name();
                    match local_name.as_ref() {
                        // a:p - paragraph
                        b"p" => {
                            in_paragraph = true;
                            current_runs.clear();
                        }
                        // a:r - text run
                        b"r" if in_paragraph => {
                            in_run = true;
                            current_text.clear();
                            current_style = TextStyle::default();
                            current_hyperlink = None;
                        }
                        // a:t - text element
                        b"t" if in_run => {
                            in_text = true;
                        }
                        // a:rPr - run properties
                        b"rPr" if in_run => {
                            in_rpr = true;
                            // Parse run properties for styling
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"b" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        current_style.bold = val != "0" && val != "false";
                                    }
                                    b"i" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        current_style.italic = val != "0" && val != "false";
                                    }
                                    b"u" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        current_style.underline = val != "none";
                                    }
                                    b"strike" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        current_style.strikethrough =
                                            val != "noStrike" && val != "0" && val != "false";
                                    }
                                    _ => {}
                                }
                            }
                        }
                        // a:hlinkClick - hyperlink (nested in a:rPr)
                        b"hlinkClick" if in_rpr => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"id" {
                                    let rel_id = String::from_utf8_lossy(&attr.value);
                                    if let Some(url) = rels.get(rel_id.as_ref()) {
                                        current_hyperlink = Some(url.clone());
                                    }
                                }
                            }
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Empty(ref e)) => {
                    let local_name = e.name().local_name();
                    match local_name.as_ref() {
                        // Handle self-closing run properties
                        b"rPr" if in_run => {
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"b" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        current_style.bold = val != "0" && val != "false";
                                    }
                                    b"i" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        current_style.italic = val != "0" && val != "false";
                                    }
                                    b"u" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        current_style.underline = val != "none";
                                    }
                                    b"strike" => {
                                        let val = String::from_utf8_lossy(&attr.value);
                                        current_style.strikethrough =
                                            val != "noStrike" && val != "0" && val != "false";
                                    }
                                    _ => {}
                                }
                            }
                        }
                        // a:hlinkClick - hyperlink (self-closing)
                        b"hlinkClick" if in_run => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"id" {
                                    let rel_id = String::from_utf8_lossy(&attr.value);
                                    if let Some(url) = rels.get(rel_id.as_ref()) {
                                        current_hyperlink = Some(url.clone());
                                    }
                                }
                            }
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Text(ref e)) => {
                    if in_text {
                        let text = e.unescape().unwrap_or_default();
                        current_text.push_str(&text);
                    }
                }
                Ok(quick_xml::events::Event::End(ref e)) => {
                    let local_name = e.name().local_name();
                    match local_name.as_ref() {
                        b"t" => {
                            in_text = false;
                        }
                        b"rPr" => {
                            in_rpr = false;
                        }
                        b"r" => {
                            if !current_text.is_empty() {
                                current_runs.push(TextRun {
                                    text: current_text.clone(),
                                    style: current_style.clone(),
                                    hyperlink: current_hyperlink.clone(),
                                    line_break: false,
                                });
                            }
                            in_run = false;
                            current_hyperlink = None;
                        }
                        b"p" => {
                            // Only add non-empty paragraphs
                            if !current_runs.is_empty() {
                                paragraphs.push(Paragraph {
                                    runs: current_runs.clone(),
                                    ..Default::default()
                                });
                            }
                            in_paragraph = false;
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => return Err(Error::XmlParse(e.to_string())),
                _ => {}
            }
            buf.clear();
        }

        Ok(paragraphs)
    }

    /// Extract resources (images, media) from the presentation.
    pub fn extract_resources(&self) -> Result<Vec<Resource>> {
        let mut resources = Vec::new();

        // Look for media files in ppt/media/
        for file in self.container.list_files() {
            if file.starts_with("ppt/media/") {
                if let Ok(data) = self.container.read_binary(&file) {
                    let filename = file.rsplit('/').next().unwrap_or(&file).to_string();
                    let ext = std::path::Path::new(&file)
                        .extension()
                        .and_then(|e| e.to_str())
                        .unwrap_or("");
                    let size = data.len();

                    resources.push(Resource {
                        resource_type: ResourceType::from_extension(ext),
                        filename: Some(filename),
                        mime_type: guess_mime_type(&file),
                        data,
                        size,
                        width: None,
                        height: None,
                        alt_text: None,
                    });
                }
            }
        }

        Ok(resources)
    }

    /// Get a reference to the container.
    pub fn container(&self) -> &OoxmlContainer {
        &self.container
    }

    /// Get the number of slides.
    pub fn slide_count(&self) -> usize {
        self.slides.len()
    }
}

/// Guess MIME type from file extension.
fn guess_mime_type(path: &str) -> Option<String> {
    let ext = path.rsplit('.').next()?.to_lowercase();
    match ext.as_str() {
        "png" => Some("image/png".to_string()),
        "jpg" | "jpeg" => Some("image/jpeg".to_string()),
        "gif" => Some("image/gif".to_string()),
        "bmp" => Some("image/bmp".to_string()),
        "tiff" | "tif" => Some("image/tiff".to_string()),
        "webp" => Some("image/webp".to_string()),
        "svg" => Some("image/svg+xml".to_string()),
        "emf" => Some("image/x-emf".to_string()),
        "wmf" => Some("image/x-wmf".to_string()),
        "mp3" => Some("audio/mpeg".to_string()),
        "wav" => Some("audio/wav".to_string()),
        "mp4" => Some("video/mp4".to_string()),
        "avi" => Some("video/x-msvideo".to_string()),
        "wmv" => Some("video/x-ms-wmv".to_string()),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_open_pptx() {
        let path = "test-files/file_example_PPT_1MB.pptx";
        if std::path::Path::new(path).exists() {
            let parser = PptxParser::open(path);
            assert!(parser.is_ok());
        }
    }

    #[test]
    fn test_parse_pptx() {
        let path = "test-files/file_example_PPT_1MB.pptx";
        if std::path::Path::new(path).exists() {
            let mut parser = PptxParser::open(path).unwrap();
            let doc = parser.parse().unwrap();

            // Should have at least one section (slide)
            assert!(!doc.sections.is_empty());
            println!("Parsed {} slides", doc.sections.len());

            // Check metadata has slide count
            assert!(doc.metadata.page_count.is_some());
        }
    }

    #[test]
    fn test_slide_count() {
        let path = "test-files/file_example_PPT_1MB.pptx";
        if std::path::Path::new(path).exists() {
            let parser = PptxParser::open(path).unwrap();
            let count = parser.slide_count();
            assert!(count > 0);
            println!("Slide count: {}", count);
        }
    }

    #[test]
    fn test_extract_text() {
        let path = "test-files/file_example_PPT_1MB.pptx";
        if std::path::Path::new(path).exists() {
            let mut parser = PptxParser::open(path).unwrap();
            let doc = parser.parse().unwrap();
            let text = doc.plain_text();

            // Should have some text content
            assert!(!text.trim().is_empty());
            println!("Extracted text length: {} chars", text.len());
            println!("First 500 chars:\n{}", &text[..text.len().min(500)]);
        }
    }

    #[test]
    fn test_extract_resources() {
        let path = "test-files/file_example_PPT_1MB.pptx";
        if std::path::Path::new(path).exists() {
            let parser = PptxParser::open(path).unwrap();
            let resources = parser.extract_resources().unwrap();

            println!("Found {} resources", resources.len());
            for res in &resources {
                println!(
                    "  - {:?}: {} ({} bytes)",
                    res.resource_type,
                    res.filename.as_deref().unwrap_or("unnamed"),
                    res.size
                );
            }
        }
    }

    #[test]
    fn test_parse_text_content() {
        // Test XML parsing directly
        let _xml = r#"<?xml version="1.0"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:txBody>
          <a:p>
            <a:r>
              <a:t>Hello World</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:r>
              <a:rPr b="1"/>
              <a:t>Bold Text</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>"#;

        let container = OoxmlContainer::from_bytes(Vec::new());
        // Can't test fully without a real container, but we can test the parse logic
        // by creating a minimal parser
        if container.is_ok() {
            // Just verify XML parsing logic compiles
        }
    }

    #[test]
    fn test_metadata() {
        let path = "test-files/file_example_PPT_1MB.pptx";
        if std::path::Path::new(path).exists() {
            let mut parser = PptxParser::open(path).unwrap();
            let doc = parser.parse().unwrap();

            println!("Title: {:?}", doc.metadata.title);
            println!("Author: {:?}", doc.metadata.author);
            println!("Page count: {:?}", doc.metadata.page_count);
        }
    }

    #[test]
    fn test_parse_tables() {
        let path = "test-files/file_example_PPT_1MB.pptx";
        if std::path::Path::new(path).exists() {
            let mut parser = PptxParser::open(path).unwrap();
            let doc = parser.parse().unwrap();

            // Find tables in the document
            let mut table_count = 0;
            for section in &doc.sections {
                for block in &section.content {
                    if let Block::Table(table) = block {
                        table_count += 1;
                        println!(
                            "Found table in {}: {} rows, {} cols",
                            section.name.as_deref().unwrap_or("unnamed"),
                            table.row_count(),
                            table.column_count()
                        );
                        // Print table content
                        for (i, row) in table.rows.iter().enumerate() {
                            let cells: Vec<String> =
                                row.cells.iter().map(|c| c.plain_text()).collect();
                            println!("  Row {}: {:?}", i, cells);
                        }
                    }
                }
            }
            println!("Total tables found: {}", table_count);
            // The test file should have at least one table (Slide 3)
            assert!(table_count > 0, "Expected at least one table in the PPTX");
        }
    }

    #[test]
    fn test_parse_hyperlinks() {
        let path = "test-files/officedissector/test/govdocs/036279.pptx";
        if std::path::Path::new(path).exists() {
            let mut parser = PptxParser::open(path).unwrap();
            let doc = parser.parse().unwrap();

            // Find hyperlinks in the document
            let mut hyperlink_count = 0;
            let mut found_email = false;
            for section in &doc.sections {
                for block in &section.content {
                    if let Block::Paragraph(para) = block {
                        for run in &para.runs {
                            if let Some(ref link) = run.hyperlink {
                                hyperlink_count += 1;
                                println!("Found hyperlink: {} -> {}", run.text, link);
                                if link.contains("ncicb@pop.nci.nih.gov") {
                                    found_email = true;
                                }
                            }
                        }
                    }
                }
            }
            println!("Total hyperlinks found: {}", hyperlink_count);
            assert!(hyperlink_count > 0, "Expected at least one hyperlink");
            assert!(
                found_email,
                "Expected to find email link ncicb@pop.nci.nih.gov"
            );
        }
    }

    #[test]
    fn test_parse_headings() {
        use crate::model::HeadingLevel;

        let path = "test-files/file_example_PPT_1MB.pptx";
        if std::path::Path::new(path).exists() {
            let mut parser = PptxParser::open(path).unwrap();
            let doc = parser.parse().unwrap();

            // Find headings in the document
            let mut h1_count = 0;
            let mut h2_count = 0;
            let mut found_lorem = false;
            for section in &doc.sections {
                for block in &section.content {
                    if let Block::Paragraph(para) = block {
                        let text = para.plain_text();
                        match para.heading {
                            HeadingLevel::H1 => {
                                h1_count += 1;
                                println!("Found H1: {}", text);
                                if text.contains("Lorem ipsum") {
                                    found_lorem = true;
                                }
                            }
                            HeadingLevel::H2 => {
                                h2_count += 1;
                                println!("Found H2: {}", text);
                            }
                            _ => {}
                        }
                    }
                }
            }
            println!("Total H1: {}, H2: {}", h1_count, h2_count);
            assert!(h1_count > 0, "Expected at least one H1 heading (title)");
            assert!(found_lorem, "Expected to find 'Lorem ipsum' as H1 title");
        }
    }
}
