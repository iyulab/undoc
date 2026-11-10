//! PPTX parser implementation.

use super::bullets::{text_marker, InheritedBullets, Marker};
use crate::charts;
use crate::container::OoxmlContainer;
use crate::error::Result;
use crate::model::{
    Block, Cell, Document, HeadingLevel, ListInfo, ListType, Metadata, Paragraph, Resource,
    ResourceRole, Row, Section, Table, TextRun, TextStyle,
};
use std::collections::HashMap;
#[cfg(not(target_arch = "wasm32"))]
use std::path::Path;

/// Slide info from presentation.xml.
#[derive(Debug, Clone)]
pub(super) struct SlideInfo {
    #[allow(dead_code)]
    id: String,
    pub(super) rel_id: String,
}

/// Parser for PPTX (PowerPoint) presentations.
pub struct PptxParser {
    pub(super) container: OoxmlContainer,
    pub(super) slides: Vec<SlideInfo>,
    pub(super) relationships: HashMap<String, String>,
}

impl PptxParser {
    /// Open a PPTX file for parsing.
    #[cfg(not(target_arch = "wasm32"))]
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
        Ok(container
            .read_required_relationships_for_part("ppt/presentation.xml")?
            .into_targets_by_id())
    }

    /// Parse presentation.xml for slide info.
    fn parse_presentation(container: &OoxmlContainer) -> Result<Vec<SlideInfo>> {
        let mut slides = Vec::new();
        let xml = container.read_xml("ppt/presentation.xml")?;

        let mut reader = crate::decode::reader_for(&xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Empty(e)) | Ok(quick_xml::events::Event::Start(e)) => {
                    let name = e.name();
                    let local_name = name.local_name();
                    if local_name.as_ref() == "sldId" {
                        let mut id = String::new();
                        let mut rel_id = String::new();

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                "id" => {
                                    id = attr.value.to_string();
                                }
                                key if key.ends_with("id") && key != "id" && key.len() > 2 => {
                                    rel_id = attr.value.to_string();
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
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(slides)
    }

    /// Parse the presentation and return a Document model.
    pub fn parse(&mut self) -> Result<Document> {
        let mut doc = Document::new();
        doc.format = crate::detect::FormatType::Pptx;

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
            let section = self.parse_slide_as_section(idx, slide)?;
            doc.add_section(section);
        }

        Ok(doc)
    }

    /// Stream sections (slides) one at a time, calling `f` for each event.
    ///
    /// See [`crate::parse_file_streaming`] for the full API contract.
    pub fn for_each_section<F>(
        &mut self,
        opts: crate::streaming::SectionStreamOptions,
        mut f: F,
    ) -> Result<()>
    where
        F: FnMut(crate::streaming::ParseEvent<'_>) -> std::ops::ControlFlow<()>,
    {
        let metadata = self.parse_metadata()?;
        let section_count = self.slides.len();

        // Extract resources once; reuse for both image_map and ResourceExtracted.
        let resources = self.extract_resources()?;
        let image_map: std::collections::HashMap<String, String> = resources
            .iter()
            .filter_map(|r| r.filename.as_ref().map(|name| (name.clone(), name.clone())))
            .collect();

        if f(crate::streaming::ParseEvent::DocumentStart {
            metadata: &metadata,
            section_count,
            image_map,
        })
        .is_break()
        {
            return Ok(());
        }

        for (idx, slide) in self.slides.clone().iter().enumerate() {
            let section_result = self.parse_slide_as_section(idx, slide);

            match section_result {
                Ok(section) => {
                    if f(crate::streaming::ParseEvent::SectionParsed(&section)).is_break() {
                        return Ok(());
                    }
                }
                Err(e) => {
                    if opts.lenient {
                        if f(crate::streaming::ParseEvent::SectionFailed {
                            index: idx,
                            error: e,
                        })
                        .is_break()
                        {
                            return Ok(());
                        }
                    } else {
                        return Err(e);
                    }
                }
            }
        }

        if f(crate::streaming::ParseEvent::DocumentEnd).is_break() {
            return Ok(());
        }

        if opts.extract_resources {
            for resource in resources {
                let name = resource
                    .filename
                    .unwrap_or_else(|| "resource.bin".to_string());
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

    /// Parse a single slide into a Section.
    fn parse_slide_as_section(&self, idx: usize, slide: &SlideInfo) -> Result<Section> {
        let mut section = Section::new(idx);
        section.name = Some(format!("Slide {}", idx + 1));

        if let Some(slide_path) = self.slide_path(slide) {
            let slide_full_rels = self
                .container
                .read_optional_relationships_for_part(&slide_path)?;
            let inherited_phs = self.build_inherited_phs(&slide_path, &slide_full_rels)?;
            let inherited_bullets = self.build_inherited_bullets(&slide_path, &slide_full_rels)?;
            let slide_rels = slide_full_rels.into_targets_by_id();

            if let Some(xml) = self.container.read_xml_optional(&slide_path)? {
                // The slide's own background picture. A background a slide takes from its
                // layout or master is template decoration, like the other layout and
                // master media, and is not referenced.
                section.background_image = crate::drawing::background_picture(&xml)
                    .and_then(|rel_id| slide_rels.get(&rel_id))
                    .map(|target| media_name(target));
                let blocks = self.parse_slide_content_with_rels(
                    &xml,
                    &slide_rels,
                    &slide_path,
                    &inherited_phs,
                    &inherited_bullets,
                )?;
                for block in blocks {
                    section.add_block(block);
                }
            }

            let notes_path = slide_path
                .replace("slides/slide", "notesSlides/notesSlide")
                .replace("slides\\slide", "notesSlides\\notesSlide");
            if let Some(xml) = self.container.read_xml_optional(&notes_path)? {
                let notes_rels = self.parse_slide_relationships(&notes_path)?;
                let notes = self.parse_notes_with_rels(&xml, &notes_rels)?;
                if !notes.is_empty() {
                    section.notes = Some(notes);
                }
            }
        }

        Ok(section)
    }

    /// Parse relationships for a specific slide/notes file.
    fn parse_slide_relationships(&self, slide_path: &str) -> Result<HashMap<String, String>> {
        self.container
            .read_optional_relationships_for_part(slide_path)
            .map(|rels| rels.into_targets_by_id())
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
        self.parse_slide_content_with_rels(
            xml,
            &HashMap::new(),
            "",
            &HashMap::new(),
            &InheritedBullets::default(),
        )
    }

    /// Parse slide XML into content blocks with relationship map for hyperlinks, images, and charts.
    fn parse_slide_content_with_rels(
        &self,
        xml: &str,
        rels: &HashMap<String, String>,
        slide_path: &str,
        inherited_phs: &HashMap<String, Vec<Paragraph>>,
        inherited_bullets: &InheritedBullets,
    ) -> Result<Vec<Block>> {
        // Each kind of content is read on its own, with the position in the slide's XML where
        // it is drawn; the blocks then follow the slide's own order -- the order its shapes are
        // listed in, the one Apache POI's slide shows give. A caption comes before or after
        // its picture, a chart after the text that introduces it, as on the slide.
        let mut placed: Vec<(u64, Block)> = Vec::new();

        for (at, para) in self.parse_text_content_excluding_tables_at(
            xml,
            rels,
            inherited_phs,
            inherited_bullets,
        )? {
            placed.push((at, Block::Paragraph(para)));
        }

        for (at, table) in self.parse_tables_at(xml, rels)? {
            placed.push((at, Block::Table(table)));
        }

        // Charts are converted to tables for RAG-ready output.
        for (at, table) in self.parse_charts(xml, rels, slide_path)? {
            placed.push((at, Block::Table(table)));
        }

        // Pictures (p:pic elements and picture-filled shapes).
        placed.extend(self.parse_images(xml, rels)?);

        // Stable: blocks at one position (the paragraphs a placeholder inherits) keep their order.
        placed.sort_by_key(|(at, _)| *at);
        Ok(placed.into_iter().map(|(_, block)| block).collect())
    }

    /// Parse images from slide XML.
    /// Images are in <p:pic> elements with <a:blip r:embed="rIdN"> referencing relationships.
    /// Each image comes with the position in `xml` where its shape ends.
    fn parse_images(&self, xml: &str, rels: &HashMap<String, String>) -> Result<Vec<(u64, Block)>> {
        let mut images = Vec::new();
        let mut reader = crate::decode::reader_for(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        // A picture is a `p:pic`; a shape filled with a picture is a `p:sp` whose `spPr`
        // holds an `a:blipFill`. Both show the image where they sit on the slide.
        let mut in_pic = false;
        let mut is_shape = false;
        let mut in_nvpicpr = false;
        let mut in_blipfill = false;
        let mut in_sppr = false;
        let mut current_descr: Option<String> = None;
        let mut current_rel_id: Option<String> = None;
        let mut current_width: Option<u32> = None;
        let mut current_height: Option<u32> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    let local_name = e.name().local_name();
                    match local_name.as_ref() {
                        // p:pic - picture element
                        "pic" | "sp" if !in_pic => {
                            in_pic = true;
                            is_shape = local_name.as_ref() == "sp";
                            in_sppr = false;
                            in_blipfill = false;
                            current_descr = None;
                            current_rel_id = None;
                            current_width = None;
                            current_height = None;
                        }
                        // p:nvPicPr - non-visual picture properties (contains name)
                        "nvPicPr" | "nvSpPr" if in_pic => {
                            in_nvpicpr = true;
                        }
                        // p:cNvPr - common non-visual properties (has name attribute)
                        "cNvPr" if in_nvpicpr => {
                            if let Some(descr) = picture_description(e) {
                                current_descr = Some(descr);
                            }
                        }
                        // p:blipFill - blip fill (contains the image reference)
                        "blipFill" if in_pic && in_sppr == is_shape => {
                            in_blipfill = true;
                        }
                        // a:blip - the actual image reference
                        "blip" if in_blipfill => {
                            for attr in e.attributes().flatten() {
                                // r:embed attribute contains the relationship ID
                                if attr.key.local_name().as_ref() == "embed" {
                                    current_rel_id = Some(attr.value.to_string());
                                }
                            }
                        }
                        // p:spPr - shape properties (contains size)
                        "spPr" if in_pic => {
                            in_sppr = true;
                        }
                        // a:ext - extent (size)
                        "ext" if in_sppr => {
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    "cx" => {
                                        if let Ok(cx) = attr.value.parse::<u32>() {
                                            current_width = Some(cx);
                                        }
                                    }
                                    "cy" => {
                                        if let Ok(cy) = attr.value.parse::<u32>() {
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
                        "cNvPr" if in_nvpicpr => {
                            if let Some(descr) = picture_description(e) {
                                current_descr = Some(descr);
                            }
                        }
                        // Handle self-closing blip
                        "blip" if in_blipfill => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == "embed" {
                                    current_rel_id = Some(attr.value.to_string());
                                }
                            }
                        }
                        // Handle self-closing ext
                        "ext" if in_sppr => {
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    "cx" => {
                                        if let Ok(cx) = attr.value.parse::<u32>() {
                                            current_width = Some(cx);
                                        }
                                    }
                                    "cy" => {
                                        if let Ok(cy) = attr.value.parse::<u32>() {
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
                        "pic" | "sp" if in_pic && (local_name.as_ref() == "sp") == is_shape => {
                            // Create image block if we have a valid relationship
                            if let Some(rel_id) = current_rel_id.take() {
                                if let Some(target) = rels.get(&rel_id) {
                                    // Extract filename from target path (e.g., "../media/image1.png" -> "image1.png")
                                    let filename =
                                        target.rsplit('/').next().unwrap_or(target).to_string();

                                    images.push((
                                        reader.buffer_position(),
                                        Block::Image {
                                            resource_id: filename,
                                            alt_text: current_descr.take(),
                                            width: current_width.take(),
                                            height: current_height.take(),
                                        },
                                    ));
                                }
                            }
                            in_pic = false;
                        }
                        "nvPicPr" | "nvSpPr" => {
                            in_nvpicpr = false;
                        }
                        "blipFill" => {
                            in_blipfill = false;
                        }
                        "spPr" => {
                            in_sppr = false;
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(images)
    }

    /// Parse charts referenced in slide relationships and convert to tables for RAG-ready output.
    /// Chart data is extracted from ppt/charts/chartN.xml files.
    ///
    /// The charts come in the order the slide draws them: where the slide first names each
    /// relationship. (A map of relationships has no order; reading it as it comes made a slide
    /// with several charts give them in a different order on each run.)
    fn parse_charts(
        &self,
        xml: &str,
        rels: &HashMap<String, String>,
        slide_path: &str,
    ) -> Result<Vec<(u64, Table)>> {
        let mut tables = Vec::new();

        // Find chart relationships (target contains "chart")
        let mut chart_rels: Vec<(&String, &String)> = rels
            .iter()
            .filter(|(_, target)| target.contains("chart"))
            .collect();
        chart_rels
            .sort_by_key(|(id, _)| (xml.find(&format!("\"{id}\"")).unwrap_or(usize::MAX), *id));
        for (id, target) in chart_rels {
            // Where the slide draws the chart: its frame names the relationship.
            let at = xml
                .find(&format!("\"{id}\""))
                .map_or(u64::MAX, |at| at as u64);
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
            let chart_xml = self.container.read_xml(&chart_path)?;
            let chart_data = charts::parse_chart_xml(&chart_xml)?;
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
                tables.push((at, table));
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
        Ok(self
            .parse_tables_at(xml, rels)?
            .into_iter()
            .map(|(_, table)| table)
            .collect())
    }

    /// The slide's tables, each with the position in `xml` where it ends.
    fn parse_tables_at(
        &self,
        xml: &str,
        rels: &HashMap<String, String>,
    ) -> Result<Vec<(u64, Table)>> {
        let mut tables = Vec::new();
        let mut reader = crate::decode::reader_for(xml);
        // Don't trim text - preserve whitespace from xml:space="preserve" elements
        reader.config_mut().trim_text(false);

        let mut buf = Vec::new();
        let mut in_table = false;
        let mut in_row = false;
        let mut in_cell = false;
        let mut in_txbody = false;
        let mut in_paragraph = false;
        let mut in_tc_pr = false;
        let mut in_ppr = false;
        let mut in_bu_blip = false;
        let mut current_level: u8 = 0;
        // The paragraph's own marker (a cell's text has no list style to inherit from here),
        // and the numbers the paragraphs of the current cell have taken.
        let mut current_bullet: Option<Marker> = None;
        let mut numbering = Numbering::default();

        let mut current_table = Table::new();
        let mut current_row = Row::new();
        let mut current_cell = Cell::new();
        // The current `a:tc` is a position another cell's merge covers — it has no cell of
        // its own in the model (see `Table::cell_columns`).
        let mut cell_covered = false;
        // The current row holds a position a vertical merge from above covers. Such a row
        // stays even when it reads as empty: dropping it would leave the merge's
        // `row_span` counting a row that is not there.
        let mut row_under_merge = false;
        let mut current_paragraphs: Vec<Paragraph> = Vec::new();
        let mut runs = RunReader::default();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e))
                    if in_paragraph && runs.element(e, false, rels) => {}
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    let local_name = e.name().local_name();
                    match local_name.as_ref() {
                        // a:tbl - table
                        "tbl" => {
                            in_table = true;
                            current_table = Table::new();
                        }
                        // a:tr - table row
                        "tr" if in_table => {
                            in_row = true;
                            row_under_merge = false;
                            current_row = Row::new();
                        }
                        // a:tc - table cell
                        "tc" if in_row => {
                            in_cell = true;
                            current_cell = Cell::new();
                            cell_covered = read_table_cell_merge(e, &mut current_cell);
                            row_under_merge |= cell_covered && is_vertically_covered(e);
                            current_paragraphs.clear();
                            numbering = Numbering::default();
                        }
                        // a:tcPr - cell properties; its fill may be a picture
                        "tcPr" if in_cell => {
                            in_tc_pr = true;
                        }
                        "blip" if in_tc_pr => {
                            if let Some(file) = blip_file(e, rels) {
                                current_cell.background_image = Some(file);
                            }
                        }
                        // a:txBody - text body in cell
                        "txBody" if in_cell => {
                            in_txbody = true;
                        }
                        // a:pPr / a:buBlip - the paragraph's own picture bullet
                        "pPr" if in_paragraph && !runs.in_run() => {
                            in_ppr = true;
                            current_level = paragraph_level(e);
                        }
                        "buBlip" if in_ppr => {
                            in_bu_blip = true;
                        }
                        name @ ("buNone" | "buChar" | "buAutoNum") if in_ppr => {
                            current_bullet = text_marker(name, e);
                        }
                        "blip" if in_bu_blip => {
                            if let Some(file) = blip_file(e, rels) {
                                current_bullet = Some(Marker::Image(file));
                            }
                        }
                        // a:p - paragraph
                        "p" if in_txbody => {
                            in_paragraph = true;
                            current_bullet = None;
                            current_level = 0;
                            in_ppr = false;
                            runs.start_paragraph();
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Empty(ref e))
                    if in_paragraph && runs.element(e, true, rels) => {}
                Ok(quick_xml::events::Event::Empty(ref e)) => {
                    let local_name = e.name().local_name();
                    match local_name.as_ref() {
                        // A self-closing a:tc: an empty cell, or a covered position.
                        "tc" if in_row => {
                            let mut cell = Cell::new();
                            let covered = read_table_cell_merge(e, &mut cell);
                            row_under_merge |= covered && is_vertically_covered(e);
                            if !covered {
                                current_row.add_cell(cell);
                            }
                        }
                        "blip" if in_tc_pr => {
                            if let Some(file) = blip_file(e, rels) {
                                current_cell.background_image = Some(file);
                            }
                        }
                        "pPr" if in_paragraph && !runs.in_run() => {
                            current_level = paragraph_level(e);
                        }
                        name @ ("buNone" | "buChar" | "buAutoNum") if in_ppr => {
                            current_bullet = text_marker(name, e);
                        }
                        "blip" if in_bu_blip => {
                            if let Some(file) = blip_file(e, rels) {
                                current_bullet = Some(Marker::Image(file));
                            }
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Text(ref e)) if runs.in_text() => {
                    runs.push_text(&crate::decode::decode_text_lossy(e));
                }
                Ok(quick_xml::events::Event::GeneralRef(ref e)) if runs.in_text() => {
                    runs.push_text(&crate::decode::resolve_general_ref(e));
                }
                Ok(quick_xml::events::Event::End(ref e)) if in_paragraph && runs.end(e) => {}
                Ok(quick_xml::events::Event::End(ref e)) => {
                    let local_name = e.name().local_name();
                    match local_name.as_ref() {
                        "p" if in_txbody => {
                            let runs = runs.finish();
                            if !runs.is_empty() {
                                let list_info =
                                    numbering.list_info(current_bullet.take(), current_level);
                                current_paragraphs.push(Paragraph {
                                    runs,
                                    list_info,
                                    ..Default::default()
                                });
                            }
                            in_paragraph = false;
                            in_ppr = false;
                        }
                        "pPr" => {
                            in_ppr = false;
                        }
                        "buBlip" => {
                            in_bu_blip = false;
                        }
                        "tcPr" => {
                            in_tc_pr = false;
                        }
                        "txBody" => {
                            in_txbody = false;
                        }
                        "tc" => {
                            if !cell_covered {
                                current_cell.content = current_paragraphs.clone();
                                current_row.add_cell(current_cell.clone());
                            }
                            in_cell = false;
                        }
                        "tr" => {
                            if !current_row.is_empty() || row_under_merge {
                                // Mark first row as header
                                if current_table.is_empty() {
                                    current_row.is_header = true;
                                }
                                current_table.add_row(current_row.clone());
                            }
                            in_row = false;
                        }
                        "tbl" => {
                            if !current_table.is_empty() {
                                tables.push((reader.buffer_position(), current_table.clone()));
                            }
                            in_table = false;
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(tables)
    }

    /// Parse text content excluding tables (paragraphs from shapes, not table cells).
    #[allow(dead_code)]
    fn parse_text_content_excluding_tables(&self, xml: &str) -> Result<Vec<Paragraph>> {
        self.parse_text_content_excluding_tables_with_rels(
            xml,
            &HashMap::new(),
            &HashMap::new(),
            &InheritedBullets::default(),
        )
    }

    /// Parse text content excluding tables with relationship map for hyperlinks.
    /// `inherited_phs` maps placeholder key → fallback paragraphs from layout/master.
    fn parse_text_content_excluding_tables_with_rels(
        &self,
        xml: &str,
        rels: &HashMap<String, String>,
        inherited_phs: &HashMap<String, Vec<Paragraph>>,
        inherited_bullets: &InheritedBullets,
    ) -> Result<Vec<Paragraph>> {
        Ok(self
            .parse_text_content_excluding_tables_at(xml, rels, inherited_phs, inherited_bullets)?
            .into_iter()
            .map(|(_, paragraph)| paragraph)
            .collect())
    }

    /// The paragraphs of the slide's shapes, each with the position in `xml` where it ends --
    /// what a slide's other contents are placed against to put the slide in the order it draws.
    fn parse_text_content_excluding_tables_at(
        &self,
        xml: &str,
        rels: &HashMap<String, String>,
        inherited_phs: &HashMap<String, Vec<Paragraph>>,
        inherited_bullets: &InheritedBullets,
    ) -> Result<Vec<(u64, Paragraph)>> {
        let mut paragraphs: Vec<(u64, Paragraph)> = Vec::new();
        let mut reader = crate::decode::reader_for(xml);
        // Don't trim text - preserve whitespace from xml:space="preserve" elements
        reader.config_mut().trim_text(false);

        let mut buf = Vec::new();
        let mut in_table = false;
        let mut table_depth = 0;
        let mut in_shape = false;
        let mut in_txbody = false;
        let mut in_paragraph = false;
        let mut runs = RunReader::default();
        let mut current_heading: HeadingLevel = HeadingLevel::None;
        // Placeholder inheritance tracking
        let mut current_ph_key: Option<String> = None;
        let mut shape_para_start: usize = 0; // paragraphs.len() when current shape started
                                             // Picture bullets (`a:buBlip`): the image of the paragraph's own bullet, and the
                                             // ones the shape's list style sets per level (`a:lstStyle/a:lvlNpPr`).
        let mut in_ppr = false;
        let mut in_bu_blip = false;
        let mut in_lst_style = false;
        let mut lst_style_level: Option<u8> = None;
        let mut shape_bullets: HashMap<u8, Option<Marker>> = HashMap::new();
        let mut current_ph_type: Option<String> = None;
        let mut current_ph_idx: Option<String> = None;
        let mut current_level: u8 = 0;
        // The paragraph's own marker, and whether it says it has none (`a:buNone`).
        let mut current_bullet: Option<Marker> = None;
        let mut bullet_overridden = false;
        // The numbers the paragraphs of the shape's text have taken, per level.
        let mut numbering = Numbering::default();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e))
                    if in_paragraph && !in_table && runs.element(e, false, rels) => {}
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    let local_name = e.name().local_name();
                    match local_name.as_ref() {
                        // Track table depth to skip table content
                        "tbl" => {
                            in_table = true;
                            table_depth += 1;
                        }
                        // p:sp - shape (also matches inner shapes inside p:grpSp groups,
                        // because quick_xml's flat event stream processes nested elements
                        // the same as top-level ones by local name)
                        "sp" if !in_table => {
                            in_shape = true;
                            current_heading = HeadingLevel::None;
                            current_ph_key = None;
                            current_ph_type = None;
                            current_ph_idx = None;
                            shape_para_start = paragraphs.len();
                            shape_bullets.clear();
                            numbering = Numbering::default();
                        }
                        "lstStyle" if in_shape && !in_table => {
                            in_lst_style = true;
                        }
                        name if in_lst_style && list_style_level(name).is_some() => {
                            lst_style_level = list_style_level(name);
                        }
                        // a:pPr - paragraph properties (level, bullet)
                        "pPr" if in_paragraph && !in_table && !runs.in_run() => {
                            in_ppr = true;
                            current_level = paragraph_level(e);
                        }
                        "buBlip" if in_ppr || lst_style_level.is_some() => {
                            in_bu_blip = true;
                        }
                        name @ ("buNone" | "buChar" | "buAutoNum") if in_ppr => {
                            bullet_overridden = true;
                            current_bullet = text_marker(name, e);
                        }
                        name @ ("buNone" | "buChar" | "buAutoNum") if lst_style_level.is_some() => {
                            if let Some(level) = lst_style_level {
                                shape_bullets.insert(level, text_marker(name, e));
                            }
                        }
                        "blip" if in_bu_blip => {
                            record_bullet(
                                e,
                                rels,
                                in_ppr,
                                lst_style_level,
                                &mut current_bullet,
                                &mut shape_bullets,
                            );
                        }
                        // p:txBody - text body in shape
                        "txBody" if in_shape && !in_table => {
                            in_txbody = true;
                        }
                        // a:p - paragraph (only if not in table, but in shape's txBody)
                        "p" if !in_table && in_txbody => {
                            in_paragraph = true;
                            runs.start_paragraph();
                            in_ppr = false;
                            current_level = 0;
                            current_bullet = None;
                            bullet_overridden = false;
                        }
                        // p:ph - placeholder type (for heading detection and inheritance key)
                        "ph" if in_shape && !in_table => {
                            let placeholder = Placeholder::of(e);
                            current_heading = placeholder.heading();
                            current_ph_key = Some(placeholder.key());
                            current_ph_type = placeholder.kind;
                            current_ph_idx = placeholder.idx;
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Empty(ref e))
                    if in_paragraph && !in_table && runs.element(e, true, rels) => {}
                Ok(quick_xml::events::Event::Empty(ref e)) => {
                    let local_name = e.name().local_name();
                    match local_name.as_ref() {
                        "pPr" if in_paragraph && !in_table && !runs.in_run() => {
                            current_level = paragraph_level(e);
                        }
                        name @ ("buNone" | "buChar" | "buAutoNum") if in_ppr => {
                            bullet_overridden = true;
                            current_bullet = text_marker(name, e);
                        }
                        name @ ("buNone" | "buChar" | "buAutoNum") if lst_style_level.is_some() => {
                            if let Some(level) = lst_style_level {
                                shape_bullets.insert(level, text_marker(name, e));
                            }
                        }
                        "blip" if in_bu_blip => {
                            record_bullet(
                                e,
                                rels,
                                in_ppr,
                                lst_style_level,
                                &mut current_bullet,
                                &mut shape_bullets,
                            );
                        }
                        // p:ph - placeholder type (self-closing)
                        "ph" if in_shape && !in_table => {
                            let placeholder = Placeholder::of(e);
                            current_heading = placeholder.heading();
                            current_ph_key = Some(placeholder.key());
                            current_ph_type = placeholder.kind;
                            current_ph_idx = placeholder.idx;
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Text(ref e)) if runs.in_text() && !in_table => {
                    runs.push_text(&crate::decode::decode_text_lossy(e));
                }
                Ok(quick_xml::events::Event::GeneralRef(ref e)) if runs.in_text() && !in_table => {
                    runs.push_text(&crate::decode::resolve_general_ref(e));
                }
                Ok(quick_xml::events::Event::End(ref e))
                    if in_paragraph && !in_table && runs.end(e) => {}
                Ok(quick_xml::events::Event::End(ref e)) => {
                    let local_name = e.name().local_name();
                    match local_name.as_ref() {
                        "tbl" => {
                            table_depth -= 1;
                            if table_depth == 0 {
                                in_table = false;
                            }
                        }
                        "pPr" => {
                            in_ppr = false;
                        }
                        "buBlip" => {
                            in_bu_blip = false;
                        }
                        "lstStyle" => {
                            in_lst_style = false;
                            lst_style_level = None;
                        }
                        name if in_lst_style && list_style_level(name).is_some() => {
                            lst_style_level = None;
                        }
                        "p" if !in_table => {
                            let runs = runs.finish();
                            if !runs.is_empty() {
                                // The paragraph's own bullet, else the shape's list style,
                                // else what the layout and master give the placeholder.
                                let marker = current_bullet.take().or_else(|| {
                                    if bullet_overridden {
                                        return None;
                                    }
                                    match shape_bullets.get(&current_level) {
                                        Some(own) => own.clone(),
                                        None => inherited_bullets
                                            .levels(
                                                current_ph_type.as_deref(),
                                                current_ph_idx.as_deref(),
                                            )
                                            .and_then(|levels| levels.get(&current_level))
                                            .cloned()
                                            .flatten()
                                            .map(|marker| match marker {
                                                Marker::Image(path) => {
                                                    Marker::Image(media_name(&path))
                                                }
                                                other => other,
                                            }),
                                    }
                                });
                                let list_info = numbering.list_info(marker, current_level);
                                paragraphs.push((
                                    reader.buffer_position(),
                                    Paragraph {
                                        runs,
                                        heading: current_heading,
                                        list_info,
                                        ..Default::default()
                                    },
                                ));
                            }
                            in_paragraph = false;
                            in_ppr = false;
                        }
                        "txBody" if !in_table => {
                            in_txbody = false;
                        }
                        "sp" if !in_table => {
                            // If this shape had a placeholder but produced no text,
                            // inherit from layout/master
                            if paragraphs.len() == shape_para_start {
                                if let Some(ref ph_key) = current_ph_key {
                                    if let Some(fallback) = inherited_phs.get(ph_key) {
                                        let at = reader.buffer_position();
                                        paragraphs.extend(fallback.iter().map(|p| (at, p.clone())));
                                    }
                                }
                            }
                            in_shape = false;
                            current_heading = HeadingLevel::None;
                            current_ph_key = None;
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => return Err(e.into()),
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
        let mut reader = crate::decode::reader_for(xml);
        // Don't trim text - preserve whitespace from xml:space="preserve" elements
        reader.config_mut().trim_text(false);

        let mut buf = Vec::new();
        let mut in_paragraph = false;
        let mut runs = RunReader::default();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    if e.name().local_name().as_ref() == "p" {
                        in_paragraph = true;
                        runs.start_paragraph();
                    } else if in_paragraph {
                        runs.element(e, false, rels);
                    }
                }
                Ok(quick_xml::events::Event::Empty(ref e)) if in_paragraph => {
                    runs.element(e, true, rels);
                }
                Ok(quick_xml::events::Event::Text(ref e)) if runs.in_text() => {
                    runs.push_text(&crate::decode::decode_text_lossy(e));
                }
                Ok(quick_xml::events::Event::GeneralRef(ref e)) if runs.in_text() => {
                    runs.push_text(&crate::decode::resolve_general_ref(e));
                }
                Ok(quick_xml::events::Event::End(ref e)) => {
                    if e.name().local_name().as_ref() == "p" {
                        let runs = runs.finish();
                        if !runs.is_empty() {
                            paragraphs.push(Paragraph {
                                runs,
                                ..Default::default()
                            });
                        }
                        in_paragraph = false;
                    } else {
                        runs.end(e);
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(paragraphs)
    }

    /// Build a map of placeholder key → fallback paragraphs from layout and master XMLs.
    /// Layout takes precedence over master; slide-defined text takes precedence over both.
    ///
    /// The layout, its relationships and the master are optional: an absent part is
    /// skipped. One that is present and unreadable is an error, as for every other
    /// optional part.
    fn build_inherited_phs(
        &self,
        slide_path: &str,
        slide_full_rels: &crate::container::Relationships,
    ) -> Result<HashMap<String, Vec<Paragraph>>> {
        const SLIDE_LAYOUT_TYPE: &str =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
        const SLIDE_MASTER_TYPE: &str =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";

        let mut inherited: HashMap<String, Vec<Paragraph>> = HashMap::new();

        // Try to find the layout from slide relationships
        let layout_path = slide_full_rels
            .get_by_type(SLIDE_LAYOUT_TYPE)
            .first()
            .map(|rel| OoxmlContainer::resolve_path(slide_path, &rel.target));

        let Some(layout_path) = layout_path else {
            return Ok(inherited);
        };

        // Parse master first (lower priority), then overlay layout
        let layout_full_rels = self
            .container
            .read_optional_relationships_for_part(&layout_path)?;

        if let Some(master_rel) = layout_full_rels.get_by_type(SLIDE_MASTER_TYPE).first() {
            let master_path = OoxmlContainer::resolve_path(&layout_path, &master_rel.target);
            if let Some(master_xml) = self.container.read_xml_optional(&master_path)? {
                inherited = parse_placeholder_texts_from_xml(&master_xml)?;
            }
        }

        if let Some(layout_xml) = self.container.read_xml_optional(&layout_path)? {
            // Layout overrides master
            for (key, paras) in parse_placeholder_texts_from_xml(&layout_xml)? {
                inherited.insert(key, paras);
            }
        }

        // Remove presentational placeholders — these are dynamic at display time
        // and would inject noise (footer text, slide numbers, dates) into every slide.
        const PRESENTATIONAL: &[&str] = &["dt", "sldNum", "ftr", "hdr", "sldImg"];
        inherited.retain(|k, _| !PRESENTATIONAL.contains(&k.as_str()));

        Ok(inherited)
    }

    /// The picture bullets the slide's placeholders inherit from its layout and master.
    fn build_inherited_bullets(
        &self,
        slide_path: &str,
        slide_full_rels: &crate::container::Relationships,
    ) -> Result<InheritedBullets> {
        const SLIDE_LAYOUT_TYPE: &str =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
        const SLIDE_MASTER_TYPE: &str =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";

        let Some(layout_rel) = slide_full_rels
            .get_by_type(SLIDE_LAYOUT_TYPE)
            .first()
            .copied()
        else {
            return Ok(InheritedBullets::default());
        };
        let layout_path = OoxmlContainer::resolve_path(slide_path, &layout_rel.target);
        let layout_rels = self
            .container
            .read_optional_relationships_for_part(&layout_path)?;
        let master_path = layout_rels
            .get_by_type(SLIDE_MASTER_TYPE)
            .first()
            .map(|rel| OoxmlContainer::resolve_path(&layout_path, &rel.target));

        let layout_xml = self.container.read_xml_optional(&layout_path)?;
        let layout_targets = layout_rels.into_targets_by_id();
        let (master_xml, master_targets) = match &master_path {
            Some(path) => (
                self.container.read_xml_optional(path)?,
                self.container
                    .read_optional_relationships_for_part(path)?
                    .into_targets_by_id(),
            ),
            None => (None, HashMap::new()),
        };

        Ok(InheritedBullets::new(
            Some((&layout_path, &layout_targets)),
            layout_xml.as_deref(),
            master_path.as_deref().map(|p| (p, &master_targets)),
            master_xml.as_deref(),
        ))
    }

    /// The media a slide draws through bullets it inherits from its layout or master:
    /// package path of each picture bullet that some paragraph of the slide ends up with.
    fn inherited_bullet_media(&self, slide_path: &str) -> Result<Vec<String>> {
        let slide_full_rels = self
            .container
            .read_optional_relationships_for_part(slide_path)?;
        let bullets = self.build_inherited_bullets(slide_path, &slide_full_rels)?;
        let Some(xml) = self.container.read_xml_optional(slide_path)? else {
            return Ok(Vec::new());
        };
        let rels = slide_full_rels.into_targets_by_id();
        let paragraphs = self.parse_text_content_excluding_tables_with_rels(
            &xml,
            &rels,
            &HashMap::new(),
            &bullets,
        )?;
        let mut paths: Vec<String> = Vec::new();
        for file in paragraphs
            .iter()
            .filter_map(|p| p.list_info.as_ref()?.marker_image.as_deref())
        {
            if let Some(path) = bullets.path_of(file) {
                if !paths.iter().any(|p| p == path) {
                    paths.push(path.to_string());
                }
            }
        }
        Ok(paths)
    }

    /// Extract resources (images, media) from the presentation.
    ///
    /// Lists the media the slides reference, keyed by file name, the way a Word document
    /// lists its image relationships: media used only by a slide layout or master, or by
    /// nothing at all, is not part of what the slides show. A picture's SVG original or
    /// HD Photo layer is listed as a companion of the image the picture shows (see
    /// [`crate::drawing`]); a file that some slide shows as a picture stays a primary.
    pub fn extract_resources(&self) -> Result<Vec<Resource>> {
        let mut resources: Vec<Resource> = Vec::new();
        let mut index: HashMap<String, usize> = HashMap::new();
        // Companion file → (role, primary file), kept only while no slide shows it.
        let mut companions: HashMap<String, (ResourceRole, String)> = HashMap::new();
        let mut shown: std::collections::HashSet<String> = std::collections::HashSet::new();

        for slide in &self.slides {
            let Some(slide_path) = self.slide_path(slide) else {
                continue;
            };
            let rels = self
                .container
                .read_optional_relationships_for_part(&slide_path)?;
            let pictures = self
                .container
                .read_xml_optional(&slide_path)?
                .map(|xml| crate::drawing::scan_pictures(&xml))
                .unwrap_or_default();

            let file_of = |rel_id: &str| -> Option<(String, String)> {
                let rel = rels.get(rel_id)?;
                let path = OoxmlContainer::resolve_path(&slide_path, &rel.target);
                let name = path.rsplit('/').next().unwrap_or(&path).to_string();
                Some((path, name))
            };

            let mut ids: Vec<&String> = rels.by_id.keys().collect();
            ids.sort();
            for rel_id in ids {
                let rel = &rels.by_id[rel_id];
                let is_media = ["/image", "/hdphoto", "/video", "/audio", "/media"]
                    .iter()
                    .any(|kind| rel.rel_type.ends_with(kind));
                if !is_media || rel.external {
                    continue;
                }
                let Some((path, name)) = file_of(rel_id) else {
                    continue;
                };

                match pictures.companions.get(rel_id.as_str()) {
                    Some((role, primary)) => {
                        if let Some((_, primary_name)) = file_of(primary) {
                            companions
                                .entry(name.clone())
                                .or_insert((*role, primary_name));
                        }
                    }
                    None => {
                        shown.insert(name.clone());
                    }
                }

                let at = match index.get(&name) {
                    Some(&at) => at,
                    None => {
                        let Ok(data) = self.container.read_binary(&path) else {
                            continue;
                        };
                        index.insert(name.clone(), resources.len());
                        resources.push(Resource::from_part(&path, data));
                        resources.len() - 1
                    }
                };
                let resource = &mut resources[at];
                if resource.alt_text.is_none() {
                    resource.alt_text = pictures.alt_texts.get(rel_id.as_str()).cloned();
                }
            }

            // A picture bullet the slide takes from its layout or master is drawn by the
            // slide's paragraphs, so its image is listed with the slide's media.
            for path in self.inherited_bullet_media(&slide_path)? {
                let name = media_name(&path);
                shown.insert(name.clone());
                if let std::collections::hash_map::Entry::Vacant(slot) = index.entry(name) {
                    let Ok(data) = self.container.read_binary(&path) else {
                        continue;
                    };
                    slot.insert(resources.len());
                    resources.push(Resource::from_part(&path, data));
                }
            }
        }

        for resource in &mut resources {
            let name = resource.filename.as_deref().unwrap_or_default();
            if shown.contains(name) {
                continue;
            }
            if let Some((role, primary)) = companions.get(name) {
                resource.role = *role;
                resource.companion_of = Some(primary.clone());
                resource.alt_text = None;
            }
        }

        Ok(resources)
    }

    /// The package path of a slide.
    fn slide_path(&self, slide: &SlideInfo) -> Option<String> {
        let target = self.relationships.get(&slide.rel_id)?;
        Some(match target.strip_prefix('/') {
            Some(stripped) => stripped.to_string(),
            None => format!("ppt/{target}"),
        })
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

/// Parse placeholder texts from a layout or master XML.
/// Returns ph_key → Vec<Paragraph> for non-empty placeholder shapes.
/// ph_key = ph type (e.g. "title") if set, else "idx:<N>".
fn parse_placeholder_texts_from_xml(xml: &str) -> Result<HashMap<String, Vec<Paragraph>>> {
    let mut result: HashMap<String, Vec<Paragraph>> = HashMap::new();

    let mut reader = crate::decode::reader_for(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    // A layout's and master's text links nowhere a slide's reader resolves: their links
    // are their own parts' relationships.
    let no_links = HashMap::new();
    let mut in_shape = false;
    let mut in_table = false;
    let mut table_depth = 0u32;
    let mut in_txbody = false;
    let mut in_paragraph = false;
    let mut runs = RunReader::default();
    let mut current_ph_key: Option<String> = None;
    let mut current_heading = HeadingLevel::None;
    let mut shape_paragraphs: Vec<Paragraph> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e))
                if in_paragraph && !in_table && runs.element(e, false, &no_links) => {}
            Ok(quick_xml::events::Event::Start(ref e)) => match e.name().local_name().as_ref() {
                "tbl" => {
                    in_table = true;
                    table_depth += 1;
                }
                "sp" if !in_table => {
                    in_shape = true;
                    current_ph_key = None;
                    current_heading = HeadingLevel::None;
                    shape_paragraphs.clear();
                }
                "txBody" if in_shape && !in_table => {
                    in_txbody = true;
                }
                "p" if in_txbody && !in_table => {
                    in_paragraph = true;
                    runs.start_paragraph();
                }
                "ph" if in_shape && !in_table => {
                    let placeholder = Placeholder::of(e);
                    current_heading = placeholder.heading();
                    current_ph_key = Some(placeholder.key());
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Empty(ref e))
                if in_paragraph && !in_table && runs.element(e, true, &no_links) => {}
            Ok(quick_xml::events::Event::Empty(ref e)) => {
                if e.name().local_name().as_ref() == "ph" && in_shape && !in_table {
                    let placeholder = Placeholder::of(e);
                    current_heading = placeholder.heading();
                    current_ph_key = Some(placeholder.key());
                }
            }
            Ok(quick_xml::events::Event::Text(ref e)) if runs.in_text() && !in_table => {
                runs.push_text(&crate::decode::decode_text_lossy(e));
            }
            Ok(quick_xml::events::Event::GeneralRef(ref e)) if runs.in_text() && !in_table => {
                runs.push_text(&crate::decode::resolve_general_ref(e));
            }
            Ok(quick_xml::events::Event::End(ref e))
                if in_paragraph && !in_table && runs.end(e) => {}
            Ok(quick_xml::events::Event::End(ref e)) => match e.name().local_name().as_ref() {
                "tbl" => {
                    table_depth -= 1;
                    if table_depth == 0 {
                        in_table = false;
                    }
                }
                "p" if !in_table => {
                    let runs = runs.finish();
                    if !runs.is_empty() {
                        shape_paragraphs.push(Paragraph {
                            runs,
                            heading: current_heading,
                            ..Default::default()
                        });
                    }
                    in_paragraph = false;
                }
                "txBody" if !in_table => {
                    in_txbody = false;
                }
                "sp" if !in_table => {
                    // Only store non-empty shapes with a placeholder key
                    if let Some(ph_key) = current_ph_key.take() {
                        if !shape_paragraphs.is_empty() {
                            result.insert(ph_key, shape_paragraphs.clone());
                        }
                    }
                    in_shape = false;
                    current_heading = HeadingLevel::None;
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Eof) => break,
            Err(e) => return Err(e.into()),
            _ => {}
        }
        buf.clear();
    }

    Ok(result)
}

/// A picture's description (`p:cNvPr descr`) — its alt text — if it has a non-blank one.
/// The shape's `name` ("Picture 3") is an editing label, not a description.
/// The list level a paragraph-properties element (`a:pPr`) sets (`lvl`, 0-based).
fn paragraph_level(e: &quick_xml::events::BytesStart<'_>) -> u8 {
    e.attributes()
        .flatten()
        .find(|a| a.key.local_name().as_ref() == "lvl")
        .and_then(|a| a.value.parse::<u8>().ok())
        .unwrap_or(0)
}

/// The 0-based level of a list-style level element (`a:lvl1pPr` … `a:lvl9pPr`).
fn list_style_level(name: &str) -> Option<u8> {
    let digit = name.strip_prefix("lvl")?.strip_suffix("pPr")?;
    digit
        .parse::<u8>()
        .ok()
        .filter(|n| (1..=9).contains(n))
        .map(|n| n - 1)
}

/// Record the picture of an `a:buBlip`'s `a:blip`: the file name of the image it embeds,
/// on the paragraph when the bullet is the paragraph's own, else on the shape's list style
/// level it sits in.
fn record_bullet(
    blip: &quick_xml::events::BytesStart<'_>,
    rels: &HashMap<String, String>,
    in_paragraph_properties: bool,
    list_style_level: Option<u8>,
    paragraph_bullet: &mut Option<Marker>,
    shape_bullets: &mut HashMap<u8, Option<Marker>>,
) {
    let Some(target) = blip
        .attributes()
        .flatten()
        .find(|a| a.key.local_name().as_ref() == "embed")
        .and_then(|a| rels.get(a.value.as_ref()))
    else {
        return;
    };
    let file = media_name(target);
    if in_paragraph_properties {
        *paragraph_bullet = Some(Marker::Image(file));
    } else if let Some(level) = list_style_level {
        shape_bullets.insert(level, Some(Marker::Image(file)));
    }
}

/// The numbers the paragraphs of one shape's text have taken, per level: automatic numbers
/// continue down consecutive paragraphs of a level (that state the same scheme and start) and
/// start over after anything that interrupts them — a paragraph without numbers at their level
/// or above, or one of a shallower level.
#[derive(Default)]
struct Numbering {
    /// For each level: the scheme, the number the run started at, and the last number given.
    levels: [Option<(String, u32, u32)>; 10],
}

impl Numbering {
    /// The list information of a paragraph at `level` with `marker`, and the numbers it uses up.
    fn list_info(&mut self, marker: Option<Marker>, level: u8) -> Option<ListInfo> {
        let at = usize::from(level).min(self.levels.len() - 1);
        // Whatever the paragraph is, deeper levels start over after it.
        for deeper in &mut self.levels[at + 1..] {
            *deeper = None;
        }
        let info = |list_type, number, marker_image| ListInfo {
            list_type,
            level,
            number,
            marker_image,
        };
        match marker {
            None => {
                self.levels[at] = None;
                None
            }
            Some(Marker::Image(file)) => {
                self.levels[at] = None;
                Some(info(ListType::Bullet, None, Some(file)))
            }
            Some(Marker::Char) => {
                self.levels[at] = None;
                Some(info(ListType::Bullet, None, None))
            }
            Some(Marker::Number { scheme, start_at }) => {
                let number = match &self.levels[at] {
                    Some((s, start, last)) if *s == scheme && *start == start_at => last + 1,
                    _ => start_at,
                };
                self.levels[at] = Some((scheme, start_at, number));
                Some(info(ListType::Numbered, Some(number), None))
            }
        }
    }
}

/// The media file an `a:blip`'s `r:embed` names, through the part's relationships.
fn blip_file(
    blip: &quick_xml::events::BytesStart<'_>,
    rels: &HashMap<String, String>,
) -> Option<String> {
    blip.attributes()
        .flatten()
        .find(|a| a.key.local_name().as_ref() == "embed")
        .and_then(|a| rels.get(a.value.as_ref()))
        .map(|target| media_name(target))
}

/// The file name of a media part: the last segment of a relationship target.
fn media_name(target: &str) -> String {
    target.rsplit('/').next().unwrap_or(target).to_string()
}

fn picture_description(e: &quick_xml::events::BytesStart<'_>) -> Option<String> {
    e.attributes()
        .flatten()
        .find(|a| a.key.local_name().as_ref() == "descr")
        .map(|a| crate::decode::attr_value(&a))
        .filter(|d| !d.trim().is_empty())
}

/// Read a DrawingML table cell's merge attributes into `cell`, returning whether the
/// cell is a covered position.
///
/// DrawingML writes every grid position as an `a:tc`: the owner of a merge carries
/// `gridSpan`/`rowSpan`, and the positions it covers are present with `hMerge="1"` or
/// `vMerge="1"`. The model records the merge once, on the owner, and gives covered
/// positions no cell.
fn read_table_cell_merge(e: &quick_xml::events::BytesStart<'_>, cell: &mut Cell) -> bool {
    let mut covered = false;
    for attr in e.attributes().flatten() {
        let value = attr.value.as_ref();
        match attr.key.local_name().as_ref() {
            "gridSpan" => cell.col_span = value.parse().unwrap_or(1).max(1),
            "rowSpan" => cell.row_span = value.parse().unwrap_or(1).max(1),
            "hMerge" | "vMerge" => covered |= is_true(value),
            _ => {}
        }
    }
    covered
}

/// Whether a covered `a:tc` is covered from a row above (`vMerge`).
fn is_vertically_covered(e: &quick_xml::events::BytesStart<'_>) -> bool {
    e.attributes()
        .flatten()
        .any(|a| a.key.local_name().as_ref() == "vMerge" && is_true(a.value.as_ref()))
}

fn is_true(value: &str) -> bool {
    value == "1" || value == "true"
}

/// Apply the formatting an `a:rPr` sets on its run -- bold, italic, underline, strikethrough --
/// to `style`. One reader for every place a run is read: a table cell's runs once kept only
/// bold and italic, and lost their underline and strikethrough.
fn apply_run_properties(e: &quick_xml::events::BytesStart<'_>, style: &mut TextStyle) {
    for attr in e.attributes().flatten() {
        let val = crate::decode::attr_value(&attr);
        match attr.key.local_name().as_ref() {
            "b" => style.bold = val != "0" && val != "false",
            "i" => style.italic = val != "0" && val != "false",
            "u" => style.underline = val != "none",
            "strike" => style.strikethrough = val != "noStrike" && val != "0" && val != "false",
            _ => {}
        }
    }
}

/// The runs of the paragraph being read. Every reader of a deck's text -- a shape's, a table
/// cell's, the notes', the placeholders' a layout and master give -- reads a paragraph's runs
/// through this one, so what one of them understands of a run they all do. (Three copies
/// had once kept apart a table cell's formatting; none of the four read a line break or a
/// field.)
#[derive(Default)]
struct RunReader {
    runs: Vec<TextRun>,
    in_run: bool,
    in_text: bool,
    text: String,
    style: TextStyle,
    hyperlink: Option<String>,
}

impl RunReader {
    /// Start reading a new paragraph.
    fn start_paragraph(&mut self) {
        *self = Self::default();
    }

    /// Whether a run or a field is open: an `a:pPr` inside one is not the paragraph's.
    fn in_run(&self) -> bool {
        self.in_run
    }

    /// Whether text read now is a run's.
    fn in_text(&self) -> bool {
        self.in_text
    }

    /// Read an element opened -- or self-closed, when `empty` -- inside the paragraph, and say
    /// whether it was one of the runs'.
    fn element(
        &mut self,
        e: &quick_xml::events::BytesStart<'_>,
        empty: bool,
        rels: &HashMap<String, String>,
    ) -> bool {
        match e.name().local_name().as_ref() {
            // A field (`a:fld` -- a slide number, a date) is a run whose text the application
            // keeps up to date; the value it last showed is its `a:t`.
            "r" | "fld" if !empty => {
                self.in_run = true;
                self.in_text = false;
                self.text.clear();
                self.style = TextStyle::default();
                self.hyperlink = None;
            }
            "t" if self.in_run && !empty => self.in_text = true,
            "rPr" if self.in_run => apply_run_properties(e, &mut self.style),
            "hlinkClick" if self.in_run => {
                if let Some(url) = link_address(e, rels) {
                    self.hyperlink = Some(url);
                }
            }
            // A line break inside the paragraph ends the line of the run before it.
            "br" if !self.in_run => match self.runs.last_mut() {
                Some(run) => run.line_break = true,
                None => self.runs.push(TextRun {
                    text: String::new(),
                    line_break: true,
                    ..TextRun::default()
                }),
            },
            _ => return false,
        }
        true
    }

    /// Text inside the open `a:t`.
    fn push_text(&mut self, text: &str) {
        self.text.push_str(text);
    }

    /// Read an element's end, and say whether it was one of the runs'.
    fn end(&mut self, e: &quick_xml::events::BytesEnd<'_>) -> bool {
        match e.name().local_name().as_ref() {
            "t" if self.in_run => self.in_text = false,
            "r" | "fld" if self.in_run => {
                if !self.text.is_empty() {
                    let hyperlink = self.hyperlink.take();
                    let style = std::mem::take(&mut self.style);
                    self.runs.push(TextRun {
                        text: std::mem::take(&mut self.text),
                        style: if hyperlink.is_some() {
                            style.in_link()
                        } else {
                            style
                        },
                        hyperlink,
                        ..TextRun::default()
                    });
                }
                self.in_run = false;
                self.in_text = false;
                self.hyperlink = None;
            }
            _ => return false,
        }
        true
    }

    /// The paragraph's runs, read to its end: none when it holds no text (a line break
    /// alone is no content), and no break after its last line.
    fn finish(&mut self) -> Vec<TextRun> {
        let mut runs = std::mem::take(&mut self.runs);
        if runs.iter().all(|run| run.text.is_empty()) {
            return Vec::new();
        }
        while runs.last().is_some_and(|run| run.text.is_empty()) {
            runs.pop();
        }
        if let Some(last) = runs.last_mut() {
            last.line_break = false;
        }
        runs
    }
}

/// A shape's placeholder (`p:ph`): its type and index.
struct Placeholder {
    kind: Option<String>,
    idx: Option<String>,
}

impl Placeholder {
    fn of(e: &quick_xml::events::BytesStart<'_>) -> Self {
        let mut kind = None;
        let mut idx = None;
        for attr in e.attributes().flatten() {
            match attr.key.local_name().as_ref() {
                "type" => kind = Some(crate::decode::attr_value(&attr)).filter(|t| !t.is_empty()),
                "idx" => idx = Some(crate::decode::attr_value(&attr)),
                _ => {}
            }
        }
        Self { kind, idx }
    }

    /// The heading a placeholder of this type makes of its text.
    fn heading(&self) -> HeadingLevel {
        match self.kind.as_deref() {
            Some("title" | "ctrTitle") => HeadingLevel::H1,
            Some("subTitle") => HeadingLevel::H2,
            _ => HeadingLevel::None,
        }
    }

    /// The key a layout's and master's text for this placeholder is kept under: its type, or
    /// its index for one that has none.
    fn key(&self) -> String {
        match &self.kind {
            Some(kind) => kind.clone(),
            None => format!("idx:{}", self.idx.as_deref().unwrap_or("0")),
        }
    }
}

/// The address an `a:hlinkClick` links its text to: the target of the relationship it names.
/// A jump to another slide of the deck (`ppaction://hlinksldjump`, `ppaction://hlinkshowjump`)
/// leads to no address the output can point at, and is no link.
fn link_address(
    e: &quick_xml::events::BytesStart<'_>,
    rels: &HashMap<String, String>,
) -> Option<String> {
    let mut id = None;
    let mut action = None;
    for attr in e.attributes().flatten() {
        match attr.key.local_name().as_ref() {
            "id" => id = Some(crate::decode::attr_value(&attr)),
            "action" => action = Some(crate::decode::attr_value(&attr)),
            _ => {}
        }
    }
    let jumps_within = action.is_some_and(|a| {
        a.starts_with("ppaction://hlinksldjump") || a.starts_with("ppaction://hlinkshowjump")
    });
    if jumps_within {
        return None;
    }
    rels.get(&id?).cloned()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::error::Error;

    fn empty_parser() -> PptxParser {
        use std::io::Cursor;
        let zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let bytes = zip.finish().unwrap().into_inner();
        PptxParser {
            container: OoxmlContainer::from_bytes(bytes).unwrap(),
            slides: Vec::new(),
            relationships: HashMap::new(),
        }
    }

    #[test]
    fn test_pptx_slide_text_entities_round_trip() {
        // Slide run text exercising the GeneralRef arm: predefined + numeric refs
        // decode, a mid-word "AT&amp;T" survives run fragmentation, and an unknown
        // &bogus; is preserved verbatim.
        let parser = empty_parser();
        let xml = r#"<?xml version="1.0"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:sp><p:txBody>
    <a:p><a:r><a:t>AT&amp;T &lt;x&gt; &#48;&#x30; &bogus; end</a:t></a:r></a:p>
  </p:txBody></p:sp></p:spTree></p:cSld>
</p:sld>"#;

        let paragraphs = parser.parse_text_content(xml).unwrap();
        let text: String = paragraphs.iter().map(|p| p.plain_text()).collect();
        assert_eq!(text, "AT&T <x> 00 &bogus; end");
    }

    #[test]
    fn test_pptx_stray_ampersand_does_not_abort() {
        // A raw `&` (ill-formed) must degrade gracefully rather than abort the slide.
        let parser = empty_parser();
        let xml = r#"<?xml version="1.0"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:sp><p:txBody>
    <a:p><a:r><a:t>R&D dept</a:t></a:r></a:p>
  </p:txBody></p:sp></p:spTree></p:cSld>
</p:sld>"#;

        let paragraphs = parser
            .parse_text_content(xml)
            .expect("stray & must not abort");
        let text: String = paragraphs.iter().map(|p| p.plain_text()).collect();
        assert_eq!(text, "R&D dept");
    }

    /// Each `a:p` becomes a paragraph with its text, and run formatting survives: the
    /// run marked `b="1"` is bold and the other is not.
    #[test]
    fn test_parse_text_content() {
        let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>Hello World</a:t></a:r></a:p><a:p><a:r><a:rPr b="1"/><a:t>Bold Text</a:t></a:r></a:p></p:txBody></p:sp>"#;
        let data = deck(&[(&slide(shape), EMPTY_RELS)], &[]);
        let doc = PptxParser::from_bytes(data).unwrap().parse().unwrap();

        let paras = paragraphs(&doc.sections[0]);
        assert_eq!(paras.len(), 2);
        assert_eq!(paras[0].plain_text(), "Hello World");
        assert!(!paras[0].runs[0].style.bold);
        assert_eq!(paras[1].plain_text(), "Bold Text");
        assert!(paras[1].runs[0].style.bold);
    }

    const EMPTY_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;

    /// A deck of `slides` (each with its relationships part), plus arbitrary extra parts
    /// such as media. The single-slide helpers below cannot express either.
    fn deck(slides: &[(&str, &str)], extra_parts: &[(&str, &[u8])]) -> Vec<u8> {
        use std::io::{Cursor, Write};
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);

        let mut overrides = String::new();
        let mut presentation_rels = String::new();
        let mut slide_ids = String::new();
        for n in 1..=slides.len() {
            overrides.push_str(&format!(
                r#"<Override PartName="/ppt/slides/slide{n}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#
            ));
            presentation_rels.push_str(&format!(
                r#"<Relationship Id="rId{n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{n}.xml"/>"#
            ));
            slide_ids.push_str(&format!(r#"<p:sldId id="{}" r:id="rId{n}"/>"#, 255 + n));
        }

        zip.start_file("[Content_Types].xml", options).unwrap();
        write!(
            zip,
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  {overrides}
</Types>"#
        )
        .unwrap();

        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"#).unwrap();

        zip.start_file("ppt/_rels/presentation.xml.rels", options)
            .unwrap();
        write!(
            zip,
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{presentation_rels}</Relationships>"#
        )
        .unwrap();

        zip.start_file("ppt/presentation.xml", options).unwrap();
        write!(
            zip,
            r#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>{slide_ids}</p:sldIdLst>
</p:presentation>"#
        )
        .unwrap();

        for (i, (slide_xml, slide_rels)) in slides.iter().enumerate() {
            let n = i + 1;
            zip.start_file(format!("ppt/slides/_rels/slide{n}.xml.rels"), options)
                .unwrap();
            zip.write_all(slide_rels.as_bytes()).unwrap();
            zip.start_file(format!("ppt/slides/slide{n}.xml"), options)
                .unwrap();
            zip.write_all(slide_xml.as_bytes()).unwrap();
        }

        for (path, data) in extra_parts {
            zip.start_file(*path, options).unwrap();
            zip.write_all(data).unwrap();
        }

        zip.finish().unwrap().into_inner()
    }

    fn slide(shapes: &str) -> String {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>{shapes}</p:spTree></p:cSld>
</p:sld>"#
        )
    }

    /// A text shape; `placeholder` is the `<p:ph>` element, or empty for a plain box.
    fn text_shape(placeholder: &str, text: &str) -> String {
        format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr>{placeholder}</p:nvPr></p:nvSpPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"#
        )
    }

    fn paragraphs(section: &crate::model::Section) -> Vec<&crate::model::Paragraph> {
        section
            .content
            .iter()
            .filter_map(|block| match block {
                crate::model::Block::Paragraph(p) => Some(p),
                _ => None,
            })
            .collect()
    }

    /// Every slide becomes one section, in order, and the metadata's page count is the
    /// number of slides.
    #[test]
    fn test_each_slide_becomes_a_section() {
        let first = slide(&text_shape("", "First slide"));
        let second = slide(&text_shape("", "Second slide"));
        let data = deck(&[(&first, EMPTY_RELS), (&second, EMPTY_RELS)], &[]);

        let mut parser = PptxParser::from_bytes(data).unwrap();
        assert_eq!(parser.slide_count(), 2);

        let doc = parser.parse().unwrap();
        assert_eq!(doc.sections.len(), 2);
        assert_eq!(doc.metadata.page_count, Some(2));
        assert_eq!(paragraphs(&doc.sections[0])[0].plain_text(), "First slide");
        assert_eq!(paragraphs(&doc.sections[1])[0].plain_text(), "Second slide");
    }

    /// Title and subtitle placeholders map to heading levels 1 and 2; a plain text box
    /// stays body text.
    #[test]
    fn test_title_placeholders_become_headings() {
        use crate::model::HeadingLevel;

        let shapes = [
            text_shape(r#"<p:ph type="title"/>"#, "Deck Title"),
            text_shape(r#"<p:ph type="subTitle" idx="1"/>"#, "The Subtitle"),
            text_shape("", "Body text"),
        ]
        .concat();
        let data = deck(&[(&slide(&shapes), EMPTY_RELS)], &[]);
        let doc = PptxParser::from_bytes(data).unwrap().parse().unwrap();

        let paras = paragraphs(&doc.sections[0]);
        let seen: Vec<(String, HeadingLevel)> =
            paras.iter().map(|p| (p.plain_text(), p.heading)).collect();
        assert_eq!(seen[0], ("Deck Title".to_string(), HeadingLevel::H1));
        assert_eq!(seen[1], ("The Subtitle".to_string(), HeadingLevel::H2));
        assert_eq!(seen[2].0, "Body text");
        assert!(!seen[2].1.is_heading());
    }

    /// A slide table becomes a table block with its cells in place, and renders as a
    /// Markdown table.
    #[test]
    fn test_slide_table_becomes_a_table() {
        let cell = |text: &str| {
            format!(
                r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>{text}</a:t></a:r></a:p></a:txBody></a:tc>"#
            )
        };
        let table = format!(
            r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tr h="370840">{}{}</a:tr><a:tr h="370840">{}{}</a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#,
            cell("Name"),
            cell("Score"),
            cell("Ada"),
            cell("99")
        );
        let data = deck(&[(&slide(&table), EMPTY_RELS)], &[]);
        let doc = PptxParser::from_bytes(data).unwrap().parse().unwrap();

        let tables: Vec<&crate::model::Table> = doc.sections[0]
            .content
            .iter()
            .filter_map(|block| match block {
                crate::model::Block::Table(t) => Some(t),
                _ => None,
            })
            .collect();
        assert_eq!(tables.len(), 1);
        let texts: Vec<Vec<String>> = tables[0]
            .rows
            .iter()
            .map(|row| row.cells.iter().map(|c| c.plain_text()).collect())
            .collect();
        assert_eq!(texts, [["Name", "Score"], ["Ada", "99"]]);

        let md =
            crate::render::to_markdown(&doc, &crate::render::RenderOptions::default()).unwrap();
        assert!(md.contains("| Name | Score |"), "markdown: {md}");
        assert!(md.contains("| Ada | 99 |"), "markdown: {md}");
    }

    /// A table cell's run keeps all of its formatting -- underline and strikethrough as well
    /// as bold and italic -- as a run in a text box does.
    #[test]
    fn test_a_table_cell_run_keeps_its_underline_and_strikethrough() {
        let cell = |rpr: &str, text: &str| {
            format!(
                r#"<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr {rpr}/><a:t>{text}</a:t></a:r></a:p></a:txBody></a:tc>"#
            )
        };
        let table = format!(
            r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tr h="370840">{}{}</a:tr><a:tr h="370840">{}{}</a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#,
            cell("", "Name"),
            cell("", "Note"),
            cell(r#"u="sng""#, "marked"),
            cell(r#"strike="sngStrike""#, "dropped")
        );
        let data = deck(&[(&slide(&table), EMPTY_RELS)], &[]);
        let doc = PptxParser::from_bytes(data).unwrap().parse().unwrap();
        let runs: Vec<&TextRun> = doc.sections[0]
            .content
            .iter()
            .filter_map(|block| match block {
                crate::model::Block::Table(t) => Some(t),
                _ => None,
            })
            .flat_map(|t| t.rows.iter())
            .flat_map(|row| row.cells.iter())
            .flat_map(|cell| cell.content.iter())
            .flat_map(|p| p.runs.iter())
            .collect();
        let run = |text: &str| runs.iter().find(|r| r.text == text).copied().unwrap();
        assert!(run("marked").style.underline);
        assert!(run("dropped").style.strikethrough);
    }

    /// DrawingML writes every grid position as an `a:tc`, marking the ones a merge covers
    /// with `hMerge`/`vMerge`. The merge is recorded on its owner, the covered positions
    /// are not cells, and the table renders one column per grid column — including a row
    /// that holds nothing but a covered position and an empty cell.
    #[test]
    fn test_slide_table_merges_are_recorded_on_their_owner() {
        let tc = |attrs: &str, text: &str| {
            format!(
                r#"<a:tc{attrs}><a:txBody><a:bodyPr/><a:p><a:r><a:t>{text}</a:t></a:r></a:p></a:txBody></a:tc>"#
            )
        };
        let empty_tc =
            |attrs: &str| format!(r#"<a:tc{attrs}><a:txBody><a:bodyPr/><a:p/></a:txBody></a:tc>"#);
        let rows = [
            // A header spanning columns 1-2, then column 3.
            format!(
                "{}{}{}",
                tc(r#" gridSpan="2""#, "Pair"),
                empty_tc(r#" hMerge="1""#),
                tc("", "C")
            ),
            // A label covering this row and the next.
            format!(
                "{}{}{}",
                tc(r#" rowSpan="2""#, "G"),
                tc("", "x"),
                tc("", "1")
            ),
            format!(
                "{}{}{}",
                empty_tc(r#" vMerge="1""#),
                empty_tc(""),
                empty_tc("")
            ),
            format!("{}{}{}", tc("", "H"), tc("", "z"), tc("", "3")),
        ];
        let body: String = rows
            .iter()
            .map(|r| format!(r#"<a:tr h="370840">{r}</a:tr>"#))
            .collect();
        let table = format!(
            r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>{body}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#
        );
        let data = deck(&[(&slide(&table), EMPTY_RELS)], &[]);
        let doc = PptxParser::from_bytes(data).unwrap().parse().unwrap();
        let table = doc.sections[0]
            .content
            .iter()
            .find_map(|block| match block {
                crate::model::Block::Table(t) => Some(t),
                _ => None,
            })
            .expect("a table");

        let texts: Vec<Vec<String>> = table
            .rows
            .iter()
            .map(|row| row.cells.iter().map(|c| c.plain_text()).collect())
            .collect();
        assert_eq!(
            texts,
            vec![
                vec!["Pair", "C"],
                vec!["G", "x", "1"],
                vec!["", ""],
                vec!["H", "z", "3"],
            ]
        );
        assert_eq!(table.rows[0].cells[0].col_span, 2);
        assert_eq!(table.rows[1].cells[0].row_span, 2);
        assert_eq!(table.cell_columns()[2], vec![1, 2]);

        let md =
            crate::render::to_markdown(&doc, &crate::render::RenderOptions::default()).unwrap();
        assert!(md.contains("| H | z | 3 |"), "markdown: {md}");
        assert!(
            md.lines()
                .filter(|l| l.starts_with('|'))
                .all(|l| l.matches('|').count() == 4),
            "every table line must have three columns:\n{md}"
        );
    }

    /// A run's hyperlink resolves through the slide's own relationships to the external URL.
    #[test]
    fn test_hyperlink_resolves_through_the_slide_relationships() {
        let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Link"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr><a:hlinkClick r:id="rIdLink"/></a:rPr><a:t>the spec</a:t></a:r></a:p></p:txBody></p:sp>"#;
        let rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/spec" TargetMode="External"/>
</Relationships>"#;
        let data = deck(&[(&slide(shape), rels)], &[]);
        let doc = PptxParser::from_bytes(data).unwrap().parse().unwrap();

        let run = paragraphs(&doc.sections[0])
            .iter()
            .flat_map(|p| p.runs.iter())
            .find(|r| r.text == "the spec")
            .expect("the linked run");
        assert_eq!(run.hyperlink.as_deref(), Some("https://example.com/spec"));
    }

    /// A link that jumps to another slide of the deck leads to no address the output can
    /// point at: its text is no link.
    #[test]
    fn test_a_jump_to_another_slide_is_no_link() {
        let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Link"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr><a:hlinkClick r:id="rIdSlide" action="ppaction://hlinksldjump"/></a:rPr><a:t>see slide 3</a:t></a:r></a:p></p:txBody></p:sp>"#;
        let rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSlide" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slide3.xml"/>
</Relationships>"#;
        let data = deck(&[(&slide(shape), rels)], &[]);
        let doc = PptxParser::from_bytes(data).unwrap().parse().unwrap();
        let run = paragraphs(&doc.sections[0])
            .iter()
            .flat_map(|p| p.runs.iter())
            .find(|r| r.text == "see slide 3")
            .expect("the run");
        assert_eq!(run.hyperlink, None);
    }

    /// A linked run's underline is how the link looks, not the author's emphasis; an
    /// underlined run outside a link keeps it.
    #[test]
    fn test_a_link_run_is_not_kept_underlined() {
        let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Link"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr u="sng"><a:hlinkClick r:id="rIdLink"/></a:rPr><a:t>the spec</a:t></a:r><a:r><a:rPr u="sng"/><a:t>marked</a:t></a:r></a:p></p:txBody></p:sp>"#;
        let rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/spec" TargetMode="External"/>
</Relationships>"#;
        let data = deck(&[(&slide(shape), rels)], &[]);
        let doc = PptxParser::from_bytes(data).unwrap().parse().unwrap();
        let runs: Vec<&TextRun> = paragraphs(&doc.sections[0])
            .iter()
            .flat_map(|p| p.runs.iter())
            .collect();
        let link = runs.iter().find(|r| r.text == "the spec").unwrap();
        assert!(link.hyperlink.is_some() && !link.style.underline);
        let marked = runs.iter().find(|r| r.text == "marked").unwrap();
        assert!(marked.style.underline);
    }

    /// A line break inside a paragraph (`a:br`) ends the line of the run before it, and a
    /// field (`a:fld` -- a slide number, a date) is read as a run of the value it last showed,
    /// formatting and all; a field's own `a:pPr` is not the paragraph's.
    #[test]
    fn test_a_line_break_and_a_field_are_read_in_shape_text() {
        let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>First line</a:t></a:r><a:br><a:rPr lang="en-US"/></a:br><a:r><a:t>Second line</a:t></a:r></a:p><a:p><a:r><a:t>Slide </a:t></a:r><a:fld id="{6A1B}" type="slidenum"><a:rPr b="1"/><a:pPr lvl="3"/><a:t>7</a:t></a:fld></a:p></p:txBody></p:sp>"#;
        let data = deck(&[(&slide(shape), EMPTY_RELS)], &[]);
        let doc = PptxParser::from_bytes(data).unwrap().parse().unwrap();
        let paras = paragraphs(&doc.sections[0]);
        assert_eq!(paras.len(), 2, "{paras:?}");
        let first: Vec<(&str, bool)> = paras[0]
            .runs
            .iter()
            .map(|r| (r.text.as_str(), r.line_break))
            .collect();
        assert_eq!(first, [("First line", true), ("Second line", false)]);
        let field = paras[1].runs.iter().find(|r| r.text == "7").unwrap();
        assert!(field.style.bold);
        assert_eq!(paras[1].plain_text(), "Slide 7");
    }

    /// A table cell's paragraphs read line breaks and fields as a text box's do.
    #[test]
    fn test_a_table_cell_reads_line_breaks_and_fields() {
        let table = r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tr h="370840"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>Top</a:t></a:r><a:br/><a:r><a:t>Bottom</a:t></a:r></a:p></a:txBody></a:tc><a:tc><a:txBody><a:bodyPr/><a:p><a:fld id="{7C2D}" type="datetime1"><a:t>10/13/2026</a:t></a:fld></a:p></a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#;
        let data = deck(&[(&slide(table), EMPTY_RELS)], &[]);
        let doc = PptxParser::from_bytes(data).unwrap().parse().unwrap();
        let cells: Vec<&Cell> = doc.sections[0]
            .content
            .iter()
            .filter_map(|block| match block {
                crate::model::Block::Table(t) => Some(t),
                _ => None,
            })
            .flat_map(|t| t.rows.iter())
            .flat_map(|row| row.cells.iter())
            .collect();
        let first = &cells[0].content[0].runs;
        assert_eq!(first[0].text, "Top");
        assert!(first[0].line_break && !first[1].line_break);
        assert_eq!(cells[1].content[0].plain_text(), "10/13/2026");
    }

    /// Notes read fields too: a note's slide-number field is text, not nothing.
    #[test]
    fn test_notes_read_a_field() {
        let notes = r#"<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>See slide </a:t></a:r><a:fld id="{1}" type="slidenum"><a:t>3</a:t></a:fld></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>"#;
        let paras = empty_parser().parse_notes(notes).unwrap();
        assert_eq!(paras.len(), 1);
        assert_eq!(paras[0].plain_text(), "See slide 3");
    }

    /// The media a slide references is listed, typed by its extension; a media part no
    /// slide references is not.
    #[test]
    fn test_extract_resources_lists_the_media_slides_reference() {
        let rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rIdVid" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="../media/clip.mp4"/>
</Relationships>"#;
        let data = deck(
            &[(&slide(&text_shape("", "x")), rels)],
            &[
                ("ppt/media/image1.png", b"\x89PNG\r\n\x1a\n"),
                ("ppt/media/clip.mp4", b"not really a video"),
                ("ppt/media/orphan.png", b"\x89PNG\r\n\x1a\n"),
            ],
        );
        let parser = PptxParser::from_bytes(data).unwrap();

        let mut resources = parser.extract_resources().unwrap();
        resources.sort_by(|a, b| a.filename.cmp(&b.filename));
        let names: Vec<_> = resources.iter().map(|r| r.filename.as_deref()).collect();
        assert_eq!(names, [Some("clip.mp4"), Some("image1.png")]);
        assert!(!resources[0].is_image());
        assert!(resources[1].is_image());
        assert_eq!(resources[1].data, b"\x89PNG\r\n\x1a\n");
    }

    /// A picture with an SVG original and an HD Photo layer is one picture: the raster image
    /// is the primary, carrying the picture's description, and the other two are its
    /// companions. The slide's image block names the primary, with the description — not
    /// the shape's name — as its alt text.
    #[test]
    fn test_picture_companions_are_marked_and_alt_text_is_the_description() {
        let pic = r#"<p:pic><p:nvPicPr><p:cNvPr id="4" name="Picture 3" descr="Quarterly revenue"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdPng"><a:extLst><a:ext uri="{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}"><a14:imgProps xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><a14:imgLayer r:embed="rIdWdp"/></a14:imgProps></a:ext><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rIdSvg"/></a:ext></a:extLst></a:blip></p:blipFill><p:spPr/></p:pic>"#;
        let rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPng" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rIdSvg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.svg"/>
  <Relationship Id="rIdWdp" Type="http://schemas.microsoft.com/office/2007/relationships/hdphoto" Target="../media/hdphoto1.wdp"/>
</Relationships>"#;
        let data = deck(
            &[(&slide(pic), rels)],
            &[
                ("ppt/media/image1.png", b"\x89PNG\r\n\x1a\n"),
                ("ppt/media/image2.svg", b"<svg/>"),
                ("ppt/media/hdphoto1.wdp", b"II\xbc\x01"),
            ],
        );
        let doc = PptxParser::from_bytes(data).unwrap().parse().unwrap();

        let png = &doc.resources["image1.png"];
        assert_eq!(png.role, ResourceRole::Primary);
        assert_eq!(png.alt_text.as_deref(), Some("Quarterly revenue"));
        let svg = &doc.resources["image2.svg"];
        assert_eq!(svg.role, ResourceRole::Alternate);
        assert_eq!(svg.companion_of.as_deref(), Some("image1.png"));
        let wdp = &doc.resources["hdphoto1.wdp"];
        assert_eq!(wdp.role, ResourceRole::Layer);
        assert_eq!(wdp.companion_of.as_deref(), Some("image1.png"));
        assert_eq!(wdp.mime_type.as_deref(), Some("image/vnd.ms-photo"));

        let images: Vec<_> = doc.sections[0]
            .content
            .iter()
            .filter_map(|b| match b {
                Block::Image {
                    resource_id,
                    alt_text,
                    ..
                } => Some((resource_id.as_str(), alt_text.as_deref())),
                _ => None,
            })
            .collect();
        assert_eq!(images, [("image1.png", Some("Quarterly revenue"))]);
    }

    /// Image rels for the picture-fill and picture-bullet tests below.
    const IMAGE_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdFill" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/fill.png"/>
  <Relationship Id="rIdBullet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/bullet.png"/>
</Relationships>"#;

    /// The image blocks of a section, as `(resource_id, alt_text)`.
    fn image_blocks(section: &Section) -> Vec<(&str, Option<&str>)> {
        section
            .content
            .iter()
            .filter_map(|b| match b {
                Block::Image {
                    resource_id,
                    alt_text,
                    ..
                } => Some((resource_id.as_str(), alt_text.as_deref())),
                _ => None,
            })
            .collect()
    }

    /// A shape filled with a picture shows that picture: the slide references it as an
    /// image, with the shape's description, and the shape's own text stays text. A shape with
    /// a plain colour fill adds nothing.
    #[test]
    fn test_shape_picture_fill_is_an_image_of_the_slide() {
        let shapes = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Card" descr="Harbour at dusk"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000" cy="3000"/></a:xfrm><a:prstGeom prst="rect"/><a:blipFill><a:blip r:embed="rIdFill"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>Caption on the card</a:t></a:r></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="3" name="Plain"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>Plain shape</a:t></a:r></a:p></p:txBody></p:sp>"#;
        let data = deck(
            &[(&slide(shapes), IMAGE_RELS)],
            &[("ppt/media/fill.png", b"\x89PNG\r\n\x1a\n")],
        );
        let doc = PptxParser::from_bytes(data).unwrap().parse().unwrap();

        let section = &doc.sections[0];
        assert_eq!(
            image_blocks(section),
            [("fill.png", Some("Harbour at dusk"))]
        );
        let texts: Vec<_> = paragraphs(section).iter().map(|p| p.plain_text()).collect();
        assert_eq!(texts, ["Caption on the card", "Plain shape"]);
        match section
            .content
            .iter()
            .find(|b| matches!(b, Block::Image { .. }))
        {
            Some(Block::Image { width, height, .. }) => {
                assert_eq!((*width, *height), (Some(4000), Some(3000)));
            }
            other => panic!("expected an image block, got {other:?}"),
        }
        let md =
            crate::render::to_markdown(&doc, &crate::render::RenderOptions::default()).unwrap();
        assert!(
            md.contains("![Harbour at dusk](fill.png)"),
            "markdown: {md}"
        );
        assert_eq!(doc.resources["fill.png"].role, ResourceRole::Primary);
    }

    /// A picture bullet is a list marker, not a picture: the paragraph is a bullet item whose
    /// `marker_image` names the image, the slide gets no image block per bullet, and the
    /// Markdown stays a plain list. A paragraph that turns the bullet off is not a list item.
    #[test]
    fn test_picture_bullet_is_referenced_by_the_list_items() {
        let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/>
<a:p><a:pPr><a:buBlip><a:blip r:embed="rIdBullet"/></a:buBlip></a:pPr><a:r><a:t>First</a:t></a:r></a:p>
<a:p><a:pPr lvl="1"><a:buBlip><a:blip r:embed="rIdBullet"/></a:buBlip></a:pPr><a:r><a:t>Nested</a:t></a:r></a:p>
<a:p><a:pPr><a:buNone/></a:pPr><a:r><a:t>No bullet</a:t></a:r></a:p>
</p:txBody></p:sp>"#;
        let data = deck(
            &[(&slide(shape), IMAGE_RELS)],
            &[("ppt/media/bullet.png", b"\x89PNG\r\n\x1a\n")],
        );
        let doc = PptxParser::from_bytes(data).unwrap().parse().unwrap();

        let section = &doc.sections[0];
        assert!(image_blocks(section).is_empty());
        let items: Vec<_> = paragraphs(section)
            .iter()
            .map(|p| {
                (
                    p.plain_text(),
                    p.list_info
                        .as_ref()
                        .map(|l| (l.list_type, l.level, l.marker_image.clone())),
                )
            })
            .collect();
        assert_eq!(
            items,
            [
                (
                    "First".to_string(),
                    Some((ListType::Bullet, 0, Some("bullet.png".to_string())))
                ),
                (
                    "Nested".to_string(),
                    Some((ListType::Bullet, 1, Some("bullet.png".to_string())))
                ),
                ("No bullet".to_string(), None),
            ]
        );
        let md =
            crate::render::to_markdown(&doc, &crate::render::RenderOptions::default()).unwrap();
        assert!(
            md.contains("- First") && md.contains("  - Nested"),
            "markdown: {md}"
        );
        assert!(!md.contains("bullet.png"), "markdown: {md}");
        let json = serde_json::to_string(&doc).unwrap();
        assert!(json.contains("\"marker_image\":\"bullet.png\""), "{json}");
        assert_eq!(doc.resources["bullet.png"].role, ResourceRole::Primary);
        assert_eq!(doc.resources["bullet.png"].alt_text, None);
    }

    /// A picture bullet set once in a shape's list style applies to the paragraphs of that
    /// level, unless a paragraph sets its own bullet.
    #[test]
    fn test_picture_bullet_from_the_shape_list_style() {
        let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/>
<a:lstStyle><a:lvl1pPr><a:buBlip><a:blip r:embed="rIdBullet"/></a:buBlip></a:lvl1pPr></a:lstStyle>
<a:p><a:r><a:t>Inherits</a:t></a:r></a:p>
<a:p><a:pPr><a:buChar char="-"/></a:pPr><a:r><a:t>Own bullet</a:t></a:r></a:p>
<a:p><a:pPr lvl="1"/><a:r><a:t>Other level</a:t></a:r></a:p>
</p:txBody></p:sp>"#;
        let data = deck(
            &[(&slide(shape), IMAGE_RELS)],
            &[("ppt/media/bullet.png", b"\x89PNG\r\n\x1a\n")],
        );
        let doc = PptxParser::from_bytes(data).unwrap().parse().unwrap();

        let markers: Vec<_> = paragraphs(&doc.sections[0])
            .iter()
            .map(|p| {
                p.list_info
                    .as_ref()
                    .and_then(|l| l.marker_image.as_deref().map(str::to_string))
            })
            .collect();
        assert_eq!(markers, [Some("bullet.png".to_string()), None, None]);
    }

    /// A picture fill or picture bullet on a slide layout is drawn by the layout, not the
    /// slide: like every other layout media it is neither listed nor referenced.
    #[test]
    fn test_layout_picture_fill_and_bullet_are_not_slide_content() {
        let slide_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"#;
        let layout_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdFill" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/fill.png"/>
</Relationships>"#;
        let layout = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="Frame"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:blipFill><a:blip r:embed="rIdFill"/></a:blipFill></p:spPr></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:pPr><a:buBlip><a:blip r:embed="rIdFill"/></a:buBlip></a:pPr><a:r><a:t>Prompt text</a:t></a:r></a:p></p:txBody></p:sp>
</p:spTree></p:cSld></p:sldLayout>"#;
        let data = deck(
            &[(&slide(&text_shape("", "body")), slide_rels)],
            &[
                ("ppt/slideLayouts/slideLayout1.xml", layout.as_bytes()),
                (
                    "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                    layout_rels.as_bytes(),
                ),
                ("ppt/media/fill.png", b"\x89PNG\r\n\x1a\n"),
            ],
        );
        let doc = PptxParser::from_bytes(data).unwrap().parse().unwrap();

        assert!(
            doc.resources.is_empty(),
            "resources: {:?}",
            doc.resources.keys()
        );
        assert!(image_blocks(&doc.sections[0]).is_empty());
    }

    const IMAGE_REL: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    const LAYOUT_REL: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
    const MASTER_REL: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";

    fn rels_xml(entries: &[(&str, &str, &str)]) -> String {
        let body: String = entries
            .iter()
            .map(|(id, ty, target)| {
                format!(r#"<Relationship Id="{id}" Type="{ty}" Target="{target}"/>"#)
            })
            .collect();
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{body}</Relationships>"#
        )
    }

    const PNG: &[u8] = b"\x89PNG\r\n\x1a\n";

    /// A slide whose own background is a picture references it from the slide (nothing in
    /// the text); a colour background references nothing. The background a layout gives
    /// the slide is template decoration, like the rest of the layout's media.
    #[test]
    fn test_slide_background_picture_is_referenced_by_the_section() {
        let with_bg = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<p:cSld><p:bg><p:bgPr><a:blipFill dpi="0"><a:blip r:embed="rIdBg"/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg><p:spTree></p:spTree></p:cSld></p:sld>"#;
        let colour_bg = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:bgPr></p:bg><p:spTree></p:spTree></p:cSld></p:sld>"#;
        let with_rels = rels_xml(&[
            ("rIdBg", IMAGE_REL, "../media/bg.png"),
            ("rIdLayout", LAYOUT_REL, "../slideLayouts/slideLayout1.xml"),
        ]);
        let layout_rels = rels_xml(&[("rIdLb", IMAGE_REL, "../media/layoutbg.png")]);
        let layout = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:bg><p:bgPr><a:blipFill><a:blip r:embed="rIdLb"/></a:blipFill></p:bgPr></p:bg><p:spTree/></p:cSld></p:sldLayout>"#;
        let colour_rels =
            rels_xml(&[("rIdLayout", LAYOUT_REL, "../slideLayouts/slideLayout1.xml")]);
        let data = deck(
            &[(with_bg, &with_rels), (colour_bg, &colour_rels)],
            &[
                ("ppt/slideLayouts/slideLayout1.xml", layout.as_bytes()),
                (
                    "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                    layout_rels.as_bytes(),
                ),
                ("ppt/media/bg.png", PNG),
                ("ppt/media/layoutbg.png", PNG),
            ],
        );
        let doc = PptxParser::from_bytes(data).unwrap().parse().unwrap();

        assert_eq!(doc.sections[0].background_image.as_deref(), Some("bg.png"));
        assert_eq!(doc.sections[1].background_image, None);
        assert!(
            doc.sections[0].content.is_empty(),
            "a background is no block"
        );
        assert!(doc.resources.contains_key("bg.png"));
        assert!(!doc.resources.contains_key("layoutbg.png"));
        let json = serde_json::to_string(&doc.sections[0]).unwrap();
        assert!(json.contains("\"background_image\":\"bg.png\""), "{json}");
        let md =
            crate::render::to_markdown(&doc, &crate::render::RenderOptions::default()).unwrap();
        assert!(!md.contains("bg.png"), "markdown: {md}");
    }

    /// A table cell filled with a picture references it from the cell; the cell's text
    /// stays its text and a picture bullet in a cell paragraph names its image.
    #[test]
    fn test_table_cell_picture_fill_and_bullet_are_referenced() {
        let table = r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><a:graphic><a:graphicData><a:tbl><a:tblGrid><a:gridCol w="100"/><a:gridCol w="100"/></a:tblGrid>
<a:tr h="10"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>Filled</a:t></a:r></a:p></a:txBody><a:tcPr><a:blipFill><a:blip r:embed="rIdCell"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></a:tcPr></a:tc><a:tc><a:txBody><a:bodyPr/><a:p><a:pPr><a:buBlip><a:blip r:embed="rIdBullet"/></a:buBlip></a:pPr><a:r><a:t>Bulleted</a:t></a:r></a:p></a:txBody><a:tcPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:tcPr></a:tc></a:tr>
<a:tr h="10"><a:tc><a:txBody><a:bodyPr/><a:p><a:endParaRPr/></a:p></a:txBody><a:tcPr><a:blipFill><a:blip r:embed="rIdCell"/></a:blipFill></a:tcPr></a:tc><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>Plain</a:t></a:r></a:p></a:txBody></a:tc></a:tr>
</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#;
        let rels = rels_xml(&[
            ("rIdCell", IMAGE_REL, "../media/cell.png"),
            ("rIdBullet", IMAGE_REL, "../media/bullet.png"),
        ]);
        let data = deck(
            &[(&slide(table), &rels)],
            &[("ppt/media/cell.png", PNG), ("ppt/media/bullet.png", PNG)],
        );
        let doc = PptxParser::from_bytes(data).unwrap().parse().unwrap();
        let Block::Table(t) = &doc.sections[0].content[0] else {
            panic!("expected a table: {:?}", doc.sections[0].content);
        };
        let cell = |r: usize, c: usize| &t.rows[r].cells[c];
        assert_eq!(cell(0, 0).background_image.as_deref(), Some("cell.png"));
        assert_eq!(cell(0, 0).plain_text(), "Filled");
        assert_eq!(cell(0, 1).background_image, None);
        assert_eq!(cell(1, 0).background_image.as_deref(), Some("cell.png"));
        assert_eq!(cell(1, 1).background_image, None);
        let bullet = cell(0, 1).content[0].list_info.as_ref().unwrap();
        assert_eq!(bullet.marker_image.as_deref(), Some("bullet.png"));
        assert!(doc.resources.contains_key("cell.png"));
        assert!(doc.resources.contains_key("bullet.png"));
        assert!(image_blocks(&doc.sections[0]).is_empty());
        let md =
            crate::render::to_markdown(&doc, &crate::render::RenderOptions::default()).unwrap();
        assert!(!md.contains(".png"), "markdown: {md}");
    }

    /// A table cell's own `a:buChar` and `a:buAutoNum` make list items just as a text box's do,
    /// with numbers that continue down the paragraphs of one cell and start over in the next;
    /// `a:buNone` and a paragraph without a marker stay plain.
    #[test]
    fn test_table_cell_text_bullets_and_numbers_make_list_items() {
        let table = r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><a:graphic><a:graphicData><a:tbl><a:tblGrid><a:gridCol w="100"/><a:gridCol w="100"/></a:tblGrid>
<a:tr h="10"><a:tc><a:txBody><a:bodyPr/><a:p><a:pPr><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:t>One</a:t></a:r></a:p><a:p><a:pPr><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:t>Two</a:t></a:r></a:p><a:p><a:pPr><a:buNone/></a:pPr><a:r><a:t>Break</a:t></a:r></a:p></a:txBody></a:tc><a:tc><a:txBody><a:bodyPr/><a:p><a:pPr><a:buChar char="&#8226;"/></a:pPr><a:r><a:t>Dot</a:t></a:r></a:p><a:p><a:r><a:t>Plain</a:t></a:r></a:p></a:txBody></a:tc></a:tr>
<a:tr h="10"><a:tc><a:txBody><a:bodyPr/><a:p><a:pPr><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:t>Again</a:t></a:r></a:p></a:txBody></a:tc><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>x</a:t></a:r></a:p></a:txBody></a:tc></a:tr>
</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#;
        let data = deck(&[(&slide(table), &rels_xml(&[]))], &[]);
        let doc = PptxParser::from_bytes(data).unwrap().parse().unwrap();
        let Block::Table(t) = &doc.sections[0].content[0] else {
            panic!("expected a table: {:?}", doc.sections[0].content);
        };
        let info = |r: usize, c: usize, p: usize| {
            t.rows[r].cells[c].content[p]
                .list_info
                .as_ref()
                .map(|l| (l.list_type, l.number))
        };
        assert_eq!(info(0, 0, 0), Some((ListType::Numbered, Some(1))));
        assert_eq!(info(0, 0, 1), Some((ListType::Numbered, Some(2))));
        assert_eq!(info(0, 0, 2), None, "a:buNone is no marker");
        assert_eq!(info(0, 1, 0), Some((ListType::Bullet, None)));
        assert_eq!(info(0, 1, 1), None);
        assert_eq!(
            info(1, 0, 0),
            Some((ListType::Numbered, Some(1))),
            "numbers restart in a cell"
        );
        assert_eq!(info(1, 1, 0), None);
        let md =
            crate::render::to_markdown(&doc, &crate::render::RenderOptions::default()).unwrap();
        assert!(
            md.contains("| 1. One<br>2. Two<br>Break | - Dot<br>Plain |"),
            "markdown: {md}"
        );
        assert!(md.contains("| 1. Again | x |"), "markdown: {md}");
    }

    /// Character bullets and automatic numbers make list items, by the same chain a picture
    /// bullet follows: master `txStyles`, then the paragraph's own properties. Numbers continue
    /// down consecutive paragraphs of a level and start over after anything that interrupts them;
    /// a shape that is no placeholder inherits nothing.
    #[test]
    fn test_text_bullets_and_numbers_make_list_items() {
        let master = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree/></p:cSld>
<p:txStyles><p:titleStyle><a:lvl1pPr><a:buNone/></a:lvl1pPr></p:titleStyle>
<p:bodyStyle><a:lvl1pPr><a:buChar char="&#8226;"/></a:lvl1pPr><a:lvl2pPr><a:buAutoNum type="arabicPeriod"/></a:lvl2pPr></p:bodyStyle>
<p:otherStyle/></p:txStyles></p:sldMaster>"#;
        let layout = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="A"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Prompt</a:t></a:r></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="4" name="T"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Prompt</a:t></a:r></a:p></p:txBody></p:sp>
</p:spTree></p:cSld></p:sldLayout>"#;
        let master_rels = rels_xml(&[]);
        let layout_rels =
            rels_xml(&[("rIdMaster", MASTER_REL, "../slideMasters/slideMaster1.xml")]);
        let shape = |nv: &str, style: &str, paras: &[(&str, &str)]| {
            let paras: String = paras
                .iter()
                .map(|(ppr, text)| {
                    format!(r#"<a:p><a:pPr{ppr}</a:pPr><a:r><a:t>{text}</a:t></a:r></a:p>"#)
                })
                .collect();
            format!(
                r#"<p:sp><p:nvSpPr><p:cNvPr id="9" name="S"/><p:cNvSpPr/><p:nvPr>{nv}</p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/>{style}{paras}</p:txBody></p:sp>"#
            )
        };
        let shapes = [
            // The body placeholder takes the master's body style: bullets, numbers on level 2.
            shape(
                r#"<p:ph type="body" idx="1"/>"#,
                "<a:lstStyle/>",
                &[
                    (">", "first"),
                    (r#" lvl="1">"#, "num a"),
                    (r#" lvl="1">"#, "num b"),
                    (">", "second"),
                    (r#" lvl="1">"#, "num c"),
                    ("><a:buNone/>", "plain"),
                    (r#" lvl="1">"#, "num d"),
                ],
            ),
            shape(
                r#"<p:ph type="title"/>"#,
                "<a:lstStyle/>",
                &[(">", "A title")],
            ),
            // A text box inherits nothing; its paragraphs state their own bullets.
            shape(
                "",
                "<a:lstStyle/>",
                &[
                    (">", "no bullet"),
                    (r#"><a:buAutoNum type="romanLcPeriod" startAt="4"/>"#, "x"),
                    (r#"><a:buAutoNum type="romanLcPeriod" startAt="4"/>"#, "y"),
                    ("><a:buChar char=\"-\"/>", "dash"),
                    (r#"><a:buAutoNum type="romanLcPeriod" startAt="4"/>"#, "z"),
                    // Another scheme, then the same scheme from another number: each starts
                    // over at its own first number.
                    (r#"><a:buAutoNum type="arabicPeriod"/>"#, "arabic"),
                    (
                        r#"><a:buAutoNum type="arabicPeriod" startAt="7"/>"#,
                        "again",
                    ),
                ],
            ),
            // A bullet the shape's list style gives its level.
            shape(
                "",
                r#"<a:lstStyle><a:lvl1pPr><a:buChar char="*"/></a:lvl1pPr></a:lstStyle>"#,
                &[(">", "styled"), ("><a:buNone/>", "cancelled")],
            ),
        ]
        .concat();
        let slide_rels = rels_xml(&[("rIdLayout", LAYOUT_REL, "../slideLayouts/slideLayout1.xml")]);
        let data = deck(
            &[(&slide(&shapes), &slide_rels)],
            &[
                ("ppt/slideLayouts/slideLayout1.xml", layout.as_bytes()),
                (
                    "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                    layout_rels.as_bytes(),
                ),
                ("ppt/slideMasters/slideMaster1.xml", master.as_bytes()),
                (
                    "ppt/slideMasters/_rels/slideMaster1.xml.rels",
                    master_rels.as_bytes(),
                ),
            ],
        );
        let doc = PptxParser::from_bytes(data).unwrap().parse().unwrap();
        let got: Vec<(String, String)> = paragraphs(&doc.sections[0])
            .iter()
            .map(|para| {
                let marker = match &para.list_info {
                    None => "none".to_string(),
                    Some(l) => match l.list_type {
                        ListType::Numbered => format!("{}:{}.", l.level, l.number.unwrap()),
                        _ => format!("{}:-", l.level),
                    },
                };
                (para.plain_text(), marker)
            })
            .collect();
        let want = [
            ("first", "0:-"),
            ("num a", "1:1."),
            ("num b", "1:2."),
            ("second", "0:-"),
            // The shallower paragraph in between started the numbers over.
            ("num c", "1:1."),
            ("plain", "none"),
            // After a paragraph without a bullet at the level above, the numbers start over too.
            ("num d", "1:1."),
            ("A title", "none"),
            ("no bullet", "none"),
            ("x", "0:4."),
            ("y", "0:5."),
            // A different kind of bullet interrupts the numbers.
            ("dash", "0:-"),
            ("z", "0:4."),
            ("arabic", "0:1."),
            ("again", "0:7."),
            ("styled", "0:-"),
            ("cancelled", "none"),
        ];
        let want: Vec<(String, String)> = want
            .iter()
            .map(|(t, m)| (t.to_string(), m.to_string()))
            .collect();
        assert_eq!(got, want);

        let md =
            crate::render::to_markdown(&doc, &crate::render::RenderOptions::default()).unwrap();
        for line in [
            "- first",
            "  1. num a",
            "  2. num b",
            "- second",
            "  1. num c",
        ] {
            assert!(
                md.lines().any(|l| l == line),
                "{line:?} missing from:\n{md}"
            );
        }
    }

    /// Picture bullets come down the placeholder chain — master `txStyles`, master
    /// placeholder, layout placeholder, slide shape, paragraph — level by level; a bullet
    /// of another kind replaces an inherited picture. The image a slide inherits is listed
    /// with the slide's media; one no slide paragraph ends up with is not.
    #[test]
    fn test_picture_bullets_are_inherited_through_the_placeholder_chain() {
        let master = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree/></p:cSld>
<p:txStyles><p:titleStyle><a:lvl1pPr><a:buNone/></a:lvl1pPr></p:titleStyle>
<p:bodyStyle><a:lvl1pPr><a:buBlip><a:blip r:embed="rIdMasterBullet"/></a:buBlip></a:lvl1pPr><a:lvl2pPr><a:buBlip><a:blip r:embed="rIdMasterBullet"/></a:buBlip></a:lvl2pPr></p:bodyStyle>
<p:otherStyle><a:lvl1pPr><a:buBlip><a:blip r:embed="rIdUnused"/></a:buBlip></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>"#;
        let master_rels = rels_xml(&[
            ("rIdMasterBullet", IMAGE_REL, "../media/master-bullet.png"),
            ("rIdUnused", IMAGE_REL, "../media/unused.png"),
        ]);
        // Layout: idx 1 sets its own level-1 picture and replaces level 2 with a character
        // bullet; idx 2 sets nothing and takes the master's.
        let layout = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="A"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:buBlip><a:blip r:embed="rIdLayoutBullet"/></a:buBlip></a:lvl1pPr><a:lvl2pPr><a:buChar char="-"/></a:lvl2pPr></a:lstStyle><a:p><a:r><a:t>Prompt</a:t></a:r></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="3" name="B"/><p:cNvSpPr/><p:nvPr><p:ph idx="2"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Prompt</a:t></a:r></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="4" name="T"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Prompt</a:t></a:r></a:p></p:txBody></p:sp>
</p:spTree></p:cSld></p:sldLayout>"#;
        let layout_rels = rels_xml(&[
            ("rIdLayoutBullet", IMAGE_REL, "../media/layout-bullet.png"),
            ("rIdMaster", MASTER_REL, "../slideMasters/slideMaster1.xml"),
        ]);
        let ph = |ph: &str, paras: &str| {
            format!(
                r#"<p:sp><p:nvSpPr><p:cNvPr id="9" name="S"/><p:cNvSpPr/><p:nvPr>{ph}</p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>{paras}</p:txBody></p:sp>"#
            )
        };
        let p = |ppr: &str, text: &str| {
            format!(r#"<a:p><a:pPr{ppr}</a:pPr><a:r><a:t>{text}</a:t></a:r></a:p>"#)
        };
        let shapes = [
            ph(
                r#"<p:ph type="body" idx="1"/>"#,
                &[
                    p(">", "layout one"),
                    p(r#" lvl="1">"#, "layout char overrides master"),
                    p("><a:buNone/>", "no bullet"),
                ]
                .concat(),
            ),
            ph(
                r#"<p:ph idx="2"/>"#,
                &[p(">", "master one"), p(r#" lvl="1">"#, "master two")].concat(),
            ),
            ph(r#"<p:ph type="title"/>"#, &p(">", "A title")),
            // A plain text box inherits nothing.
            ph("", &p(">", "text box")),
        ]
        .concat();
        let slide_rels = rels_xml(&[("rIdLayout", LAYOUT_REL, "../slideLayouts/slideLayout1.xml")]);
        let data = deck(
            &[(&slide(&shapes), &slide_rels)],
            &[
                ("ppt/slideLayouts/slideLayout1.xml", layout.as_bytes()),
                (
                    "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                    layout_rels.as_bytes(),
                ),
                ("ppt/slideMasters/slideMaster1.xml", master.as_bytes()),
                (
                    "ppt/slideMasters/_rels/slideMaster1.xml.rels",
                    master_rels.as_bytes(),
                ),
                ("ppt/media/layout-bullet.png", PNG),
                ("ppt/media/master-bullet.png", PNG),
                ("ppt/media/unused.png", PNG),
            ],
        );
        let doc = PptxParser::from_bytes(data).unwrap().parse().unwrap();
        let got: Vec<(String, Option<(u8, String)>)> = paragraphs(&doc.sections[0])
            .iter()
            .map(|para| {
                (
                    para.plain_text(),
                    para.list_info
                        .as_ref()
                        .map(|l| (l.level, l.marker_image.clone().unwrap_or_default())),
                )
            })
            .collect();
        let some = |level: u8, img: &str| Some((level, img.to_string()));
        assert_eq!(
            got,
            [
                ("layout one".to_string(), some(0, "layout-bullet.png")),
                // The layout's character replaces the master's picture: a plain bullet, with no
                // picture to trace.
                ("layout char overrides master".to_string(), some(1, "")),
                ("no bullet".to_string(), None),
                ("master one".to_string(), some(0, "master-bullet.png")),
                ("master two".to_string(), some(1, "master-bullet.png")),
                ("A title".to_string(), None),
                ("text box".to_string(), None),
            ]
        );
        assert!(doc.resources.contains_key("layout-bullet.png"));
        assert!(doc.resources.contains_key("master-bullet.png"));
        assert!(
            !doc.resources.contains_key("unused.png"),
            "resources: {:?}",
            doc.resources.keys()
        );
    }

    /// A one-slide deck whose slide points at `slideLayout1.xml`, with the layout's own
    /// relationships part holding `layout_rels`.
    fn deck_with_layout(layout_rels: &[u8]) -> Vec<u8> {
        let slide_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"#;
        let layout: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#;
        deck(
            &[(&slide(&text_shape("", "body")), slide_rels)],
            &[
                ("ppt/slideLayouts/slideLayout1.xml", layout),
                ("ppt/slideLayouts/_rels/slideLayout1.xml.rels", layout_rels),
            ],
        )
    }

    /// Control: with readable layout relationships the deck parses, so the test below is
    /// about the relationships part and not about the layout path being skipped.
    #[test]
    fn test_deck_with_a_layout_parses() {
        let doc = PptxParser::from_bytes(deck_with_layout(EMPTY_RELS.as_bytes()))
            .unwrap()
            .parse()
            .unwrap();

        assert_eq!(paragraphs(&doc.sections[0])[0].plain_text(), "body");
    }

    /// A layout's relationships part is optional, but one that is present and unreadable is
    /// damage, not absence -- the contract every other optional part in this crate keeps.
    /// It used to be read as "no relationships", which silently dropped inherited
    /// placeholders.
    #[test]
    fn test_malformed_layout_relationships_are_reported() {
        let data = deck_with_layout(b"<Relationships>caf\xe9</Relationships>");

        let err = PptxParser::from_bytes(data).unwrap().parse().unwrap_err();

        assert!(matches!(err, Error::Encoding(_)), "got {err:?}");
    }

    /// Helper to create a minimal PPTX in memory with given slide XML content.
    fn create_minimal_pptx(slide_xml: &str) -> Vec<u8> {
        create_minimal_pptx_with_relationships(
            slide_xml,
            Some(
                r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>"#,
            ),
            Some(
                r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#,
            ),
        )
    }

    fn create_minimal_pptx_with_notes(slide_xml: &str, notes_xml: &str) -> Vec<u8> {
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
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>
</Types>"#).unwrap();

        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"#).unwrap();

        zip.start_file("ppt/_rels/presentation.xml.rels", options)
            .unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>"#).unwrap();

        zip.start_file("ppt/presentation.xml", options).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
  </p:sldIdLst>
</p:presentation>"#,
        )
        .unwrap();

        zip.start_file("ppt/slides/_rels/slide1.xml.rels", options)
            .unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#,
        )
        .unwrap();

        zip.start_file("ppt/slides/slide1.xml", options).unwrap();
        zip.write_all(slide_xml.as_bytes()).unwrap();

        zip.start_file("ppt/notesSlides/notesSlide1.xml", options)
            .unwrap();
        zip.write_all(notes_xml.as_bytes()).unwrap();

        zip.finish().unwrap().into_inner()
    }

    fn create_minimal_pptx_with_relationships(
        slide_xml: &str,
        presentation_rels_xml: Option<&str>,
        slide_rels_xml: Option<&str>,
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
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>"#).unwrap();

        // _rels/.rels
        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"#).unwrap();

        if let Some(presentation_rels_xml) = presentation_rels_xml {
            // ppt/_rels/presentation.xml.rels
            zip.start_file("ppt/_rels/presentation.xml.rels", options)
                .unwrap();
            zip.write_all(presentation_rels_xml.as_bytes()).unwrap();
        }

        // ppt/presentation.xml
        zip.start_file("ppt/presentation.xml", options).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
  </p:sldIdLst>
</p:presentation>"#,
        )
        .unwrap();

        if let Some(slide_rels_xml) = slide_rels_xml {
            // ppt/slides/_rels/slide1.xml.rels
            zip.start_file("ppt/slides/_rels/slide1.xml.rels", options)
                .unwrap();
            zip.write_all(slide_rels_xml.as_bytes()).unwrap();
        }

        // ppt/slides/slide1.xml
        zip.start_file("ppt/slides/slide1.xml", options).unwrap();
        zip.write_all(slide_xml.as_bytes()).unwrap();

        zip.finish().unwrap().into_inner()
    }

    #[test]
    fn test_pptx_chart_invalid_numeric_value_propagates_error() {
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

        let slide_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <p:cSld><p:spTree>
    <p:graphicFrame><a:graphic><a:graphicData>
      <c:chart r:id="rIdChart"/>
    </a:graphicData></a:graphic></p:graphicFrame>
  </p:spTree></p:cSld>
</p:sld>"#;

        let slide_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>"#;

        let presentation_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst>
</p:presentation>"#;

        let presentation_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSlide" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
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
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
</Types>"#).unwrap();

        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"#).unwrap();

        zip.start_file("ppt/presentation.xml", options).unwrap();
        zip.write_all(presentation_xml.as_bytes()).unwrap();

        zip.start_file("ppt/_rels/presentation.xml.rels", options)
            .unwrap();
        zip.write_all(presentation_rels.as_bytes()).unwrap();

        zip.start_file("ppt/slides/slide1.xml", options).unwrap();
        zip.write_all(slide_xml.as_bytes()).unwrap();

        zip.start_file("ppt/slides/_rels/slide1.xml.rels", options)
            .unwrap();
        zip.write_all(slide_rels.as_bytes()).unwrap();

        zip.start_file("ppt/charts/chart1.xml", options).unwrap();
        zip.write_all(chart_xml.as_bytes()).unwrap();

        let data = zip.finish().unwrap().into_inner();
        let mut parser = PptxParser::from_bytes(data).unwrap();
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

    /// A presentation of one slide that draws the charts named by `drawn` (relationship ids, in
    /// drawing order), each relationship `rIdN` being `ppt/charts/chartN.xml` whose single
    /// series is named `SeriesN`.
    fn presentation_with_charts(drawn: &[u32]) -> Vec<u8> {
        use std::io::{Cursor, Write};
        use zip::write::SimpleFileOptions;

        let frames: String = drawn
            .iter()
            .map(|n| {
                format!(
                    r#"<p:graphicFrame><a:graphic><a:graphicData><c:chart r:id="rId{n}"/></a:graphicData></a:graphic></p:graphicFrame>"#
                )
            })
            .collect();
        let slide_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <p:cSld><p:spTree>{frames}</p:spTree></p:cSld>
</p:sld>"#
        );
        let mut sorted = drawn.to_vec();
        sorted.sort_unstable();
        let slide_rels: String = sorted
            .iter()
            .map(|n| {
                format!(
                    r#"<Relationship Id="rId{n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart{n}.xml"/>"#
                )
            })
            .collect();

        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let options =
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
        let mut put = |name: &str, body: &str| {
            zip.start_file(name, options).unwrap();
            zip.write_all(body.as_bytes()).unwrap();
        };
        put(
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
</Types>"#,
        );
        put(
            "_rels/.rels",
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"#,
        );
        put(
            "ppt/presentation.xml",
            r#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst>
</p:presentation>"#,
        );
        put(
            "ppt/_rels/presentation.xml.rels",
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSlide" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>"#,
        );
        put("ppt/slides/slide1.xml", &slide_xml);
        put(
            "ppt/slides/_rels/slide1.xml.rels",
            &format!(
                r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{slide_rels}</Relationships>"#
            ),
        );
        for n in &sorted {
            put(
                &format!("ppt/charts/chart{n}.xml"),
                &format!(
                    r#"<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart><c:plotArea><c:lineChart>
    <c:ser>
      <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Series{n}</c:v></c:pt></c:strCache></c:strRef></c:tx>
      <c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt></c:strCache></c:strRef></c:cat>
      <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>{n}</c:v></c:pt></c:numCache></c:numRef></c:val>
    </c:ser>
  </c:lineChart></c:plotArea></c:chart>
</c:chartSpace>"#
                ),
            );
        }
        zip.finish().unwrap().into_inner()
    }

    /// Reading the same presentation again gives the same output: charts drawn in an order that
    /// is neither the order of their ids nor any order a map hands out, and picture bullets
    /// that make several resources.
    #[test]
    fn test_the_same_presentation_gives_the_same_output_every_time() {
        let data = presentation_with_charts(&[5, 2, 4, 1, 3]);
        crate::assert_deterministic("pptx", || {
            PptxParser::from_bytes(data.clone())
                .unwrap()
                .parse()
                .unwrap()
        });
    }

    #[test]
    fn test_charts_come_in_the_order_the_slide_draws_them() {
        // Drawn 3, 1, 4, 2: not the order of their ids, and not any order a map of the
        // relationships hands out — which differs from one run to the next, so the same file
        // is read several times.
        for _ in 0..8 {
            let mut parser =
                PptxParser::from_bytes(presentation_with_charts(&[3, 1, 4, 2])).unwrap();
            let text = parser.parse().unwrap().plain_text();
            let at = |n: u32| {
                text.find(&format!("Series{n}"))
                    .unwrap_or_else(|| panic!("Series{n} missing from: {text}"))
            };
            assert!(
                at(3) < at(1) && at(1) < at(4) && at(4) < at(2),
                "charts out of drawing order: {text}"
            );
        }
    }

    /// A slide's text, pictures and tables follow the order its shapes are listed in, not one
    /// kind after another: the caption sits next to its picture and the text drawn after a
    /// table stays after it.
    #[test]
    fn test_slide_content_follows_the_drawing_order_of_its_shapes() {
        let shape = |id: u32, text: &str| {
            format!(
                r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="S{id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"#
            )
        };
        let pic = |id: u32, rel: &str| {
            format!(
                r#"<p:pic><p:nvPicPr><p:cNvPr id="{id}" name="P{id}" descr="pic {rel}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="{rel}"/></p:blipFill><p:spPr/></p:pic>"#
            )
        };
        let table = r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="T"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><a:graphic><a:graphicData><a:tbl><a:tblGrid><a:gridCol w="100"/></a:tblGrid><a:tr h="10"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>Cell</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#;
        // Listed: text, picture, text, table, text, picture.
        let shapes = [
            shape(2, "Title"),
            pic(3, "rIdA"),
            shape(4, "Caption"),
            table.to_string(),
            shape(5, "After"),
            pic(6, "rIdB"),
        ]
        .concat();
        let rels = rels_xml(&[
            ("rIdA", IMAGE_REL, "../media/a.png"),
            ("rIdB", IMAGE_REL, "../media/b.png"),
        ]);
        let data = deck(
            &[(&slide(&shapes), &rels)],
            &[("ppt/media/a.png", PNG), ("ppt/media/b.png", PNG)],
        );
        let doc = PptxParser::from_bytes(data).unwrap().parse().unwrap();
        let order: Vec<String> = doc.sections[0]
            .content
            .iter()
            .map(|block| match block {
                Block::Paragraph(p) => p.plain_text(),
                Block::Table(t) => format!("table:{}", t.rows[0].cells[0].plain_text()),
                Block::Image { resource_id, .. } => format!("image:{resource_id}"),
                other => panic!("unexpected block: {other:?}"),
            })
            .collect();
        assert_eq!(
            order,
            [
                "Title",
                "image:a.png",
                "Caption",
                "table:Cell",
                "After",
                "image:b.png"
            ]
        );
    }

    #[test]
    fn test_pptx_missing_chart_part_propagates_error() {
        use std::io::{Cursor, Write};
        use zip::write::SimpleFileOptions;

        let slide_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <p:cSld><p:spTree>
    <p:graphicFrame><a:graphic><a:graphicData>
      <c:chart r:id="rIdChart"/>
    </a:graphicData></a:graphic></p:graphicFrame>
  </p:spTree></p:cSld>
</p:sld>"#;

        let slide_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>"#;

        let presentation_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst>
</p:presentation>"#;

        let presentation_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSlide" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
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
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
</Types>"#).unwrap();

        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"#).unwrap();

        zip.start_file("ppt/presentation.xml", options).unwrap();
        zip.write_all(presentation_xml.as_bytes()).unwrap();

        zip.start_file("ppt/_rels/presentation.xml.rels", options)
            .unwrap();
        zip.write_all(presentation_rels.as_bytes()).unwrap();

        zip.start_file("ppt/slides/slide1.xml", options).unwrap();
        zip.write_all(slide_xml.as_bytes()).unwrap();

        zip.start_file("ppt/slides/_rels/slide1.xml.rels", options)
            .unwrap();
        zip.write_all(slide_rels.as_bytes()).unwrap();

        let data = zip.finish().unwrap().into_inner();
        let mut parser = PptxParser::from_bytes(data).unwrap();
        let err = parser
            .parse()
            .expect_err("missing referenced chart part must surface");

        match err {
            Error::MissingComponent(path) => assert_eq!(path, "ppt/charts/chart1.xml"),
            other => panic!("expected MissingComponent, got {other:?}"),
        }
    }

    #[test]
    fn test_pptx_slide_table_preserves_raw_malformed_entity() {
        let slide_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:graphicFrame><a:graphic><a:graphicData>
      <a:tbl>
        <a:tr><a:tc><a:txBody><a:p><a:r><a:t>Cell &bogus; text</a:t></a:r></a:p></a:txBody></a:tc></a:tr>
      </a:tbl>
    </a:graphicData></a:graphic></p:graphicFrame>
  </p:spTree></p:cSld>
</p:sld>"#;

        let data = create_minimal_pptx(slide_xml);
        let mut parser = PptxParser::from_bytes(data).unwrap();
        let doc = parser.parse().unwrap();
        assert!(
            doc.plain_text().contains("Cell &bogus; text"),
            "expected raw malformed entity preserved in slide table, got: {}",
            doc.plain_text()
        );
    }

    #[test]
    fn test_pptx_slide_text_preserves_raw_malformed_entity() {
        let slide_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:sp><p:txBody>
      <a:p><a:r><a:t>Slide &bogus; body</a:t></a:r></a:p>
    </p:txBody></p:sp>
  </p:spTree></p:cSld>
</p:sld>"#;

        let data = create_minimal_pptx(slide_xml);
        let mut parser = PptxParser::from_bytes(data).unwrap();
        let doc = parser.parse().unwrap();
        assert!(
            doc.plain_text().contains("Slide &bogus; body"),
            "expected raw malformed entity preserved in slide text, got: {}",
            doc.plain_text()
        );
    }

    #[test]
    fn test_pptx_notes_preserves_raw_malformed_entity() {
        let slide_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree/></p:cSld>
</p:sld>"#;

        let notes_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
         xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:sp><p:txBody>
      <a:p><a:r><a:t>Note &bogus; body</a:t></a:r></a:p>
    </p:txBody></p:sp>
  </p:spTree></p:cSld>
</p:notes>"#;

        let data = create_minimal_pptx_with_notes(slide_xml, notes_xml);
        let mut parser = PptxParser::from_bytes(data).unwrap();
        let doc = parser.parse().unwrap();
        let notes = doc.sections[0]
            .notes
            .as_ref()
            .expect("notes should be parsed");
        let notes_text = notes
            .iter()
            .map(Paragraph::plain_text)
            .collect::<Vec<_>>()
            .join("\n");
        assert!(
            notes_text.contains("Note &bogus; body"),
            "expected raw malformed entity preserved in notes, got: {}",
            notes_text
        );
    }

    #[test]
    fn test_pptx_requires_presentation_relationships() {
        let slide_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#;

        let data = create_minimal_pptx_with_relationships(slide_xml, None, None);
        let err = PptxParser::from_bytes(data)
            .err()
            .expect("missing presentation relationships should fail");

        match err {
            Error::MissingComponent(path) => assert_eq!(path, "ppt/_rels/presentation.xml.rels"),
            other => panic!("expected missing presentation rels error, got {other:?}"),
        }
    }

    #[test]
    fn test_pptx_rejects_malformed_presentation_relationships() {
        let slide_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#;

        let data = create_minimal_pptx_with_relationships(slide_xml, Some("<Relationships"), None);
        let err = PptxParser::from_bytes(data)
            .err()
            .expect("malformed presentation relationships should fail");

        match err {
            Error::XmlParseWithContext { location, .. } => {
                assert_eq!(location, "ppt/_rels/presentation.xml.rels")
            }
            other => panic!("expected malformed presentation rels error, got {other:?}"),
        }
    }

    #[test]
    fn test_pptx_non_utf8_presentation_is_error() {
        use std::io::{Cursor, Write};
        use zip::write::SimpleFileOptions;

        let buf = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(buf);
        let options =
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);

        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
</Types>"#).unwrap();

        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"#).unwrap();

        zip.start_file("ppt/_rels/presentation.xml.rels", options)
            .unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#,
        )
        .unwrap();

        zip.start_file("ppt/presentation.xml", options).unwrap();
        zip.write_all(b"<?xml version=\"1.0\"?><presentation>Caf\xe9</presentation>")
            .unwrap();

        let data = zip.finish().unwrap().into_inner();
        let err = match PptxParser::from_bytes(data) {
            Ok(_) => panic!("non-UTF-8 presentation must surface Error::Encoding"),
            Err(err) => err,
        };
        assert!(
            matches!(err, Error::Encoding(_)),
            "expected Error::Encoding, got {err:?}"
        );
    }

    fn create_minimal_pptx_with_malformed_part(malformed_part_path: &str) -> Vec<u8> {
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
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>
</Types>"#).unwrap();

        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"#).unwrap();

        zip.start_file("ppt/_rels/presentation.xml.rels", options)
            .unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>"#).unwrap();

        zip.start_file("ppt/presentation.xml", options).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
  </p:sldIdLst>
</p:presentation>"#,
        )
        .unwrap();

        zip.start_file("ppt/slides/_rels/slide1.xml.rels", options)
            .unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#,
        )
        .unwrap();

        let valid_slide = br#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree/></p:cSld>
</p:sld>"#;
        let valid_notes = br#"<?xml version="1.0" encoding="UTF-8"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
         xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree/></p:cSld>
</p:notes>"#;
        let malformed = b"<?xml version=\"1.0\"?><root>Caf\xe9</root>";

        zip.start_file("ppt/slides/slide1.xml", options).unwrap();
        if malformed_part_path == "ppt/slides/slide1.xml" {
            zip.write_all(malformed).unwrap();
        } else {
            zip.write_all(valid_slide).unwrap();
        }

        zip.start_file("ppt/notesSlides/notesSlide1.xml", options)
            .unwrap();
        if malformed_part_path == "ppt/notesSlides/notesSlide1.xml" {
            zip.write_all(malformed).unwrap();
        } else {
            zip.write_all(valid_notes).unwrap();
        }

        zip.finish().unwrap().into_inner()
    }

    #[test]
    fn test_pptx_non_utf8_optional_parts_surface_encoding_error() {
        // Slide and notes parts with malformed (non-UTF-8/UTF-16) byte content
        // must surface Error::Encoding instead of being silently dropped.
        for part_path in &["ppt/slides/slide1.xml", "ppt/notesSlides/notesSlide1.xml"] {
            let data = create_minimal_pptx_with_malformed_part(part_path);
            let mut parser = PptxParser::from_bytes(data).expect("constructor must succeed");
            let err = match parser.parse() {
                Ok(_) => panic!("malformed {part_path} must surface Error::Encoding"),
                Err(err) => err,
            };
            assert!(
                matches!(err, Error::Encoding(_)),
                "expected Error::Encoding for {part_path}, got {err:?}"
            );
        }
    }

    #[test]
    fn test_pptx_allows_missing_optional_slide_relationships() {
        let slide_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree/></p:cSld>
</p:sld>"#;

        let data = create_minimal_pptx_with_relationships(
            slide_xml,
            Some(
                r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"></Relationship>
</Relationships>"#,
            ),
            None,
        );
        let mut parser = PptxParser::from_bytes(data).unwrap();
        let doc = parser.parse().unwrap();

        assert_eq!(doc.sections.len(), 1);
    }

    #[test]
    fn test_pptx_rejects_malformed_optional_slide_relationships() {
        let slide_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:txBody>
          <a:p><a:r><a:t>Hello from slide</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>"#;

        let data = create_minimal_pptx_with_relationships(
            slide_xml,
            Some(
                r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>"#,
            ),
            Some("<Relationships"),
        );
        let mut parser = PptxParser::from_bytes(data).unwrap();
        let err = parser
            .parse()
            .expect_err("malformed optional slide relationships should fail");

        match err {
            Error::XmlParseWithContext { location, .. } => {
                assert_eq!(location, "ppt/slides/_rels/slide1.xml.rels")
            }
            other => panic!("expected malformed optional slide rels error, got {other:?}"),
        }
    }

    #[test]
    fn test_parse_grouped_shapes() {
        let slide_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:txBody>
          <a:p><a:r><a:t>Top level shape</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:grpSp>
        <p:nvGrpSpPr>
          <p:cNvPr id="10" name="Group 1"/>
          <p:cNvGrpSpPr/>
          <p:nvPr/>
        </p:nvGrpSpPr>
        <p:grpSpPr/>
        <p:sp>
          <p:txBody>
            <a:p><a:r><a:t>Grouped shape 1</a:t></a:r></a:p>
          </p:txBody>
        </p:sp>
        <p:sp>
          <p:txBody>
            <a:p><a:r><a:t>Grouped shape 2</a:t></a:r></a:p>
          </p:txBody>
        </p:sp>
      </p:grpSp>
    </p:spTree>
  </p:cSld>
</p:sld>"#;

        let data = create_minimal_pptx(slide_xml);
        let mut parser = PptxParser::from_bytes(data).unwrap();
        let doc = parser.parse().unwrap();
        let text = doc.plain_text();

        assert!(
            text.contains("Top level shape"),
            "Should contain top-level shape text, got: {}",
            text
        );
        assert!(
            text.contains("Grouped shape 1"),
            "Should contain first grouped shape text, got: {}",
            text
        );
        assert!(
            text.contains("Grouped shape 2"),
            "Should contain second grouped shape text, got: {}",
            text
        );
    }

    #[test]
    fn test_parse_nested_grouped_shapes() {
        let slide_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:grpSp>
        <p:nvGrpSpPr><p:cNvPr id="10" name="Outer Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
        <p:grpSpPr/>
        <p:sp>
          <p:txBody>
            <a:p><a:r><a:t>Outer group text</a:t></a:r></a:p>
          </p:txBody>
        </p:sp>
        <p:grpSp>
          <p:nvGrpSpPr><p:cNvPr id="20" name="Inner Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
          <p:grpSpPr/>
          <p:sp>
            <p:txBody>
              <a:p><a:r><a:t>Inner group text</a:t></a:r></a:p>
            </p:txBody>
          </p:sp>
        </p:grpSp>
      </p:grpSp>
    </p:spTree>
  </p:cSld>
</p:sld>"#;

        let data = create_minimal_pptx(slide_xml);
        let mut parser = PptxParser::from_bytes(data).unwrap();
        let doc = parser.parse().unwrap();
        let text = doc.plain_text();

        assert!(
            text.contains("Outer group text"),
            "Should contain outer group shape text, got: {}",
            text
        );
        assert!(
            text.contains("Inner group text"),
            "Should contain inner (nested) group shape text, got: {}",
            text
        );
    }

    #[test]
    fn test_pptx_slide_mixed_entities_preserve_legitimate_and_malformed() {
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
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>"#).unwrap();

            zip.start_file("_rels/.rels", options).unwrap();
            zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"#).unwrap();

            zip.start_file("ppt/_rels/presentation.xml.rels", options)
                .unwrap();
            zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>"#).unwrap();

            zip.start_file("ppt/presentation.xml", options).unwrap();
            zip.write_all(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
</p:presentation>"#,
            )
            .unwrap();

            zip.start_file("ppt/slides/slide1.xml", options).unwrap();
            zip.write_all(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:sp><p:txBody><a:p><a:r><a:t>A &amp; B &bogus; C</a:t></a:r></a:p></p:txBody></p:sp>
  </p:spTree></p:cSld>
</p:sld>"#,
            )
            .unwrap();

            zip.finish().unwrap();
        }

        let mut parser = PptxParser::from_bytes(buf).expect("parser opens");
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
    fn test_pptx_missing_presentation_surfaces_missing_component() {
        use std::io::Write;

        let mut buf = Vec::new();
        {
            let cursor = std::io::Cursor::new(&mut buf);
            let mut zip = zip::ZipWriter::new(cursor);
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);

            zip.start_file("[Content_Types].xml", options).unwrap();
            zip.write_all(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
</Types>"#,
            )
            .unwrap();

            zip.start_file("_rels/.rels", options).unwrap();
            zip.write_all(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"#,
            )
            .unwrap();

            // ppt/presentation.xml INTENTIONALLY ABSENT.
            // Empty ppt/_rels/presentation.xml.rels so the required rels
            // read succeeds and we reach the presentation-missing path.
            zip.start_file("ppt/_rels/presentation.xml.rels", options)
                .unwrap();
            zip.write_all(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#,
            )
            .unwrap();

            zip.finish().unwrap();
        }

        let err = PptxParser::from_bytes(buf)
            .err()
            .expect("must fail on missing presentation");
        match err {
            Error::MissingComponent(path) => {
                assert_eq!(path, "ppt/presentation.xml");
            }
            other => panic!("expected MissingComponent(\"ppt/presentation.xml\"), got {other:?}"),
        }
    }
}
