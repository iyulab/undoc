//! DOCX parser implementation.

use crate::charts;
use crate::container::OoxmlContainer;
use crate::error::{Error, Result};
use crate::model::{
    Block, Cell, CellAlignment, Document, ListInfo, ListType, Metadata, Paragraph, Resource,
    ResourceType, Row, Section, Table, TextAlignment, TextRun, TextStyle, VerticalAlignment,
};

use super::numbering::NumberingMap;
use super::styles::StyleMap;

/// Parser for DOCX (Word) documents.
pub struct DocxParser {
    container: OoxmlContainer,
    styles: StyleMap,
    numbering: NumberingMap,
    relationships: crate::container::Relationships,
}

impl DocxParser {
    /// Open a DOCX file for parsing.
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
        // Parse styles
        let styles = if let Ok(xml) = container.read_xml("word/styles.xml") {
            StyleMap::parse(&xml)?
        } else {
            StyleMap::default()
        };

        // Parse numbering
        let numbering = if let Ok(xml) = container.read_xml("word/numbering.xml") {
            NumberingMap::parse(&xml)?
        } else {
            NumberingMap::default()
        };

        // Parse document relationships
        let relationships = container
            .read_relationships("word/document.xml")
            .unwrap_or_default();

        Ok(Self {
            container,
            styles,
            numbering,
            relationships,
        })
    }

    /// Parse the document and return a Document model.
    pub fn parse(&mut self) -> Result<Document> {
        let mut doc = Document::new();

        // Parse metadata
        doc.metadata = self.parse_metadata()?;

        // Parse main document content
        let mut main_section = self.parse_document_xml()?;

        // Parse charts and add as tables for RAG-ready output
        let chart_tables = self.parse_charts()?;
        for table in chart_tables {
            main_section.add_block(Block::Table(table));
        }

        doc.add_section(main_section);

        // Extract resources (images)
        self.extract_resources(&mut doc)?;

        Ok(doc)
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
                if let Ok(chart_xml) = self.container.read_xml(&chart_path) {
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
                                                first_cell.content.push(Paragraph::with_text(
                                                    format!("{} ({})", original, title),
                                                ));
                                            }
                                        }
                                    }
                                }
                                tables.push(table);
                            }
                        }
                        Err(_) => {
                            // Chart parsing failed, skip
                        }
                    }
                }
            }
        }

        Ok(tables)
    }

    /// Parse the main document.xml content.
    fn parse_document_xml(&mut self) -> Result<Section> {
        let xml = self.container.read_xml("word/document.xml")?;
        let mut section = Section::new(0);

        let mut reader = quick_xml::Reader::from_str(&xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut in_body = false;
        let mut paragraph_xml = String::new();
        let mut table_xml = String::new();
        let mut in_paragraph = false;
        let mut table_depth: u32 = 0; // Track nested table depth

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    let name = e.name();
                    match name.as_ref() {
                        b"w:body" => {
                            in_body = true;
                        }
                        b"w:p" if in_body && table_depth == 0 => {
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
                        b"w:tbl" if in_body => {
                            if table_depth == 0 {
                                // Start collecting table XML
                                table_xml.clear();
                            }
                            table_depth += 1;
                            table_xml.push_str("<w:tbl>");
                        }
                        _ => {
                            if in_paragraph {
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
                    if in_paragraph {
                        let name = e.name();
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
                        let name = e.name();
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
                        let text = e.unescape().unwrap_or_default();
                        paragraph_xml.push_str(&escape_xml(&text));
                    } else if table_depth > 0 {
                        let text = e.unescape().unwrap_or_default();
                        table_xml.push_str(&escape_xml(&text));
                    }
                }
                Ok(quick_xml::events::Event::End(ref e)) => {
                    let name = e.name();
                    match name.as_ref() {
                        b"w:body" => {
                            in_body = false;
                        }
                        b"w:p" if in_paragraph && table_depth == 0 => {
                            paragraph_xml.push_str("</w:p>");
                            if let Ok(para) = self.parse_paragraph(&paragraph_xml) {
                                section.add_block(Block::Paragraph(para));
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
                Err(e) => return Err(Error::XmlParse(e.to_string())),
                _ => {}
            }
            buf.clear();
        }

        Ok(section)
    }

    /// Parse a single paragraph element.
    fn parse_paragraph(&mut self, xml: &str) -> Result<Paragraph> {
        use crate::model::InlineImage;

        let mut para = Paragraph::new();
        let mut reader = quick_xml::Reader::from_str(xml);
        // Don't trim text - preserve whitespace from xml:space="preserve" elements
        reader.config_mut().trim_text(false);

        let mut buf = Vec::new();
        let mut in_ppr = false;
        let mut in_rpr = false;
        let mut in_run = false;
        let mut in_text = false; // Track w:t elements (regular text)
        let mut in_instr_text = false; // Track w:instrText elements (field codes to skip)
        let mut in_drawing = false; // Track w:drawing elements for images
        let mut current_style = TextStyle::default();
        let mut current_hyperlink: Option<String> = None;
        let mut current_image_alt: Option<String> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => match e.name().as_ref() {
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
                    // Line break handling
                    b"w:br" if in_run => {
                        // Mark the last run as having a line break, or create an empty run with line break
                        if let Some(last_run) = para.runs.last_mut() {
                            last_run.line_break = true;
                        } else {
                            // No previous run, create an empty run with line break
                            para.runs.push(TextRun {
                                text: String::new(),
                                style: current_style.clone(),
                                hyperlink: None,
                                line_break: true,
                            });
                        }
                    }
                    _ => {}
                },
                Ok(quick_xml::events::Event::Text(ref e)) => {
                    // Only extract text from w:t elements, skip w:instrText (field codes)
                    if in_run && in_text && !in_instr_text {
                        let text = e.unescape().unwrap_or_default().to_string();
                        if !text.is_empty() {
                            let run = TextRun {
                                text,
                                style: current_style.clone(),
                                hyperlink: current_hyperlink.clone(),
                                line_break: false,
                            };
                            para.runs.push(run);
                        }
                    }
                }
                Ok(quick_xml::events::Event::End(ref e)) => match e.name().as_ref() {
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
                    _ => {}
                },
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => return Err(Error::XmlParse(e.to_string())),
                _ => {}
            }
            buf.clear();
        }

        // Parse numbering (list info)
        para.list_info = self.parse_list_info(xml);

        Ok(para)
    }

    /// Parse list info from paragraph XML.
    fn parse_list_info(&mut self, xml: &str) -> Option<ListInfo> {
        let mut reader = quick_xml::Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut num_id: Option<String> = None;
        let mut level: u8 = 0;
        let mut in_num_pr = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    if e.name().as_ref() == b"w:numPr" {
                        in_num_pr = true;
                    }
                }
                Ok(quick_xml::events::Event::Empty(ref e)) => {
                    if in_num_pr {
                        match e.name().as_ref() {
                            b"w:numId" => {
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"w:val" {
                                        num_id =
                                            Some(String::from_utf8_lossy(&attr.value).to_string());
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
                }
                Ok(quick_xml::events::Event::End(ref e)) => {
                    if e.name().as_ref() == b"w:numPr" {
                        in_num_pr = false;
                    }
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
        let mut reader = quick_xml::Reader::from_str(xml);
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
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Text(ref e)) => {
                    // If we're inside a nested table, just collect XML
                    if nested_table_depth > 0 {
                        let text = e.unescape().unwrap_or_default();
                        nested_table_xml.push_str(&escape_xml(&text));
                        continue;
                    }

                    // Only extract text from w:t elements, skip w:instrText (field codes)
                    if in_run && in_text && !in_instr_text {
                        let text = e.unescape().unwrap_or_default().to_string();
                        if !text.is_empty() {
                            if let Some(ref mut para) = current_paragraph {
                                let run = TextRun {
                                    text,
                                    style: current_style.clone(),
                                    hyperlink: None,
                                    line_break: false,
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
                                if let Some(ref mut row) = current_row {
                                    row.cells.push(cell);
                                }
                            }
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
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => return Err(Error::XmlParse(e.to_string())),
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
}
