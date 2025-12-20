//! DOCX parser implementation.

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
        let main_section = self.parse_document_xml()?;
        doc.add_section(main_section);

        // Extract resources (images)
        self.extract_resources(&mut doc)?;

        Ok(doc)
    }

    /// Parse document metadata from docProps/core.xml.
    fn parse_metadata(&self) -> Result<Metadata> {
        let mut meta = Metadata::default();

        if let Ok(xml) = self.container.read_xml("docProps/core.xml") {
            let mut reader = quick_xml::Reader::from_str(&xml);
            reader.config_mut().trim_text(true);

            let mut buf = Vec::new();
            let mut current_element: Option<String> = None;

            loop {
                match reader.read_event_into(&mut buf) {
                    Ok(quick_xml::events::Event::Start(e)) => {
                        let name = e.name();
                        current_element = Some(
                            String::from_utf8_lossy(name.local_name().as_ref()).to_string(),
                        );
                    }
                    Ok(quick_xml::events::Event::Text(e)) => {
                        if let Some(ref elem) = current_element {
                            let text = e.unescape().unwrap_or_default().to_string();
                            match elem.as_str() {
                                "title" => meta.title = Some(text),
                                "creator" => meta.author = Some(text),
                                "subject" => meta.subject = Some(text),
                                "description" => meta.description = Some(text),
                                "keywords" => {
                                    meta.keywords = text
                                        .split(|c| c == ',' || c == ';')
                                        .map(|s| s.trim().to_string())
                                        .filter(|s| !s.is_empty())
                                        .collect();
                                }
                                "created" => meta.created = Some(text),
                                "modified" => meta.modified = Some(text),
                                _ => {}
                            }
                        }
                    }
                    Ok(quick_xml::events::Event::End(_)) => {
                        current_element = None;
                    }
                    Ok(quick_xml::events::Event::Eof) => break,
                    Err(_) => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(meta)
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
        let mut in_table = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    let name = e.name();
                    match name.as_ref() {
                        b"w:body" => {
                            in_body = true;
                        }
                        b"w:p" if in_body && !in_table => {
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
                            in_table = true;
                            table_xml.clear();
                            table_xml.push_str("<w:tbl>");
                        }
                        _ => {
                            if in_paragraph {
                                paragraph_xml.push('<');
                                paragraph_xml
                                    .push_str(&String::from_utf8_lossy(name.as_ref()));
                                for attr in e.attributes().flatten() {
                                    paragraph_xml.push_str(&format!(
                                        " {}=\"{}\"",
                                        String::from_utf8_lossy(attr.key.as_ref()),
                                        String::from_utf8_lossy(&attr.value)
                                    ));
                                }
                                paragraph_xml.push('>');
                            } else if in_table {
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
                    } else if in_table {
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
                    } else if in_table {
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
                        b"w:p" if in_paragraph && !in_table => {
                            paragraph_xml.push_str("</w:p>");
                            if let Ok(para) = self.parse_paragraph(&paragraph_xml) {
                                section.add_block(Block::Paragraph(para));
                            }
                            in_paragraph = false;
                        }
                        b"w:tbl" if in_table => {
                            table_xml.push_str("</w:tbl>");
                            if let Ok(table) = self.parse_table(&table_xml) {
                                section.add_block(Block::Table(table));
                            }
                            in_table = false;
                        }
                        _ => {
                            if in_paragraph {
                                paragraph_xml.push_str("</");
                                paragraph_xml
                                    .push_str(&String::from_utf8_lossy(name.as_ref()));
                                paragraph_xml.push('>');
                            } else if in_table {
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
        let mut para = Paragraph::new();
        let mut reader = quick_xml::Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut in_ppr = false;
        let mut in_rpr = false;
        let mut in_run = false;
        let mut current_style = TextStyle::default();
        let mut current_hyperlink: Option<String> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    match e.name().as_ref() {
                        b"w:pPr" => in_ppr = true,
                        b"w:rPr" => in_rpr = true,
                        b"w:r" => {
                            in_run = true;
                            current_style = TextStyle::default();
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
                    }
                }
                Ok(quick_xml::events::Event::Empty(ref e)) => {
                    match e.name().as_ref() {
                        b"w:pStyle" if in_ppr => {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"w:val" {
                                    let style_id = String::from_utf8_lossy(&attr.value);
                                    para.style_id = Some(style_id.to_string());
                                    para.heading = self.styles.get_heading_level(&style_id);
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
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Text(ref e)) => {
                    if in_run {
                        let text = e.unescape().unwrap_or_default().to_string();
                        if !text.is_empty() {
                            let run = TextRun {
                                text,
                                style: current_style.clone(),
                                hyperlink: current_hyperlink.clone(),
                            };
                            para.runs.push(run);
                        }
                    }
                }
                Ok(quick_xml::events::Event::End(ref e)) => {
                    match e.name().as_ref() {
                        b"w:pPr" => in_ppr = false,
                        b"w:rPr" => in_rpr = false,
                        b"w:r" => in_run = false,
                        b"w:hyperlink" => current_hyperlink = None,
                        _ => {}
                    }
                }
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
    fn parse_table(&self, xml: &str) -> Result<Table> {
        let mut table = Table::new();
        let mut reader = quick_xml::Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut in_row = false;
        let mut in_cell = false;
        let mut in_paragraph = false;
        let mut current_row: Option<Row> = None;
        let mut cell_text = String::new();
        let mut is_header_row = false;
        let mut col_span = 1u32;
        let mut row_span = 1u32;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    match e.name().as_ref() {
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
                            cell_text.clear();
                            col_span = 1;
                            row_span = 1;
                        }
                        b"w:p" if in_cell => {
                            in_paragraph = true;
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Empty(ref e)) => {
                    match e.name().as_ref() {
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
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Text(ref e)) => {
                    if in_paragraph && in_cell {
                        let text = e.unescape().unwrap_or_default();
                        cell_text.push_str(&text);
                    }
                }
                Ok(quick_xml::events::Event::End(ref e)) => {
                    match e.name().as_ref() {
                        b"w:tr" => {
                            if let Some(mut row) = current_row.take() {
                                row.is_header = is_header_row;
                                table.add_row(row);
                            }
                            in_row = false;
                        }
                        b"w:tc" => {
                            if row_span > 0 {
                                let cell = Cell {
                                    content: vec![Paragraph::with_text(&cell_text)],
                                    col_span,
                                    row_span,
                                    alignment: CellAlignment::Left,
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
                        b"w:p" => {
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
