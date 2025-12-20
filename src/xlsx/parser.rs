//! XLSX parser implementation.

use crate::container::OoxmlContainer;
use crate::error::{Error, Result};
use crate::model::{
    Block, Cell, CellAlignment, Document, Metadata, Paragraph, Row, Section, Table, TextRun,
};
use std::collections::HashMap;
use std::path::Path;

use super::shared_strings::SharedStrings;

/// Sheet info from workbook.xml.
#[derive(Debug, Clone)]
struct SheetInfo {
    name: String,
    #[allow(dead_code)]
    sheet_id: String,
    rel_id: String,
}

/// Parser for XLSX (Excel) workbooks.
pub struct XlsxParser {
    container: OoxmlContainer,
    shared_strings: SharedStrings,
    sheets: Vec<SheetInfo>,
    relationships: HashMap<String, String>,
}

impl XlsxParser {
    /// Open an XLSX file for parsing.
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
        // Parse shared strings
        let shared_strings = if let Ok(xml) = container.read_xml("xl/sharedStrings.xml") {
            SharedStrings::parse(&xml)?
        } else {
            SharedStrings::default()
        };

        // Parse workbook relationships
        let relationships = Self::parse_workbook_rels(&container)?;

        // Parse workbook for sheet info
        let sheets = Self::parse_workbook(&container)?;

        Ok(Self {
            container,
            shared_strings,
            sheets,
            relationships,
        })
    }

    /// Parse workbook relationships.
    fn parse_workbook_rels(container: &OoxmlContainer) -> Result<HashMap<String, String>> {
        let mut rels = HashMap::new();

        if let Ok(xml) = container.read_xml("xl/_rels/workbook.xml.rels") {
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

    /// Parse workbook.xml for sheet info.
    fn parse_workbook(container: &OoxmlContainer) -> Result<Vec<SheetInfo>> {
        let mut sheets = Vec::new();

        if let Ok(xml) = container.read_xml("xl/workbook.xml") {
            let mut reader = quick_xml::Reader::from_str(&xml);
            reader.config_mut().trim_text(true);

            let mut buf = Vec::new();

            loop {
                match reader.read_event_into(&mut buf) {
                    Ok(quick_xml::events::Event::Empty(e))
                    | Ok(quick_xml::events::Event::Start(e)) => {
                        if e.name().as_ref() == b"sheet" {
                            let mut name = String::new();
                            let mut sheet_id = String::new();
                            let mut rel_id = String::new();

                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"name" => {
                                        name = String::from_utf8_lossy(&attr.value).to_string();
                                    }
                                    b"sheetId" => {
                                        sheet_id =
                                            String::from_utf8_lossy(&attr.value).to_string();
                                    }
                                    b"r:id" => {
                                        rel_id = String::from_utf8_lossy(&attr.value).to_string();
                                    }
                                    _ => {}
                                }
                            }

                            if !name.is_empty() {
                                sheets.push(SheetInfo {
                                    name,
                                    sheet_id,
                                    rel_id,
                                });
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

        Ok(sheets)
    }

    /// Parse the workbook and return a Document model.
    pub fn parse(&mut self) -> Result<Document> {
        let mut doc = Document::new();

        // Parse metadata
        doc.metadata = self.parse_metadata()?;

        // Parse each sheet as a section with a table
        for (idx, sheet) in self.sheets.clone().iter().enumerate() {
            let mut section = Section::new(idx);
            section.name = Some(sheet.name.clone());

            // Get the sheet path from relationships
            if let Some(target) = self.relationships.get(&sheet.rel_id) {
                let sheet_path = if target.starts_with('/') {
                    target[1..].to_string()
                } else {
                    format!("xl/{}", target)
                };

                if let Ok(xml) = self.container.read_xml(&sheet_path) {
                    if let Ok(table) = self.parse_sheet(&xml) {
                        section.add_block(Block::Table(table));
                    }
                }
            }

            doc.add_section(section);
        }

        Ok(doc)
    }

    /// Parse metadata from docProps/core.xml.
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

        // Set sheet count
        meta.page_count = Some(self.sheets.len() as u32);

        Ok(meta)
    }

    /// Parse a worksheet XML into a table.
    fn parse_sheet(&self, xml: &str) -> Result<Table> {
        let mut table = Table::new();
        let mut reader = quick_xml::Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut in_row = false;
        let mut in_cell = false;
        let mut in_value = false;
        let mut current_row: Option<Row> = None;
        let mut current_cell_type: Option<String> = None;
        let mut current_cell_value = String::new();
        let mut is_first_row = true;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    match e.name().as_ref() {
                        b"row" => {
                            in_row = true;
                            current_row = Some(Row {
                                cells: Vec::new(),
                                is_header: is_first_row,
                                height: None,
                            });
                        }
                        b"c" if in_row => {
                            in_cell = true;
                            current_cell_type = None;
                            current_cell_value.clear();

                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"t" {
                                    current_cell_type =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }
                        b"v" if in_cell => {
                            in_value = true;
                        }
                        b"t" if in_cell => {
                            // Inline string
                            in_value = true;
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Text(ref e)) => {
                    if in_value {
                        let text = e.unescape().unwrap_or_default();
                        current_cell_value.push_str(&text);
                    }
                }
                Ok(quick_xml::events::Event::End(ref e)) => {
                    match e.name().as_ref() {
                        b"row" => {
                            if let Some(row) = current_row.take() {
                                if !row.cells.is_empty() {
                                    table.add_row(row);
                                }
                            }
                            in_row = false;
                            is_first_row = false;
                        }
                        b"c" => {
                            // Resolve the cell value
                            let value = self.resolve_cell_value(
                                &current_cell_value,
                                current_cell_type.as_deref(),
                            );

                            let cell = Cell {
                                content: vec![Paragraph {
                                    runs: vec![TextRun::plain(&value)],
                                    ..Default::default()
                                }],
                                col_span: 1,
                                row_span: 1,
                                alignment: CellAlignment::Left,
                                vertical_alignment: Default::default(),
                                is_header: is_first_row,
                                background: None,
                            };

                            if let Some(ref mut row) = current_row {
                                row.cells.push(cell);
                            }

                            in_cell = false;
                        }
                        b"v" | b"t" => {
                            in_value = false;
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

    /// Resolve a cell value based on its type.
    fn resolve_cell_value(&self, value: &str, cell_type: Option<&str>) -> String {
        match cell_type {
            Some("s") => {
                // Shared string index
                if let Ok(idx) = value.parse::<usize>() {
                    self.shared_strings.get(idx).unwrap_or("").to_string()
                } else {
                    value.to_string()
                }
            }
            Some("b") => {
                // Boolean
                if value == "1" {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
                }
            }
            Some("e") => {
                // Error
                format!("#ERROR:{}", value)
            }
            Some("str") | Some("inlineStr") => {
                // Inline string
                value.to_string()
            }
            _ => {
                // Number or general
                value.to_string()
            }
        }
    }

    /// Get a reference to the container.
    pub fn container(&self) -> &OoxmlContainer {
        &self.container
    }

    /// Get the number of sheets.
    pub fn sheet_count(&self) -> usize {
        self.sheets.len()
    }

    /// Get sheet names.
    pub fn sheet_names(&self) -> Vec<&str> {
        self.sheets.iter().map(|s| s.name.as_str()).collect()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_open_xlsx() {
        let path = "test-files/file_example_XLSX_5000.xlsx";
        if std::path::Path::new(path).exists() {
            let parser = XlsxParser::open(path);
            assert!(parser.is_ok());
        }
    }

    #[test]
    fn test_parse_xlsx() {
        let path = "test-files/file_example_XLSX_5000.xlsx";
        if std::path::Path::new(path).exists() {
            let mut parser = XlsxParser::open(path).unwrap();
            let doc = parser.parse().unwrap();

            // Should have at least one section (sheet)
            assert!(!doc.sections.is_empty());

            // First section should have a table
            if let Some(Block::Table(table)) = doc.sections[0].content.first() {
                assert!(!table.rows.is_empty());
                // Check first row is header
                assert!(table.rows[0].is_header);
            }
        }
    }

    #[test]
    fn test_sheet_names() {
        let path = "test-files/file_example_XLSX_5000.xlsx";
        if std::path::Path::new(path).exists() {
            let parser = XlsxParser::open(path).unwrap();
            let names = parser.sheet_names();
            assert!(!names.is_empty());
        }
    }

    #[test]
    fn test_shared_strings() {
        let path = "test-files/file_example_XLSX_5000.xlsx";
        if std::path::Path::new(path).exists() {
            let mut parser = XlsxParser::open(path).unwrap();
            let doc = parser.parse().unwrap();

            // Get plain text and check for expected content
            let text = doc.plain_text();
            assert!(text.contains("First Name"));
            assert!(text.contains("Last Name"));
        }
    }
}
