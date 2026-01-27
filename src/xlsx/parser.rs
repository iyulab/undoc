//! XLSX parser implementation.

use crate::container::OoxmlContainer;
use crate::error::{Error, Result};
use crate::model::{
    Block, Cell, CellAlignment, Document, Metadata, Paragraph, Row, Section, Table, TextRun,
};
use std::collections::HashMap;
use std::path::Path;

use super::shared_strings::SharedStrings;
use super::styles::Styles;

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
    styles: Styles,
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

        // Parse styles for number formats
        let styles = if let Ok(xml) = container.read_xml("xl/styles.xml") {
            Styles::parse(&xml)
        } else {
            Styles::default()
        };

        // Parse workbook relationships
        let relationships = Self::parse_workbook_rels(&container)?;

        // Parse workbook for sheet info
        let sheets = Self::parse_workbook(&container)?;

        Ok(Self {
            container,
            shared_strings,
            styles,
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
                                        sheet_id = String::from_utf8_lossy(&attr.value).to_string();
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
                let sheet_path = if let Some(stripped) = target.strip_prefix('/') {
                    stripped.to_string()
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
        // Use shared metadata parsing from container
        let mut meta = self.container.parse_core_metadata()?;
        // Set sheet count
        meta.page_count = Some(self.sheets.len() as u32);
        Ok(meta)
    }

    /// Parse merge cells information from worksheet XML.
    fn parse_merge_cells(xml: &str) -> HashMap<String, (u32, u32)> {
        let mut merge_map = HashMap::new();
        let mut reader = quick_xml::Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Empty(ref e))
                | Ok(quick_xml::events::Event::Start(ref e)) => {
                    if e.name().as_ref() == b"mergeCell" {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"ref" {
                                let range = String::from_utf8_lossy(&attr.value);
                                // Parse range like "A1:C3" or "G11:H11"
                                if let Some((start, end)) = range.split_once(':') {
                                    if let (Some((start_col, start_row)), Some((end_col, end_row))) =
                                        (Self::parse_cell_ref(start), Self::parse_cell_ref(end))
                                    {
                                        let col_span = end_col - start_col + 1;
                                        let row_span = end_row - start_row + 1;
                                        merge_map.insert(start.to_uppercase(), (col_span, row_span));
                                    }
                                }
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

        merge_map
    }

    /// Parse cell reference like "A1" into (column, row) where column is 0-indexed.
    fn parse_cell_ref(cell_ref: &str) -> Option<(u32, u32)> {
        let cell_ref = cell_ref.to_uppercase();
        let mut col_str = String::new();
        let mut row_str = String::new();

        for c in cell_ref.chars() {
            if c.is_ascii_alphabetic() {
                col_str.push(c);
            } else if c.is_ascii_digit() {
                row_str.push(c);
            }
        }

        if col_str.is_empty() || row_str.is_empty() {
            return None;
        }

        // Convert column letters to number (A=0, B=1, ..., Z=25, AA=26, ...)
        let mut col: u32 = 0;
        for c in col_str.chars() {
            col = col * 26 + (c as u32 - 'A' as u32 + 1);
        }
        col -= 1; // Make it 0-indexed

        let row: u32 = row_str.parse().ok()?;

        Some((col, row))
    }

    /// Parse a worksheet XML into a table.
    fn parse_sheet(&self, xml: &str) -> Result<Table> {
        // First pass: parse merge cells
        let merge_map = Self::parse_merge_cells(xml);

        let mut table = Table::new();
        let mut reader = quick_xml::Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut in_row = false;
        let mut in_cell = false;
        let mut in_value = false;
        let mut current_row: Option<Row> = None;
        let mut current_cell_type: Option<String> = None;
        let mut current_cell_ref: Option<String> = None;
        let mut current_cell_style: Option<usize> = None;
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
                            current_cell_ref = None;
                            current_cell_style = None;
                            current_cell_value.clear();

                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"t" => {
                                        current_cell_type =
                                            Some(String::from_utf8_lossy(&attr.value).to_string());
                                    }
                                    b"r" => {
                                        // Cell reference like "A1", "B2", etc.
                                        current_cell_ref =
                                            Some(String::from_utf8_lossy(&attr.value).to_uppercase());
                                    }
                                    b"s" => {
                                        // Style index for number format detection
                                        current_cell_style = String::from_utf8_lossy(&attr.value)
                                            .parse()
                                            .ok();
                                    }
                                    _ => {}
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
                                current_cell_style,
                            );

                            // Look up merge info for this cell
                            let (col_span, row_span) = current_cell_ref
                                .as_ref()
                                .and_then(|r| merge_map.get(r))
                                .copied()
                                .unwrap_or((1, 1));

                            let cell = Cell {
                                content: vec![Paragraph {
                                    runs: vec![TextRun::plain(&value)],
                                    ..Default::default()
                                }],
                                nested_tables: Vec::new(),
                                col_span,
                                row_span,
                                alignment: CellAlignment::Left,
                                vertical_alignment: Default::default(),
                                is_header: is_first_row,
                                background: None,
                            };

                            if let Some(ref mut row) = current_row {
                                row.cells.push(cell);
                            }

                            in_cell = false;
                            current_cell_ref = None;
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

    /// Resolve a cell value based on its type and style.
    fn resolve_cell_value(
        &self,
        value: &str,
        cell_type: Option<&str>,
        style_index: Option<usize>,
    ) -> String {
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
                // Number or general - check for date format
                if let Some(style_idx) = style_index {
                    if let Some(num_fmt_id) = self.styles.get_num_fmt_id(style_idx) {
                        if self.styles.is_date_format(num_fmt_id) {
                            // Try to parse as date
                            if let Ok(serial) = value.parse::<f64>() {
                                if let Some(date_str) = Styles::serial_to_date(serial) {
                                    return date_str;
                                }
                            }
                        }
                    }
                }
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

    #[test]
    fn test_merged_cells() {
        let path = "test-files/Basic Invoice.xlsx";
        if std::path::Path::new(path).exists() {
            let mut parser = XlsxParser::open(path).unwrap();
            let doc = parser.parse().unwrap();

            // Find merged cells
            let mut found_merged = false;
            for section in &doc.sections {
                for block in &section.content {
                    if let Block::Table(table) = block {
                        for row in &table.rows {
                            for cell in &row.cells {
                                if cell.col_span > 1 || cell.row_span > 1 {
                                    found_merged = true;
                                    println!(
                                        "Found merged cell: col_span={}, row_span={}, text='{}'",
                                        cell.col_span,
                                        cell.row_span,
                                        cell.plain_text()
                                    );
                                }
                            }
                        }
                    }
                }
            }
            assert!(found_merged, "Expected to find merged cells in Basic Invoice.xlsx");
        }
    }

    #[test]
    fn test_parse_cell_ref() {
        // Test cell reference parsing
        assert_eq!(XlsxParser::parse_cell_ref("A1"), Some((0, 1)));
        assert_eq!(XlsxParser::parse_cell_ref("B2"), Some((1, 2)));
        assert_eq!(XlsxParser::parse_cell_ref("Z1"), Some((25, 1)));
        assert_eq!(XlsxParser::parse_cell_ref("AA1"), Some((26, 1)));
        assert_eq!(XlsxParser::parse_cell_ref("AB1"), Some((27, 1)));
        assert_eq!(XlsxParser::parse_cell_ref("AZ1"), Some((51, 1)));
        assert_eq!(XlsxParser::parse_cell_ref("BA1"), Some((52, 1)));
    }

    #[test]
    fn test_date_formatting() {
        // Test that styles are correctly parsed and dates are formatted
        use crate::xlsx::styles::Styles;

        // Test parsing styles.xml content
        let styles_xml = r#"<?xml version="1.0"?>
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <numFmts count="1">
                    <numFmt numFmtId="177" formatCode="mmmm\ d\,\ yyyy"/>
                </numFmts>
                <cellXfs count="2">
                    <xf numFmtId="0"/>
                    <xf numFmtId="177"/>
                </cellXfs>
            </styleSheet>"#;

        let styles = Styles::parse(styles_xml);

        // Style index 1 should have numFmtId 177 (date format)
        assert_eq!(styles.get_num_fmt_id(1), Some(177));
        assert!(styles.is_date_format(177));

        // Test date conversion
        assert_eq!(Styles::serial_to_date(44197.0), Some("2021-01-01".to_string()));
    }
}
