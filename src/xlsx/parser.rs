//! XLSX parser implementation.

use crate::container::OoxmlContainer;
use crate::error::{Error, Result};
use crate::model::{
    Block, Cell, CellAlignment, Document, Metadata, Paragraph, Resource, ResourceType, Row,
    Section, Table, TextRun,
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
                    // Parse sheet-level relationships for hyperlinks
                    let rels_path = Self::rels_path_for(&sheet_path);
                    let hyperlink_map = if let Ok(rels_xml) =
                        self.container.read_xml(&rels_path)
                    {
                        let sheet_rels = Self::parse_relationships(&rels_xml);
                        Self::parse_hyperlinks(&xml, &sheet_rels)
                    } else {
                        HashMap::new()
                    };

                    if let Ok(table) = self.parse_sheet(&xml, &hyperlink_map) {
                        section.add_block(Block::Table(table));
                    }

                    // Parse drawing images linked to this sheet
                    let images = self.parse_sheet_drawing_images(&sheet_path);
                    for image in images {
                        section.add_block(image);
                    }
                }
            }

            doc.add_section(section);
        }

        // Extract resources (images, media)
        self.extract_resources(&mut doc)?;

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
                                    if let (
                                        Some((start_col, start_row)),
                                        Some((end_col, end_row)),
                                    ) = (Self::parse_cell_ref(start), Self::parse_cell_ref(end))
                                    {
                                        let col_span = end_col - start_col + 1;
                                        let row_span = end_row - start_row + 1;
                                        merge_map
                                            .insert(start.to_uppercase(), (col_span, row_span));
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
    ///
    /// `hyperlink_map` maps uppercase cell references (e.g. "A1") to URLs.
    /// When a cell matches, all its TextRuns get the hyperlink URL set.
    fn parse_sheet(&self, xml: &str, hyperlink_map: &HashMap<String, String>) -> Result<Table> {
        // First pass: parse merge cells
        let merge_map = Self::parse_merge_cells(xml);

        let mut table = Table::new();
        let mut reader = quick_xml::Reader::from_str(xml);
        // IMPORTANT: Don't trim text - preserve whitespace from xml:space="preserve" elements
        // Excel cell values may contain significant leading/trailing spaces
        reader.config_mut().trim_text(false);

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
                                        current_cell_ref = Some(
                                            String::from_utf8_lossy(&attr.value).to_uppercase(),
                                        );
                                    }
                                    b"s" => {
                                        // Style index for number format detection
                                        current_cell_style =
                                            String::from_utf8_lossy(&attr.value).parse().ok();
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

                            // Check if this cell has a hyperlink
                            let hyperlink_url = current_cell_ref
                                .as_ref()
                                .and_then(|r| hyperlink_map.get(r))
                                .cloned();

                            let mut text_run = TextRun::plain(&value);
                            if let Some(url) = hyperlink_url {
                                text_run.hyperlink = Some(url);
                            }

                            let cell = Cell {
                                content: vec![Paragraph {
                                    runs: vec![text_run],
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

    /// Parse `<hyperlinks>` section from worksheet XML and resolve URLs via sheet rels.
    ///
    /// Returns a map of uppercase cell reference (e.g. "A1") to URL string.
    fn parse_hyperlinks(
        xml: &str,
        sheet_rels: &HashMap<String, (String, String)>,
    ) -> HashMap<String, String> {
        let mut hyperlinks = HashMap::new();
        let mut reader = quick_xml::Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut in_hyperlinks = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    match e.name().as_ref() {
                        b"hyperlinks" => {
                            in_hyperlinks = true;
                        }
                        b"hyperlink" if in_hyperlinks => {
                            Self::collect_hyperlink_attrs(e, sheet_rels, &mut hyperlinks);
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Empty(ref e))
                    if in_hyperlinks && e.name().as_ref() == b"hyperlink" =>
                {
                    Self::collect_hyperlink_attrs(e, sheet_rels, &mut hyperlinks);
                }
                Ok(quick_xml::events::Event::End(ref e)) => {
                    if e.name().as_ref() == b"hyperlinks" {
                        break; // Done with hyperlinks section
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }

        hyperlinks
    }

    /// Extract cell ref and r:id from a `<hyperlink>` element and insert into map.
    fn collect_hyperlink_attrs(
        e: &quick_xml::events::BytesStart<'_>,
        sheet_rels: &HashMap<String, (String, String)>,
        hyperlinks: &mut HashMap<String, String>,
    ) {
        let mut cell_ref = String::new();
        let mut r_id = String::new();

        for attr in e.attributes().flatten() {
            match attr.key.as_ref() {
                b"ref" => {
                    cell_ref = String::from_utf8_lossy(&attr.value).to_uppercase();
                }
                b"r:id" => {
                    r_id = String::from_utf8_lossy(&attr.value).to_string();
                }
                _ => {}
            }
        }

        if !cell_ref.is_empty() && !r_id.is_empty() {
            if let Some((rel_type, target)) = sheet_rels.get(&r_id) {
                if rel_type.contains("hyperlink") {
                    hyperlinks.insert(cell_ref, target.clone());
                }
            }
        }
    }

    /// Parse drawing images linked to a sheet.
    ///
    /// Follows the OOXML relationship chain:
    /// 1. sheet rels → find drawing relationship → drawing XML path
    /// 2. drawing rels → rId → media filename mapping
    /// 3. drawing XML → xdr:pic > xdr:blipFill > a:blip r:embed → rId
    fn parse_sheet_drawing_images(&self, sheet_path: &str) -> Vec<Block> {
        // Step 1: Parse sheet-level rels to find drawing target
        let rels_path = Self::rels_path_for(sheet_path);
        let sheet_rels = match self.container.read_xml(&rels_path) {
            Ok(xml) => Self::parse_relationships(&xml),
            Err(_) => return Vec::new(),
        };

        // Find drawing relationships (Type contains "drawing")
        let sheet_dir = sheet_path.rsplit_once('/').map(|(d, _)| d).unwrap_or("");
        let mut all_images = Vec::new();

        for (rel_type, target) in sheet_rels.values() {
            if !rel_type.contains("drawing") {
                continue;
            }

            // Resolve the drawing path relative to the sheet directory
            let drawing_path = Self::resolve_relative_path(sheet_dir, target);

            // Step 2: Parse drawing rels for rId → media filename mapping
            let drawing_rels_path = Self::rels_path_for(&drawing_path);
            let drawing_rels = match self.container.read_xml(&drawing_rels_path) {
                Ok(xml) => {
                    let rels = Self::parse_relationships(&xml);
                    // Build rId → filename map (only image relationships)
                    let drawing_dir = drawing_path.rsplit_once('/').map(|(d, _)| d).unwrap_or("");
                    rels.into_iter()
                        .filter(|(_, (t, _))| t.contains("image"))
                        .map(|(id, (_, tgt))| {
                            let resolved = Self::resolve_relative_path(drawing_dir, &tgt);
                            let filename = resolved.rsplit('/').next().unwrap_or(&resolved).to_string();
                            (id, filename)
                        })
                        .collect::<HashMap<String, String>>()
                }
                Err(_) => continue,
            };

            if drawing_rels.is_empty() {
                continue;
            }

            // Step 3: Parse drawing XML for pic elements with a:blip references
            if let Ok(xml) = self.container.read_xml(&drawing_path) {
                if let Ok(images) = Self::parse_drawing_images(&xml, &drawing_rels) {
                    all_images.extend(images);
                }
            }
        }

        all_images
    }

    /// Build the _rels path for a given part path.
    /// e.g., "xl/worksheets/sheet1.xml" → "xl/worksheets/_rels/sheet1.xml.rels"
    fn rels_path_for(part_path: &str) -> String {
        if let Some((dir, file)) = part_path.rsplit_once('/') {
            format!("{}/_rels/{}.rels", dir, file)
        } else {
            format!("_rels/{}.rels", part_path)
        }
    }

    /// Resolve a relative path (e.g., "../drawings/drawing1.xml") against a base directory.
    fn resolve_relative_path(base_dir: &str, relative: &str) -> String {
        if relative.starts_with('/') {
            return relative.trim_start_matches('/').to_string();
        }

        let mut parts: Vec<&str> = if base_dir.is_empty() {
            Vec::new()
        } else {
            base_dir.split('/').collect()
        };

        for segment in relative.split('/') {
            match segment {
                ".." => { parts.pop(); }
                "." | "" => {}
                s => parts.push(s),
            }
        }

        parts.join("/")
    }

    /// Parse a relationships XML file and return a map of id → (type, target).
    fn parse_relationships(xml: &str) -> HashMap<String, (String, String)> {
        let mut rels = HashMap::new();
        let mut reader = quick_xml::Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Empty(ref e))
                | Ok(quick_xml::events::Event::Start(ref e)) => {
                    if e.name().as_ref() == b"Relationship" {
                        let mut id = String::new();
                        let mut rel_type = String::new();
                        let mut target = String::new();

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"Id" => id = String::from_utf8_lossy(&attr.value).to_string(),
                                b"Type" => rel_type = String::from_utf8_lossy(&attr.value).to_string(),
                                b"Target" => target = String::from_utf8_lossy(&attr.value).to_string(),
                                _ => {}
                            }
                        }

                        if !id.is_empty() && !target.is_empty() {
                            rels.insert(id, (rel_type, target));
                        }
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }

        rels
    }

    /// Parse drawing XML for pic elements and create Block::Image for each.
    ///
    /// Structure: xdr:pic > xdr:nvPicPr > xdr:cNvPr[@name]
    ///            xdr:pic > xdr:blipFill > a:blip[@r:embed]
    ///            xdr:pic > xdr:spPr > a:xfrm > a:ext[@cx, @cy]
    fn parse_drawing_images(
        xml: &str,
        rels: &HashMap<String, String>,
    ) -> Result<Vec<Block>> {
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
                        b"pic" => {
                            in_pic = true;
                            current_name = None;
                            current_rel_id = None;
                            current_width = None;
                            current_height = None;
                        }
                        b"nvPicPr" if in_pic => {
                            in_nvpicpr = true;
                        }
                        b"cNvPr" if in_nvpicpr => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"name" {
                                    current_name =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }
                        b"blipFill" if in_pic => {
                            in_blipfill = true;
                        }
                        b"blip" if in_blipfill => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"embed" {
                                    current_rel_id =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }
                        b"spPr" if in_pic => {
                            in_sppr = true;
                        }
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
                        b"cNvPr" if in_nvpicpr => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"name" {
                                    current_name =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }
                        b"blip" if in_blipfill => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"embed" {
                                    current_rel_id =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }
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
                            if let Some(rel_id) = current_rel_id.take() {
                                if let Some(filename) = rels.get(&rel_id) {
                                    images.push(Block::Image {
                                        resource_id: filename.clone(),
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

    /// Extract resources (images, media) from the workbook.
    fn extract_resources(&self, doc: &mut Document) -> Result<()> {
        for file in self.container.list_files() {
            if file.starts_with("xl/media/") {
                if let Ok(data) = self.container.read_binary(&file) {
                    let filename = file.rsplit('/').next().unwrap_or(&file).to_string();
                    let ext = std::path::Path::new(&file)
                        .extension()
                        .and_then(|e| e.to_str())
                        .unwrap_or("");
                    let size = data.len();

                    let resource = Resource {
                        resource_type: ResourceType::from_extension(ext),
                        filename: Some(filename.clone()),
                        mime_type: guess_mime_type(&file),
                        data,
                        size,
                        width: None,
                        height: None,
                        alt_text: None,
                    };
                    doc.resources.insert(filename, resource);
                }
            }
        }

        Ok(())
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

/// Guess MIME type from file path.
fn guess_mime_type(path: &str) -> Option<String> {
    let ext = path.rsplit('.').next()?.to_lowercase();
    match ext.as_str() {
        "png" => Some("image/png".to_string()),
        "jpg" | "jpeg" => Some("image/jpeg".to_string()),
        "gif" => Some("image/gif".to_string()),
        "bmp" => Some("image/bmp".to_string()),
        "tiff" | "tif" => Some("image/tiff".to_string()),
        "svg" => Some("image/svg+xml".to_string()),
        "wmf" => Some("image/x-wmf".to_string()),
        "emf" => Some("image/x-emf".to_string()),
        _ => None,
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
            assert!(
                found_merged,
                "Expected to find merged cells in Basic Invoice.xlsx"
            );
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
        assert_eq!(
            Styles::serial_to_date(44197.0),
            Some("2021-01-01".to_string())
        );
    }

    #[test]
    fn test_rels_path_for() {
        assert_eq!(
            XlsxParser::rels_path_for("xl/worksheets/sheet1.xml"),
            "xl/worksheets/_rels/sheet1.xml.rels"
        );
        assert_eq!(
            XlsxParser::rels_path_for("xl/drawings/drawing1.xml"),
            "xl/drawings/_rels/drawing1.xml.rels"
        );
    }

    #[test]
    fn test_resolve_relative_path() {
        assert_eq!(
            XlsxParser::resolve_relative_path("xl/worksheets", "../drawings/drawing1.xml"),
            "xl/drawings/drawing1.xml"
        );
        assert_eq!(
            XlsxParser::resolve_relative_path("xl/drawings", "../media/image1.png"),
            "xl/media/image1.png"
        );
        assert_eq!(
            XlsxParser::resolve_relative_path("", "xl/media/image1.png"),
            "xl/media/image1.png"
        );
    }

    #[test]
    fn test_parse_relationships() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
                <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="../printerSettings/printerSettings1.bin"/>
            </Relationships>"#;

        let rels = XlsxParser::parse_relationships(xml);
        assert_eq!(rels.len(), 2);

        let (rel_type, target) = rels.get("rId1").unwrap();
        assert!(rel_type.contains("drawing"));
        assert_eq!(target, "../drawings/drawing1.xml");

        let (rel_type, _) = rels.get("rId2").unwrap();
        assert!(rel_type.contains("printerSettings"));
    }

    #[test]
    fn test_parse_drawing_images() {
        // Synthetic drawing XML with two images
        let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <xdr:twoCellAnchor>
                    <xdr:pic>
                        <xdr:nvPicPr>
                            <xdr:cNvPr id="2" name="Logo"/>
                            <xdr:cNvPicPr/>
                        </xdr:nvPicPr>
                        <xdr:blipFill>
                            <a:blip r:embed="rId1"/>
                        </xdr:blipFill>
                        <xdr:spPr>
                            <a:xfrm>
                                <a:off x="0" y="0"/>
                                <a:ext cx="914400" cy="457200"/>
                            </a:xfrm>
                        </xdr:spPr>
                    </xdr:pic>
                </xdr:twoCellAnchor>
                <xdr:twoCellAnchor>
                    <xdr:pic>
                        <xdr:nvPicPr>
                            <xdr:cNvPr id="3" name="Chart Screenshot"/>
                            <xdr:cNvPicPr/>
                        </xdr:nvPicPr>
                        <xdr:blipFill>
                            <a:blip r:embed="rId2"/>
                        </xdr:blipFill>
                        <xdr:spPr>
                            <a:xfrm>
                                <a:off x="100" y="100"/>
                                <a:ext cx="1828800" cy="1371600"/>
                            </a:xfrm>
                        </xdr:spPr>
                    </xdr:pic>
                </xdr:twoCellAnchor>
            </xdr:wsDr>"#;

        let mut rels = HashMap::new();
        rels.insert("rId1".to_string(), "image1.png".to_string());
        rels.insert("rId2".to_string(), "image2.jpeg".to_string());

        let images = XlsxParser::parse_drawing_images(drawing_xml, &rels).unwrap();
        assert_eq!(images.len(), 2);

        match &images[0] {
            Block::Image { resource_id, alt_text, width, height } => {
                assert_eq!(resource_id, "image1.png");
                assert_eq!(alt_text.as_deref(), Some("Logo"));
                assert_eq!(*width, Some(914400));
                assert_eq!(*height, Some(457200));
            }
            other => panic!("Expected Block::Image, got {:?}", other),
        }

        match &images[1] {
            Block::Image { resource_id, alt_text, width, height } => {
                assert_eq!(resource_id, "image2.jpeg");
                assert_eq!(alt_text.as_deref(), Some("Chart Screenshot"));
                assert_eq!(*width, Some(1828800));
                assert_eq!(*height, Some(1371600));
            }
            other => panic!("Expected Block::Image, got {:?}", other),
        }
    }

    #[test]
    fn test_parse_drawing_images_empty_blip() {
        // Test with self-closing blip element (empty element form)
        let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <xdr:oneCellAnchor>
                    <xdr:pic>
                        <xdr:nvPicPr>
                            <xdr:cNvPr id="2" name="Pic1"/>
                            <xdr:cNvPicPr/>
                        </xdr:nvPicPr>
                        <xdr:blipFill>
                            <a:blip r:embed="rId1"></a:blip>
                        </xdr:blipFill>
                        <xdr:spPr/>
                    </xdr:pic>
                </xdr:oneCellAnchor>
            </xdr:wsDr>"#;

        let mut rels = HashMap::new();
        rels.insert("rId1".to_string(), "image1.png".to_string());

        let images = XlsxParser::parse_drawing_images(drawing_xml, &rels).unwrap();
        assert_eq!(images.len(), 1);

        match &images[0] {
            Block::Image { resource_id, alt_text, .. } => {
                assert_eq!(resource_id, "image1.png");
                assert_eq!(alt_text.as_deref(), Some("Pic1"));
            }
            other => panic!("Expected Block::Image, got {:?}", other),
        }
    }

    #[test]
    fn test_parse_drawing_no_images() {
        // Drawing with only shapes (no pic elements) — like existing test files
        let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                <xdr:twoCellAnchor>
                    <xdr:sp macro="" textlink="">
                        <xdr:nvSpPr>
                            <xdr:cNvPr id="5" name="Rectangle 0"/>
                            <xdr:cNvSpPr/>
                        </xdr:nvSpPr>
                        <xdr:spPr>
                            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                        </xdr:spPr>
                    </xdr:sp>
                </xdr:twoCellAnchor>
            </xdr:wsDr>"#;

        let rels = HashMap::new();
        let images = XlsxParser::parse_drawing_images(drawing_xml, &rels).unwrap();
        assert!(images.is_empty());
    }

    #[test]
    fn test_xlsx_with_drawing_no_images() {
        // Test that existing test files with drawings but no images still parse correctly
        let path = "test-files/Auto Expense Report.xlsx";
        if std::path::Path::new(path).exists() {
            let mut parser = XlsxParser::open(path).unwrap();
            let doc = parser.parse().unwrap();

            // Should parse without errors and have sections
            assert!(!doc.sections.is_empty());

            // No Block::Image should be present (this file has shapes, not images)
            for section in &doc.sections {
                for block in &section.content {
                    assert!(
                        !matches!(block, Block::Image { .. }),
                        "Expected no images in Auto Expense Report.xlsx"
                    );
                }
            }
        }
    }

    #[test]
    fn test_parse_hyperlinks() {
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <sheetData>
                    <row r="1">
                        <c r="A1" t="s"><v>0</v></c>
                        <c r="B1" t="s"><v>1</v></c>
                    </row>
                </sheetData>
                <hyperlinks>
                    <hyperlink ref="A1" r:id="rId1"/>
                    <hyperlink ref="B1" r:id="rId2" display="Example"/>
                </hyperlinks>
            </worksheet>"#;

        let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
                <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://rust-lang.org" TargetMode="External"/>
                <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
            </Relationships>"#;

        let sheet_rels = XlsxParser::parse_relationships(rels_xml);
        let hyperlinks = XlsxParser::parse_hyperlinks(sheet_xml, &sheet_rels);

        assert_eq!(hyperlinks.len(), 2);
        assert_eq!(hyperlinks.get("A1").unwrap(), "https://example.com");
        assert_eq!(hyperlinks.get("B1").unwrap(), "https://rust-lang.org");
    }

    #[test]
    fn test_parse_hyperlinks_empty() {
        // Sheet with no hyperlinks section
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData>
                    <row r="1">
                        <c r="A1" t="s"><v>0</v></c>
                    </row>
                </sheetData>
            </worksheet>"#;

        let sheet_rels = HashMap::new();
        let hyperlinks = XlsxParser::parse_hyperlinks(sheet_xml, &sheet_rels);
        assert!(hyperlinks.is_empty());
    }

    #[test]
    fn test_parse_hyperlinks_non_hyperlink_rels_ignored() {
        // hyperlink element references a non-hyperlink relationship (should be ignored)
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <sheetData/>
                <hyperlinks>
                    <hyperlink ref="A1" r:id="rId1"/>
                </hyperlinks>
            </worksheet>"#;

        let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
            </Relationships>"#;

        let sheet_rels = XlsxParser::parse_relationships(rels_xml);
        let hyperlinks = XlsxParser::parse_hyperlinks(sheet_xml, &sheet_rels);
        assert!(hyperlinks.is_empty());
    }
}
