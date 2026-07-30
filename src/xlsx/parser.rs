//! XLSX parser implementation.

use std::borrow::Cow;

use crate::container::OoxmlContainer;
use crate::decode::{normalize_line_endings, resolve_general_ref};
use crate::error::{Error, Result};
use crate::model::{
    Block, Cell, CellAlignment, Document, InlineImage, Metadata, Paragraph, Resource, ResourceType,
    Row, Section, Table, TextRun,
};
use std::collections::HashMap;
#[cfg(not(target_arch = "wasm32"))]
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

struct BuildSheetCellContext<'a> {
    merge_map: &'a HashMap<String, (u32, u32)>,
    hyperlink_map: &'a HashMap<String, String>,
    comment_map: &'a HashMap<String, String>,
    is_header: bool,
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
        // Parse shared strings — absent is OK, malformed bytes must surface.
        let shared_strings = match container.read_xml_optional("xl/sharedStrings.xml")? {
            Some(xml) => SharedStrings::parse(&xml)?,
            None => SharedStrings::default(),
        };

        // Parse styles for number formats — absent is OK, malformed bytes must surface.
        let styles = match container.read_xml_optional("xl/styles.xml")? {
            Some(xml) => Styles::parse(&xml),
            None => Styles::default(),
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
        Ok(container
            .read_required_relationships_for_part("xl/workbook.xml")?
            .into_targets_by_id())
    }

    /// Parse workbook.xml for sheet info.
    fn parse_workbook(container: &OoxmlContainer) -> Result<Vec<SheetInfo>> {
        let mut sheets = Vec::new();
        let xml = container.read_xml("xl/workbook.xml")?;

        let mut reader = crate::decode::reader_for(&xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Empty(e)) | Ok(quick_xml::events::Event::Start(e))
                    if e.name().as_ref() == b"sheet" =>
                {
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
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(sheets)
    }

    /// Parse the workbook and return a Document model.
    pub fn parse(&mut self) -> Result<Document> {
        let mut doc = Document::new();
        doc.format = crate::detect::FormatType::Xlsx;

        // Parse metadata
        doc.metadata = self.parse_metadata()?;

        // Resolve "Place in Cell" rich-value images once per workbook
        let rich_value_images = self.parse_rich_value_images()?;

        // Parse each sheet as a section with a table
        for (idx, sheet) in self.sheets.clone().iter().enumerate() {
            let section = self.parse_sheet_as_section(idx, sheet, &rich_value_images)?;
            doc.add_section(section);
        }

        // Extract resources (images, media)
        self.extract_resources(&mut doc)?;

        Ok(doc)
    }

    /// Stream sections (sheets) one at a time, calling `f` for each event.
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
        let section_count = self.sheets.len();

        // Build image_map from resources before streaming sections.
        let mut dummy_doc = Document::new();
        self.extract_resources(&mut dummy_doc)?;
        let image_map: std::collections::HashMap<String, String> = dummy_doc
            .resources
            .iter()
            .filter_map(|(id, r)| r.filename.as_ref().map(|name| (id.clone(), name.clone())))
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

        let rich_value_images = self.parse_rich_value_images()?;

        for (idx, sheet) in self.sheets.clone().iter().enumerate() {
            let section_result = self.parse_sheet_as_section(idx, sheet, &rich_value_images);

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
            for (id, resource) in dummy_doc.resources {
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

    /// Parse a single sheet into a Section.
    fn parse_sheet_as_section(
        &self,
        idx: usize,
        sheet: &SheetInfo,
        rich_value_images: &HashMap<u32, String>,
    ) -> Result<Section> {
        let mut section = Section::new(idx);
        section.name = Some(sheet.name.clone());

        if let Some(target) = self.relationships.get(&sheet.rel_id) {
            let sheet_path = if let Some(stripped) = target.strip_prefix('/') {
                stripped.to_string()
            } else {
                format!("xl/{}", target)
            };

            if let Some(xml) = self.container.read_xml_optional(&sheet_path)? {
                let sheet_dir = sheet_path.rsplit_once('/').map(|(d, _)| d).unwrap_or("");
                let sheet_rels = self.read_optional_relationships_for_part(&sheet_path)?;
                let hyperlink_map = Self::parse_hyperlinks(&xml, &sheet_rels);
                let comment_map =
                    Self::find_and_parse_comments(&self.container, &sheet_rels, sheet_dir)?;
                let table =
                    self.parse_sheet(&xml, &hyperlink_map, &comment_map, rich_value_images)?;
                section.add_block(Block::Table(table));

                let images = self.parse_sheet_drawing_images(&sheet_path)?;
                for image in images {
                    section.add_block(image);
                }
            }
        }

        Ok(section)
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
        let mut reader = crate::decode::reader_for(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Empty(ref e))
                | Ok(quick_xml::events::Event::Start(ref e))
                    if e.name().as_ref() == b"mergeCell" =>
                {
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
    ///
    /// `comment_map` maps uppercase cell references (e.g. "A1") to comment text.
    /// When a cell matches, the comment is appended as an italic TextRun.
    fn parse_sheet(
        &self,
        xml: &str,
        hyperlink_map: &HashMap<String, String>,
        comment_map: &HashMap<String, String>,
        rich_value_images: &HashMap<u32, String>,
    ) -> Result<Table> {
        // First pass: parse merge cells
        let merge_map = Self::parse_merge_cells(xml);

        let mut table = Table::new();
        let mut reader = crate::decode::reader_for(xml);
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
        let mut current_cell_vm: Option<u32> = None;
        let mut current_cell_value = String::new();
        let mut is_first_row = true;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => match e.name().as_ref() {
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
                        current_cell_vm = None;
                        current_cell_value.clear();
                        Self::parse_cell_attributes(
                            e,
                            &mut current_cell_type,
                            &mut current_cell_ref,
                            &mut current_cell_style,
                            &mut current_cell_vm,
                        );
                    }
                    b"v" if in_cell => {
                        in_value = true;
                    }
                    b"t" if in_cell => {
                        // Inline string
                        in_value = true;
                    }
                    _ => {}
                },
                Ok(quick_xml::events::Event::Empty(ref e)) => match e.name().as_ref() {
                    b"c" if in_row => {
                        current_cell_type = None;
                        current_cell_ref = None;
                        current_cell_style = None;
                        current_cell_vm = None;
                        current_cell_value.clear();
                        Self::parse_cell_attributes(
                            e,
                            &mut current_cell_type,
                            &mut current_cell_ref,
                            &mut current_cell_style,
                            &mut current_cell_vm,
                        );

                        let mut cell = self.build_sheet_cell(
                            &current_cell_value,
                            current_cell_type.as_deref(),
                            current_cell_style,
                            current_cell_ref.as_deref(),
                            BuildSheetCellContext {
                                merge_map: &merge_map,
                                hyperlink_map,
                                comment_map,
                                is_header: is_first_row,
                            },
                        )?;
                        Self::attach_rich_value_image(
                            &mut cell,
                            current_cell_vm,
                            current_cell_type.as_deref(),
                            rich_value_images,
                        );

                        if let Some(ref mut row) = current_row {
                            Self::push_cell_with_row_local_spacing(
                                row,
                                cell,
                                current_cell_ref.as_deref(),
                                is_first_row,
                            );
                        }

                        current_cell_ref = None;
                    }
                    _ => {}
                },
                Ok(quick_xml::events::Event::Text(ref e)) if in_value => {
                    let text = crate::decode::decode_text_lossy(e);
                    current_cell_value.push_str(&text);
                }
                // quick-xml 0.40+ splits entity refs out of Text; without this
                // arm a cell value like "A &amp; B" or an in-cell "&#10;" break
                // would silently lose the entity.
                Ok(quick_xml::events::Event::GeneralRef(ref e)) if in_value => {
                    current_cell_value.push_str(&resolve_general_ref(e));
                }
                Ok(quick_xml::events::Event::End(ref e)) => match e.name().as_ref() {
                    b"row" => {
                        if let Some(row) = current_row.take() {
                            table.add_row(row);
                        }
                        in_row = false;
                        is_first_row = false;
                    }
                    b"c" => {
                        // Collapse CR that re-entered via &#13;/&#xD; refs (Excel
                        // in-cell breaks). A CRLF pair arrives as two refs, so
                        // normalize the whole accumulated value once here.
                        let cell_value =
                            normalize_line_endings(Cow::Borrowed(&current_cell_value)).into_owned();
                        let mut cell = self.build_sheet_cell(
                            &cell_value,
                            current_cell_type.as_deref(),
                            current_cell_style,
                            current_cell_ref.as_deref(),
                            BuildSheetCellContext {
                                merge_map: &merge_map,
                                hyperlink_map,
                                comment_map,
                                is_header: is_first_row,
                            },
                        )?;
                        Self::attach_rich_value_image(
                            &mut cell,
                            current_cell_vm,
                            current_cell_type.as_deref(),
                            rich_value_images,
                        );

                        if let Some(ref mut row) = current_row {
                            Self::push_cell_with_row_local_spacing(
                                row,
                                cell,
                                current_cell_ref.as_deref(),
                                is_first_row,
                            );
                        }

                        in_cell = false;
                        current_cell_ref = None;
                        current_cell_vm = None;
                    }
                    b"v" | b"t" => {
                        in_value = false;
                    }
                    _ => {}
                },
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        // Trim trailing rows where every cell is content-empty (no text, no value).
        // These arise from formatting-only rows, deleted-but-kept rows, and conditional
        // formatting applied to entire columns — all common causes of massive row inflation.
        while table.rows.last().is_some_and(|r| r.is_empty()) {
            table.rows.pop();
        }

        // Trim trailing columns where every row has empty cells.
        // Find the rightmost column index that has content in any row.
        // Uses physical cell indices (one slot per Cell, regardless of col_span),
        // matching how push_cell_with_row_local_spacing builds the cells vec.
        let max_content_col = table
            .rows
            .iter()
            .flat_map(|r| {
                r.cells
                    .iter()
                    .enumerate()
                    .filter(|(_, c)| !c.is_empty())
                    .map(|(i, _)| i)
            })
            .max();

        if let Some(max_col) = max_content_col {
            let keep = max_col + 1;
            for row in &mut table.rows {
                row.cells.truncate(keep);
            }
        }

        Ok(table)
    }

    fn parse_cell_attributes(
        e: &quick_xml::events::BytesStart<'_>,
        current_cell_type: &mut Option<String>,
        current_cell_ref: &mut Option<String>,
        current_cell_style: &mut Option<usize>,
        current_cell_vm: &mut Option<u32>,
    ) {
        for attr in e.attributes().flatten() {
            match attr.key.as_ref() {
                b"t" => {
                    *current_cell_type = Some(String::from_utf8_lossy(&attr.value).to_string());
                }
                b"r" => {
                    *current_cell_ref = Some(String::from_utf8_lossy(&attr.value).to_uppercase());
                }
                b"s" => {
                    *current_cell_style = String::from_utf8_lossy(&attr.value).parse().ok();
                }
                b"vm" => {
                    *current_cell_vm = String::from_utf8_lossy(&attr.value).parse().ok();
                }
                _ => {}
            }
        }
    }

    fn build_sheet_cell(
        &self,
        current_cell_value: &str,
        current_cell_type: Option<&str>,
        current_cell_style: Option<usize>,
        current_cell_ref: Option<&str>,
        context: BuildSheetCellContext<'_>,
    ) -> Result<Cell> {
        let value =
            self.resolve_cell_value(current_cell_value, current_cell_type, current_cell_style)?;

        let (col_span, row_span) = current_cell_ref
            .and_then(|r| context.merge_map.get(r))
            .copied()
            .unwrap_or((1, 1));

        let hyperlink_url = current_cell_ref
            .and_then(|r| context.hyperlink_map.get(r))
            .cloned();

        let mut text_run = TextRun::plain(&value);
        if let Some(url) = hyperlink_url {
            text_run.hyperlink = Some(url);
        }

        let mut runs = vec![text_run];

        if let Some(comment_text) = current_cell_ref.and_then(|r| context.comment_map.get(r)) {
            let mut comment_run = TextRun::plain(format!(" [Comment: {}]", comment_text));
            comment_run.style.italic = true;
            runs.push(comment_run);
        }

        Ok(Cell {
            content: vec![Paragraph {
                runs,
                ..Default::default()
            }],
            nested_tables: Vec::new(),
            col_span,
            row_span,
            alignment: CellAlignment::Left,
            vertical_alignment: Default::default(),
            is_header: context.is_header,
            background: None,
        })
    }

    fn push_cell_with_row_local_spacing(
        row: &mut Row,
        cell: Cell,
        cell_ref: Option<&str>,
        is_header: bool,
    ) {
        if let Some((target_col, _)) = cell_ref.and_then(Self::parse_cell_ref) {
            let current_cols = row.effective_columns();

            // Guardrail: reconstruct gaps only inside the current row up to the highest
            // explicitly referenced column in that row. This is not whole-sheet densification.
            for _ in current_cols..target_col as usize {
                row.cells.push(Cell {
                    is_header,
                    ..Cell::new()
                });
            }
        }

        row.cells.push(cell);
    }

    /// Resolve a cell value based on its type and style.
    fn resolve_cell_value(
        &self,
        value: &str,
        cell_type: Option<&str>,
        style_index: Option<usize>,
    ) -> Result<String> {
        match cell_type {
            Some("s") => {
                // Shared string index
                if let Ok(idx) = value.parse::<usize>() {
                    self.shared_strings
                        .get(idx)
                        .map(str::to_string)
                        .ok_or_else(|| {
                            Error::InvalidData(format!("shared string index out of range: {idx}"))
                        })
                } else {
                    Ok(value.to_string())
                }
            }
            Some("b") => {
                // Boolean
                if value == "1" {
                    Ok("TRUE".to_string())
                } else {
                    Ok("FALSE".to_string())
                }
            }
            Some("e") => {
                // Error
                Ok(format!("#ERROR:{}", value))
            }
            Some("str") | Some("inlineStr") => {
                // Inline string
                Ok(value.to_string())
            }
            _ => {
                // Number or general - check for date format
                if let Some(style_idx) = style_index {
                    if let Some(num_fmt_id) = self.styles.get_num_fmt_id(style_idx) {
                        if self.styles.is_date_format(num_fmt_id) {
                            // Try to parse as date
                            if let Ok(serial) = value.parse::<f64>() {
                                if let Some(date_str) = Styles::serial_to_date(serial) {
                                    return Ok(date_str);
                                }
                            }
                        }
                    }
                }
                Ok(value.to_string())
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
        let mut reader = crate::decode::reader_for(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut in_hyperlinks = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => match e.name().as_ref() {
                    b"hyperlinks" => {
                        in_hyperlinks = true;
                    }
                    b"hyperlink" if in_hyperlinks => {
                        Self::collect_hyperlink_attrs(e, sheet_rels, &mut hyperlinks);
                    }
                    _ => {}
                },
                Ok(quick_xml::events::Event::Empty(ref e))
                    if in_hyperlinks && e.name().as_ref() == b"hyperlink" =>
                {
                    Self::collect_hyperlink_attrs(e, sheet_rels, &mut hyperlinks);
                }
                Ok(quick_xml::events::Event::End(ref e)) if e.name().as_ref() == b"hyperlinks" => {
                    break; // Done with hyperlinks section
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

    /// Parse comments XML and return a map of uppercase cell reference to comment text.
    ///
    /// The XML structure is:
    /// ```xml
    /// <comments>
    ///   <commentList>
    ///     <comment ref="A1" authorId="0">
    ///       <text><r><t>Comment text</t></r></text>
    ///     </comment>
    ///   </commentList>
    /// </comments>
    /// ```
    fn parse_comments_xml(xml: &str) -> HashMap<String, String> {
        let mut comments = HashMap::new();
        let mut reader = crate::decode::reader_for(xml);
        // Preserve whitespace inside <t>: quick-xml 0.40+ fragments a single <t>
        // into Text/GeneralRef events, and trimming each fragment would drop the
        // spaces adjacent to an entity ("Bad &bogus; comment" → "Bad&bogus;comment").
        // Inter-element indentation is already excluded by the `in_t` guard.
        reader.config_mut().trim_text(false);
        let mut buf = Vec::new();

        let mut in_comment = false;
        let mut in_text = false;
        let mut in_t = false;
        let mut current_ref = String::new();
        let mut current_text = String::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    let local = e.name().local_name();
                    match local.as_ref() {
                        b"comment" => {
                            in_comment = true;
                            current_ref.clear();
                            current_text.clear();
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"ref" {
                                    current_ref =
                                        String::from_utf8_lossy(&attr.value).to_uppercase();
                                }
                            }
                        }
                        b"text" if in_comment => {
                            in_text = true;
                        }
                        b"t" if in_text => {
                            in_t = true;
                            // Separate consecutive <t> runs with a space. This must
                            // happen at the run boundary, not per Text event: under
                            // quick-xml 0.40+ a single <t> is fragmented into
                            // Text/GeneralRef events, and joining those with spaces
                            // would corrupt the text (e.g. "A &amp; B" → "A &  B").
                            if !current_text.is_empty() {
                                current_text.push(' ');
                            }
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Text(ref e)) if in_t => {
                    current_text.push_str(&crate::decode::decode_text_lossy(e));
                }
                Ok(quick_xml::events::Event::GeneralRef(ref e)) if in_t => {
                    current_text.push_str(&resolve_general_ref(e));
                }
                Ok(quick_xml::events::Event::End(ref e)) => {
                    let local = e.name().local_name();
                    match local.as_ref() {
                        b"comment" => {
                            if !current_ref.is_empty() && !current_text.is_empty() {
                                let text = normalize_line_endings(Cow::Borrowed(&current_text))
                                    .into_owned();
                                comments.insert(current_ref.clone(), text);
                            }
                            in_comment = false;
                            in_text = false;
                            in_t = false;
                        }
                        b"text" => {
                            in_text = false;
                        }
                        b"t" => {
                            in_t = false;
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

        comments
    }

    /// Find the comments file for a sheet via its relationships, read it, and parse.
    ///
    /// A missing comments part is treated as "no comments", but malformed
    /// content surfaces as `Error::Encoding` / other read errors so corruption
    /// is never silently dropped.
    fn find_and_parse_comments(
        container: &OoxmlContainer,
        sheet_rels: &HashMap<String, (String, String)>,
        sheet_dir: &str,
    ) -> Result<HashMap<String, String>> {
        // Look for a relationship whose type contains "comments"
        for (rel_type, target) in sheet_rels.values() {
            if !rel_type.contains("comments") {
                continue;
            }
            let comments_path = Self::resolve_relative_path(sheet_dir, target);
            if let Some(comments_xml) = container.read_xml_optional(&comments_path)? {
                return Ok(Self::parse_comments_xml(&comments_xml));
            }
        }
        Ok(HashMap::new())
    }

    /// Parse drawing images linked to a sheet.
    ///
    /// Follows the OOXML relationship chain:
    /// 1. sheet rels → find drawing relationship → drawing XML path
    /// 2. drawing rels → rId → media filename mapping
    /// 3. drawing XML → xdr:pic > xdr:blipFill > a:blip r:embed → rId
    fn parse_sheet_drawing_images(&self, sheet_path: &str) -> Result<Vec<Block>> {
        // Step 1: Parse sheet-level rels to find drawing target
        let sheet_rels = self.read_optional_relationships_for_part(sheet_path)?;

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
            let drawing_dir = drawing_path.rsplit_once('/').map(|(d, _)| d).unwrap_or("");
            let drawing_rels = self
                .read_optional_relationships_for_part(&drawing_path)?
                .into_iter()
                .filter(|(_, (t, _))| t.contains("image"))
                .map(|(id, (_, tgt))| {
                    let resolved = Self::resolve_relative_path(drawing_dir, &tgt);
                    let filename = resolved.rsplit('/').next().unwrap_or(&resolved).to_string();
                    (id, filename)
                })
                .collect::<HashMap<String, String>>();

            if drawing_rels.is_empty() {
                continue;
            }

            // Step 3: Parse drawing XML for pic elements with a:blip references.
            // Missing drawing XML is OK; malformed bytes must surface.
            if let Some(xml) = self.container.read_xml_optional(&drawing_path)? {
                if let Ok(images) = Self::parse_drawing_images(&xml, &drawing_rels) {
                    all_images.extend(images);
                }
            }
        }

        Ok(all_images)
    }

    /// Build the _rels path for a given part path.
    /// e.g., "xl/worksheets/sheet1.xml" → "xl/worksheets/_rels/sheet1.xml.rels"
    #[cfg(test)]
    fn rels_path_for(part_path: &str) -> String {
        if let Some((dir, file)) = part_path.rsplit_once('/') {
            format!("{}/_rels/{}.rels", dir, file)
        } else {
            format!("_rels/{}.rels", part_path)
        }
    }

    fn read_optional_relationships_for_part(
        &self,
        part_path: &str,
    ) -> Result<HashMap<String, (String, String)>> {
        self.container
            .read_optional_relationships_for_part(part_path)
            .map(|rels| rels.into_type_targets_by_id())
    }

    /// Resolve "Place in Cell" (rich-value) images to a `vm` → media-filename map.
    ///
    /// Excel stores in-cell images as error cells (`t="e" vm="N"` with a
    /// `#VALUE!` placeholder) whose real content lives in a chain of
    /// workbook-level parts:
    ///
    /// 1. `xl/metadata.xml` — `vm` (1-based) → valueMetadata bk → rc[@v] →
    ///    futureMetadata\[XLRICHVALUE\] bk → `xlrd:rvb[@i]` = rich-value index
    /// 2. `xl/richData/rdrichvaluestructure.xml` — structure key order; the
    ///    `_rvRel:LocalImageIdentifier` key position selects the `<v>` slot
    /// 3. `xl/richData/rdrichvalue.xml` — rich value `<rv s="…">` with `<v>`
    ///    values; the selected slot is an index into richValueRel
    /// 4. `xl/richData/richValueRel.xml` + its rels — index → r:id → media path
    ///
    /// Any missing link in the chain yields an empty map (graceful
    /// degradation for foreign producers); malformed XML surfaces as an error.
    fn parse_rich_value_images(&self) -> Result<HashMap<u32, String>> {
        let wb_rels = self.read_optional_relationships_for_part("xl/workbook.xml")?;
        let part_by_type = |suffix: &str| -> Option<String> {
            wb_rels
                .values()
                .find(|(rel_type, _)| rel_type.ends_with(suffix))
                .map(|(_, target)| Self::resolve_relative_path("xl", target))
        };

        let (Some(metadata_path), Some(rv_path), Some(structure_path), Some(rel_path)) = (
            part_by_type("/sheetMetadata"),
            part_by_type("/rdRichValue"),
            part_by_type("/rdRichValueStructure"),
            part_by_type("/richValueRel"),
        ) else {
            return Ok(HashMap::new());
        };
        let (Some(metadata_xml), Some(rv_xml), Some(structure_xml), Some(rel_xml)) = (
            self.container.read_xml_optional(&metadata_path)?,
            self.container.read_xml_optional(&rv_path)?,
            self.container.read_xml_optional(&structure_path)?,
            self.container.read_xml_optional(&rel_path)?,
        ) else {
            return Ok(HashMap::new());
        };

        let vm_to_rv = Self::parse_value_metadata(&metadata_xml)?;
        let image_key_positions = Self::parse_rich_value_structures(&structure_xml)?;
        let rich_values = Self::parse_rich_values(&rv_xml)?;
        let rel_ids = Self::parse_rich_value_rel_ids(&rel_xml)?;
        let rel_targets = self
            .read_optional_relationships_for_part(&rel_path)?
            .into_iter()
            .map(|(id, (_, target))| (id, target))
            .collect::<HashMap<_, _>>();

        let mut result = HashMap::new();
        for (vm, rv_index) in vm_to_rv {
            let Some((structure_index, values)) = rich_values.get(rv_index) else {
                continue;
            };
            let Some(&Some(key_pos)) = image_key_positions.get(*structure_index) else {
                continue;
            };
            let Some(rel_index) = values.get(key_pos).and_then(|v| v.parse::<usize>().ok()) else {
                continue;
            };
            let Some(target) = rel_ids.get(rel_index).and_then(|id| rel_targets.get(id)) else {
                continue;
            };
            let filename = target.rsplit('/').next().unwrap_or(target).to_string();
            result.insert(vm, filename);
        }

        Ok(result)
    }

    /// Parse `xl/metadata.xml` into a `vm` (1-based) → rich-value-index map.
    fn parse_value_metadata(xml: &str) -> Result<HashMap<u32, usize>> {
        let mut reader = crate::decode::reader_for(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();

        let mut metadata_type_names: Vec<String> = Vec::new();
        let mut future_rvb: Vec<usize> = Vec::new(); // futureMetadata bk order → rvb i
        let mut value_rcs: Vec<(u32, usize)> = Vec::new(); // valueMetadata bk order → (rc t, rc v)
        let mut in_future_xlrichvalue = false;
        let mut in_value_metadata = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e))
                | Ok(quick_xml::events::Event::Empty(ref e)) => {
                    match e.name().local_name().as_ref() {
                        b"metadataType" => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"name" {
                                    metadata_type_names
                                        .push(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }
                        b"futureMetadata" => {
                            in_future_xlrichvalue = e.attributes().flatten().any(|attr| {
                                attr.key.local_name().as_ref() == b"name"
                                    && attr.value.as_ref() == b"XLRICHVALUE"
                            });
                        }
                        b"rvb" if in_future_xlrichvalue => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"i" {
                                    if let Ok(i) =
                                        String::from_utf8_lossy(&attr.value).parse::<usize>()
                                    {
                                        future_rvb.push(i);
                                    }
                                }
                            }
                        }
                        b"valueMetadata" => in_value_metadata = true,
                        b"rc" if in_value_metadata => {
                            let mut t = None;
                            let mut v = None;
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"t" => t = String::from_utf8_lossy(&attr.value).parse().ok(),
                                    b"v" => v = String::from_utf8_lossy(&attr.value).parse().ok(),
                                    _ => {}
                                }
                            }
                            if let (Some(t), Some(v)) = (t, v) {
                                value_rcs.push((t, v));
                            }
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::End(ref e)) => match e.name().local_name().as_ref() {
                    b"futureMetadata" => in_future_xlrichvalue = false,
                    b"valueMetadata" => in_value_metadata = false,
                    _ => {}
                },
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => return Err(Error::xml_parse_with_context(e.to_string(), "metadata")),
                _ => {}
            }
            buf.clear();
        }

        // rc[@t] is a 1-based index into metadataTypes; only XLRICHVALUE
        // records reference rich values.
        let xlrichvalue_type = metadata_type_names
            .iter()
            .position(|n| n == "XLRICHVALUE")
            .map(|p| (p + 1) as u32);

        let mut map = HashMap::new();
        if let Some(rv_type) = xlrichvalue_type {
            for (idx, (t, v)) in value_rcs.iter().enumerate() {
                if *t == rv_type {
                    if let Some(&rv_index) = future_rvb.get(*v) {
                        map.insert((idx + 1) as u32, rv_index);
                    }
                }
            }
        }
        Ok(map)
    }

    /// Parse `rdrichvaluestructure.xml`: per structure (in order), the
    /// position of the `_rvRel:LocalImageIdentifier` key, if any.
    fn parse_rich_value_structures(xml: &str) -> Result<Vec<Option<usize>>> {
        let mut reader = crate::decode::reader_for(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();

        let mut structures: Vec<Option<usize>> = Vec::new();
        let mut current_key_index = 0usize;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e))
                | Ok(quick_xml::events::Event::Empty(ref e)) => {
                    match e.name().local_name().as_ref() {
                        b"s" => {
                            structures.push(None);
                            current_key_index = 0;
                        }
                        b"k" => {
                            let is_image_key = e.attributes().flatten().any(|attr| {
                                attr.key.local_name().as_ref() == b"n"
                                    && attr.value.as_ref() == b"_rvRel:LocalImageIdentifier"
                            });
                            if is_image_key {
                                if let Some(last) = structures.last_mut() {
                                    *last = Some(current_key_index);
                                }
                            }
                            current_key_index += 1;
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => {
                    return Err(Error::xml_parse_with_context(
                        e.to_string(),
                        "rdrichvaluestructure",
                    ))
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(structures)
    }

    /// Parse `rdrichvalue.xml`: ordered rich values as (structure index, v slots).
    fn parse_rich_values(xml: &str) -> Result<Vec<(usize, Vec<String>)>> {
        let mut reader = crate::decode::reader_for(xml);
        // Preserve intra-<v> whitespace so entity-adjacent spaces survive
        // quick-xml 0.40+ fragmentation; the `in_v` guard already excludes
        // inter-element indentation.
        reader.config_mut().trim_text(false);
        let mut buf = Vec::new();

        let mut values: Vec<(usize, Vec<String>)> = Vec::new();
        let mut in_v = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    match e.name().local_name().as_ref() {
                        b"rv" => {
                            let mut s = 0usize;
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"s" {
                                    s = String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
                                }
                            }
                            values.push((s, Vec::new()));
                        }
                        b"v" => {
                            in_v = true;
                            if let Some((_, slots)) = values.last_mut() {
                                slots.push(String::new());
                            }
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Text(ref e)) if in_v => {
                    if let Some(slot) = values.last_mut().and_then(|(_, s)| s.last_mut()) {
                        slot.push_str(&crate::decode::decode_text_lossy(e));
                    }
                }
                Ok(quick_xml::events::Event::GeneralRef(ref e)) if in_v => {
                    if let Some(slot) = values.last_mut().and_then(|(_, s)| s.last_mut()) {
                        slot.push_str(&resolve_general_ref(e));
                    }
                }
                Ok(quick_xml::events::Event::End(ref e))
                    if e.name().local_name().as_ref() == b"v" =>
                {
                    in_v = false;
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => return Err(Error::xml_parse_with_context(e.to_string(), "rdrichvalue")),
                _ => {}
            }
            buf.clear();
        }

        Ok(values)
    }

    /// Parse `richValueRel.xml`: ordered relationship IDs.
    fn parse_rich_value_rel_ids(xml: &str) -> Result<Vec<String>> {
        let mut reader = crate::decode::reader_for(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();

        let mut ids = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e))
                | Ok(quick_xml::events::Event::Empty(ref e))
                    if e.name().local_name().as_ref() == b"rel" =>
                {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"id" {
                            ids.push(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => return Err(Error::xml_parse_with_context(e.to_string(), "richValueRel")),
                _ => {}
            }
            buf.clear();
        }
        Ok(ids)
    }

    /// Attach a resolved rich-value image to a finished cell.
    ///
    /// The `#VALUE!` placeholder of an error-typed cell is what Excel shows
    /// without rich-value support; the image is the real content, so the
    /// placeholder runs are dropped.
    fn attach_rich_value_image(
        cell: &mut Cell,
        vm: Option<u32>,
        cell_type: Option<&str>,
        rich_value_images: &HashMap<u32, String>,
    ) {
        let Some(filename) = vm.and_then(|v| rich_value_images.get(&v)) else {
            return;
        };
        if cell_type == Some("e") {
            for para in &mut cell.content {
                para.runs.clear();
            }
        }
        let image = InlineImage {
            resource_id: filename.clone(),
            alt_text: None,
            width: None,
            height: None,
        };
        match cell.content.first_mut() {
            Some(para) => para.images.push(image),
            None => cell.content.push(Paragraph {
                images: vec![image],
                ..Default::default()
            }),
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
                ".." => {
                    parts.pop();
                }
                "." | "" => {}
                s => parts.push(s),
            }
        }

        parts.join("/")
    }

    /// Parse a relationships XML file and return a map of id → (type, target).
    #[cfg(test)]
    fn parse_relationships(xml: &str) -> HashMap<String, (String, String)> {
        crate::container::parse_relationships_xml(xml, "worksheet relationships")
            .expect("worksheet relationships should parse in tests")
            .into_type_targets_by_id()
    }

    /// Parse drawing XML for pic elements and create Block::Image for each.
    ///
    /// Structure: xdr:pic > xdr:nvPicPr > xdr:cNvPr[@name]
    ///            xdr:pic > xdr:blipFill > a:blip[@r:embed]
    ///            xdr:pic > xdr:spPr > a:xfrm > a:ext[@cx, @cy]
    fn parse_drawing_images(xml: &str, rels: &HashMap<String, String>) -> Result<Vec<Block>> {
        let mut images = Vec::new();
        let mut reader = crate::decode::reader_for(xml);
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
                Err(e) => return Err(e.into()),
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
    use std::io::{Cursor, Write};

    fn create_test_zip(entries: &[(&str, &str)]) -> Vec<u8> {
        let buf = Vec::new();
        let cursor = Cursor::new(buf);
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        for (name, content) in entries {
            zip.start_file(*name, options).unwrap();
            zip.write_all(content.as_bytes()).unwrap();
        }
        zip.finish().unwrap().into_inner()
    }

    fn test_parser() -> XlsxParser {
        let container = OoxmlContainer::from_bytes(create_test_zip(&[])).unwrap();
        XlsxParser {
            container,
            shared_strings: SharedStrings::default(),
            styles: Styles::default(),
            sheets: Vec::new(),
            relationships: HashMap::new(),
        }
    }

    #[test]
    fn test_parse_sheet_preserves_sparse_cell_positions() {
        let parser = test_parser();
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData>
                    <row r="1">
                        <c r="A1" t="inlineStr"><is><t>left</t></is></c>
                        <c r="C1" t="inlineStr"><is><t>right</t></is></c>
                    </row>
                </sheetData>
            </worksheet>"#;

        let table = parser
            .parse_sheet(sheet_xml, &HashMap::new(), &HashMap::new(), &HashMap::new())
            .unwrap();

        assert_eq!(table.rows.len(), 1);
        assert_eq!(table.rows[0].cells.len(), 3);
        assert_eq!(table.rows[0].cells[0].plain_text(), "left");
        assert_eq!(table.rows[0].cells[1].plain_text(), "");
        assert_eq!(table.rows[0].cells[2].plain_text(), "right");
    }

    #[test]
    fn test_parse_sheet_inline_string_entities_and_cr_refs() {
        // Inline-string cell exercising the cell-value loop's GeneralRef arm and
        // the flush-point CR normalization: entities decode, a CRLF pair encoded
        // as two &#13;&#10; refs collapses to a single LF.
        let parser = test_parser();
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData>
                    <row r="1">
                        <c r="A1" t="inlineStr"><is><t>a&#13;&#10;b &amp; c &bogus; d</t></is></c>
                    </row>
                </sheetData>
            </worksheet>"#;

        let table = parser
            .parse_sheet(sheet_xml, &HashMap::new(), &HashMap::new(), &HashMap::new())
            .unwrap();

        assert_eq!(table.rows[0].cells[0].plain_text(), "a\nb & c &bogus; d");
    }

    #[test]
    fn test_parse_sheet_stray_ampersand_in_skipped_element_does_not_abort() {
        // A raw `&` (ill-formed XML from a non-conforming producer) inside a
        // <f> formula — which the loop skips — must not abort the whole sheet.
        // reader_for enables allow_dangling_amp so tokenization degrades
        // gracefully instead of returning Err(IllFormed(UnclosedReference)).
        let parser = test_parser();
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData>
                    <row r="1">
                        <c r="A1"><f>IF(a&b,1,0)</f><v>5</v></c>
                    </row>
                </sheetData>
            </worksheet>"#;

        let table = parser
            .parse_sheet(sheet_xml, &HashMap::new(), &HashMap::new(), &HashMap::new())
            .expect("stray & must not abort parsing");
        assert_eq!(table.rows[0].cells[0].plain_text(), "5");
    }

    #[test]
    fn test_parse_sheet_trims_formula_only_trailing_row() {
        // The parser does not retain formula presence on the cell model: a formula-only
        // cell with no computed <v> is content-empty. Under issue #6's fix it is therefore
        // trimmed together with any other trailing content-empty row.
        let parser = test_parser();
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData>
                    <row r="1">
                        <c r="A1"><f>SUM(B1:C1)</f></c>
                    </row>
                </sheetData>
            </worksheet>"#;

        let table = parser
            .parse_sheet(sheet_xml, &HashMap::new(), &HashMap::new(), &HashMap::new())
            .unwrap();

        assert_eq!(table.rows.len(), 0);
    }

    #[test]
    fn test_parse_sheet_trims_trailing_rows_with_explicit_empty_cells() {
        // Reproduces issue #6: rows with explicit empty cell references (e.g. from
        // conditional formatting) must be trimmed from the trailing end.
        let parser = test_parser();
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData>
                    <row r="1">
                        <c r="A1" t="inlineStr"><is><t>data</t></is></c>
                    </row>
                    <row r="2">
                        <c r="A2"/>
                        <c r="B2"/>
                    </row>
                    <row r="3">
                        <c r="A3"/>
                    </row>
                </sheetData>
            </worksheet>"#;

        let table = parser
            .parse_sheet(sheet_xml, &HashMap::new(), &HashMap::new(), &HashMap::new())
            .unwrap();

        assert_eq!(
            table.rows.len(),
            1,
            "trailing empty-cell rows must be trimmed"
        );
        assert_eq!(table.rows[0].cells[0].plain_text(), "data");
    }

    #[test]
    fn test_parse_sheet_preserves_formula_only_sparse_middle_cell_position() {
        let parser = test_parser();
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData>
                    <row r="1">
                        <c r="A1" t="inlineStr"><is><t>left</t></is></c>
                        <c r="C1"><f>SUM(A1:B1)</f></c>
                        <c r="E1" t="inlineStr"><is><t>right</t></is></c>
                    </row>
                </sheetData>
            </worksheet>"#;

        let table = parser
            .parse_sheet(sheet_xml, &HashMap::new(), &HashMap::new(), &HashMap::new())
            .unwrap();

        assert_eq!(table.rows.len(), 1);
        assert_eq!(table.rows[0].cells.len(), 5);
        assert_eq!(table.rows[0].cells[0].plain_text(), "left");
        assert_eq!(table.rows[0].cells[1].plain_text(), "");
        assert_eq!(table.rows[0].cells[2].plain_text(), "");
        assert_eq!(table.rows[0].cells[3].plain_text(), "");
        assert_eq!(table.rows[0].cells[4].plain_text(), "right");
    }

    #[test]
    fn test_parse_sheet_preserves_explicit_empty_cell_position() {
        // An explicit empty cell in a MIDDLE position (not trailing) must be preserved.
        // Trailing empty columns are trimmed by the post-parse column-trim pass, but
        // interior gaps between content cells must remain to keep column alignment.
        let parser = test_parser();
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData>
                    <row r="1">
                        <c r="A1" t="inlineStr"><is><t>value</t></is></c>
                        <c r="C1"/>
                        <c r="E1" t="inlineStr"><is><t>end</t></is></c>
                    </row>
                </sheetData>
            </worksheet>"#;

        let table = parser
            .parse_sheet(sheet_xml, &HashMap::new(), &HashMap::new(), &HashMap::new())
            .unwrap();

        assert_eq!(table.rows.len(), 1);
        // E1 (index 4) is the rightmost content cell → keep = 5
        // Interior gaps B1 and C1/D1 are preserved.
        assert_eq!(table.rows[0].cells.len(), 5);
        assert_eq!(table.rows[0].cells[0].plain_text(), "value");
        assert_eq!(table.rows[0].cells[1].plain_text(), "");
        assert_eq!(table.rows[0].cells[2].plain_text(), "");
        assert_eq!(table.rows[0].cells[3].plain_text(), "");
        assert_eq!(table.rows[0].cells[4].plain_text(), "end");
    }

    #[test]
    fn test_parse_sheet_trims_trailing_empty_columns() {
        // Columns C, D (gap), E contain only empty cells — they should be trimmed.
        // Data is only in columns A and B.
        let parser = test_parser();
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData>
                    <row r="1">
                        <c r="A1" t="inlineStr"><is><t>H1</t></is></c>
                        <c r="B1" t="inlineStr"><is><t>H2</t></is></c>
                        <c r="D1"/>
                        <c r="E1"/>
                    </row>
                    <row r="2">
                        <c r="A2" t="inlineStr"><is><t>val</t></is></c>
                        <c r="B2" t="inlineStr"><is><t>42</t></is></c>
                        <c r="D2"/>
                        <c r="E2"/>
                    </row>
                </sheetData>
            </worksheet>"#;

        let table = parser
            .parse_sheet(sheet_xml, &HashMap::new(), &HashMap::new(), &HashMap::new())
            .unwrap();

        assert_eq!(table.rows.len(), 2);
        // Rightmost non-empty column is B (index 1), so each row should have 2 cells.
        // Note: push_cell_with_row_local_spacing adds gap cell for C before D,
        // so before trim each row has [A, B, gap-C, D, E] = 5 cells.
        assert_eq!(
            table.rows[0].cells.len(),
            2,
            "trailing empty columns (C gap, D, E) must be trimmed; got {} cells",
            table.rows[0].cells.len()
        );
        assert_eq!(table.rows[0].cells[0].plain_text(), "H1");
        assert_eq!(table.rows[0].cells[1].plain_text(), "H2");
        assert_eq!(table.rows[1].cells.len(), 2);
    }

    #[test]
    fn test_parse_sheet_preserves_middle_empty_row() {
        // Empty rows between data rows must NOT be trimmed — only trailing ones.
        let parser = test_parser();
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData>
                    <row r="1">
                        <c r="A1" t="inlineStr"><is><t>first</t></is></c>
                    </row>
                    <row r="2">
                        <c r="A2"/>
                    </row>
                    <row r="3">
                        <c r="A3" t="inlineStr"><is><t>third</t></is></c>
                    </row>
                </sheetData>
            </worksheet>"#;

        let table = parser
            .parse_sheet(sheet_xml, &HashMap::new(), &HashMap::new(), &HashMap::new())
            .unwrap();

        assert_eq!(
            table.rows.len(),
            3,
            "middle empty row must be preserved, only trailing empty rows are trimmed"
        );
        assert_eq!(table.rows[0].cells[0].plain_text(), "first");
        assert_eq!(table.rows[2].cells[0].plain_text(), "third");
    }

    #[test]
    fn test_parse_sheet_sparse_high_column_reconstruction_is_row_local() {
        let parser = test_parser();
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData>
                    <row r="1">
                        <c r="BA1" t="inlineStr"><is><t>far</t></is></c>
                    </row>
                    <row r="2">
                        <c r="A2" t="inlineStr"><is><t>near</t></is></c>
                    </row>
                </sheetData>
            </worksheet>"#;

        let table = parser
            .parse_sheet(sheet_xml, &HashMap::new(), &HashMap::new(), &HashMap::new())
            .unwrap();

        assert_eq!(table.rows.len(), 2);
        assert_eq!(table.rows[0].cells.len(), 53);
        assert_eq!(table.rows[0].cells[52].plain_text(), "far");
        assert_eq!(table.rows[1].cells.len(), 1);
        assert_eq!(table.rows[1].cells[0].plain_text(), "near");
    }

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

    fn create_minimal_xlsx(
        workbook_rels_xml: Option<&str>,
        sheet_rels_xml: Option<&str>,
    ) -> Vec<u8> {
        create_minimal_xlsx_with_parts(
            workbook_rels_xml,
            sheet_rels_xml,
            None,
            r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Hello</t></is></c>
    </row>
  </sheetData>
</worksheet>"#,
        )
    }

    fn create_minimal_xlsx_with_parts(
        workbook_rels_xml: Option<&str>,
        sheet_rels_xml: Option<&str>,
        shared_strings_xml: Option<&str>,
        sheet_xml: &str,
    ) -> Vec<u8> {
        use std::io::{Cursor, Write};

        let buf = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(buf);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);

        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>"#,
        )
        .unwrap();

        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
        )
        .unwrap();

        if let Some(workbook_rels_xml) = workbook_rels_xml {
            zip.start_file("xl/_rels/workbook.xml.rels", options)
                .unwrap();
            zip.write_all(workbook_rels_xml.as_bytes()).unwrap();
        }

        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#,
        )
        .unwrap();

        if let Some(sheet_rels_xml) = sheet_rels_xml {
            zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options)
                .unwrap();
            zip.write_all(sheet_rels_xml.as_bytes()).unwrap();
        }

        if let Some(shared_strings_xml) = shared_strings_xml {
            zip.start_file("xl/sharedStrings.xml", options).unwrap();
            zip.write_all(shared_strings_xml.as_bytes()).unwrap();
        }

        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(sheet_xml.as_bytes()).unwrap();

        zip.finish().unwrap().into_inner()
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
    fn test_cell_carriage_returns_do_not_break_markdown_table() {
        // Excel/openpyxl store in-cell CR and CRLF as &#13; character
        // references. A bare CR leaking into a Markdown pipe table splits
        // the row across physical lines (issue #7).
        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>h1</t></is></c>
      <c r="B1" t="inlineStr"><is><t>h2</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>cr</t></is></c>
      <c r="B2" t="inlineStr"><is><t>line one&#13;line two</t></is></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>crlf</t></is></c>
      <c r="B3" t="inlineStr"><is><t>line one&#13;&#10;line two</t></is></c>
    </row>
  </sheetData>
</worksheet>"#;
        let data = create_minimal_xlsx_with_parts(Some(workbook_rels), None, None, sheet_xml);

        let mut parser = XlsxParser::from_bytes(data).unwrap();
        let doc = parser.parse().unwrap();
        let md =
            crate::render::to_markdown(&doc, &crate::render::RenderOptions::default()).unwrap();

        assert!(!md.contains('\r'), "bare CR leaked into markdown: {md:?}");
        for label in ["cr", "crlf"] {
            let row_line = md
                .lines()
                .find(|l| l.starts_with(&format!("| {label} ")))
                .unwrap_or_else(|| panic!("row '{label}' missing in:\n{md}"));
            assert!(
                row_line.contains("line one") && row_line.contains("line two"),
                "row '{label}' split across lines:\n{md}"
            );
        }
    }

    /// Build an XLSX containing one "Place in Cell" rich-value image in B1.
    /// Mirrors the part layout Excel writes: cell `t="e" vm="1"` with a
    /// `#VALUE!` placeholder, resolved through xl/metadata.xml and the
    /// xl/richData/* chain to xl/media/image1.png.
    fn create_rich_value_image_xlsx() -> Vec<u8> {
        create_test_zip(&[
            (
                "[Content_Types].xml",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"#,
            ),
            (
                "_rels/.rels",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
            ),
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/>
  <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue" Target="richData/rdrichvalue.xml"/>
  <Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure" Target="richData/rdrichvaluestructure.xml"/>
  <Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRel" Target="richData/richValueRel.xml"/>
</Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>label</t></is></c>
      <c r="B1" t="e" vm="1"><v>#VALUE!</v></c>
    </row>
  </sheetData>
</worksheet>"#,
            ),
            (
                "xl/metadata.xml",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1"/>
  </metadataTypes>
  <futureMetadata name="XLRICHVALUE" count="1">
    <bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="0"/></ext></extLst></bk>
  </futureMetadata>
  <valueMetadata count="1">
    <bk><rc t="1" v="0"/></bk>
  </valueMetadata>
</metadata>"#,
            ),
            (
                "xl/richData/rdrichvaluestructure.xml",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">
  <s t="_localImage">
    <k n="_rvRel:LocalImageIdentifier" t="i"/>
    <k n="CalcOrigin" t="i"/>
  </s>
</rvStructures>"#,
            ),
            (
                "xl/richData/rdrichvalue.xml",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">
  <rv s="0"><v>0</v><v>5</v></rv>
</rvData>"#,
            ),
            (
                "xl/richData/richValueRel.xml",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
</richValueRels>"#,
            ),
            (
                "xl/richData/_rels/richValueRel.xml.rels",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#,
            ),
            ("xl/media/image1.png", "fake-png-bytes"),
        ])
    }

    #[test]
    fn test_rich_value_in_cell_image_extracted() {
        let data = create_rich_value_image_xlsx();
        let mut parser = XlsxParser::from_bytes(data).unwrap();
        let doc = parser.parse().unwrap();

        let table = doc.sections[0]
            .content
            .iter()
            .find_map(|b| match b {
                Block::Table(t) => Some(t),
                _ => None,
            })
            .expect("table missing");
        let cell = &table.rows[0].cells[1];

        assert!(
            !cell.plain_text().contains("#VALUE!"),
            "placeholder error text must be replaced by the image: {:?}",
            cell.plain_text()
        );
        let images: Vec<_> = cell.content.iter().flat_map(|p| &p.images).collect();
        assert_eq!(images.len(), 1, "in-cell image not extracted");
        assert_eq!(images[0].resource_id, "image1.png");

        let md =
            crate::render::to_markdown(&doc, &crate::render::RenderOptions::default()).unwrap();
        assert!(
            md.contains("![") && md.contains("image1.png"),
            "markdown must reference the in-cell image:\n{md}"
        );
        assert!(
            doc.resources.contains_key("image1.png"),
            "media file must be extracted as a resource"
        );
    }

    #[test]
    fn test_rich_value_vm_without_rich_data_parts_keeps_placeholder() {
        // A vm attribute with no resolvable richData chain (foreign producer,
        // stripped file) must degrade gracefully: keep the cell text, no panic.
        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="e" vm="1"><v>#VALUE!</v></c>
    </row>
  </sheetData>
</worksheet>"#;
        let data = create_minimal_xlsx_with_parts(Some(workbook_rels), None, None, sheet_xml);

        let mut parser = XlsxParser::from_bytes(data).unwrap();
        let doc = parser.parse().unwrap();

        let table = doc.sections[0]
            .content
            .iter()
            .find_map(|b| match b {
                Block::Table(t) => Some(t),
                _ => None,
            })
            .expect("table missing");
        assert!(table.rows[0].cells[0].plain_text().contains("#VALUE!"));
    }

    #[test]
    fn test_xlsx_requires_workbook_relationships() {
        let data = create_minimal_xlsx(None, None);
        let err = XlsxParser::from_bytes(data)
            .err()
            .expect("missing workbook relationships should fail");

        match err {
            Error::MissingComponent(path) => assert_eq!(path, "xl/_rels/workbook.xml.rels"),
            other => panic!("expected missing workbook rels error, got {other:?}"),
        }
    }

    #[test]
    fn test_xlsx_rejects_malformed_workbook_relationships() {
        let data = create_minimal_xlsx(Some("<Relationships"), None);
        let err = XlsxParser::from_bytes(data)
            .err()
            .expect("malformed workbook relationships should fail");

        match err {
            Error::XmlParseWithContext { location, .. } => {
                assert_eq!(location, "xl/_rels/workbook.xml.rels")
            }
            other => panic!("expected malformed workbook rels error, got {other:?}"),
        }
    }

    #[test]
    fn test_xlsx_non_utf8_workbook_is_error() {
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
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>"#).unwrap();

        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#).unwrap();

        zip.start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#,
        )
        .unwrap();

        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(b"<?xml version=\"1.0\"?><workbook>Caf\xe9</workbook>")
            .unwrap();

        let data = zip.finish().unwrap().into_inner();
        let err = match XlsxParser::from_bytes(data) {
            Ok(_) => panic!("non-UTF-8 workbook must surface Error::Encoding"),
            Err(err) => err,
        };

        assert!(
            matches!(err, Error::Encoding(_)),
            "expected Error::Encoding, got {err:?}"
        );
    }

    fn create_minimal_xlsx_with_malformed_optional_part(extra_part_path: &str) -> Vec<u8> {
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
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>"#).unwrap();

        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#).unwrap();

        zip.start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#,
        )
        .unwrap();

        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets/>
</workbook>"#,
        )
        .unwrap();

        zip.start_file(extra_part_path, options).unwrap();
        zip.write_all(b"<?xml version=\"1.0\"?><root>Caf\xe9</root>")
            .unwrap();

        zip.finish().unwrap().into_inner()
    }

    #[test]
    fn test_xlsx_non_utf8_optional_parts_surface_encoding_error() {
        // Optional XLSX parts (shared strings, styles) with malformed byte
        // content must surface Error::Encoding instead of being silently
        // treated as absent.
        for part_path in &["xl/sharedStrings.xml", "xl/styles.xml"] {
            let data = create_minimal_xlsx_with_malformed_optional_part(part_path);
            let err = match XlsxParser::from_bytes(data) {
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
    fn test_xlsx_allows_missing_optional_sheet_relationships() {
        let data = create_minimal_xlsx(
            Some(
                r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"></Relationship>
</Relationships>"#,
            ),
            None,
        );
        let mut parser = XlsxParser::from_bytes(data).unwrap();
        let doc = parser.parse().unwrap();

        assert_eq!(doc.sections.len(), 1);
    }

    #[test]
    fn test_xlsx_rejects_malformed_optional_sheet_relationships() {
        let data = create_minimal_xlsx(
            Some(
                r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#,
            ),
            Some("<Relationships"),
        );
        let mut parser = XlsxParser::from_bytes(data).unwrap();
        let err = parser
            .parse()
            .expect_err("malformed optional sheet relationships should fail");

        match err {
            Error::XmlParseWithContext { location, .. } => {
                assert_eq!(location, "xl/worksheets/_rels/sheet1.xml.rels")
            }
            other => panic!("expected malformed optional sheet rels error, got {other:?}"),
        }
    }

    #[test]
    fn test_xlsx_out_of_range_shared_string_index_is_error() {
        let data = create_minimal_xlsx_with_parts(
            Some(
                r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#,
            ),
            None,
            Some(
                r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <si><t>Only entry</t></si>
</sst>"#,
            ),
            r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>3</v></c>
    </row>
  </sheetData>
</worksheet>"#,
        );
        let mut parser = XlsxParser::from_bytes(data).unwrap();
        let err = parser
            .parse()
            .expect_err("out-of-range shared string index should fail");

        match err {
            Error::InvalidData(message) => {
                assert!(message.contains("shared string index out of range: 3"))
            }
            other => panic!("expected invalid shared string error, got {other:?}"),
        }
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
            Block::Image {
                resource_id,
                alt_text,
                width,
                height,
            } => {
                assert_eq!(resource_id, "image1.png");
                assert_eq!(alt_text.as_deref(), Some("Logo"));
                assert_eq!(*width, Some(914400));
                assert_eq!(*height, Some(457200));
            }
            other => panic!("Expected Block::Image, got {:?}", other),
        }

        match &images[1] {
            Block::Image {
                resource_id,
                alt_text,
                width,
                height,
            } => {
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
            Block::Image {
                resource_id,
                alt_text,
                ..
            } => {
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

    #[test]
    fn test_parse_comments_xml_basic() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <authors>
                    <author>John Doe</author>
                </authors>
                <commentList>
                    <comment ref="A1" authorId="0">
                        <text>
                            <r><t>This is a comment</t></r>
                        </text>
                    </comment>
                    <comment ref="B5" authorId="0">
                        <text>
                            <r><t>Another comment</t></r>
                        </text>
                    </comment>
                </commentList>
            </comments>"#;

        let comments = XlsxParser::parse_comments_xml(xml);
        assert_eq!(comments.len(), 2);
        assert_eq!(comments.get("A1").unwrap(), "This is a comment");
        assert_eq!(comments.get("B5").unwrap(), "Another comment");
    }

    #[test]
    fn test_parse_comments_xml_multiple_runs() {
        // Comment with multiple <r><t>...</t></r> runs should concatenate text
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <authors><author>Author</author></authors>
                <commentList>
                    <comment ref="C3" authorId="0">
                        <text>
                            <r><t>First part</t></r>
                            <r><t>Second part</t></r>
                        </text>
                    </comment>
                </commentList>
            </comments>"#;

        let comments = XlsxParser::parse_comments_xml(xml);
        assert_eq!(comments.len(), 1);
        assert_eq!(comments.get("C3").unwrap(), "First part Second part");
    }

    #[test]
    fn test_parse_comments_xml_empty() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <authors><author>Author</author></authors>
                <commentList/>
            </comments>"#;

        let comments = XlsxParser::parse_comments_xml(xml);
        assert!(comments.is_empty());
    }

    #[test]
    fn test_parse_comments_xml_case_insensitive_ref() {
        // Cell ref should be uppercased
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <authors><author>Author</author></authors>
                <commentList>
                    <comment ref="a1" authorId="0">
                        <text><r><t>Lowercase ref</t></r></text>
                    </comment>
                </commentList>
            </comments>"#;

        let comments = XlsxParser::parse_comments_xml(xml);
        assert_eq!(comments.len(), 1);
        assert!(comments.contains_key("A1"));
        assert_eq!(comments.get("A1").unwrap(), "Lowercase ref");
    }

    #[test]
    fn test_parse_comments_xml_malformed_entity_preserves_raw_text() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <authors><author>Author</author></authors>
                <commentList>
                    <comment ref="A1" authorId="0">
                        <text><r><t>Bad &bogus; comment</t></r></text>
                    </comment>
                </commentList>
            </comments>"#;

        let comments = XlsxParser::parse_comments_xml(xml);
        assert_eq!(comments.get("A1").unwrap(), "Bad &bogus; comment");
    }

    #[test]
    fn test_find_and_parse_comments_resolves_path() {
        // Verify find_and_parse_comments looks for "comments" in relationship types
        let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
                <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
            </Relationships>"#;

        let sheet_rels = XlsxParser::parse_relationships(rels_xml);

        // Verify the relationship is correctly identified as comments type
        let (rel_type, target) = sheet_rels.get("rId2").unwrap();
        assert!(rel_type.contains("comments"));
        assert_eq!(target, "../comments1.xml");

        // Verify path resolution
        let resolved = XlsxParser::resolve_relative_path("xl/worksheets", target);
        assert_eq!(resolved, "xl/comments1.xml");
    }

    #[test]
    fn test_xlsx_inline_str_mixed_entities_preserve_legitimate_and_malformed() {
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
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"#).unwrap();

            zip.start_file("_rels/.rels", options).unwrap();
            zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#).unwrap();

            zip.start_file("xl/_rels/workbook.xml.rels", options)
                .unwrap();
            zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#).unwrap();

            zip.start_file("xl/workbook.xml", options).unwrap();
            zip.write_all(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#,
            )
            .unwrap();

            zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
            zip.write_all(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="str"><v>A &amp; B &bogus; C</v></c></row>
  </sheetData>
</worksheet>"#,
            )
            .unwrap();

            zip.finish().unwrap();
        }

        let mut parser = XlsxParser::from_bytes(buf).expect("parser opens");
        let doc = parser.parse().expect("document parses");
        let text = doc.plain_text();
        assert!(
            text.contains("A & B &bogus; C"),
            "expected legitimate decoded + malformed preserved; got {text:?}"
        );
        assert!(
            !text.contains("A &amp; B"),
            "legitimate entity must not remain escaped; got {text:?}"
        );
    }

    #[test]
    fn test_xlsx_missing_workbook_surfaces_missing_component() {
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
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>"#,
            )
            .unwrap();

            zip.start_file("_rels/.rels", options).unwrap();
            zip.write_all(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
            )
            .unwrap();

            // xl/workbook.xml INTENTIONALLY ABSENT.
            // xl/_rels/workbook.xml.rels must be present — otherwise the earlier
            // read_required_relationships_for_part call in from_container would
            // fail first, masking the workbook-missing error we're testing.
            zip.start_file("xl/_rels/workbook.xml.rels", options)
                .unwrap();
            zip.write_all(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#,
            )
            .unwrap();

            zip.finish().unwrap();
        }

        let err = XlsxParser::from_bytes(buf)
            .err()
            .expect("must fail on missing workbook");
        match err {
            Error::MissingComponent(path) => {
                assert_eq!(path, "xl/workbook.xml");
            }
            other => panic!("expected MissingComponent(\"xl/workbook.xml\"), got {other:?}"),
        }
    }
}
