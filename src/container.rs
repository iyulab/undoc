//! ZIP container abstraction for OOXML documents.

use crate::decode::{normalize_line_endings, resolve_general_ref};
use crate::error::{Error, Result};
use crate::model::Metadata;
use std::borrow::Cow;
use std::cell::RefCell;
use std::collections::HashMap;
#[cfg(not(target_arch = "wasm32"))]
use std::fs::File;
#[cfg(not(target_arch = "wasm32"))]
use std::io::BufReader;
use std::io::{Cursor, Read, Seek};
use std::path::Path;

/// Classify a failure to open an entry inside the archive.
///
/// Only a genuinely absent entry is a missing component — that is the case
/// [`OoxmlContainer::read_xml_optional`] turns into `Ok(None)`. Every other reason (a
/// damaged container, one that needs a password) has to keep its own classification,
/// or the optional-part path silently degrades a broken file into a valid-empty
/// result — exactly what that method promises not to do.
fn entry_error(path: &str, err: zip::result::ZipError) -> Error {
    match err {
        zip::result::ZipError::FileNotFound => Error::MissingComponent(path.to_string()),
        other => Error::from(other),
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RelationshipPolicy {
    Required,
    Optional,
}

/// A relationship entry from a .rels file.
#[derive(Debug, Clone)]
pub struct Relationship {
    /// Relationship ID (e.g., "rId1")
    pub id: String,
    /// Relationship type URI
    pub rel_type: String,
    /// Target path (relative or absolute)
    pub target: String,
    /// Whether the target is external
    pub external: bool,
}

/// Collection of relationships parsed from a .rels file.
#[derive(Debug, Clone, Default)]
pub struct Relationships {
    /// Map from relationship ID to relationship data
    pub by_id: HashMap<String, Relationship>,
    /// Map from relationship type to list of relationships
    pub by_type: HashMap<String, Vec<Relationship>>,
}

impl Relationships {
    /// Create a new empty relationships collection.
    pub fn new() -> Self {
        Self::default()
    }

    /// Get a relationship by ID.
    pub fn get(&self, id: &str) -> Option<&Relationship> {
        self.by_id.get(id)
    }

    /// Get relationships by type.
    pub fn get_by_type(&self, rel_type: &str) -> Vec<&Relationship> {
        self.by_type
            .get(rel_type)
            .map(|v| v.iter().collect())
            .unwrap_or_default()
    }

    /// Add a relationship.
    pub fn add(&mut self, rel: Relationship) {
        self.by_type
            .entry(rel.rel_type.clone())
            .or_default()
            .push(rel.clone());
        self.by_id.insert(rel.id.clone(), rel);
    }

    /// Consume the collection into an ID → target map.
    pub fn into_targets_by_id(self) -> HashMap<String, String> {
        self.by_id
            .into_iter()
            .map(|(id, rel)| (id, rel.target))
            .collect()
    }

    /// Consume the collection into an ID → (type, target) map.
    pub fn into_type_targets_by_id(self) -> HashMap<String, (String, String)> {
        self.by_id
            .into_iter()
            .map(|(id, rel)| (id, (rel.rel_type, rel.target)))
            .collect()
    }
}

/// Restate the XML declaration's encoding as UTF-8 after a decode.
///
/// Once the bytes have been decoded into a Rust `String` the content *is* UTF-8, so a
/// declaration still naming the source encoding contradicts the string carrying it.
/// Callers get the part as text and may re-serialize or re-read it; leaving the stale
/// name in place hands them a document that describes itself wrongly.
///
/// Whatever the declared value was, it is no longer the truth — so the value is replaced
/// rather than matched against a list of spellings (`UTF-16`, `utf-16le`, `UTF-16BE`, …).
fn fix_xml_encoding_declaration(content: &str) -> String {
    if !content.starts_with("<?xml") {
        return content.to_string();
    }
    let Some(end_decl) = content.find("?>") else {
        return content.to_string();
    };
    let (decl, rest) = content.split_at(end_decl + 2);

    let Some(value_start) = decl.find("encoding=") else {
        return content.to_string();
    };
    let after = &decl[value_start + "encoding=".len()..];
    let Some(quote) = after.chars().next().filter(|c| *c == '"' || *c == '\'') else {
        return content.to_string();
    };
    let Some(value_len) = after[1..].find(quote) else {
        return content.to_string();
    };

    format!(
        "{}encoding={quote}UTF-8{quote}{}{}",
        &decl[..value_start],
        &after[1 + value_len + 1..],
        rest
    )
}

/// OOXML container abstraction over a ZIP archive.
///
/// Provides methods to read XML files, binary data, and relationships
/// from an Office Open XML document.
pub struct OoxmlContainer {
    archive: RefCell<zip::ZipArchive<Cursor<Vec<u8>>>>,
    /// Cached package-level relationships (used in Phase 2+)
    #[allow(dead_code)]
    package_rels: Option<Relationships>,
}

/// Decode XML bytes handling different encodings (UTF-8, UTF-16 LE/BE).
///
/// OOXML files are typically UTF-8 encoded, but some (especially older
/// or non-standard documents) may use UTF-16 encoding.
///
/// A BOM decides the encoding outright. Without one, UTF-16 is detected *before* UTF-8 is
/// attempted, because BOM-less UTF-16 of ASCII text is a valid UTF-8 byte sequence: every
/// second byte is NUL, and NUL is a legal UTF-8 encoding of U+0000. Trying UTF-8 first
/// would therefore succeed and yield a string whose element names are interleaved with
/// NUL — no XML reader matches anything in it, and the part parses to nothing at all. The
/// detection cannot misfire on real UTF-8 XML, since XML 1.0 forbids U+0000 in content.
pub fn decode_xml_bytes(bytes: &[u8]) -> Result<String> {
    // Check for BOM (Byte Order Mark)
    if bytes.len() >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF {
        // UTF-8 BOM: EF BB BF - skip BOM and decode as UTF-8
        return String::from_utf8(bytes[3..].to_vec())
            .map_err(|e| Error::Encoding(format!("UTF-8 XML part has invalid bytes: {e}")));
    }

    if bytes.len() >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE {
        // UTF-16 LE BOM: FF FE
        let content = decode_utf16_le(&bytes[2..])?;
        // Fix XML declaration encoding to UTF-8 since we've already converted
        return Ok(fix_xml_encoding_declaration(&content));
    }

    if bytes.len() >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF {
        // UTF-16 BE BOM: FE FF
        let content = decode_utf16_be(&bytes[2..])?;
        // Fix XML declaration encoding to UTF-8 since we've already converted
        return Ok(fix_xml_encoding_declaration(&content));
    }

    // No BOM. XML 1.0 requires one on UTF-16 entities, so this is already non-conformant
    // input — but it is recoverable, and the NUL pattern says which half of the code unit
    // holds the ASCII byte.
    if bytes.len() >= 4 && bytes[1] == 0 && bytes[3] == 0 {
        let content = decode_utf16_le(bytes)?;
        return Ok(fix_xml_encoding_declaration(&content));
    }
    if bytes.len() >= 4 && bytes[0] == 0 && bytes[2] == 0 {
        let content = decode_utf16_be(bytes)?;
        return Ok(fix_xml_encoding_declaration(&content));
    }

    String::from_utf8(bytes.to_vec())
        .map_err(|_| Error::Encoding("XML part is not valid UTF-8 or UTF-16".to_string()))
}

/// Decode UTF-16 Little Endian bytes to String.
fn decode_utf16_le(bytes: &[u8]) -> Result<String> {
    if !bytes.len().is_multiple_of(2) {
        return Err(Error::Encoding(
            "UTF-16 little-endian XML has an odd byte length".to_string(),
        ));
    }

    let u16_iter = (0..bytes.len())
        .step_by(2)
        .map(|i| u16::from_le_bytes([bytes[i], bytes[i + 1]]));

    char::decode_utf16(u16_iter)
        .collect::<std::result::Result<String, _>>()
        .map_err(|e| Error::Encoding(format!("UTF-16 little-endian XML is malformed: {e}")))
}

/// Decode UTF-16 Big Endian bytes to String.
fn decode_utf16_be(bytes: &[u8]) -> Result<String> {
    if !bytes.len().is_multiple_of(2) {
        return Err(Error::Encoding(
            "UTF-16 big-endian XML has an odd byte length".to_string(),
        ));
    }

    let u16_iter = (0..bytes.len())
        .step_by(2)
        .map(|i| u16::from_be_bytes([bytes[i], bytes[i + 1]]));

    char::decode_utf16(u16_iter)
        .collect::<std::result::Result<String, _>>()
        .map_err(|e| Error::Encoding(format!("UTF-16 big-endian XML is malformed: {e}")))
}

impl OoxmlContainer {
    fn rels_path_for_part(part_path: &str) -> String {
        if part_path.is_empty() || part_path == "/" {
            "_rels/.rels".to_string()
        } else {
            let path = Path::new(part_path);
            let parent = path.parent().unwrap_or(Path::new(""));
            let filename = path.file_name().unwrap_or_default().to_string_lossy();
            format!("{}/_rels/{}.rels", parent.display(), filename)
        }
    }

    /// Open an OOXML container from a file path.
    ///
    /// # Example
    ///
    /// ```no_run
    /// use undoc::container::OoxmlContainer;
    ///
    /// let container = OoxmlContainer::open("document.docx")?;
    /// # Ok::<(), undoc::Error>(())
    /// ```
    #[cfg(not(target_arch = "wasm32"))]
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let file = File::open(path.as_ref())?;
        let mut reader = BufReader::new(file);
        let mut data = Vec::new();
        reader.read_to_end(&mut data)?;
        Self::from_bytes(data)
    }

    /// Create an OOXML container from a byte vector.
    pub fn from_bytes(data: Vec<u8>) -> Result<Self> {
        let cursor = Cursor::new(data);
        let archive = zip::ZipArchive::new(cursor)?;
        Ok(Self {
            archive: RefCell::new(archive),
            package_rels: None,
        })
    }

    /// Create an OOXML container from a reader.
    pub fn from_reader<R: Read + Seek>(mut reader: R) -> Result<Self> {
        let mut data = Vec::new();
        reader.read_to_end(&mut data)?;
        Self::from_bytes(data)
    }

    /// Read an XML file from the archive as a string.
    ///
    /// Handles different encodings:
    /// - UTF-8 (with or without BOM)
    /// - UTF-16 LE (with BOM: FF FE)
    /// - UTF-16 BE (with BOM: FE FF)
    pub fn read_xml(&self, path: &str) -> Result<String> {
        let mut archive = self.archive.borrow_mut();
        let mut file = archive.by_name(path).map_err(|e| entry_error(path, e))?;

        // Read raw bytes first
        let mut bytes = Vec::new();
        file.read_to_end(&mut bytes)?;

        // Detect encoding from BOM
        let content = decode_xml_bytes(&bytes)?;
        Ok(content)
    }

    /// Read an optional XML part.
    ///
    /// Returns `Ok(None)` when the part is absent (`Error::MissingComponent`),
    /// but surfaces any other error — including `Error::Encoding` for malformed
    /// byte content — so that corrupted optional parts are never silently
    /// degraded into a valid-empty result.
    pub fn read_xml_optional(&self, path: &str) -> Result<Option<String>> {
        match self.read_xml(path) {
            Ok(content) => Ok(Some(content)),
            Err(Error::MissingComponent(_)) => Ok(None),
            Err(e) => Err(e),
        }
    }

    /// Read a binary file from the archive.
    pub fn read_binary(&self, path: &str) -> Result<Vec<u8>> {
        let mut archive = self.archive.borrow_mut();
        let mut file = archive.by_name(path).map_err(|e| entry_error(path, e))?;
        let mut data = Vec::new();
        file.read_to_end(&mut data)?;
        Ok(data)
    }

    /// Check if a file exists in the archive.
    pub fn exists(&self, path: &str) -> bool {
        let archive = self.archive.borrow();
        let result = archive.file_names().any(|n| n == path);
        result
    }

    /// List all files in the archive.
    pub fn list_files(&self) -> Vec<String> {
        let archive = self.archive.borrow();
        archive.file_names().map(String::from).collect()
    }

    /// List files matching a prefix.
    pub fn list_files_with_prefix(&self, prefix: &str) -> Vec<String> {
        let archive = self.archive.borrow();
        archive
            .file_names()
            .filter(|n| n.starts_with(prefix))
            .map(String::from)
            .collect()
    }

    /// Read and parse relationships from a .rels file.
    ///
    /// Legacy helper: treats missing relationship parts as optional.
    pub fn read_relationships(&self, part_path: &str) -> Result<Relationships> {
        self.read_optional_relationships_for_part(part_path)
    }

    /// Read relationships for a required OOXML part.
    pub fn read_required_relationships_for_part(&self, part_path: &str) -> Result<Relationships> {
        let rels_path = Self::rels_path_for_part(part_path);
        self.read_relationships_at_path(&rels_path, RelationshipPolicy::Required)
    }

    /// Read relationships for an optional OOXML part.
    pub fn read_optional_relationships_for_part(&self, part_path: &str) -> Result<Relationships> {
        let rels_path = Self::rels_path_for_part(part_path);
        self.read_relationships_at_path(&rels_path, RelationshipPolicy::Optional)
    }

    fn read_relationships_at_path(
        &self,
        rels_path: &str,
        policy: RelationshipPolicy,
    ) -> Result<Relationships> {
        let content = match self.read_xml(rels_path) {
            Ok(content) => content,
            Err(Error::MissingComponent(_)) if policy == RelationshipPolicy::Optional => {
                return Ok(Relationships::new());
            }
            Err(err) => return Err(err),
        };

        parse_relationships_xml(&content, rels_path)
    }

    /// Read package-level relationships (_rels/.rels).
    pub fn read_package_relationships(&self) -> Result<Relationships> {
        self.read_relationships_at_path("_rels/.rels", RelationshipPolicy::Required)
    }

    /// Parse core metadata from docProps/core.xml.
    ///
    /// This is common to all OOXML formats (DOCX, XLSX, PPTX).
    pub fn parse_core_metadata(&self) -> Result<Metadata> {
        let mut meta = Metadata::default();

        match self.read_xml("docProps/core.xml") {
            Ok(xml) => {
                let mut reader = crate::decode::reader_for(&xml);
                reader.config_mut().trim_text(false);

                let mut buf = Vec::new();
                let mut current_element: Option<String> = None;
                // Accumulate an element's text across events. quick-xml 0.40+ splits
                // a value like "Tom &amp; Jerry" into Text/GeneralRef/Text, so the
                // value must be built up and flushed at the element's End rather than
                // assigned per Text event (which would drop all but the last fragment).
                let mut current_text = String::new();

                loop {
                    match reader.read_event_into(&mut buf) {
                        Ok(quick_xml::events::Event::Start(e)) => {
                            let name = e.name();
                            current_element = Some(
                                String::from_utf8_lossy(name.local_name().as_ref()).to_string(),
                            );
                            current_text.clear();
                        }
                        Ok(quick_xml::events::Event::Text(e)) if current_element.is_some() => {
                            current_text.push_str(&metadata_text_or_raw(&e, "docProps/core.xml")?);
                        }
                        Ok(quick_xml::events::Event::GeneralRef(e))
                            if current_element.is_some() =>
                        {
                            current_text.push_str(&resolve_general_ref(&e));
                        }
                        Ok(quick_xml::events::Event::End(_)) => {
                            if let Some(ref elem) = current_element {
                                let text = normalize_line_endings(Cow::Borrowed(&current_text))
                                    .into_owned();
                                match elem.as_str() {
                                    "title" => meta.title = Some(text),
                                    "creator" => meta.author = Some(text),
                                    "subject" => meta.subject = Some(text),
                                    "description" => meta.description = Some(text),
                                    "keywords" => {
                                        meta.keywords = text
                                            .split([',', ';'])
                                            .map(|s| s.trim().to_string())
                                            .filter(|s| !s.is_empty())
                                            .collect();
                                    }
                                    "created" => meta.created = Some(text),
                                    "modified" => meta.modified = Some(text),
                                    "lastModifiedBy" => meta.last_modified_by = Some(text),
                                    _ => {}
                                }
                            }
                            current_element = None;
                            current_text.clear();
                        }
                        Ok(quick_xml::events::Event::Eof) => break,
                        Err(e) => {
                            return Err(Error::xml_parse_with_context(
                                e.to_string(),
                                "docProps/core.xml",
                            ))
                        }
                        _ => {}
                    }
                    buf.clear();
                }
            }
            Err(Error::MissingComponent(_)) => {}
            Err(err) => return Err(err),
        }

        // Enrich with app.xml metadata
        self.parse_app_metadata(&mut meta)?;

        Ok(meta)
    }

    /// Parse extended metadata from docProps/app.xml.
    ///
    /// Extracts Application, Pages, Words, and Slides properties.
    /// Only sets `page_count` and `word_count` if not already populated.
    fn parse_app_metadata(&self, meta: &mut Metadata) -> Result<()> {
        let xml = match self.read_xml("docProps/app.xml") {
            Ok(xml) => xml,
            Err(Error::MissingComponent(_)) => return Ok(()),
            Err(err) => return Err(err),
        };

        let mut reader = crate::decode::reader_for(&xml);
        reader.config_mut().trim_text(false);

        let mut buf = Vec::new();
        let mut current_element: Option<String> = None;
        // Accumulate per element and flush at End — see parse_core_metadata for
        // why per-Text assignment is unsafe under quick-xml 0.40+.
        let mut current_text = String::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(e)) => {
                    let name = e.name();
                    current_element =
                        Some(String::from_utf8_lossy(name.local_name().as_ref()).to_string());
                    current_text.clear();
                }
                Ok(quick_xml::events::Event::Text(e)) if current_element.is_some() => {
                    current_text.push_str(&metadata_text_or_raw(&e, "docProps/app.xml")?);
                }
                Ok(quick_xml::events::Event::GeneralRef(e)) if current_element.is_some() => {
                    current_text.push_str(&resolve_general_ref(&e));
                }
                Ok(quick_xml::events::Event::End(_)) => {
                    if let Some(ref elem) = current_element {
                        let text =
                            normalize_line_endings(Cow::Borrowed(&current_text)).into_owned();
                        match elem.as_str() {
                            "Application" => meta.application = Some(text),
                            "Pages" if meta.page_count.is_none() => {
                                meta.page_count = text.trim().parse::<u32>().ok();
                            }
                            "Words" if meta.word_count.is_none() => {
                                meta.word_count = text.trim().parse::<u32>().ok();
                            }
                            "Slides" if meta.page_count.is_none() => {
                                meta.page_count = text.trim().parse::<u32>().ok();
                            }
                            _ => {}
                        }
                    }
                    current_element = None;
                    current_text.clear();
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => {
                    return Err(Error::xml_parse_with_context(
                        e.to_string(),
                        "docProps/app.xml",
                    ))
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(())
    }

    /// Resolve a relative path from a base path.
    pub fn resolve_path(base: &str, relative: &str) -> String {
        if let Some(stripped) = relative.strip_prefix('/') {
            return stripped.to_string();
        }

        let base_path = Path::new(base);
        let base_dir = base_path.parent().unwrap_or(Path::new(""));

        let mut result = base_dir.to_path_buf();
        for component in Path::new(relative).components() {
            match component {
                std::path::Component::ParentDir => {
                    result.pop();
                }
                std::path::Component::Normal(c) => {
                    result.push(c);
                }
                _ => {}
            }
        }

        result.to_string_lossy().replace('\\', "/")
    }
}

fn metadata_text_or_raw(text: &quick_xml::events::BytesText<'_>, location: &str) -> Result<String> {
    crate::decode::decode_text_strict(text, location)
}

fn parse_relationship_element(
    e: &quick_xml::events::BytesStart<'_>,
    location: &str,
) -> Result<Relationship> {
    let mut id = String::new();
    let mut rel_type = String::new();
    let mut target = String::new();
    let mut external = false;

    for attr in e.attributes().flatten() {
        match attr.key.as_ref() {
            b"Id" => id = String::from_utf8_lossy(attr.value.as_ref()).to_string(),
            b"Type" => rel_type = String::from_utf8_lossy(attr.value.as_ref()).to_string(),
            b"Target" => target = String::from_utf8_lossy(attr.value.as_ref()).to_string(),
            b"TargetMode" => {
                external =
                    String::from_utf8_lossy(attr.value.as_ref()).eq_ignore_ascii_case("external")
            }
            _ => {}
        }
    }

    let mut missing = Vec::new();
    if id.is_empty() {
        missing.push("Id");
    }
    if rel_type.is_empty() {
        missing.push("Type");
    }
    if target.is_empty() {
        missing.push("Target");
    }

    if !missing.is_empty() {
        return Err(Error::xml_parse_with_context(
            format!(
                "Relationship element missing required attribute(s): {}",
                missing.join(", ")
            ),
            location,
        ));
    }

    Ok(Relationship {
        id,
        rel_type,
        target,
        external,
    })
}

pub(crate) fn parse_relationships_xml(content: &str, location: &str) -> Result<Relationships> {
    if content.trim().is_empty() {
        return Err(Error::xml_parse_with_context(
            "relationship file is empty",
            location,
        ));
    }

    let mut rels = Relationships::new();
    let mut reader = crate::decode::reader_for(content);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Empty(e)) | Ok(quick_xml::events::Event::Start(e))
                if e.name().as_ref() == b"Relationship" =>
            {
                rels.add(parse_relationship_element(&e, location)?);
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(e) => return Err(Error::xml_parse_with_context(e.to_string(), location)),
            _ => {}
        }
        buf.clear();
    }

    Ok(rels)
}

impl std::fmt::Debug for OoxmlContainer {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("OoxmlContainer")
            .field("files", &self.list_files().len())
            .finish()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_resolve_path() {
        assert_eq!(
            OoxmlContainer::resolve_path("word/document.xml", "../media/image1.png"),
            "media/image1.png"
        );
        assert_eq!(
            OoxmlContainer::resolve_path("word/document.xml", "styles.xml"),
            "word/styles.xml"
        );
        assert_eq!(
            OoxmlContainer::resolve_path("xl/worksheets/sheet1.xml", "../sharedStrings.xml"),
            "xl/sharedStrings.xml"
        );
        assert_eq!(
            OoxmlContainer::resolve_path("ppt/slides/slide1.xml", "/ppt/media/image1.png"),
            "ppt/media/image1.png"
        );
    }

    #[test]
    fn test_relationships_collection() {
        let mut rels = Relationships::new();
        rels.add(Relationship {
            id: "rId1".to_string(),
            rel_type: "http://test/type1".to_string(),
            target: "target1.xml".to_string(),
            external: false,
        });
        rels.add(Relationship {
            id: "rId2".to_string(),
            rel_type: "http://test/type1".to_string(),
            target: "target2.xml".to_string(),
            external: false,
        });

        assert!(rels.get("rId1").is_some());
        assert!(rels.get("rId3").is_none());
        assert_eq!(rels.get_by_type("http://test/type1").len(), 2);
    }

    #[test]
    fn test_relationship_projection_helpers_preserve_targets_and_types() {
        let mut rels = Relationships::new();
        rels.add(Relationship {
            id: "rId1".to_string(),
            rel_type: "urn:test:type".to_string(),
            target: "../media/image1.png".to_string(),
            external: false,
        });

        let target_only = rels.clone().into_targets_by_id();
        assert_eq!(
            target_only.get("rId1"),
            Some(&"../media/image1.png".to_string())
        );

        let typed = rels.into_type_targets_by_id();
        assert_eq!(
            typed.get("rId1"),
            Some(&(
                "urn:test:type".to_string(),
                "../media/image1.png".to_string()
            ))
        );
    }

    fn create_container_with_files(files: &[(&str, &str)]) -> OoxmlContainer {
        use std::io::{Cursor, Write};

        let buf = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(buf);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);

        for (path, contents) in files {
            zip.start_file(*path, options).unwrap();
            zip.write_all(contents.as_bytes()).unwrap();
        }

        OoxmlContainer::from_bytes(zip.finish().unwrap().into_inner()).unwrap()
    }

    fn create_container_with_binary_files(files: &[(&str, &[u8])]) -> OoxmlContainer {
        use std::io::{Cursor, Write};

        let buf = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(buf);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);

        for (path, contents) in files {
            zip.start_file(*path, options).unwrap();
            zip.write_all(contents).unwrap();
        }

        OoxmlContainer::from_bytes(zip.finish().unwrap().into_inner()).unwrap()
    }

    /// Build a ZIP whose entry is flagged as password-protected.
    ///
    /// Bit 0 of the general-purpose bit flag marks an entry encrypted, and it lives in
    /// both the local header and the central directory. Flipping it is enough for the
    /// ZIP layer to refuse the entry as needing a password — which is the condition we
    /// need to reproduce, without pulling in a crypto feature just to write one.
    fn create_password_protected_zip_bytes(path: &str) -> Vec<u8> {
        use std::io::{Cursor, Write};

        let buf = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(buf);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        zip.start_file(path, options).unwrap();
        zip.write_all(b"<Types/>").unwrap();
        let mut bytes = zip.finish().unwrap().into_inner();

        // Local file header: signature at 0, flag at offset 6.
        bytes[6] |= 1;
        // Central directory header: flag at offset 8 from its signature.
        let cd = bytes
            .windows(4)
            .position(|w| w == [0x50, 0x4B, 0x01, 0x02])
            .expect("central directory header");
        bytes[cd + 8] |= 1;

        bytes
    }

    /// A container that needs a password must not be reported as one that is merely
    /// missing the part we asked for — the two call for completely different responses
    /// from whoever is handling the failure.
    #[test]
    fn test_password_protected_entry_is_reported_as_encrypted() {
        let bytes = create_password_protected_zip_bytes("word/document.xml");
        let container = OoxmlContainer::from_bytes(bytes).unwrap();

        let err = container.read_xml("word/document.xml").unwrap_err();

        assert_eq!(err.kind(), crate::ErrorKind::Encrypted, "got: {err}");
    }

    /// The other half of the same contract: a genuinely absent part stays a missing
    /// component, so `read_xml_optional` keeps treating it as "not there".
    #[test]
    fn test_absent_entry_is_still_a_missing_component() {
        let container = create_container_with_files(&[("word/document.xml", "<w:document/>")]);

        let err = container.read_xml("word/styles.xml").unwrap_err();
        assert_eq!(err.kind(), crate::ErrorKind::MissingComponent);
        assert!(matches!(
            container.read_xml_optional("word/styles.xml"),
            Ok(None)
        ));
    }

    /// And the promise `read_xml_optional` makes in its own doc comment: a part that
    /// exists but cannot be read is not silently downgraded to "absent".
    #[test]
    fn test_optional_part_does_not_swallow_an_unreadable_entry() {
        let bytes = create_password_protected_zip_bytes("word/styles.xml");
        let container = OoxmlContainer::from_bytes(bytes).unwrap();

        let result = container.read_xml_optional("word/styles.xml");

        assert!(
            matches!(&result, Err(e) if e.kind() == crate::ErrorKind::Encrypted),
            "expected the failure to surface, got: {result:?}"
        );
    }

    #[test]
    fn test_parse_relationships_preserves_target_mode_and_start_forms() {
        let rels = parse_relationships_xml(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="urn:test:inline" Target="item1.xml"></Relationship>
  <Relationship Id="rId2" Type="urn:test:external" Target="https://example.com" TargetMode="External"/>
</Relationships>"#,
            "test.rels",
        )
        .unwrap();

        assert_eq!(rels.by_id.len(), 2);
        assert_eq!(rels.by_id["rId1"].target, "item1.xml");
        assert!(!rels.by_id["rId1"].external);
        assert_eq!(rels.by_id["rId2"].target, "https://example.com");
        assert!(rels.by_id["rId2"].external);
    }

    #[test]
    fn test_optional_relationships_allow_missing_but_not_malformed_content() {
        let missing = create_container_with_files(&[]);
        let rels = missing
            .read_optional_relationships_for_part("ppt/slides/slide1.xml")
            .unwrap();
        assert!(rels.by_id.is_empty());

        let malformed =
            create_container_with_files(&[("ppt/slides/_rels/slide1.xml.rels", "<Relationships")]);
        let err = malformed
            .read_optional_relationships_for_part("ppt/slides/slide1.xml")
            .unwrap_err();

        match err {
            Error::XmlParseWithContext { location, .. } => {
                assert_eq!(location, "ppt/slides/_rels/slide1.xml.rels")
            }
            other => panic!("expected malformed optional rels error, got {other:?}"),
        }
    }

    #[test]
    fn test_required_relationships_error_when_missing() {
        let container = create_container_with_files(&[]);
        let err = container
            .read_required_relationships_for_part("word/document.xml")
            .unwrap_err();

        match err {
            Error::MissingComponent(path) => assert_eq!(path, "word/_rels/document.xml.rels"),
            other => panic!("expected missing required rels error, got {other:?}"),
        }
    }

    #[test]
    fn test_open_docx() {
        let path = "test-files/file-sample_1MB.docx";
        if std::path::Path::new(path).exists() {
            let container = OoxmlContainer::open(path).unwrap();
            assert!(container.exists("[Content_Types].xml"));
            assert!(container.exists("word/document.xml"));

            let files = container.list_files();
            assert!(!files.is_empty());

            // Test relationships parsing
            let rels = container.read_package_relationships().unwrap();
            assert!(!rels.by_id.is_empty());
        }
    }

    #[test]
    fn test_open_xlsx() {
        let path = "test-files/file_example_XLSX_5000.xlsx";
        if std::path::Path::new(path).exists() {
            let container = OoxmlContainer::open(path).unwrap();
            assert!(container.exists("[Content_Types].xml"));
            assert!(container.exists("xl/workbook.xml"));

            let xl_files = container.list_files_with_prefix("xl/");
            assert!(!xl_files.is_empty());
        }
    }

    #[test]
    fn test_utf16_xml_reading() {
        let path = "test-files/officedissector/test/unit_test/testdocs/testutf16.docx";
        if std::path::Path::new(path).exists() {
            let container = OoxmlContainer::open(path).unwrap();

            // Read Content_Types.xml (UTF-16 encoded)
            let content = container
                .read_xml("[Content_Types].xml")
                .expect("Should read UTF-16 XML");
            assert!(
                content.contains("ContentType"),
                "Content should contain ContentType"
            );
            // Verify UTF-16 was decoded to UTF-8 (no null bytes in ASCII range)
            assert!(
                !content.starts_with("\0"),
                "Should not start with null byte"
            );
            assert!(
                content.starts_with("<?xml"),
                "Should start with XML declaration"
            );

            // Read document.xml (UTF-16 encoded)
            let doc_xml = container
                .read_xml("word/document.xml")
                .expect("Should read UTF-16 document.xml");
            assert!(
                doc_xml.contains("w:document"),
                "Should contain w:document element"
            );
            // Verify content is readable
            assert!(
                doc_xml.contains("Footnote in section"),
                "Should contain document text"
            );
        }
    }

    #[test]
    fn test_utf16_decoding_function() {
        // Test UTF-16 LE with BOM
        let utf16_le = b"\xFF\xFE<\0?\0x\0m\0l\0>\0";
        let result = decode_xml_bytes(utf16_le).expect("Should decode UTF-16 LE");
        assert_eq!(result, "<?xml>");

        // Test UTF-16 BE with BOM
        let utf16_be = b"\xFE\xFF\0<\0?\0x\0m\0l\0>";
        let result = decode_xml_bytes(utf16_be).expect("Should decode UTF-16 BE");
        assert_eq!(result, "<?xml>");

        // Test UTF-8 BOM
        let utf8_bom = b"\xEF\xBB\xBF<?xml>";
        let result = decode_xml_bytes(utf8_bom).expect("Should decode UTF-8 with BOM");
        assert_eq!(result, "<?xml>");

        // Test UTF-8 without BOM
        let utf8_plain = b"<?xml>";
        let result = decode_xml_bytes(utf8_plain).expect("Should decode UTF-8 without BOM");
        assert_eq!(result, "<?xml>");
    }

    #[test]
    fn test_decode_xml_bytes_rejects_invalid_non_utf8_input() {
        let cp1252_xml = b"<?xml version=\"1.0\"?><root>Caf\xe9 & \x93quotes\x94</root>";
        let err = decode_xml_bytes(cp1252_xml).unwrap_err();

        assert!(matches!(err, Error::Encoding(_)));
    }

    /// BOM-less UTF-16 of ASCII text is a *valid* UTF-8 byte sequence — every second byte
    /// is NUL, which UTF-8 accepts as U+0000. Decoding it as UTF-8 therefore succeeds and
    /// returns element names no reader can match, so the part parses to nothing without
    /// an error. Detection has to come before the UTF-8 attempt, not after it fails.
    #[test]
    fn test_decode_xml_bytes_detects_bom_less_utf16_of_pure_ascii() {
        let ascii = "<?xml version=\"1.0\"?><root>plain</root>";

        let le: Vec<u8> = ascii.encode_utf16().flat_map(u16::to_le_bytes).collect();
        assert_eq!(
            decode_xml_bytes(&le).expect("BOM-less UTF-16 LE must decode"),
            ascii
        );

        let be: Vec<u8> = ascii.encode_utf16().flat_map(u16::to_be_bytes).collect();
        assert_eq!(
            decode_xml_bytes(&be).expect("BOM-less UTF-16 BE must decode"),
            ascii
        );
    }

    /// An encoding failure is an encoding failure at every entry point. Reporting some of
    /// them as I/O sends consumers a discriminant that says the disk was at fault.
    #[test]
    fn test_encoding_failures_are_classified_as_encoding_everywhere() {
        // A lone high surrogate: well-formed UTF-16 code units, not a valid scalar.
        let mut lone_surrogate = vec![0xFF, 0xFE];
        lone_surrogate.extend_from_slice(&0xD800u16.to_le_bytes());
        lone_surrogate.extend_from_slice(&0x0041u16.to_le_bytes());
        assert_eq!(
            decode_xml_bytes(&lone_surrogate).unwrap_err().kind(),
            crate::error::ErrorKind::Encoding
        );

        let mut lone_surrogate_be = vec![0xFE, 0xFF];
        lone_surrogate_be.extend_from_slice(&0xD800u16.to_be_bytes());
        lone_surrogate_be.extend_from_slice(&0x0041u16.to_be_bytes());
        assert_eq!(
            decode_xml_bytes(&lone_surrogate_be).unwrap_err().kind(),
            crate::error::ErrorKind::Encoding
        );

        // UTF-8 BOM followed by bytes that are not UTF-8.
        let bad_after_bom = b"\xEF\xBB\xBF<root>Caf\xe9</root>";
        assert_eq!(
            decode_xml_bytes(bad_after_bom).unwrap_err().kind(),
            crate::error::ErrorKind::Encoding
        );
    }

    /// After decoding, the string is UTF-8 whatever the declaration claimed — including
    /// spellings a fixed list of literals would miss.
    #[test]
    fn test_decoded_part_declares_the_encoding_it_actually_has() {
        for declared in ["UTF-16", "utf-16", "UTF-16LE", "utf-16be", "ISO-8859-1"] {
            let xml = format!("<?xml version=\"1.0\" encoding=\"{declared}\"?><root/>");
            let bytes: Vec<u8> = std::iter::once(0xFEFFu16)
                .chain(xml.encode_utf16())
                .flat_map(u16::to_le_bytes)
                .collect();

            let decoded = decode_xml_bytes(&bytes).expect("must decode");
            assert_eq!(
                decoded, "<?xml version=\"1.0\" encoding=\"UTF-8\"?><root/>",
                "declared {declared} was left in place"
            );
        }
    }

    /// A declaration with no encoding attribute, and one with none at all, must come back
    /// untouched rather than rewritten or dropped.
    #[test]
    fn test_declaration_without_an_encoding_is_left_alone() {
        for xml in ["<?xml version=\"1.0\"?><root/>", "<root/>"] {
            let bytes: Vec<u8> = std::iter::once(0xFEFFu16)
                .chain(xml.encode_utf16())
                .flat_map(u16::to_le_bytes)
                .collect();
            assert_eq!(decode_xml_bytes(&bytes).expect("must decode"), xml);
        }
    }

    #[test]
    fn test_parse_core_metadata_preserves_whitespace_and_raw_malformed_text() {
        let container = create_container_with_files(&[(
            "docProps/core.xml",
            r#"<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>  Spaced &bogus; Title  </dc:title>
</cp:coreProperties>"#,
        )]);

        let meta = container.parse_core_metadata().unwrap();
        assert_eq!(meta.title.as_deref(), Some("  Spaced &bogus; Title  "));
    }

    #[test]
    fn test_parse_app_metadata_preserves_whitespace_and_raw_malformed_text() {
        let container = create_container_with_files(&[(
            "docProps/app.xml",
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>  App &bogus; Name  </Application>
</Properties>"#,
        )]);

        let meta = container.parse_core_metadata().unwrap();
        assert_eq!(meta.application.as_deref(), Some("  App &bogus; Name  "));
    }

    #[test]
    fn test_parse_core_metadata_propagates_invalid_xml_bytes() {
        let container = create_container_with_binary_files(&[(
            "docProps/core.xml",
            b"<?xml version=\"1.0\"?><cp:coreProperties xmlns:cp=\"urn:test\"><cp:title>Caf\xe9</cp:title></cp:coreProperties>",
        )]);

        let err = container.parse_core_metadata().unwrap_err();
        assert!(matches!(err, Error::Encoding(_)));
    }

    #[test]
    fn test_utf16_full_parse() {
        let path = "test-files/officedissector/test/unit_test/testdocs/testutf16.docx";
        if std::path::Path::new(path).exists() {
            // First test reading individual files
            let container = OoxmlContainer::open(path).unwrap();

            // Test reading various XML files
            for file_path in [
                "word/styles.xml",
                "word/numbering.xml",
                "word/document.xml",
                "docProps/core.xml",
                "word/footnotes.xml",
                "word/endnotes.xml",
            ] {
                match container.read_xml(file_path) {
                    Ok(content) => {
                        println!(
                            "{}: {} bytes, empty={}",
                            file_path,
                            content.len(),
                            content.trim().is_empty()
                        );
                        // Print first 100 chars to verify encoding
                        if !content.is_empty() {
                            let preview = &content[..content.len().min(100)];
                            println!("  Preview: {}", preview.replace('\n', "\\n"));
                        }
                    }
                    Err(e) => {
                        println!("{}: ERROR - {:?}", file_path, e);
                    }
                }
            }

            // Read raw bytes first
            println!("\n=== Testing raw styles.xml ===");
            match container.read_binary("word/styles.xml") {
                Ok(data) => {
                    println!("Raw bytes: {} bytes", data.len());
                    println!("First 10 bytes: {:02x?}", &data[..10.min(data.len())]);
                    println!(
                        "Last 10 bytes: {:02x?}",
                        &data[data.len().saturating_sub(10)..]
                    );

                    // Try decode manually
                    let decoded = decode_xml_bytes(&data).expect("decode failed");
                    println!("Decoded: {} chars", decoded.len());
                    println!(
                        "Decoded first 100: {:?}",
                        &decoded[..100.min(decoded.len())]
                    );
                    println!(
                        "Decoded last 100: {:?}",
                        &decoded[decoded.len().saturating_sub(100)..]
                    );
                    let null_count = decoded.bytes().filter(|&b| b == 0).count();
                    println!("Null bytes after decode: {}", null_count);
                }
                Err(e) => println!("read_binary ERROR: {:?}", e),
            }

            // Read styles.xml once and analyze
            println!("\n=== Testing StyleMap ===");
            match container.read_xml("word/styles.xml") {
                Ok(xml) => {
                    println!("Read styles.xml: {} bytes", xml.len());

                    // Print first and last characters
                    let first_100 = &xml[..xml.len().min(100)];
                    let last_100 = if xml.len() > 100 {
                        &xml[xml.len() - 100..]
                    } else {
                        &xml
                    };
                    println!("First 100: {:?}", first_100);
                    println!("Last 100: {:?}", last_100);

                    // Check for null bytes
                    let null_count = xml.bytes().filter(|&b| b == 0).count();
                    println!("Null bytes in string: {}", null_count);

                    // Try parsing
                    match crate::docx::styles::StyleMap::parse(&xml) {
                        Ok(styles) => println!("Styles OK: {} styles", styles.styles.len()),
                        Err(e) => println!("Styles ERROR: {:?}", e),
                    }
                }
                Err(e) => {
                    println!("read_xml ERROR: {:?}", e);
                }
            }

            // Test step by step: DOCX parser init
            println!("\n=== Testing DocxParser ===");
            match crate::docx::DocxParser::open(path) {
                Ok(mut parser) => {
                    println!("DocxParser init OK");
                    match parser.parse() {
                        Ok(doc) => {
                            println!("Parse OK: {} sections", doc.sections.len());
                            println!(
                                "Text: {}",
                                &doc.plain_text()[..doc.plain_text().len().min(200)]
                            );
                        }
                        Err(e) => {
                            println!("Parse ERROR: {:?}", e);
                        }
                    }
                }
                Err(e) => {
                    println!("DocxParser init ERROR: {:?}", e);
                }
            }
        }
    }

    #[test]
    fn test_open_pptx() {
        let path = "test-files/file_example_PPT_1MB.pptx";
        if std::path::Path::new(path).exists() {
            let container = OoxmlContainer::open(path).unwrap();
            assert!(container.exists("[Content_Types].xml"));
            assert!(container.exists("ppt/presentation.xml"));

            let slides = container.list_files_with_prefix("ppt/slides/");
            assert!(!slides.is_empty());
        }
    }

    /// Helper: create a minimal OOXML ZIP archive with given entries.
    fn create_test_zip(entries: &[(&str, &str)]) -> Vec<u8> {
        use std::io::Write;
        let buf = Vec::new();
        let cursor = Cursor::new(buf);
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        for (name, content) in entries {
            zip.start_file(*name, options).unwrap();
            zip.write_all(content.as_bytes()).unwrap();
        }
        let cursor = zip.finish().unwrap();
        cursor.into_inner()
    }

    #[test]
    fn test_parse_core_metadata_last_modified_by() {
        let core_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/">
  <dc:title>Test Title</dc:title>
  <dc:creator>Author Name</dc:creator>
  <cp:lastModifiedBy>Editor Name</cp:lastModifiedBy>
  <dcterms:created>2024-01-01T00:00:00Z</dcterms:created>
</cp:coreProperties>"#;

        let data = create_test_zip(&[
            ("[Content_Types].xml", "<Types/>"),
            ("docProps/core.xml", core_xml),
        ]);

        let container = OoxmlContainer::from_bytes(data).unwrap();
        let meta = container.parse_core_metadata().unwrap();

        assert_eq!(meta.title.as_deref(), Some("Test Title"));
        assert_eq!(meta.author.as_deref(), Some("Author Name"));
        assert_eq!(meta.last_modified_by.as_deref(), Some("Editor Name"));
        assert_eq!(meta.created.as_deref(), Some("2024-01-01T00:00:00Z"));
    }

    #[test]
    fn test_parse_app_metadata_basic() {
        let app_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Microsoft Office Word</Application>
  <Pages>5</Pages>
  <Words>1234</Words>
</Properties>"#;

        let data = create_test_zip(&[
            ("[Content_Types].xml", "<Types/>"),
            ("docProps/app.xml", app_xml),
        ]);

        let container = OoxmlContainer::from_bytes(data).unwrap();
        let meta = container.parse_core_metadata().unwrap();

        assert_eq!(meta.application.as_deref(), Some("Microsoft Office Word"));
        assert_eq!(meta.page_count, Some(5));
        assert_eq!(meta.word_count, Some(1234));
    }

    #[test]
    fn test_parse_app_metadata_slides() {
        let app_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Microsoft Office PowerPoint</Application>
  <Slides>10</Slides>
</Properties>"#;

        let data = create_test_zip(&[
            ("[Content_Types].xml", "<Types/>"),
            ("docProps/app.xml", app_xml),
        ]);

        let container = OoxmlContainer::from_bytes(data).unwrap();
        let meta = container.parse_core_metadata().unwrap();

        assert_eq!(
            meta.application.as_deref(),
            Some("Microsoft Office PowerPoint")
        );
        assert_eq!(meta.page_count, Some(10));
        assert_eq!(meta.word_count, None);
    }

    #[test]
    fn test_parse_app_metadata_pages_does_not_override_slides() {
        // When both Pages and Slides are present, the first one encountered wins
        // (Pages comes before Slides in the XML), which is correct since
        // DOCX has Pages, PPTX has Slides — they don't coexist in practice.
        let app_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Test</Application>
  <Pages>3</Pages>
  <Slides>10</Slides>
</Properties>"#;

        let data = create_test_zip(&[
            ("[Content_Types].xml", "<Types/>"),
            ("docProps/app.xml", app_xml),
        ]);

        let container = OoxmlContainer::from_bytes(data).unwrap();
        let meta = container.parse_core_metadata().unwrap();

        // Pages is parsed first, so page_count = 3; Slides won't override it
        assert_eq!(meta.page_count, Some(3));
    }

    #[test]
    fn test_parse_app_metadata_missing_file() {
        // No app.xml — should not fail, just leave fields as None
        let data = create_test_zip(&[("[Content_Types].xml", "<Types/>")]);

        let container = OoxmlContainer::from_bytes(data).unwrap();
        let meta = container.parse_core_metadata().unwrap();

        assert_eq!(meta.application, None);
        assert_eq!(meta.page_count, None);
        assert_eq!(meta.word_count, None);
    }

    #[test]
    fn test_parse_combined_core_and_app_metadata() {
        let core_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>My Document</dc:title>
  <dc:creator>Jane Doe</dc:creator>
  <cp:lastModifiedBy>John Smith</cp:lastModifiedBy>
</cp:coreProperties>"#;

        let app_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>LibreOffice</Application>
  <Pages>12</Pages>
  <Words>5000</Words>
</Properties>"#;

        let data = create_test_zip(&[
            ("[Content_Types].xml", "<Types/>"),
            ("docProps/core.xml", core_xml),
            ("docProps/app.xml", app_xml),
        ]);

        let container = OoxmlContainer::from_bytes(data).unwrap();
        let meta = container.parse_core_metadata().unwrap();

        assert_eq!(meta.title.as_deref(), Some("My Document"));
        assert_eq!(meta.author.as_deref(), Some("Jane Doe"));
        assert_eq!(meta.last_modified_by.as_deref(), Some("John Smith"));
        assert_eq!(meta.application.as_deref(), Some("LibreOffice"));
        assert_eq!(meta.page_count, Some(12));
        assert_eq!(meta.word_count, Some(5000));
    }

    #[test]
    fn test_docx_metadata_with_app() {
        let path = "test-files/file-sample_1MB.docx";
        if std::path::Path::new(path).exists() {
            let container = OoxmlContainer::open(path).unwrap();
            let meta = container.parse_core_metadata().unwrap();

            // Verify parse_core_metadata completes without error and
            // enriches with app.xml data when available.
            // The specific test file may not have all app.xml fields,
            // so just verify no errors occur.
            println!("Application: {:?}", meta.application);
            println!("Page count: {:?}", meta.page_count);
            println!("Word count: {:?}", meta.word_count);
            println!("Last modified by: {:?}", meta.last_modified_by);
        }
    }
}
