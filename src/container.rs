//! ZIP container abstraction for OOXML documents.

use crate::error::{Error, Result};
use crate::model::Metadata;
use std::cell::RefCell;
use std::collections::HashMap;
use std::fs::File;
use std::io::{BufReader, Cursor, Read, Seek};
use std::path::Path;

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
}

/// Fix XML encoding declaration from UTF-16 to UTF-8.
///
/// When we decode UTF-16 XML to a Rust String (UTF-8), the XML declaration
/// still says encoding="UTF-16". This causes quick-xml to fail when it tries
/// to re-interpret the already-decoded UTF-8 string as UTF-16.
fn fix_xml_encoding_declaration(content: &str) -> String {
    // Replace encoding="UTF-16" with encoding="UTF-8" in XML declaration
    if content.starts_with("<?xml") {
        if let Some(end_decl) = content.find("?>") {
            let decl = &content[..end_decl + 2];
            let rest = &content[end_decl + 2..];

            // Replace UTF-16 with UTF-8 (case insensitive)
            let fixed_decl = decl
                .replace("encoding=\"UTF-16\"", "encoding=\"UTF-8\"")
                .replace("encoding='UTF-16'", "encoding='UTF-8'")
                .replace("encoding=\"utf-16\"", "encoding=\"UTF-8\"")
                .replace("encoding='utf-16'", "encoding='UTF-8'");

            return format!("{}{}", fixed_decl, rest);
        }
    }
    content.to_string()
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
pub fn decode_xml_bytes(bytes: &[u8]) -> Result<String> {
    // Check for BOM (Byte Order Mark)
    if bytes.len() >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF {
        // UTF-8 BOM: EF BB BF - skip BOM and decode as UTF-8
        return String::from_utf8(bytes[3..].to_vec())
            .map_err(|e| Error::Io(std::io::Error::new(std::io::ErrorKind::InvalidData, e)));
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

    // No BOM - try UTF-8 first, then attempt UTF-16 detection
    match String::from_utf8(bytes.to_vec()) {
        Ok(s) => Ok(s),
        Err(_) => {
            // Try to detect UTF-16 by checking for common XML patterns
            // UTF-16 LE typically has null bytes in odd positions for ASCII
            if bytes.len() >= 4 && bytes[1] == 0 && bytes[3] == 0 {
                decode_utf16_le(bytes)
            } else if bytes.len() >= 4 && bytes[0] == 0 && bytes[2] == 0 {
                decode_utf16_be(bytes)
            } else {
                // Fall back to lossy UTF-8 conversion
                Ok(String::from_utf8_lossy(bytes).into_owned())
            }
        }
    }
}

/// Decode UTF-16 Little Endian bytes to String.
fn decode_utf16_le(bytes: &[u8]) -> Result<String> {
    // Ensure even number of bytes
    let len = bytes.len() & !1;

    let u16_iter = (0..len)
        .step_by(2)
        .map(|i| u16::from_le_bytes([bytes[i], bytes[i + 1]]));

    char::decode_utf16(u16_iter)
        .collect::<std::result::Result<String, _>>()
        .map_err(|e| Error::Io(std::io::Error::new(std::io::ErrorKind::InvalidData, e)))
}

/// Decode UTF-16 Big Endian bytes to String.
fn decode_utf16_be(bytes: &[u8]) -> Result<String> {
    // Ensure even number of bytes
    let len = bytes.len() & !1;

    let u16_iter = (0..len)
        .step_by(2)
        .map(|i| u16::from_be_bytes([bytes[i], bytes[i + 1]]));

    char::decode_utf16(u16_iter)
        .collect::<std::result::Result<String, _>>()
        .map_err(|e| Error::Io(std::io::Error::new(std::io::ErrorKind::InvalidData, e)))
}

impl OoxmlContainer {
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
        let mut file = archive
            .by_name(path)
            .map_err(|_| Error::MissingComponent(path.to_string()))?;

        // Read raw bytes first
        let mut bytes = Vec::new();
        file.read_to_end(&mut bytes)?;

        // Detect encoding from BOM
        let content = decode_xml_bytes(&bytes)?;
        Ok(content)
    }

    /// Read a binary file from the archive.
    pub fn read_binary(&self, path: &str) -> Result<Vec<u8>> {
        let mut archive = self.archive.borrow_mut();
        let mut file = archive
            .by_name(path)
            .map_err(|_| Error::MissingComponent(path.to_string()))?;
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
    pub fn read_relationships(&self, part_path: &str) -> Result<Relationships> {
        // Build the rels path
        let rels_path = if part_path.is_empty() || part_path == "/" {
            "_rels/.rels".to_string()
        } else {
            let path = Path::new(part_path);
            let parent = path.parent().unwrap_or(Path::new(""));
            let filename = path.file_name().unwrap_or_default().to_string_lossy();
            format!("{}/_rels/{}.rels", parent.display(), filename)
        };

        self.parse_relationships(&rels_path)
    }

    /// Read package-level relationships (_rels/.rels).
    pub fn read_package_relationships(&self) -> Result<Relationships> {
        self.parse_relationships("_rels/.rels")
    }

    /// Parse core metadata from docProps/core.xml.
    ///
    /// This is common to all OOXML formats (DOCX, XLSX, PPTX).
    pub fn parse_core_metadata(&self) -> Result<Metadata> {
        let mut meta = Metadata::default();

        if let Ok(xml) = self.read_xml("docProps/core.xml") {
            let mut reader = quick_xml::Reader::from_str(&xml);
            reader.config_mut().trim_text(true);

            let mut buf = Vec::new();
            let mut current_element: Option<String> = None;

            loop {
                match reader.read_event_into(&mut buf) {
                    Ok(quick_xml::events::Event::Start(e)) => {
                        let name = e.name();
                        current_element =
                            Some(String::from_utf8_lossy(name.local_name().as_ref()).to_string());
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
                                        .split([',', ';'])
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

    /// Parse a relationships file.
    fn parse_relationships(&self, rels_path: &str) -> Result<Relationships> {
        let content = match self.read_xml(rels_path) {
            Ok(c) => c,
            Err(_) => return Ok(Relationships::new()),
        };

        // Handle empty content
        if content.trim().is_empty() {
            return Ok(Relationships::new());
        }

        let mut rels = Relationships::new();
        let mut reader = quick_xml::Reader::from_str(&content);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Empty(e)) if e.name().as_ref() == b"Relationship" => {
                    let mut id = String::new();
                    let mut rel_type = String::new();
                    let mut target = String::new();
                    let mut external = false;

                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"Id" => id = String::from_utf8_lossy(&attr.value).to_string(),
                            b"Type" => rel_type = String::from_utf8_lossy(&attr.value).to_string(),
                            b"Target" => target = String::from_utf8_lossy(&attr.value).to_string(),
                            b"TargetMode" => {
                                external = String::from_utf8_lossy(&attr.value).to_lowercase()
                                    == "external"
                            }
                            _ => {}
                        }
                    }

                    if !id.is_empty() {
                        rels.add(Relationship {
                            id,
                            rel_type,
                            target,
                            external,
                        });
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => return Err(Error::XmlParse(e.to_string())),
                _ => {}
            }
            buf.clear();
        }

        Ok(rels)
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
                        if content.len() > 0 {
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
}
