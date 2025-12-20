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
    pub fn read_xml(&self, path: &str) -> Result<String> {
        let mut archive = self.archive.borrow_mut();
        let mut file = archive
            .by_name(path)
            .map_err(|_| Error::MissingComponent(path.to_string()))?;
        let mut content = String::new();
        file.read_to_string(&mut content)?;
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

    /// Parse a relationships file.
    fn parse_relationships(&self, rels_path: &str) -> Result<Relationships> {
        let content = match self.read_xml(rels_path) {
            Ok(c) => c,
            Err(_) => return Ok(Relationships::new()),
        };

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
                                external =
                                    String::from_utf8_lossy(&attr.value).to_lowercase() == "external"
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
        if relative.starts_with('/') {
            return relative[1..].to_string();
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
