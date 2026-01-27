//! Format detection for Office Open XML documents.

use crate::container::decode_xml_bytes;
use crate::error::{Error, Result};
use std::fs::File;
use std::io::{BufReader, Read, Seek};
use std::path::Path;

/// ZIP file magic bytes: PK\x03\x04
const ZIP_MAGIC: [u8; 4] = [0x50, 0x4B, 0x03, 0x04];

/// Content type for DOCX main document part.
const DOCX_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

/// Content type for XLSX workbook part.
const XLSX_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";

/// Content type for PPTX presentation part.
const PPTX_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";

/// Detected Office document format.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormatType {
    /// Microsoft Word document (.docx)
    Docx,
    /// Microsoft Excel workbook (.xlsx)
    Xlsx,
    /// Microsoft PowerPoint presentation (.pptx)
    Pptx,
}

impl FormatType {
    /// Returns the file extension for this format.
    pub fn extension(&self) -> &'static str {
        match self {
            FormatType::Docx => "docx",
            FormatType::Xlsx => "xlsx",
            FormatType::Pptx => "pptx",
        }
    }

    /// Returns a human-readable name for this format.
    pub fn name(&self) -> &'static str {
        match self {
            FormatType::Docx => "Word Document",
            FormatType::Xlsx => "Excel Workbook",
            FormatType::Pptx => "PowerPoint Presentation",
        }
    }
}

impl std::fmt::Display for FormatType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", self.name())
    }
}

/// Detect the format type from a file path.
///
/// This function reads the file, verifies it's a valid ZIP archive,
/// and inspects the `[Content_Types].xml` to determine the specific format.
///
/// # Example
///
/// ```no_run
/// use undoc::detect::detect_format_from_path;
///
/// let format = detect_format_from_path("document.docx")?;
/// println!("Detected format: {}", format);
/// # Ok::<(), undoc::Error>(())
/// ```
pub fn detect_format_from_path(path: impl AsRef<Path>) -> Result<FormatType> {
    let file = File::open(path.as_ref())?;
    let reader = BufReader::new(file);
    detect_format_from_reader(reader)
}

/// Detect the format type from a byte slice.
///
/// # Example
///
/// ```no_run
/// use undoc::detect::detect_format_from_bytes;
///
/// let data = std::fs::read("document.docx")?;
/// let format = detect_format_from_bytes(&data)?;
/// # Ok::<(), undoc::Error>(())
/// ```
pub fn detect_format_from_bytes(data: &[u8]) -> Result<FormatType> {
    // Check magic bytes first
    if data.len() < 4 || data[..4] != ZIP_MAGIC {
        return Err(Error::UnknownFormat);
    }

    let cursor = std::io::Cursor::new(data);
    detect_format_from_reader(cursor)
}

/// Detect the format type from a reader.
pub fn detect_format_from_reader<R: Read + Seek>(reader: R) -> Result<FormatType> {
    let mut archive = zip::ZipArchive::new(reader)?;

    // Try to read [Content_Types].xml
    let content_types = match archive.by_name("[Content_Types].xml") {
        Ok(mut file) => {
            let mut bytes = Vec::new();
            file.read_to_end(&mut bytes)?;
            decode_xml_bytes(&bytes)?
        }
        Err(_) => {
            return Err(Error::MissingComponent("[Content_Types].xml".to_string()));
        }
    };

    // Check content types to determine format
    if content_types.contains(DOCX_CONTENT_TYPE) {
        Ok(FormatType::Docx)
    } else if content_types.contains(XLSX_CONTENT_TYPE) {
        Ok(FormatType::Xlsx)
    } else if content_types.contains(PPTX_CONTENT_TYPE) {
        Ok(FormatType::Pptx)
    } else {
        // Fallback: check for format-specific folders
        detect_by_folder_structure(&mut archive)
    }
}

/// Fallback detection by checking folder structure.
fn detect_by_folder_structure<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
) -> Result<FormatType> {
    let names: Vec<String> = archive.file_names().map(String::from).collect();

    // Check for format-specific paths
    let has_word = names.iter().any(|n| n.starts_with("word/"));
    let has_xl = names.iter().any(|n| n.starts_with("xl/"));
    let has_ppt = names.iter().any(|n| n.starts_with("ppt/"));

    match (has_word, has_xl, has_ppt) {
        (true, false, false) => Ok(FormatType::Docx),
        (false, true, false) => Ok(FormatType::Xlsx),
        (false, false, true) => Ok(FormatType::Pptx),
        _ => Err(Error::UnknownFormat),
    }
}

/// Check if data starts with ZIP magic bytes.
pub fn is_zip_file(data: &[u8]) -> bool {
    data.len() >= 4 && data[..4] == ZIP_MAGIC
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_format_type_display() {
        assert_eq!(FormatType::Docx.to_string(), "Word Document");
        assert_eq!(FormatType::Xlsx.to_string(), "Excel Workbook");
        assert_eq!(FormatType::Pptx.to_string(), "PowerPoint Presentation");
    }

    #[test]
    fn test_format_type_extension() {
        assert_eq!(FormatType::Docx.extension(), "docx");
        assert_eq!(FormatType::Xlsx.extension(), "xlsx");
        assert_eq!(FormatType::Pptx.extension(), "pptx");
    }

    #[test]
    fn test_is_zip_file() {
        assert!(is_zip_file(&[0x50, 0x4B, 0x03, 0x04, 0x00]));
        assert!(!is_zip_file(&[0x00, 0x00, 0x00, 0x00]));
        assert!(!is_zip_file(&[0x50, 0x4B])); // Too short
    }

    #[test]
    fn test_detect_invalid_data() {
        let result = detect_format_from_bytes(&[0x00, 0x00, 0x00, 0x00]);
        assert!(matches!(result, Err(Error::UnknownFormat)));
    }

    #[test]
    fn test_detect_docx_from_file() {
        let path = "test-files/file-sample_1MB.docx";
        if std::path::Path::new(path).exists() {
            let result = detect_format_from_path(path);
            assert!(result.is_ok());
            assert_eq!(result.unwrap(), FormatType::Docx);
        }
    }

    #[test]
    fn test_detect_xlsx_from_file() {
        let path = "test-files/file_example_XLSX_5000.xlsx";
        if std::path::Path::new(path).exists() {
            let result = detect_format_from_path(path);
            assert!(result.is_ok());
            assert_eq!(result.unwrap(), FormatType::Xlsx);
        }
    }

    #[test]
    fn test_detect_pptx_from_file() {
        let path = "test-files/file_example_PPT_1MB.pptx";
        if std::path::Path::new(path).exists() {
            let result = detect_format_from_path(path);
            assert!(result.is_ok());
            assert_eq!(result.unwrap(), FormatType::Pptx);
        }
    }
}
