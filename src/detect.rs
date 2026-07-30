//! Format detection for Office Open XML documents.

use crate::container::decode_xml_bytes;
use crate::error::{Error, Result};
use serde::{Deserialize, Serialize};
#[cfg(not(target_arch = "wasm32"))]
use std::fs::File;
#[cfg(not(target_arch = "wasm32"))]
use std::io::BufReader;
use std::io::{Read, Seek};
#[cfg(not(target_arch = "wasm32"))]
use std::path::Path;

/// ZIP file magic bytes: PK\x03\x04
const ZIP_MAGIC: [u8; 4] = [0x50, 0x4B, 0x03, 0x04];

/// Compound File Binary (OLE2) magic bytes: D0 CF 11 E0 A1 B1 1A E1.
///
/// Two kinds of file arrive with this header: the legacy binary Office formats
/// (.doc/.xls/.ppt) and an OOXML document protected with ECMA-376 encryption, whose ZIP
/// package is wrapped inside a CFB container. The header alone cannot tell them apart —
/// [`classify_cfb_container`] walks the directory, because the two answers send a caller
/// in opposite directions: one looks for a converter, the other for a password.
const CFB_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

/// Stream that marks a CFB container as an ECMA-376 encrypted OOXML package.
///
/// Such a file also carries `EncryptionInfo` (the key derivation data), but the payload
/// stream is the one that says "there is a document in here, behind a password".
const ENCRYPTED_PACKAGE_STREAM: &str = "/EncryptedPackage";

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
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum FormatType {
    /// Microsoft Word document (.docx)
    #[default]
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
#[cfg(not(target_arch = "wasm32"))]
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
    detect_format_from_reader(std::io::Cursor::new(data))
}

/// Classify the container by its leading bytes.
///
/// Runs before the ZIP layer gets involved, so that a file we can *recognise* but not
/// open is reported as such instead of surfacing as a damaged archive — which would send
/// the caller off to repair a file that is not broken.
fn classify_container_magic<R: Read + Seek>(reader: &mut R) -> Result<()> {
    let mut head = [0u8; CFB_MAGIC.len()];
    let mut filled = 0;
    while filled < head.len() {
        match reader.read(&mut head[filled..])? {
            0 => break,
            n => filled += n,
        }
    }
    reader.seek(std::io::SeekFrom::Start(0))?;

    if filled == CFB_MAGIC.len() && head == CFB_MAGIC {
        return Err(classify_cfb_container(reader));
    }

    if filled < ZIP_MAGIC.len() || head[..ZIP_MAGIC.len()] != ZIP_MAGIC {
        return Err(Error::UnknownFormat);
    }

    Ok(())
}

/// Say which kind of CFB container this is, having established that it is one.
///
/// An ECMA-376 encrypted OOXML document is a CFB wrapper around the encrypted ZIP
/// package; a legacy binary Office file is a CFB holding its own well-known streams. The
/// header is identical, so only the directory distinguishes them — and the distinction is
/// the whole point: `Encrypted` tells the caller to supply a password, while
/// `UnsupportedFormat` tells them to convert the file. Reporting the disjunction leaves
/// them to guess.
///
/// Never returns `Ok`: neither kind can be opened by this library. The naming of the
/// legacy format is best-effort — an unrecognised CFB is still reported as unsupported,
/// which is what the header proved.
fn classify_cfb_container<R: Read + Seek>(reader: &mut R) -> Error {
    let container = match cfb::CompoundFile::open(reader) {
        Ok(container) => container,
        // A CFB header whose directory will not parse. Still an Office container, still
        // unopenable; guessing which kind would claim more than was established.
        Err(_) => {
            return Error::UnsupportedFormat(
                "OLE/CFB container whose directory could not be read — a legacy binary \
                 Office format (.doc/.xls/.ppt) or an ECMA-376 encrypted document"
                    .to_string(),
            )
        }
    };

    if container.exists(ENCRYPTED_PACKAGE_STREAM) {
        return Error::Encrypted;
    }

    // Well-known root streams of the pre-2007 binary formats. Checked only to make the
    // message specific; absence of all three does not make the file openable.
    let legacy = if container.exists("/WordDocument") {
        Some("Word 97-2003 (.doc)")
    } else if container.exists("/Workbook") || container.exists("/Book") {
        Some("Excel 97-2003 (.xls)")
    } else if container.exists("/PowerPoint Document") {
        Some("PowerPoint 97-2003 (.ppt)")
    } else {
        None
    };

    Error::UnsupportedFormat(match legacy {
        Some(format) => format!("legacy binary Office format: {format}"),
        None => "OLE/CFB container that is not a recognised Office document or encrypted \
             OOXML package"
            .to_string(),
    })
}

/// Detect the format type from a reader.
///
/// The leading bytes are inspected before the ZIP layer is involved, so the reader is
/// rewound to absolute position 0 rather than continued from wherever the caller left
/// it, and the container must *begin* with a recognised signature — a ZIP archive with
/// data prepended to it (a self-extracting stub, say) is reported as
/// [`Error::UnknownFormat`] instead of being recovered from its central directory.
pub fn detect_format_from_reader<R: Read + Seek>(reader: R) -> Result<FormatType> {
    let mut reader = reader;
    classify_container_magic(&mut reader)?;

    let mut archive = zip::ZipArchive::new(reader)?;

    // Try to read [Content_Types].xml
    let content_types = match archive.by_name("[Content_Types].xml") {
        Ok(mut file) => {
            let mut bytes = Vec::new();
            file.read_to_end(&mut bytes)?;
            decode_xml_bytes(&bytes)?
        }
        // Only an absent part is a missing component. A container that is damaged or
        // needs a password must say so — reporting either as "the part isn't there"
        // points the caller at the wrong problem.
        Err(zip::result::ZipError::FileNotFound) => {
            return Err(Error::MissingComponent("[Content_Types].xml".to_string()));
        }
        Err(e) => return Err(Error::from(e)),
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

    /// A legacy binary Office file — or an ECMA-376 encrypted one — is a container we
    /// recognise and cannot open. Reporting it as unrecognised sends the caller looking
    /// for a file-type problem they do not have, and letting it reach the ZIP layer
    /// reports it as damaged, which is worse: it invites a pointless repair attempt.
    #[test]
    fn test_ole_cfb_container_is_unsupported_not_unknown_or_damaged() {
        let mut data = CFB_MAGIC.to_vec();
        data.extend_from_slice(&[0u8; 64]);

        let err = detect_format_from_bytes(&data).unwrap_err();

        assert_eq!(
            err.kind(),
            crate::ErrorKind::UnsupportedFormat,
            "got: {err}"
        );
    }

    /// Build a CFB container holding the named root streams, and nothing else.
    fn cfb_with_streams(names: &[&str]) -> Vec<u8> {
        let mut container = cfb::CompoundFile::create(std::io::Cursor::new(Vec::new()))
            .expect("create CFB container");
        for name in names {
            container.create_stream(name).expect("create stream");
        }
        container.flush().expect("flush CFB container");
        container.into_inner().into_inner()
    }

    /// The distinction this branch exists for. Both files carry the CFB header, and the
    /// two answers send the caller in opposite directions — a password versus a
    /// converter — so reporting the disjunction leaves them to guess.
    #[test]
    fn test_ecma376_encrypted_package_is_encrypted_not_unsupported() {
        let data = cfb_with_streams(&["/EncryptedPackage", "/EncryptionInfo"]);

        // Same header as a legacy binary file, so the directory is what decides.
        assert_eq!(data[..CFB_MAGIC.len()], CFB_MAGIC);

        let err = detect_format_from_bytes(&data).unwrap_err();

        assert_eq!(err.kind(), crate::ErrorKind::Encrypted, "got: {err}");
        assert!(matches!(err, Error::Encrypted));
    }

    #[test]
    fn test_legacy_binary_office_names_the_format_it_found() {
        for (stream, expected) in [
            ("/WordDocument", "Word 97-2003"),
            ("/Workbook", "Excel 97-2003"),
            ("/PowerPoint Document", "PowerPoint 97-2003"),
        ] {
            let err = detect_format_from_bytes(&cfb_with_streams(&[stream])).unwrap_err();

            assert_eq!(
                err.kind(),
                crate::ErrorKind::UnsupportedFormat,
                "{stream} got: {err}"
            );
            let message = err.to_string();
            assert!(
                message.contains(expected),
                "{stream} should be named in: {message}"
            );
        }
    }

    /// A CFB that is neither encrypted OOXML nor a format we can name stays unsupported.
    /// Naming is best-effort; the header is what was proved.
    #[test]
    fn test_unrecognised_cfb_is_still_unsupported() {
        let err = detect_format_from_bytes(&cfb_with_streams(&["/SomethingElse"])).unwrap_err();

        assert_eq!(
            err.kind(),
            crate::ErrorKind::UnsupportedFormat,
            "got: {err}"
        );
    }

    /// Truncated input must not be mistaken for a CFB header on a prefix match.
    #[test]
    fn test_truncated_cfb_prefix_is_unknown_format() {
        let err = detect_format_from_bytes(&CFB_MAGIC[..4]).unwrap_err();

        assert!(matches!(err, Error::UnknownFormat));
    }

    /// The ZIP path must keep working through the new magic-byte gate, including the
    /// damaged-archive case that still has to read as damaged.
    #[test]
    fn test_zip_magic_still_reaches_the_archive_layer() {
        let mut data = ZIP_MAGIC.to_vec();
        data.extend_from_slice(b"not a real central directory");

        let err = detect_format_from_bytes(&data).unwrap_err();

        assert_eq!(err.kind(), crate::ErrorKind::ZipArchive, "got: {err}");
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

    #[test]
    fn test_format_type_default_is_docx() {
        assert_eq!(FormatType::default(), FormatType::Docx);
    }

    #[test]
    fn test_format_type_serde_roundtrip() {
        let json = serde_json::to_string(&FormatType::Xlsx).unwrap();
        assert_eq!(json, "\"xlsx\"");
        let back: FormatType = serde_json::from_str(&json).unwrap();
        assert_eq!(back, FormatType::Xlsx);
    }
}
