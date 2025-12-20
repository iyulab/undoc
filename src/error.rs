//! Error types for the undoc library.

use std::io;
use thiserror::Error;

/// Result type alias for undoc operations.
pub type Result<T> = std::result::Result<T, Error>;

/// Errors that can occur during document processing.
#[derive(Error, Debug)]
pub enum Error {
    /// I/O error during file operations.
    #[error("I/O error: {0}")]
    Io(#[from] io::Error),

    /// The file format could not be determined.
    #[error("Unknown file format")]
    UnknownFormat,

    /// The file format is recognized but not supported.
    #[error("Unsupported format: {0}")]
    UnsupportedFormat(String),

    /// Error reading ZIP archive.
    #[error("ZIP archive error: {0}")]
    ZipArchive(String),

    /// Error parsing XML content.
    #[error("XML parse error: {0}")]
    XmlParse(String),

    /// Invalid or malformed data in the document.
    #[error("Invalid data: {0}")]
    InvalidData(String),

    /// A required document component is missing.
    #[error("Missing component: {0}")]
    MissingComponent(String),

    /// Error during text encoding conversion.
    #[error("Encoding error: {0}")]
    Encoding(String),

    /// A referenced style was not found.
    #[error("Style not found: {0}")]
    StyleNotFound(String),

    /// A referenced resource was not found.
    #[error("Resource not found: {0}")]
    ResourceNotFound(String),

    /// The document is encrypted and cannot be processed.
    #[error("Document is encrypted")]
    Encrypted,

    /// Error during rendering.
    #[error("Render error: {0}")]
    Render(String),
}

impl From<zip::result::ZipError> for Error {
    fn from(err: zip::result::ZipError) -> Self {
        Error::ZipArchive(err.to_string())
    }
}

impl From<quick_xml::Error> for Error {
    fn from(err: quick_xml::Error) -> Self {
        Error::XmlParse(err.to_string())
    }
}

impl From<quick_xml::DeError> for Error {
    fn from(err: quick_xml::DeError) -> Self {
        Error::XmlParse(err.to_string())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_error_display() {
        let err = Error::UnknownFormat;
        assert_eq!(err.to_string(), "Unknown file format");

        let err = Error::UnsupportedFormat("legacy .doc".to_string());
        assert_eq!(err.to_string(), "Unsupported format: legacy .doc");
    }

    #[test]
    fn test_error_from_io() {
        let io_err = io::Error::new(io::ErrorKind::NotFound, "file not found");
        let err: Error = io_err.into();
        assert!(matches!(err, Error::Io(_)));
    }
}
