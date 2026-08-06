//! Error types for the undoc library.

use std::io;
use thiserror::Error;

/// Result type alias for undoc operations.
pub type Result<T> = std::result::Result<T, Error>;

/// Stable classification of an [`Error`], for consumers that must branch on the
/// *reason* a call failed rather than match on its message text.
///
/// Every [`Error`] variant maps onto one of these, so the mapping needs no judgement
/// and stays obvious as the error type grows. The discriminants are explicit and part
/// of the public contract: they cross the C-ABI boundary as `undoc_last_error_kind`
/// return values, so **existing values must never be renumbered** — a new failure
/// reason takes the next free number instead. Treat an unrecognised value as a
/// generic failure rather than as an error.
///
/// Values `100` and above are reserved for FFI-boundary reasons that have no core
/// `Error` counterpart (null arguments, caught panics, output that cannot cross the
/// ABI); see the `UNDOC_ERROR_*` constants in the `ffi` module.
///
/// This enum is `#[non_exhaustive]`: match it with a `_ =>` arm and treat an unfamiliar
/// reason as a generic failure. That is the same contract the C, C# and Python surfaces
/// document, and it is what lets a later release name a new reason without breaking
/// callers.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
#[repr(i32)]
pub enum ErrorKind {
    /// A failure with no more specific classification.
    ///
    /// [`Error::kind`] never returns this — no core variant maps to it. It exists so
    /// that bindings and consumers have a value for a failure that did not come from
    /// this library and therefore carries no classification. It must not be confused
    /// with success, which is the absence of an error (`0`).
    Other = 1,
    /// [`Error::Io`]
    Io = 2,
    /// [`Error::UnknownFormat`]
    UnknownFormat = 3,
    /// [`Error::UnsupportedFormat`]
    UnsupportedFormat = 4,
    /// [`Error::ZipArchive`]
    ZipArchive = 5,
    /// [`Error::XmlParse`] and [`Error::XmlParseWithContext`] — the same reason,
    /// differing only in whether location context was available.
    XmlParse = 6,
    /// [`Error::InvalidData`]
    InvalidData = 7,
    /// [`Error::MissingComponent`]
    MissingComponent = 8,
    /// [`Error::Encoding`]
    Encoding = 9,
    /// [`Error::StyleNotFound`]
    StyleNotFound = 10,
    /// [`Error::ResourceNotFound`]
    ResourceNotFound = 11,
    /// [`Error::Encrypted`]
    Encrypted = 12,
    /// [`Error::Render`]
    Render = 13,
}

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
    #[error("XML parse error in {location}: {message}")]
    XmlParseWithContext {
        /// The error message
        message: String,
        /// Location context (file path, element name, etc.)
        location: String,
    },

    /// Error parsing XML content (legacy, no context).
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
    /// Preserve the reason the container could not be read.
    ///
    /// The ZIP layer already distinguishes a damaged archive from an unsupported one,
    /// from a missing entry, from one that needs a password. Collapsing all of them
    /// into a single variant would throw that away here, at the one place it is still
    /// known — and no amount of work further up could recover it.
    fn from(err: zip::result::ZipError) -> Self {
        use zip::result::ZipError;

        match err {
            ZipError::Io(e) => Error::Io(e),
            ZipError::InvalidPassword => Error::Encrypted,
            // The ZIP spec has no dedicated "encrypted" status: a password-protected
            // entry surfaces as an unsupported archive with this specific message.
            ZipError::UnsupportedArchive(ZipError::PASSWORD_REQUIRED) => Error::Encrypted,
            ZipError::UnsupportedArchive(what) => Error::UnsupportedFormat(what.to_string()),
            ZipError::FileNotFound => {
                Error::MissingComponent("entry not present in archive".to_string())
            }
            ZipError::InvalidArchive(what) => Error::ZipArchive(what.to_string()),
            // `ZipError` is not `#[non_exhaustive]` today, but treat an unfamiliar
            // variant as a damaged container rather than failing to compile a
            // consumer's build on a dependency bump.
            other => Error::ZipArchive(other.to_string()),
        }
    }
}

impl From<quick_xml::Error> for Error {
    /// Preserve whether the XML failed to *arrive*, to *decode*, or to *parse*.
    ///
    /// These are three different problems for whoever has to act on the failure, and
    /// only this conversion still knows which one happened.
    fn from(err: quick_xml::Error) -> Self {
        match err {
            quick_xml::Error::Io(e) => Error::Io(io::Error::new(e.kind(), e.to_string())),
            quick_xml::Error::Encoding(e) => Error::Encoding(e.to_string()),
            other => Error::XmlParse(other.to_string()),
        }
    }
}

impl Error {
    /// Create an XML parse error with context information.
    ///
    /// # Arguments
    /// * `message` - The error message
    /// * `location` - Context such as file path or element being parsed
    ///
    /// # Example
    /// ```ignore
    /// Error::xml_parse_with_context("Invalid element", "word/document.xml")
    /// ```
    pub fn xml_parse_with_context(message: impl Into<String>, location: impl Into<String>) -> Self {
        Error::XmlParseWithContext {
            message: message.into(),
            location: location.into(),
        }
    }

    /// Classify this error into a stable [`ErrorKind`].
    ///
    /// Lets a caller branch on *why* an operation failed without matching on the
    /// message text. The message stays the human-readable diagnostic; the kind is
    /// additive.
    ///
    /// # Example
    /// ```
    /// use undoc::{Error, ErrorKind};
    ///
    /// let err = Error::UnknownFormat;
    /// assert_eq!(err.kind(), ErrorKind::UnknownFormat);
    /// ```
    pub fn kind(&self) -> ErrorKind {
        match self {
            Error::Io(_) => ErrorKind::Io,
            Error::UnknownFormat => ErrorKind::UnknownFormat,
            Error::UnsupportedFormat(_) => ErrorKind::UnsupportedFormat,
            Error::ZipArchive(_) => ErrorKind::ZipArchive,
            Error::XmlParseWithContext { .. } => ErrorKind::XmlParse,
            Error::XmlParse(_) => ErrorKind::XmlParse,
            Error::InvalidData(_) => ErrorKind::InvalidData,
            Error::MissingComponent(_) => ErrorKind::MissingComponent,
            Error::Encoding(_) => ErrorKind::Encoding,
            Error::StyleNotFound(_) => ErrorKind::StyleNotFound,
            Error::ResourceNotFound(_) => ErrorKind::ResourceNotFound,
            Error::Encrypted => ErrorKind::Encrypted,
            Error::Render(_) => ErrorKind::Render,
        }
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

    #[test]
    fn test_error_kind_mapping() {
        let io_err = io::Error::new(io::ErrorKind::NotFound, "file not found");
        assert_eq!(Error::from(io_err).kind(), ErrorKind::Io);
        assert_eq!(Error::UnknownFormat.kind(), ErrorKind::UnknownFormat);
        assert_eq!(
            Error::UnsupportedFormat("legacy .doc".into()).kind(),
            ErrorKind::UnsupportedFormat
        );
        assert_eq!(
            Error::ZipArchive("bad central directory".into()).kind(),
            ErrorKind::ZipArchive
        );
        assert_eq!(
            Error::InvalidData("bad".into()).kind(),
            ErrorKind::InvalidData
        );
        assert_eq!(
            Error::MissingComponent("[Content_Types].xml".into()).kind(),
            ErrorKind::MissingComponent
        );
        assert_eq!(
            Error::Encoding("bad utf-8".into()).kind(),
            ErrorKind::Encoding
        );
        assert_eq!(
            Error::StyleNotFound("Heading1".into()).kind(),
            ErrorKind::StyleNotFound
        );
        assert_eq!(
            Error::ResourceNotFound("rId1".into()).kind(),
            ErrorKind::ResourceNotFound
        );
        assert_eq!(Error::Encrypted.kind(), ErrorKind::Encrypted);
        assert_eq!(Error::Render("bad table".into()).kind(), ErrorKind::Render);
    }

    /// Both XML variants are the same failure reason, so they must share one number —
    /// otherwise the ABI carries two values for one reason.
    #[test]
    fn test_xml_parse_variants_share_one_kind() {
        assert_eq!(
            Error::XmlParse("unexpected eof".into()).kind(),
            ErrorKind::XmlParse
        );
        assert_eq!(
            Error::xml_parse_with_context("unexpected eof", "word/document.xml").kind(),
            ErrorKind::XmlParse
        );
    }

    /// The ZIP layer knows more than "something went wrong with the archive", and this
    /// is the only place that knowledge still exists. Constructing the `ZipError`
    /// values directly tests every branch without having to fabricate an archive that
    /// provokes each one — including the password cases, which are otherwise awkward to
    /// produce.
    #[test]
    fn test_zip_error_reasons_are_not_collapsed() {
        use zip::result::ZipError;

        assert_eq!(
            Error::from(ZipError::InvalidArchive("no central directory")).kind(),
            ErrorKind::ZipArchive,
            "a damaged container stays a damaged container"
        );
        assert_eq!(
            Error::from(ZipError::UnsupportedArchive("zip64 with disks")).kind(),
            ErrorKind::UnsupportedFormat,
            "unsupported is not the same problem as damaged"
        );
        assert_eq!(
            Error::from(ZipError::InvalidPassword).kind(),
            ErrorKind::Encrypted
        );
        assert_eq!(
            Error::from(ZipError::UnsupportedArchive(ZipError::PASSWORD_REQUIRED)).kind(),
            ErrorKind::Encrypted,
            "the ZIP spec reports a password-protected entry as unsupported"
        );
        assert_eq!(
            Error::from(ZipError::FileNotFound).kind(),
            ErrorKind::MissingComponent
        );
        assert_eq!(
            Error::from(ZipError::Io(io::Error::new(
                io::ErrorKind::UnexpectedEof,
                "eof"
            )))
            .kind(),
            ErrorKind::Io
        );
    }

    /// XML that fails to arrive, to decode, and to parse are three different problems.
    #[test]
    fn test_xml_error_reasons_are_not_collapsed() {
        let io_backed: Error = quick_xml::Error::Io(std::sync::Arc::new(io::Error::new(
            io::ErrorKind::UnexpectedEof,
            "eof mid-element",
        )))
        .into();
        assert_eq!(io_backed.kind(), ErrorKind::Io);

        let malformed: Error =
            quick_xml::Error::Syntax(quick_xml::errors::SyntaxError::UnclosedTag).into();
        assert_eq!(malformed.kind(), ErrorKind::XmlParse);
    }

    /// Serializing output is rendering. Reporting it as an XML parse failure sent a
    /// caller looking at the input document for a problem that is not there.
    #[test]
    fn test_json_serialization_failure_is_classified_as_render() {
        let err = Error::Render("JSON serialization: recursion limit".to_string());
        assert_eq!(err.kind(), ErrorKind::Render);
    }

    // The discriminants are a public ABI contract: they cross the C boundary as
    // `undoc_last_error_kind` values. Pinning every one of them here — via the same
    // macro the sibling crates use — is what makes an accidental renumbering a test
    // failure instead of a silent consumer break.
    uncore::assert_stable_kinds! {
        ErrorKind, test_error_kind_discriminants_are_stable,
        Other = 1,
        Io = 2,
        UnknownFormat = 3,
        UnsupportedFormat = 4,
        ZipArchive = 5,
        XmlParse = 6,
        InvalidData = 7,
        MissingComponent = 8,
        Encoding = 9,
        StyleNotFound = 10,
        ResourceNotFound = 11,
        Encrypted = 12,
        Render = 13,
    }
}
