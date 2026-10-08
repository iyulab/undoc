//! An ECMA-376 encrypted document must reach the caller as *encrypted*.
//!
//! Both an encrypted OOXML package and a legacy binary Office file arrive as an OLE/CFB
//! container with an identical header, and the two answers send the caller in opposite
//! directions: one supplies a password, the other looks for a converter. Reporting the
//! disjunction leaves them to guess.
//!
//! These go through the public entry points rather than the detection helper, because
//! that is where a consumer meets the behaviour — `parse_bytes` runs detection first, so
//! the classification has to survive that path.

use undoc::{parse_bytes, Error, ErrorKind};

/// Build a CFB container holding the named root streams, and nothing else.
///
/// A real encrypted package also carries the ciphertext; the classification depends on
/// the directory, not the payload, so the streams are left empty on purpose.
///
/// Reaches `cfb` through the crate's unconditional `[dependencies]` entry rather than a
/// dev-dependency. If that ever moves behind a feature, this file stops compiling — add
/// the dev-dependency then, rather than wondering where the crate went.
fn cfb_with_streams(names: &[&str]) -> Vec<u8> {
    let mut container =
        cfb::CompoundFile::create(std::io::Cursor::new(Vec::new())).expect("create CFB container");
    for name in names {
        container.create_stream(name).expect("create stream");
    }
    container.flush().expect("flush CFB container");
    container.into_inner().into_inner()
}

const CFB_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

#[test]
fn parsing_an_ecma376_encrypted_package_reports_encrypted() {
    let data = cfb_with_streams(&["/EncryptedPackage", "/EncryptionInfo"]);

    let err = parse_bytes(&data).expect_err("an encrypted package cannot be parsed");

    assert_eq!(err.kind(), ErrorKind::Encrypted, "got: {err}");
    assert!(matches!(err, Error::Encrypted));
}

/// The same header, the opposite answer — this is the pair the classification exists for.
/// A container that holds no Office document the library recognises is unsupported.
#[test]
fn parsing_an_unrecognised_cfb_container_reports_an_unsupported_format() {
    let other = cfb_with_streams(&["/Contents"]);
    let encrypted = cfb_with_streams(&["/EncryptedPackage"]);

    assert_eq!(other[..CFB_MAGIC.len()], CFB_MAGIC);
    assert_eq!(encrypted[..CFB_MAGIC.len()], CFB_MAGIC);

    let err = parse_bytes(&other).expect_err("no Office document in it");

    assert_eq!(err.kind(), ErrorKind::UnsupportedFormat, "got: {err}");
}

/// An Excel 5.0/95 workbook (`Book` stream) is read, so an empty one is a damaged workbook.
#[test]
fn an_empty_excel_95_book_stream_is_a_damaged_workbook() {
    let err = parse_bytes(&cfb_with_streams(&["/Book"])).unwrap_err();

    assert_eq!(err.kind(), ErrorKind::InvalidData, "got: {err}");
}

/// A Word 97-2003 document is read, so a Word stream with nothing in it is a damaged
/// document — not an encrypted one, and not an unsupported format.
#[test]
fn an_empty_word_stream_is_a_damaged_document() {
    let err = parse_bytes(&cfb_with_streams(&["/WordDocument"])).unwrap_err();

    assert_eq!(err.kind(), ErrorKind::InvalidData, "got: {err}");
}

/// An unrecognised value must not be collapsed into a familiar one, and neither of these
/// two is the other. Guards the pair rather than each kind alone, because the defect this
/// closes was reporting *both* as unsupported.
#[test]
fn the_two_cfb_answers_are_distinguishable() {
    let encrypted = parse_bytes(&cfb_with_streams(&["/EncryptedPackage"])).unwrap_err();
    let legacy = parse_bytes(&cfb_with_streams(&["/Contents"])).unwrap_err();

    assert_ne!(encrypted.kind(), legacy.kind());
}
