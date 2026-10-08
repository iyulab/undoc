//! Text in a Windows code page — what the 97-2003 binary formats and their predecessors store
//! wherever they do not use UTF-16: compressed Word text, document property strings, and every
//! string of an Excel workbook older than Excel 97.
//!
//! Without the `codepages` feature only the code pages that need no tables are decoded
//! (Windows-1252, UTF-16, UTF-8), and any byte string that is plain ASCII.

/// `CODEPAGE` values that name no Windows code page: Excel's own identifiers for the code
/// pages of BIFF2–BIFF4 ([MS-XLS] 2.4.52).
const APPLE_ROMAN: u16 = 32768;
const WINDOWS_ANSI: u16 = 32769;

/// Decode `bytes` in `codepage`, or `None` when the code page is unknown or the bytes are not
/// valid in it — for text that is better left out than shown as the wrong characters.
pub(crate) fn decode_strict(bytes: &[u8], codepage: u16) -> Option<String> {
    let codepage = match codepage {
        APPLE_ROMAN => 10000,
        WINDOWS_ANSI => 1252,
        other => other,
    };
    match codepage {
        // CP_WINUNICODE, CP_UTF8 (65001 read as a signed 16-bit value is -535).
        1200 => {
            let units: Vec<u16> = bytes
                .as_chunks::<2>()
                .0
                .iter()
                .map(|&c| u16::from_le_bytes(c))
                .collect();
            Some(String::from_utf16_lossy(&units))
        }
        65001 => String::from_utf8(bytes.to_vec()).ok(),
        1252 => Some(bytes.iter().map(|&b| windows_1252(b)).collect()),
        _ if bytes.is_ascii() => Some(String::from_utf8_lossy(bytes).into_owned()),
        _ => decode_with_tables(bytes, codepage, false),
    }
}

/// Decode `bytes` in `codepage`, never dropping text: a byte the code page cannot account for
/// — or any non-ASCII byte when the code page itself is unknown or its table not compiled in —
/// becomes U+FFFD. For content, where a visible gap is better than a missing cell.
#[cfg(feature = "xls")]
pub(crate) fn decode_lossy(bytes: &[u8], codepage: u16) -> String {
    if let Some(text) = decode_strict(bytes, codepage) {
        return text;
    }
    decode_with_tables(bytes, codepage, true).unwrap_or_else(|| {
        bytes
            .iter()
            .map(|&b| if b.is_ascii() { b as char } else { '\u{FFFD}' })
            .collect()
    })
}

/// The code page a font's `charset` ([MS-XLS] 2.5.82, the GDI `LOGFONT` character sets)
/// writes in; `None` for the ANSI, default and symbol sets, which do not narrow it.
#[cfg(feature = "xls")]
pub(crate) fn from_charset(charset: u8) -> Option<u16> {
    Some(match charset {
        128 => 932,  // SHIFTJIS
        129 => 949,  // HANGUL
        134 => 936,  // GB2312
        136 => 950,  // CHINESEBIG5
        161 => 1253, // GREEK
        162 => 1254, // TURKISH
        163 => 1258, // VIETNAMESE
        177 => 1255, // HEBREW
        178 => 1256, // ARABIC
        186 => 1257, // BALTIC
        204 => 1251, // RUSSIAN
        222 => 874,  // THAI
        238 => 1250, // EASTEUROPE
        77 => 10000, // MAC
        _ => return None,
    })
}

/// One byte of Windows-1252.
pub(crate) fn windows_1252(byte: u8) -> char {
    const HIGH: [u16; 32] = [
        0x20AC, 0x0081, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, 0x02C6, 0x2030, 0x0160,
        0x2039, 0x0152, 0x008D, 0x017D, 0x008F, 0x0090, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022,
        0x2013, 0x2014, 0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x009D, 0x017E, 0x0178,
    ];
    match byte {
        0x80..=0x9F => char::from_u32(HIGH[(byte - 0x80) as usize] as u32).unwrap_or('\u{FFFD}'),
        _ => byte as char,
    }
}

#[cfg(feature = "codepages")]
fn decode_with_tables(bytes: &[u8], codepage: u16, lossy: bool) -> Option<String> {
    use encoding_rs::*;
    let encoding: &'static Encoding = match codepage {
        874 => WINDOWS_874,
        932 => SHIFT_JIS,
        936 => GBK,
        949 => EUC_KR,
        950 => BIG5,
        1250 => WINDOWS_1250,
        1251 => WINDOWS_1251,
        1253 => WINDOWS_1253,
        1254 => WINDOWS_1254,
        1255 => WINDOWS_1255,
        1256 => WINDOWS_1256,
        1257 => WINDOWS_1257,
        1258 => WINDOWS_1258,
        10000 => MACINTOSH,
        _ => return None,
    };
    let (text, _, had_errors) = encoding.decode(bytes);
    (lossy || !had_errors).then(|| text.into_owned())
}

#[cfg(not(feature = "codepages"))]
fn decode_with_tables(_bytes: &[u8], _codepage: u16, _lossy: bool) -> Option<String> {
    None
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn windows_1252_maps_the_windows_block_and_passes_latin1_through() {
        assert_eq!(windows_1252(b'A'), 'A');
        assert_eq!(windows_1252(0x93), '\u{201C}');
        assert_eq!(windows_1252(0x80), '\u{20AC}');
        assert_eq!(windows_1252(0xE9), 'é');
        assert_eq!(decode_strict(&[0x80], WINDOWS_ANSI).as_deref(), Some("€"));
    }

    #[cfg(feature = "xls")]
    #[test]
    fn strict_leaves_out_what_lossy_marks() {
        // An unknown code page: strict refuses the non-ASCII byte, lossy marks it.
        assert_eq!(decode_strict(b"ab\xFF", 4242), None);
        assert_eq!(decode_lossy(b"ab\xFF", 4242), "ab\u{FFFD}");
        assert_eq!(decode_lossy(b"plain", 4242), "plain");
    }

    #[cfg(all(feature = "codepages", feature = "xls"))]
    #[test]
    fn multi_byte_code_pages_decode_with_tables() {
        let euc_kr = [0xBA, 0xB8, 0xB0, 0xED, 0xBC, 0xAD]; // 보고서
        assert_eq!(decode_strict(&euc_kr, 949).as_deref(), Some("보고서"));
        // A truncated double-byte character: strict refuses, lossy keeps the rest.
        assert_eq!(decode_strict(&euc_kr[..5], 949), None);
        assert_eq!(decode_lossy(&euc_kr[..5], 949), "보고\u{FFFD}");
        assert_eq!(
            decode_strict(b"caf\x8E", APPLE_ROMAN).as_deref(),
            Some("café")
        );
    }

    #[cfg(feature = "xls")]
    #[test]
    fn charsets_name_their_code_pages() {
        assert_eq!(from_charset(129), Some(949));
        assert_eq!(from_charset(0), None);
        assert_eq!(from_charset(2), None);
    }
}
