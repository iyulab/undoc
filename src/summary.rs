//! Document properties of the 97-2003 binary formats: the `\u{5}SummaryInformation` property
//! set ([MS-OLEPS]) every one of them carries, read into the same [`Metadata`] the Office Open
//! XML readers fill from `docProps/core.xml` and `app.xml`.

use std::io::{Read, Seek};

use crate::model::Metadata;

const STREAM: &str = "/\u{5}SummaryInformation";

const VT_I2: u16 = 0x0002;
const VT_I4: u16 = 0x0003;
const VT_LPSTR: u16 = 0x001E;
const VT_LPWSTR: u16 = 0x001F;
const VT_FILETIME: u16 = 0x0040;

const PID_CODEPAGE: u32 = 1;
const PID_TITLE: u32 = 2;
const PID_SUBJECT: u32 = 3;
const PID_AUTHOR: u32 = 4;
const PID_KEYWORDS: u32 = 5;
const PID_COMMENTS: u32 = 6;
const PID_LAST_AUTHOR: u32 = 8;
const PID_CREATED: u32 = 12;
const PID_SAVED: u32 = 13;
const PID_PAGE_COUNT: u32 = 14;
const PID_WORD_COUNT: u32 = 15;
const PID_APP_NAME: u32 = 18;

/// Read a container's summary properties. Absent or unreadable properties leave the metadata
/// empty: they describe the document, they are never a reason not to read it.
pub(crate) fn read<F: Read + Seek>(container: &mut cfb::CompoundFile<F>) -> Metadata {
    let mut data = Vec::new();
    let read = container
        .open_stream(STREAM)
        .and_then(|mut stream| stream.read_to_end(&mut data));
    match read {
        Ok(_) => parse(&data),
        Err(_) => Metadata::default(),
    }
}

fn u16_at(d: &[u8], at: usize) -> Option<u16> {
    d.get(at..at + 2).map(|b| u16::from_le_bytes([b[0], b[1]]))
}

fn u32_at(d: &[u8], at: usize) -> Option<u32> {
    d.get(at..at + 4)
        .map(|b| u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
}

/// One property's value.
enum Value {
    Int(i64),
    Text(String),
    Time(String),
}

fn parse(data: &[u8]) -> Metadata {
    let mut meta = Metadata::default();
    // PropertySetStream: byte order, version, system id, CLSID, set count, then the first
    // set's FMTID and offset — SummaryInformation is the first (and only) set.
    let Some(section) = u32_at(data, 44).map(|o| o as usize) else {
        return meta;
    };
    let Some(count) = u32_at(data, section + 4) else {
        return meta;
    };
    let entries: Vec<(u32, usize)> = (0..count.min(1024) as usize)
        .filter_map(|i| {
            let at = section + 8 + i * 8;
            Some((u32_at(data, at)?, section + u32_at(data, at + 4)? as usize))
        })
        .collect();

    // Strings are in the set's code page, which is itself a property.
    let codepage = entries
        .iter()
        .find(|(id, _)| *id == PID_CODEPAGE)
        .and_then(|&(_, at)| u16_at(data, at + 4))
        .unwrap_or(1252);

    for &(id, at) in &entries {
        let Some(value) = value_at(data, at, codepage) else {
            continue;
        };
        match (id, value) {
            (PID_TITLE, Value::Text(t)) => meta.title = Some(t),
            (PID_SUBJECT, Value::Text(t)) => meta.subject = Some(t),
            (PID_AUTHOR, Value::Text(t)) => meta.author = Some(t),
            (PID_COMMENTS, Value::Text(t)) => meta.description = Some(t),
            (PID_LAST_AUTHOR, Value::Text(t)) => meta.last_modified_by = Some(t),
            (PID_APP_NAME, Value::Text(t)) => meta.application = Some(t),
            (PID_KEYWORDS, Value::Text(t)) => {
                meta.keywords = t
                    .split([';', ','])
                    .map(str::trim)
                    .filter(|k| !k.is_empty())
                    .map(str::to_string)
                    .collect();
            }
            (PID_CREATED, Value::Time(t)) => meta.created = Some(t),
            (PID_SAVED, Value::Time(t)) => meta.modified = Some(t),
            (PID_PAGE_COUNT, Value::Int(n)) if n > 0 => meta.page_count = Some(n as u32),
            (PID_WORD_COUNT, Value::Int(n)) if n > 0 => meta.word_count = Some(n as u32),
            _ => {}
        }
    }
    meta
}

fn value_at(data: &[u8], at: usize, codepage: u16) -> Option<Value> {
    let kind = u16_at(data, at)?;
    let body = at + 4;
    match kind {
        VT_I2 => Some(Value::Int(u16_at(data, body)? as i16 as i64)),
        VT_I4 => Some(Value::Int(u32_at(data, body)? as i32 as i64)),
        VT_LPSTR => {
            let len = u32_at(data, body)? as usize;
            let bytes = data.get(body + 4..body + 4 + len)?;
            let text = decode(bytes, codepage)?;
            let text = text.trim_end_matches('\0').trim().to_string();
            (!text.is_empty()).then_some(Value::Text(text))
        }
        VT_LPWSTR => {
            let count = u32_at(data, body)? as usize;
            let units: Vec<u16> = data
                .get(body + 4..body + 4 + count * 2)?
                .as_chunks::<2>()
                .0
                .iter()
                .map(|&c| u16::from_le_bytes(c))
                .take_while(|&u| u != 0)
                .collect();
            let text = String::from_utf16_lossy(&units).trim().to_string();
            (!text.is_empty()).then_some(Value::Text(text))
        }
        VT_FILETIME => {
            let ticks = u32_at(data, body)? as u64 | (u32_at(data, body + 4)? as u64) << 32;
            filetime(ticks).map(Value::Time)
        }
        _ => None,
    }
}

/// A FILETIME (100 ns ticks since 1601-01-01 UTC) as `YYYY-MM-DDThh:mm:ssZ`, the form the
/// .docx reader reports. Zero — "never set" — is none.
fn filetime(ticks: u64) -> Option<String> {
    const UNIX_EPOCH_SECONDS: i64 = 11_644_473_600;
    if ticks == 0 {
        return None;
    }
    let secs = (ticks / 10_000_000) as i64 - UNIX_EPOCH_SECONDS;
    let (days, rem) = (secs.div_euclid(86_400), secs.rem_euclid(86_400));
    // Civil date from days since 1970-01-01 (Howard Hinnant's algorithm).
    let z = days + 719_468;
    let era = z.div_euclid(146_097);
    let doe = z - era * 146_097;
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let day = doy - (153 * mp + 2) / 5 + 1;
    let month = if mp < 10 { mp + 3 } else { mp - 9 };
    let year = yoe + era * 400 + if month <= 2 { 1 } else { 0 };
    Some(format!(
        "{year:04}-{month:02}-{day:02}T{:02}:{:02}:{:02}Z",
        rem / 3600,
        rem % 3600 / 60,
        rem % 60
    ))
}

/// Decode a property string in a Windows code page. Without the `codepages` feature only the
/// code pages that need no tables are decoded — and any string that is plain ASCII; anything
/// else is left out rather than shown as the wrong characters.
fn decode(bytes: &[u8], codepage: u16) -> Option<String> {
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
        _ => decode_with_tables(bytes, codepage),
    }
}

fn windows_1252(byte: u8) -> char {
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
fn decode_with_tables(bytes: &[u8], codepage: u16) -> Option<String> {
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
    (!had_errors).then(|| text.into_owned())
}

#[cfg(not(feature = "codepages"))]
fn decode_with_tables(_bytes: &[u8], _codepage: u16) -> Option<String> {
    None
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A SummaryInformation stream with the given (pid, type, value bytes) properties.
    pub(crate) fn stream(props: &[(u32, u16, Vec<u8>)]) -> Vec<u8> {
        let mut out = vec![0xFE, 0xFF, 0, 0];
        out.extend([0u8; 4 + 16]);
        out.extend(1u32.to_le_bytes());
        out.extend([0u8; 16]); // FMTID
        out.extend(48u32.to_le_bytes());
        let mut body = Vec::new();
        let mut table = Vec::new();
        let head = 8 + props.len() * 8;
        for (id, kind, value) in props {
            table.extend(id.to_le_bytes());
            table.extend(((head + body.len()) as u32).to_le_bytes());
            body.extend(kind.to_le_bytes());
            body.extend([0, 0]);
            body.extend(value);
            while body.len() % 4 != 0 {
                body.push(0);
            }
        }
        out.extend(((head + body.len()) as u32).to_le_bytes());
        out.extend((props.len() as u32).to_le_bytes());
        out.extend(table);
        out.extend(body);
        out
    }

    fn lpstr(bytes: &[u8]) -> Vec<u8> {
        let mut v = ((bytes.len() + 1) as u32).to_le_bytes().to_vec();
        v.extend(bytes);
        v.push(0);
        v
    }

    #[test]
    fn summary_properties_fill_the_metadata() {
        let created: u64 = (1_700_000_000 + 11_644_473_600) * 10_000_000;
        let data = stream(&[
            (PID_CODEPAGE, VT_I2, 1252u16.to_le_bytes().to_vec()),
            (PID_TITLE, VT_LPSTR, lpstr(b"Annual \x93report\x94")),
            (PID_AUTHOR, VT_LPSTR, lpstr(b"A. Writer")),
            (PID_KEYWORDS, VT_LPSTR, lpstr(b"finance; 2023, audit")),
            (PID_CREATED, VT_FILETIME, created.to_le_bytes().to_vec()),
            (PID_PAGE_COUNT, VT_I4, 12i32.to_le_bytes().to_vec()),
        ]);
        let meta = parse(&data);
        assert_eq!(meta.title.as_deref(), Some("Annual \u{201C}report\u{201D}"));
        assert_eq!(meta.author.as_deref(), Some("A. Writer"));
        assert_eq!(meta.keywords, ["finance", "2023", "audit"]);
        assert_eq!(meta.created.as_deref(), Some("2023-11-14T22:13:20Z"));
        assert_eq!(meta.page_count, Some(12));
    }

    #[test]
    fn utf8_code_page_and_empty_strings() {
        let data = stream(&[
            (PID_CODEPAGE, VT_I2, (65001u16).to_le_bytes().to_vec()),
            (PID_TITLE, VT_LPSTR, lpstr("보고서".as_bytes())),
            (PID_SUBJECT, VT_LPSTR, lpstr(b"")),
        ]);
        let meta = parse(&data);
        assert_eq!(meta.title.as_deref(), Some("보고서"));
        assert_eq!(meta.subject, None);
    }

    #[cfg(feature = "codepages")]
    #[test]
    fn multi_byte_code_pages_decode_with_tables() {
        // "보고서" in EUC-KR (949).
        let data = stream(&[
            (PID_CODEPAGE, VT_I2, 949u16.to_le_bytes().to_vec()),
            (
                PID_TITLE,
                VT_LPSTR,
                lpstr(&[0xBA, 0xB8, 0xB0, 0xED, 0xBC, 0xAD]),
            ),
        ]);
        assert_eq!(parse(&data).title.as_deref(), Some("보고서"));
    }

    #[test]
    fn a_never_set_time_is_none() {
        assert_eq!(filetime(0), None);
        assert_eq!(
            filetime(116_444_736_000_000_000).as_deref(),
            Some("1970-01-01T00:00:00Z")
        );
    }
}
