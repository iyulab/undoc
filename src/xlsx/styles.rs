//! XLSX styles parsing for number formats.

use std::collections::HashMap;

/// Styles information parsed from xl/styles.xml.
#[derive(Debug, Default)]
pub struct Styles {
    /// Custom number formats: numFmtId -> formatCode
    num_fmts: HashMap<u32, String>,
    /// Cell style formats: style index -> numFmtId
    cell_xfs: Vec<u32>,
}

impl Styles {
    /// Parse styles from xl/styles.xml content.
    pub fn parse(xml: &str) -> Self {
        let mut styles = Self::default();
        let mut reader = quick_xml::Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut in_num_fmts = false;
        let mut in_cell_xfs = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    match e.name().as_ref() {
                        b"numFmts" => in_num_fmts = true,
                        b"cellXfs" => in_cell_xfs = true,
                        b"xf" if in_cell_xfs => {
                            // Extract numFmtId from xf element
                            let mut num_fmt_id: u32 = 0;
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"numFmtId" {
                                    if let Ok(id) = String::from_utf8_lossy(&attr.value).parse() {
                                        num_fmt_id = id;
                                    }
                                }
                            }
                            styles.cell_xfs.push(num_fmt_id);
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Empty(ref e)) => {
                    match e.name().as_ref() {
                        b"numFmt" if in_num_fmts => {
                            let mut num_fmt_id: Option<u32> = None;
                            let mut format_code = String::new();
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"numFmtId" => {
                                        num_fmt_id =
                                            String::from_utf8_lossy(&attr.value).parse().ok();
                                    }
                                    b"formatCode" => {
                                        format_code =
                                            String::from_utf8_lossy(&attr.value).to_string();
                                    }
                                    _ => {}
                                }
                            }
                            if let Some(id) = num_fmt_id {
                                styles.num_fmts.insert(id, format_code);
                            }
                        }
                        b"xf" if in_cell_xfs => {
                            // Empty xf element (self-closing)
                            let mut num_fmt_id: u32 = 0;
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"numFmtId" {
                                    if let Ok(id) = String::from_utf8_lossy(&attr.value).parse() {
                                        num_fmt_id = id;
                                    }
                                }
                            }
                            styles.cell_xfs.push(num_fmt_id);
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::End(ref e)) => match e.name().as_ref() {
                    b"numFmts" => in_num_fmts = false,
                    b"cellXfs" => in_cell_xfs = false,
                    _ => {}
                },
                Ok(quick_xml::events::Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }

        styles
    }

    /// Get the numFmtId for a cell style index.
    pub fn get_num_fmt_id(&self, style_index: usize) -> Option<u32> {
        self.cell_xfs.get(style_index).copied()
    }

    /// Check if a numFmtId represents a date format.
    pub fn is_date_format(&self, num_fmt_id: u32) -> bool {
        // Built-in date formats (Excel standard)
        // 14-22: Date formats
        // 45-47: Time formats
        if (14..=22).contains(&num_fmt_id) || (45..=47).contains(&num_fmt_id) {
            return true;
        }

        // Check custom formats for date patterns
        if let Some(format_code) = self.num_fmts.get(&num_fmt_id) {
            return Self::is_date_format_code(format_code);
        }

        false
    }

    /// Check if a format code string represents a date format.
    fn is_date_format_code(format_code: &str) -> bool {
        // Date patterns: d, m, y (case insensitive, not in quotes or brackets)
        // Time patterns: h, s (case insensitive)
        // We need to exclude patterns in square brackets [Red] or quotes "text"

        let mut in_bracket = false;
        let mut in_quote = false;
        let mut prev_char = '\0';

        for c in format_code.chars() {
            match c {
                '[' if !in_quote => in_bracket = true,
                ']' if !in_quote => in_bracket = false,
                '"' => in_quote = !in_quote,
                _ if !in_bracket && !in_quote => {
                    // Check for date/time patterns
                    let lower = c.to_ascii_lowercase();
                    match lower {
                        // 'd' for day, 'm' for month (but not 'mm:ss' which is minutes)
                        'd' => return true,
                        'y' => return true,
                        // 'h' for hour indicates time, which is often stored as fractional day
                        // But we mainly want date, so check for 'm' after 'd' or before 'd'
                        'm' => {
                            // 'm' could be month or minute
                            // If preceded by 'd' or 'y', it's likely month
                            // If preceded by 'h' or followed by 's', it's likely minute
                            // For simplicity, check surrounding context
                            let lower_prev = prev_char.to_ascii_lowercase();
                            if lower_prev == 'd' || lower_prev == 'y' {
                                return true; // Month after day/year
                            }
                            // Could also be month at start or standalone
                            // Check if format contains 'd' or 'y' anywhere
                            let lower_format = format_code.to_lowercase();
                            if lower_format.contains('d') || lower_format.contains('y') {
                                return true;
                            }
                        }
                        _ => {}
                    }
                }
                _ => {}
            }
            prev_char = c;
        }

        false
    }

    /// Convert Excel serial date number to ISO 8601 date string.
    pub fn serial_to_date(serial: f64) -> Option<String> {
        // Excel date system: days since December 30, 1899
        // (Excel incorrectly treats 1900 as a leap year for Lotus 1-2-3 compatibility)

        if serial < 0.0 {
            return None;
        }

        // Handle the "Lotus 1-2-3" bug: Excel thinks Feb 29, 1900 exists
        // Serial 60 = Feb 29, 1900 (doesn't exist)
        // Serial 61 = Mar 1, 1900
        let adjusted_serial = if serial > 60.0 { serial - 1.0 } else { serial };

        // Days since January 1, 1900 (day 1 = Jan 1, 1900)
        let days = adjusted_serial.floor() as i64;

        // Convert to date
        // January 1, 1900 is day 1
        // Using a simple calculation:
        // Base date: 1899-12-31 (so day 1 = 1900-01-01)

        // Calculate year, month, day
        let (year, month, day) = days_to_ymd(days)?;

        // Check if there's a time component
        let time_fraction = serial.fract();
        if time_fraction > 0.0001 {
            // Has time component
            let total_seconds = (time_fraction * 86400.0).round() as u32;
            let hours = total_seconds / 3600;
            let minutes = (total_seconds % 3600) / 60;
            let seconds = total_seconds % 60;
            Some(format!(
                "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}",
                year, month, day, hours, minutes, seconds
            ))
        } else {
            Some(format!("{:04}-{:02}-{:02}", year, month, day))
        }
    }
}

/// Convert days since December 31, 1899 to (year, month, day).
fn days_to_ymd(days: i64) -> Option<(i32, u32, u32)> {
    if days < 1 {
        return None;
    }

    // Start from 1900-01-01, which is serial day 1
    let mut year = 1900;
    let mut remaining_days = days;

    // Year loop
    loop {
        let days_in_year = if is_leap_year(year) { 366 } else { 365 };
        if remaining_days <= days_in_year {
            break;
        }
        remaining_days -= days_in_year;
        year += 1;
    }

    // Month loop
    let months_days = if is_leap_year(year) {
        [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    } else {
        [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    };

    let mut month = 1u32;
    for &days_in_month in &months_days {
        if remaining_days <= days_in_month as i64 {
            break;
        }
        remaining_days -= days_in_month as i64;
        month += 1;
    }

    let day = remaining_days.max(1) as u32;

    Some((year, month, day))
}

/// Check if a year is a leap year.
fn is_leap_year(year: i32) -> bool {
    (year % 4 == 0 && year % 100 != 0) || (year % 400 == 0)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_builtin_date_formats() {
        let styles = Styles::default();

        // Built-in date formats (14-22)
        assert!(styles.is_date_format(14)); // m/d/yyyy
        assert!(styles.is_date_format(15)); // d-mmm-yy
        assert!(styles.is_date_format(16)); // d-mmm
        assert!(styles.is_date_format(17)); // mmm-yy
        assert!(styles.is_date_format(22)); // m/d/yy h:mm

        // Not date formats
        assert!(!styles.is_date_format(0)); // General
        assert!(!styles.is_date_format(1)); // 0
        assert!(!styles.is_date_format(2)); // 0.00
    }

    #[test]
    fn test_custom_date_format_detection() {
        assert!(Styles::is_date_format_code("mmmm\\ d\\,\\ yyyy"));
        assert!(Styles::is_date_format_code("yyyy-mm-dd"));
        assert!(Styles::is_date_format_code("d/m/yy"));
        assert!(Styles::is_date_format_code("[$-409]mmmm\\ d\\,\\ yyyy;@"));

        // Not date formats
        assert!(!Styles::is_date_format_code("0.00"));
        assert!(!Styles::is_date_format_code("#,##0"));
        assert!(!Styles::is_date_format_code("\"$\"#,##0.00"));
    }

    #[test]
    fn test_serial_to_date() {
        // Excel serial dates
        assert_eq!(Styles::serial_to_date(1.0), Some("1900-01-01".to_string()));
        assert_eq!(Styles::serial_to_date(2.0), Some("1900-01-02".to_string()));
        assert_eq!(Styles::serial_to_date(59.0), Some("1900-02-28".to_string()));
        // Note: serial 60 is the fake Feb 29, 1900
        assert_eq!(Styles::serial_to_date(61.0), Some("1900-03-01".to_string()));

        // More recent dates
        assert_eq!(
            Styles::serial_to_date(44197.0),
            Some("2021-01-01".to_string())
        );
        assert_eq!(
            Styles::serial_to_date(45658.0),
            Some("2025-01-01".to_string())
        );

        // With time component
        assert_eq!(
            Styles::serial_to_date(44197.5),
            Some("2021-01-01T12:00:00".to_string())
        );
    }
}
