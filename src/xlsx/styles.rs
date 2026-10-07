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
        let mut reader = crate::decode::reader_for(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut in_num_fmts = false;
        let mut in_cell_xfs = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    match e.name().as_ref() {
                        "numFmts" => in_num_fmts = true,
                        "cellXfs" => in_cell_xfs = true,
                        "xf" if in_cell_xfs => {
                            // Extract numFmtId from xf element
                            let mut num_fmt_id: u32 = 0;
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == "numFmtId" {
                                    if let Ok(id) = attr.value.parse() {
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
                        "numFmt" if in_num_fmts => {
                            let mut num_fmt_id: Option<u32> = None;
                            let mut format_code = String::new();
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    "numFmtId" => {
                                        num_fmt_id = attr.value.parse().ok();
                                    }
                                    "formatCode" => {
                                        format_code = attr.value.to_string();
                                    }
                                    _ => {}
                                }
                            }
                            if let Some(id) = num_fmt_id {
                                styles.num_fmts.insert(id, format_code);
                            }
                        }
                        "xf" if in_cell_xfs => {
                            // Empty xf element (self-closing)
                            let mut num_fmt_id: u32 = 0;
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == "numFmtId" {
                                    if let Ok(id) = attr.value.parse() {
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
                    "numFmts" => in_num_fmts = false,
                    "cellXfs" => in_cell_xfs = false,
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
        if crate::sheet::is_builtin_date_format(num_fmt_id) {
            return true;
        }

        // Check custom formats for date patterns
        if let Some(format_code) = self.num_fmts.get(&num_fmt_id) {
            return crate::sheet::is_date_format_code(format_code);
        }

        false
    }

    /// Check if a format code string represents a date format.
    #[cfg(test)]
    fn is_date_format_code(format_code: &str) -> bool {
        crate::sheet::is_date_format_code(format_code)
    }

    /// Convert Excel serial date number to ISO 8601 date string.
    pub fn serial_to_date(serial: f64) -> Option<String> {
        crate::sheet::serial_to_date(serial)
    }
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
