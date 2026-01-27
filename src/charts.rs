//! PPTX chart parsing module.
//!
//! Extracts data from chart XML files and converts them to tables for RAG-ready output.

use crate::error::{Error, Result};
use crate::model::{Cell, Row, Table};

/// Parsed chart data
#[derive(Debug, Clone)]
pub struct ChartData {
    /// Chart title (if available)
    pub title: Option<String>,
    /// Category labels (X-axis)
    pub categories: Vec<String>,
    /// Series data
    pub series: Vec<ChartSeries>,
}

/// A data series in a chart
#[derive(Debug, Clone)]
pub struct ChartSeries {
    /// Series name (legend label)
    pub name: String,
    /// Data values
    pub values: Vec<f64>,
}

impl ChartData {
    /// Convert chart data to a Table for markdown rendering
    pub fn to_table(&self) -> Table {
        let mut table = Table::new();

        // Build header row: Category | Series1 | Series2 | ...
        let mut header_cells = vec![Cell::header("Category")];
        for series in &self.series {
            header_cells.push(Cell::header(&series.name));
        }
        let mut header = Row::header(header_cells);
        header.is_header = true;
        table.add_row(header);

        // Build data rows
        for (i, category) in self.categories.iter().enumerate() {
            let mut cells = vec![Cell::with_text(category)];
            for series in &self.series {
                let value = series.values.get(i).copied().unwrap_or(0.0);
                // Format number: remove trailing zeros
                let formatted = format_number(value);
                cells.push(Cell::with_text(&formatted));
            }
            table.add_row(Row {
                cells,
                is_header: false,
                height: None,
            });
        }

        table
    }

    /// Check if chart data is empty
    pub fn is_empty(&self) -> bool {
        self.categories.is_empty() || self.series.is_empty()
    }
}

/// Format a number, removing unnecessary trailing zeros
fn format_number(n: f64) -> String {
    if n.fract() == 0.0 {
        format!("{:.0}", n)
    } else {
        // Remove trailing zeros
        let s = format!("{:.6}", n);
        s.trim_end_matches('0').trim_end_matches('.').to_string()
    }
}

/// Parse chart XML to extract data
pub fn parse_chart_xml(xml: &str) -> Result<ChartData> {
    let mut reader = quick_xml::Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut chart_data = ChartData {
        title: None,
        categories: Vec::new(),
        series: Vec::new(),
    };

    let mut buf = Vec::new();

    // State tracking
    let mut in_ser = false;
    let mut in_tx = false; // Series name
    let mut in_cat = false; // Categories
    let mut in_val = false; // Values
    let mut in_str_cache = false;
    let mut in_num_cache = false;
    let mut in_pt = false;
    let mut in_v = false;

    let mut current_series_name = String::new();
    let mut current_values: Vec<f64> = Vec::new();
    let mut current_text = String::new();
    let mut pt_idx: Option<usize> = None;

    // Temporary storage for categories (only from first series)
    let mut temp_categories: Vec<String> = Vec::new();
    let mut categories_captured = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e)) => {
                let local_name = e.name().local_name();
                match local_name.as_ref() {
                    b"ser" => {
                        in_ser = true;
                        current_series_name.clear();
                        current_values.clear();
                    }
                    b"tx" if in_ser => {
                        in_tx = true;
                    }
                    b"cat" if in_ser => {
                        in_cat = true;
                    }
                    b"val" if in_ser => {
                        in_val = true;
                    }
                    b"strCache" => {
                        in_str_cache = true;
                    }
                    b"numCache" => {
                        in_num_cache = true;
                    }
                    b"pt" => {
                        in_pt = true;
                        // Get idx attribute
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"idx" {
                                if let Ok(idx) =
                                    String::from_utf8_lossy(&attr.value).parse::<usize>()
                                {
                                    pt_idx = Some(idx);
                                }
                            }
                        }
                    }
                    b"v" if in_pt => {
                        in_v = true;
                        current_text.clear();
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::End(ref e)) => {
                let local_name = e.name().local_name();
                match local_name.as_ref() {
                    b"ser" => {
                        // Save series if we have data
                        if !current_series_name.is_empty() || !current_values.is_empty() {
                            let name = if current_series_name.is_empty() {
                                format!("Series {}", chart_data.series.len() + 1)
                            } else {
                                current_series_name.clone()
                            };
                            chart_data.series.push(ChartSeries {
                                name,
                                values: current_values.clone(),
                            });
                        }

                        // Capture categories from first series
                        if !categories_captured && !temp_categories.is_empty() {
                            chart_data.categories = temp_categories.clone();
                            categories_captured = true;
                        }
                        temp_categories.clear();

                        in_ser = false;
                    }
                    b"tx" => {
                        in_tx = false;
                    }
                    b"cat" => {
                        in_cat = false;
                    }
                    b"val" => {
                        in_val = false;
                    }
                    b"strCache" => {
                        in_str_cache = false;
                    }
                    b"numCache" => {
                        in_num_cache = false;
                    }
                    b"pt" => {
                        in_pt = false;
                        pt_idx = None;
                    }
                    b"v" => {
                        if in_v {
                            // Process the value based on context
                            if in_tx && in_str_cache {
                                // Series name
                                current_series_name = current_text.trim().to_string();
                            } else if in_cat && in_str_cache {
                                // Category label
                                temp_categories.push(current_text.trim().to_string());
                            } else if in_val && in_num_cache {
                                // Numeric value
                                if let Ok(val) = current_text.trim().parse::<f64>() {
                                    // Ensure vector is large enough
                                    if let Some(idx) = pt_idx {
                                        while current_values.len() <= idx {
                                            current_values.push(0.0);
                                        }
                                        current_values[idx] = val;
                                    } else {
                                        current_values.push(val);
                                    }
                                }
                            }
                        }
                        in_v = false;
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Text(ref e)) => {
                if in_v {
                    if let Ok(text) = e.unescape() {
                        current_text.push_str(&text);
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(e) => return Err(Error::XmlParse(e.to_string())),
            _ => {}
        }
        buf.clear();
    }

    Ok(chart_data)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_bar_chart() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:tx>
            <c:strRef>
              <c:strCache>
                <c:pt idx="0"><c:v>2010</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:tx>
          <c:cat>
            <c:strRef>
              <c:strCache>
                <c:pt idx="0"><c:v>Q1</c:v></c:pt>
                <c:pt idx="1"><c:v>Q2</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:numCache>
                <c:pt idx="0"><c:v>100</c:v></c:pt>
                <c:pt idx="1"><c:v>150</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:ser>
          <c:tx>
            <c:strRef>
              <c:strCache>
                <c:pt idx="0"><c:v>2011</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:tx>
          <c:cat>
            <c:strRef>
              <c:strCache>
                <c:pt idx="0"><c:v>Q1</c:v></c:pt>
                <c:pt idx="1"><c:v>Q2</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:numCache>
                <c:pt idx="0"><c:v>120</c:v></c:pt>
                <c:pt idx="1"><c:v>180</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let chart_data = parse_chart_xml(xml).unwrap();

        assert_eq!(chart_data.categories, vec!["Q1", "Q2"]);
        assert_eq!(chart_data.series.len(), 2);
        assert_eq!(chart_data.series[0].name, "2010");
        assert_eq!(chart_data.series[0].values, vec![100.0, 150.0]);
        assert_eq!(chart_data.series[1].name, "2011");
        assert_eq!(chart_data.series[1].values, vec![120.0, 180.0]);
    }

    #[test]
    fn test_chart_to_table() {
        let chart_data = ChartData {
            title: Some("Revenue".to_string()),
            categories: vec!["Q1".to_string(), "Q2".to_string()],
            series: vec![
                ChartSeries {
                    name: "2010".to_string(),
                    values: vec![100.0, 150.0],
                },
                ChartSeries {
                    name: "2011".to_string(),
                    values: vec![120.0, 180.0],
                },
            ],
        };

        let table = chart_data.to_table();

        assert_eq!(table.row_count(), 3); // header + 2 data rows
        assert_eq!(table.column_count(), 3); // Category + 2 series

        // Check header
        assert_eq!(table.rows[0].cells[0].plain_text(), "Category");
        assert_eq!(table.rows[0].cells[1].plain_text(), "2010");
        assert_eq!(table.rows[0].cells[2].plain_text(), "2011");

        // Check data
        assert_eq!(table.rows[1].cells[0].plain_text(), "Q1");
        assert_eq!(table.rows[1].cells[1].plain_text(), "100");
        assert_eq!(table.rows[2].cells[0].plain_text(), "Q2");
    }

    #[test]
    fn test_format_number() {
        assert_eq!(format_number(100.0), "100");
        assert_eq!(format_number(8.3), "8.3");
        assert_eq!(format_number(8.300000), "8.3");
        assert_eq!(format_number(12.345678), "12.345678");
    }
}
