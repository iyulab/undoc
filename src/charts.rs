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
    pub values: Vec<Option<f64>>,
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
                let formatted = series
                    .values
                    .get(i)
                    .and_then(|value| *value)
                    .map(format_number)
                    .unwrap_or_default();
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
    n.to_string()
}

/// Parse chart XML to extract data
pub fn parse_chart_xml(xml: &str) -> Result<ChartData> {
    let mut reader = quick_xml::Reader::from_str(xml);
    reader.config_mut().trim_text(false);

    let mut chart_data = ChartData {
        title: None,
        categories: Vec::new(),
        series: Vec::new(),
    };

    let mut buf = Vec::new();

    // State tracking
    let mut in_title = false;
    let mut in_ser = false;
    let mut in_tx = false; // Series name / title text
    let mut in_cat = false; // Categories
    let mut in_val = false; // Values
    let mut in_pt = false;
    let mut in_text_node = false;

    let mut current_series_name = String::new();
    let mut current_values: Vec<Option<f64>> = Vec::new();
    let mut current_title = String::new();
    let mut current_text = String::new();
    let mut current_point_text = String::new();
    let mut pt_idx: Option<usize> = None;

    // Temporary storage for categories (only from first series)
    let mut temp_categories: Vec<String> = Vec::new();
    let mut categories_captured = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e)) => {
                let local_name = e.name().local_name();
                match local_name.as_ref() {
                    b"title" => {
                        in_title = true;
                        current_title.clear();
                    }
                    b"ser" => {
                        in_ser = true;
                        current_series_name.clear();
                        current_values.clear();
                    }
                    b"tx" if in_ser || in_title => {
                        in_tx = true;
                    }
                    b"cat" if in_ser => {
                        in_cat = true;
                    }
                    b"val" if in_ser => {
                        in_val = true;
                    }
                    b"pt" => {
                        in_pt = true;
                        current_point_text.clear();
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
                    b"v" | b"t" => {
                        in_text_node = true;
                        current_text.clear();
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::End(ref e)) => {
                let local_name = e.name().local_name();
                match local_name.as_ref() {
                    b"title" => {
                        let title = current_title.trim();
                        if !title.is_empty() {
                            chart_data.title = Some(title.to_string());
                        }
                        in_title = false;
                    }
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
                    b"pt" => {
                        let point_text = current_point_text.trim();

                        if in_title && !point_text.is_empty() {
                            current_title.push_str(point_text);
                        } else if in_ser && in_tx && !point_text.is_empty() {
                            current_series_name.push_str(point_text);
                        } else if in_cat && !point_text.is_empty() {
                            temp_categories.push(point_text.to_string());
                        } else if in_val && !point_text.is_empty() {
                            let val = point_text.parse::<f64>().map_err(|_| {
                                Error::InvalidData(format!(
                                    "invalid chart numeric value: {point_text}"
                                ))
                            })?;
                            if let Some(idx) = pt_idx {
                                while current_values.len() <= idx {
                                    current_values.push(None);
                                }
                                current_values[idx] = Some(val);
                            } else {
                                current_values.push(Some(val));
                            }
                        }

                        in_pt = false;
                        pt_idx = None;
                        current_point_text.clear();
                    }
                    b"v" | b"t" => {
                        if in_text_node {
                            if in_pt {
                                current_point_text.push_str(&current_text);
                            } else if in_title {
                                current_title.push_str(&current_text);
                            } else if in_ser && in_tx {
                                current_series_name.push_str(&current_text);
                            }
                        }
                        in_text_node = false;
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Text(ref e)) => {
                if in_text_node {
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
        assert_eq!(chart_data.series[0].values, vec![Some(100.0), Some(150.0)]);
        assert_eq!(chart_data.series[1].name, "2011");
        assert_eq!(chart_data.series[1].values, vec![Some(120.0), Some(180.0)]);
    }

    #[test]
    fn test_chart_to_table() {
        let chart_data = ChartData {
            title: Some("Revenue".to_string()),
            categories: vec!["Q1".to_string(), "Q2".to_string()],
            series: vec![
                ChartSeries {
                    name: "2010".to_string(),
                    values: vec![Some(100.0), Some(150.0)],
                },
                ChartSeries {
                    name: "2011".to_string(),
                    values: vec![Some(120.0), Some(180.0)],
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
    fn test_parse_chart_title_from_rich_text() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:p>
            <a:r><a:t>Revenue</a:t></a:r>
            <a:r><a:t> Growth</a:t></a:r>
          </a:p>
        </c:rich>
      </c:tx>
    </c:title>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:tx>
            <c:strRef>
              <c:strCache>
                <c:pt idx="0"><c:v>2024</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:tx>
          <c:cat>
            <c:strRef>
              <c:strCache>
                <c:pt idx="0"><c:v>Q1</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:numCache>
                <c:pt idx="0"><c:v>42</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let chart_data = parse_chart_xml(xml).unwrap();

        assert_eq!(chart_data.title.as_deref(), Some("Revenue Growth"));
    }

    #[test]
    fn test_chart_to_table_keeps_missing_values_blank() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:ser>
          <c:tx>
            <c:strRef>
              <c:strCache>
                <c:pt idx="0"><c:v>Series A</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:tx>
          <c:cat>
            <c:strRef>
              <c:strCache>
                <c:pt idx="0"><c:v>Q1</c:v></c:pt>
                <c:pt idx="1"><c:v>Q2</c:v></c:pt>
                <c:pt idx="2"><c:v>Q3</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:numCache>
                <c:pt idx="0"><c:v>100</c:v></c:pt>
                <c:pt idx="2"><c:v>150</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let chart_data = parse_chart_xml(xml).unwrap();
        let table = chart_data.to_table();

        assert_eq!(table.rows[2].cells[0].plain_text(), "Q2");
        assert_eq!(table.rows[2].cells[1].plain_text(), "");
    }

    #[test]
    fn test_parse_chart_invalid_numeric_value_errors() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:ser>
          <c:tx>
            <c:strRef>
              <c:strCache>
                <c:pt idx="0"><c:v>Series A</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:tx>
          <c:cat>
            <c:strRef>
              <c:strCache>
                <c:pt idx="0"><c:v>Q1</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:numCache>
                <c:pt idx="0"><c:v>not-a-number</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let err = parse_chart_xml(xml).unwrap_err();
        assert!(matches!(
            err,
            Error::InvalidData(message) if message == "invalid chart numeric value: not-a-number"
        ));
    }

    #[test]
    fn test_format_number() {
        assert_eq!(format_number(100.0), "100");
        assert_eq!(format_number(8.3), "8.3");
        assert_eq!(format_number(8.300000), "8.3");
        assert_eq!(format_number(12.345678), "12.345678");
        assert_eq!(format_number(12.3456789), "12.3456789");
    }
}
