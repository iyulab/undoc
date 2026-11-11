//! Charts: the data of a chart part painted as the plain chart it describes -- columns and bars
//! (clustered, stacked, percent-stacked), lines with markers, areas, and pies and doughnuts --
//! inside the chart's frame, with its title, legend, value axis and category labels.
//!
//! It is a stand-in, not PowerPoint's drawing: layout, fonts, trend lines and
//! effects are not reproduced, so a chart painted here counts as approximated
//! (`approximated_charts`). A chart of a kind this does not paint (bubble, stock,
//! surface) or one without data counts as not drawn (`charts`).

use super::*;

const C: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";

/// The ink the chart's text and axes are drawn in (Office's `595959`) and its gridlines (`D9D9D9`).
const INK: Rgba = Rgba {
    r: 0.349,
    g: 0.349,
    b: 0.349,
    a: 1.0,
};
const WHITE: Rgba = Rgba {
    r: 1.0,
    g: 1.0,
    b: 1.0,
    a: 1.0,
};
const GRID: Rgba = Rgba {
    r: 0.851,
    g: 0.851,
    b: 0.851,
    a: 1.0,
};

#[derive(Debug, Clone, Copy, PartialEq)]
enum Kind {
    Bar,
    Line,
    Area,
    Pie,
    Scatter,
    Radar,
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum Grouping {
    Clustered,
    Stacked,
    Percent,
}

/// What a series labels its points with, and the number format those labels take.
#[derive(Debug, Default, Clone)]
struct Labels {
    value: bool,
    percent: bool,
    category: bool,
    format: Option<String>,
}

#[derive(Debug)]
struct Series {
    name: String,
    values: Vec<Option<f64>>,
    /// A scatter series' x values, by point.
    xs: Vec<Option<f64>>,
    /// A scatter series draws a line through its points.
    line: bool,
    labels: Labels,
    /// The number format the cached values carry (`General` is none).
    source_format: Option<String>,
    color: Rgba,
    /// A pie's color per slice.
    slices: Vec<Rgba>,
    marker: bool,
}

#[derive(Debug)]
struct Plot {
    kind: Kind,
    horizontal: bool,
    grouping: Grouping,
    /// Bars: the gap between categories and the overlap of a category's bars, in percent.
    gap: f32,
    overlap: f32,
    /// Lines and radar: whether markers are drawn.
    markers: bool,
    /// Radar: the areas are filled.
    filled: bool,
    /// A doughnut's hole as a fraction of the radius; 0 for a pie.
    hole: f32,
    /// Degrees clockwise from twelve o'clock the first slice starts at.
    first_angle: f32,
    series: Vec<Series>,
}

#[derive(Debug)]
struct Chart {
    title: Option<String>,
    categories: Vec<String>,
    plots: Vec<Plot>,
    /// `r`, `l`, `t` or `b`; `None` for no legend.
    legend: Option<char>,
    min: Option<f64>,
    max: Option<f64>,
    /// The value axis's number format.
    axis_format: Option<String>,
    /// Plots on a second pair of axes, with that value axis's bounds and number format.
    secondary: Vec<Plot>,
    secondary_min: Option<f64>,
    secondary_max: Option<f64>,
    secondary_format: Option<String>,
    background: Option<Rgba>,
    /// The text sizes (points) the chart declares: axes and labels, the legend, the title.
    text_size: f32,
    legend_size: f32,
    title_size: f32,
}

/// Office's default chart text size (points) and title size.
const TEXT_SIZE: f32 = 10.0;
const TITLE_SIZE: f32 = 14.0;

/// A declared text size kept within what a frame `height` points tall can hold: never past a sixth
/// of it (but never under the default either), so a large declared size cannot leave the plot no
/// room.
fn text_points(declared: f32, height: f32) -> f32 {
    declared.clamp(6.0, 40.0).min((height / 6.0).max(TEXT_SIZE))
}

/// The font size (points) the text properties (`txPr`) or rich text under `node` declare.
fn declared_size(node: Option<Node>) -> Option<f32> {
    node?
        .descendants()
        .filter(|n| {
            n.tag_name().namespace() == Some(A) && matches!(n.tag_name().name(), "defRPr" | "rPr")
        })
        .find_map(|n| n.attribute("sz")?.parse::<f32>().ok())
        .map(|hundredths| hundredths / 100.0)
        .filter(|points| points.is_finite() && *points > 0.0)
}

fn cnode<'a, 'i>(node: Node<'a, 'i>, name: &str) -> Option<Node<'a, 'i>> {
    child(node, C, name)
}

/// The `val` attribute of the child `name`.
fn cval<'a>(node: Node<'a, '_>, name: &str) -> Option<&'a str> {
    cnode(node, name)?.attribute("val")
}

/// The points of a data reference (`numRef`/`strRef`/`numLit`/`strLit`), by index.
fn points(node: Node) -> Vec<Option<String>> {
    let Some(cache) = node.descendants().find(|n| {
        n.is_element()
            && matches!(
                n.tag_name().name(),
                "numCache" | "strCache" | "numLit" | "strLit"
            )
    }) else {
        return Vec::new();
    };
    let declared: usize = cval(cache, "ptCount")
        .and_then(|v| v.parse().ok())
        .unwrap_or(0);
    let mut out = vec![None; declared.min(100_000)];
    for pt in cache.children().filter(|n| n.has_tag_name((C, "pt"))) {
        let Some(index) = pt.attribute("idx").and_then(|v| v.parse::<usize>().ok()) else {
            continue;
        };
        if index >= 100_000 {
            continue;
        }
        if index >= out.len() {
            out.resize(index + 1, None);
        }
        out[index] = cnode(pt, "v")
            .and_then(|v| v.text())
            .map(|t| t.trim().to_string());
    }
    out
}

/// The ink that reads on `background`: white on a dark one, the chart's ink on a light one.
fn contrast(background: Rgba) -> Rgba {
    if background.r * 0.299 + background.g * 0.587 + background.b * 0.114 < 0.5 {
        WHITE
    } else {
        INK
    }
}

impl Labels {
    /// The text of a value label.
    fn render(&self, value: f64) -> String {
        match &self.format {
            Some(code) => format_code(value, code),
            None => format_value(value, false),
        }
    }
}

/// The labels of a `c:dLbls` element; a format that follows the source cells is the series'.
fn read_labels(node: Node, source: &Option<String>) -> Labels {
    if cval(node, "delete") == Some("1") {
        return Labels::default();
    }
    let flag = |name: &str| cval(node, name) == Some("1");
    let explicit = cnode(node, "numFmt").and_then(|n| {
        (n.attribute("sourceLinked") != Some("1"))
            .then(|| n.attribute("formatCode").map(str::to_string))
            .flatten()
    });
    Labels {
        value: flag("showVal"),
        percent: flag("showPercent"),
        category: flag("showCatName"),
        format: explicit.or_else(|| source.clone()),
    }
}

/// `value` written the way the number format `code` asks for: digits after the point, a
/// thousands separator, a percent scale, and the literal text around the number. Anything else a
/// format can say (conditions, dates, exponents) is left out.
fn format_code(value: f64, code: &str) -> String {
    let section = code.split(';').next().unwrap_or("").trim();
    if section.is_empty() || section.eq_ignore_ascii_case("general") {
        return format_value(value, false);
    }
    let (mut prefix, mut suffix) = (String::new(), String::new());
    let (mut percent, mut thousands, mut point, mut seen) = (false, false, false, false);
    let mut decimals = 0usize;
    let mut chars = section.chars();
    while let Some(c) = chars.next() {
        let literal: Option<String> = match c {
            '"' => Some(chars.by_ref().take_while(|c| *c != '"').collect()),
            '\\' => chars.next().map(String::from),
            '_' | '*' => {
                chars.next();
                None
            }
            '[' => {
                chars.by_ref().find(|c| *c == ']');
                None
            }
            '0' | '#' | '?' => {
                seen = true;
                if point {
                    decimals += 1;
                }
                None
            }
            '.' if seen || chars.clone().next().is_some_and(|n| matches!(n, '0' | '#')) => {
                point = true;
                None
            }
            ',' if seen => {
                thousands = true;
                None
            }
            '%' => {
                percent = true;
                None
            }
            c if c.is_ascii_alphabetic() => None,
            other => Some(other.to_string()),
        };
        if let Some(text) = literal {
            if seen {
                suffix.push_str(&text);
            } else {
                prefix.push_str(&text);
            }
        }
    }
    let scaled = if percent { value * 100.0 } else { value };
    let digits = format!("{:.*}", decimals, scaled.abs());
    let (whole, fraction) = digits.split_once('.').unwrap_or((&digits, ""));
    let whole = if thousands {
        let mut out = String::new();
        for (i, ch) in whole.chars().enumerate() {
            if i > 0 && (whole.len() - i) % 3 == 0 {
                out.push(',');
            }
            out.push(ch);
        }
        out
    } else {
        whole.to_string()
    };
    let sign = if scaled < 0.0 && digits.chars().any(|c| c.is_ascii_digit() && c != '0') {
        "-"
    } else {
        ""
    };
    let fraction = if fraction.is_empty() {
        String::new()
    } else {
        format!(".{fraction}")
    };
    format!(
        "{sign}{prefix}{whole}{fraction}{}{suffix}",
        if percent { "%" } else { "" }
    )
}

/// The color Office gives the `index`-th series or slice of a chart whose style names none: the
/// theme's six accents, then the same accents darker and lighter.
fn auto_color(colors: &Colors, index: usize) -> Rgba {
    let accent = colors
        .scheme(&format!("accent{}", index % 6 + 1))
        .unwrap_or(Rgba {
            r: 0.5,
            g: 0.5,
            b: 0.5,
            a: 1.0,
        });
    match index / 6 {
        0 => accent,
        round if round % 2 == 1 => shaded(accent, -0.3),
        _ => shaded(accent, 0.4),
    }
}

/// The fill a series element names in `c:spPr`: its solid fill, or for a line its line's.
fn explicit_color(sp_pr: Option<Node>, kind: Kind, colors: &Colors) -> Option<Rgba> {
    let sp_pr = sp_pr?;
    let fill = if matches!(kind, Kind::Line | Kind::Scatter) {
        child(child(sp_pr, A, "ln")?, A, "solidFill")?
    } else {
        child(sp_pr, A, "solidFill")?
    };
    color_child(fill, colors, None)
}

fn read_chart(doc: &roxmltree::Document, colors: &Colors) -> Option<Chart> {
    let space = doc.root_element();
    let chart = cnode(space, "chart")?;
    let area = cnode(chart, "plotArea")?;
    let mut plots: Vec<Plot> = Vec::new();
    let mut categories: Vec<String> = Vec::new();
    let mut axes: Option<Vec<&str>> = None;
    let mut second_axes: Option<Vec<&str>> = None;
    let mut secondary_plots: Vec<Plot> = Vec::new();
    let mut index = 0usize;
    for node in area.children().filter(|n| n.is_element()) {
        let kind = match node.tag_name().name() {
            "barChart" | "bar3DChart" => Kind::Bar,
            "lineChart" | "line3DChart" => Kind::Line,
            "areaChart" | "area3DChart" => Kind::Area,
            "pieChart" | "pie3DChart" | "doughnutChart" => Kind::Pie,
            "scatterChart" => Kind::Scatter,
            "radarChart" => Kind::Radar,
            _ => continue,
        };
        // Plots on a second pair of axes are not painted: their scale is another one.
        let ids: Vec<&str> = node
            .children()
            .filter(|c| c.has_tag_name((C, "axId")))
            .filter_map(|c| c.attribute("val"))
            .collect();
        let second = match &axes {
            None => {
                axes = Some(ids);
                false
            }
            Some(first) if *first == ids => false,
            Some(_) => match &second_axes {
                None => {
                    second_axes = Some(ids);
                    true
                }
                Some(s) if *s == ids => true,
                // A third pair of axes is not painted.
                Some(_) => continue,
            },
        };
        let mut series = Vec::new();
        for ser in node.children().filter(|c| c.has_tag_name((C, "ser"))) {
            let value_node = cnode(ser, if kind == Kind::Scatter { "yVal" } else { "val" });
            let values: Vec<Option<f64>> = value_node
                .map(points)
                .unwrap_or_default()
                .into_iter()
                .map(|v| {
                    v.and_then(|t| t.parse::<f64>().ok())
                        .filter(|v| v.is_finite())
                })
                .collect();
            if categories.is_empty() {
                if let Some(cat) = cnode(ser, "cat") {
                    categories = points(cat)
                        .into_iter()
                        .map(Option::unwrap_or_default)
                        .collect();
                }
            }
            let name = cnode(ser, "tx")
                .and_then(|tx| {
                    points(tx)
                        .into_iter()
                        .next()
                        .flatten()
                        .or_else(|| cnode(tx, "v")?.text().map(|t| t.trim().to_string()))
                })
                .filter(|n| !n.is_empty())
                .unwrap_or_else(|| format!("Series {}", index + 1));
            let sp_pr = cnode(ser, "spPr");
            let color = explicit_color(sp_pr, kind, colors)
                .or_else(|| {
                    cnode(ser, "marker")
                        .and_then(|m| explicit_color(cnode(m, "spPr"), Kind::Area, colors))
                })
                .unwrap_or_else(|| auto_color(colors, if kind == Kind::Pie { 0 } else { index }));
            let slices = if kind == Kind::Pie {
                (0..values.len())
                    .map(|i| {
                        ser.children()
                            .filter(|d| d.has_tag_name((C, "dPt")))
                            .find(|d| cval(*d, "idx").and_then(|v| v.parse().ok()) == Some(i))
                            .and_then(|d| explicit_color(cnode(d, "spPr"), kind, colors))
                            .unwrap_or_else(|| auto_color(colors, i))
                    })
                    .collect()
            } else {
                Vec::new()
            };
            let marker = cnode(ser, "marker")
                .and_then(|m| cval(m, "symbol"))
                .is_none_or(|s| s != "none");
            let xs: Vec<Option<f64>> = if kind == Kind::Scatter {
                cnode(ser, "xVal")
                    .map(points)
                    .unwrap_or_default()
                    .into_iter()
                    .map(|v| {
                        v.and_then(|t| t.parse::<f64>().ok())
                            .filter(|v| v.is_finite())
                    })
                    .collect()
            } else {
                Vec::new()
            };
            let ln_hidden = cnode(ser, "spPr")
                .and_then(|s| child(s, A, "ln"))
                .is_some_and(|ln| child(ln, A, "noFill").is_some());
            let line = kind == Kind::Scatter
                && !ln_hidden
                && cval(node, "scatterStyle")
                    .is_some_and(|s| s.contains("ine") || s.contains("mooth"));
            let source_format = value_node
                .and_then(|v| {
                    v.descendants()
                        .find(|n| n.has_tag_name((C, "formatCode")))
                        .and_then(|n| n.text())
                })
                .map(|t| t.trim().to_string())
                .filter(|t| !t.is_empty() && !t.eq_ignore_ascii_case("general"));
            let labels = cnode(ser, "dLbls")
                .or_else(|| cnode(node, "dLbls"))
                .map(|l| read_labels(l, &source_format))
                .unwrap_or_default();
            series.push(Series {
                name,
                values,
                xs,
                line,
                labels,
                source_format,
                color,
                slices,
                marker,
            });
            index += 1;
        }
        let number = |name: &str, default: f32| {
            cval(node, name)
                .and_then(|v| v.parse::<f32>().ok())
                .unwrap_or(default)
        };
        (if second {
            &mut secondary_plots
        } else {
            &mut plots
        })
        .push(Plot {
            kind,
            horizontal: cval(node, "barDir") == Some("bar"),
            grouping: match cval(node, "grouping") {
                Some("stacked") => Grouping::Stacked,
                Some("percentStacked") => Grouping::Percent,
                _ => Grouping::Clustered,
            },
            gap: number("gapWidth", 150.0),
            overlap: number("overlap", 0.0),
            markers: if kind == Kind::Radar {
                cval(node, "radarStyle") == Some("marker")
            } else {
                cval(node, "marker") != Some("0")
            },
            filled: kind == Kind::Radar && cval(node, "radarStyle") == Some("filled"),
            hole: if node.tag_name().name() == "doughnutChart" {
                number("holeSize", 50.0) / 100.0
            } else {
                0.0
            },
            first_angle: number("firstSliceAng", 0.0),
            series,
        });
    }
    if plots.iter().all(|p| p.series.is_empty()) {
        return None;
    }
    let first_series_name = plots
        .iter()
        .flat_map(|p| p.series.iter())
        .map(|s| s.name.clone())
        .collect::<Vec<_>>();
    let title = cnode(chart, "title").and_then(|t| {
        let text: String = t
            .descendants()
            .filter(|n| n.has_tag_name((A, "t")))
            .filter_map(|n| n.text())
            .collect();
        let text = text.trim().to_string();
        if !text.is_empty() {
            Some(text)
        } else if cval(chart, "autoTitleDeleted") != Some("1") && first_series_name.len() == 1 {
            // An automatic title is the name of the one series.
            first_series_name.first().cloned()
        } else {
            None
        }
    });
    let legend = cnode(chart, "legend")
        .filter(|l| cval(*l, "delete") != Some("1"))
        .map(|l| match cval(l, "legendPos") {
            Some(p) if p.starts_with('b') => 'b',
            Some(p) if p.starts_with('t') && p != "tr" => 't',
            Some("l") => 'l',
            _ => 'r',
        });
    let scaling = area
        .children()
        .find(|n| n.has_tag_name((C, "valAx")))
        .and_then(|ax| cnode(ax, "scaling"));
    let bound = |name: &str| {
        scaling
            .and_then(|s| cval(s, name))
            .and_then(|v| v.parse::<f64>().ok())
            .filter(|v| v.is_finite())
    };
    let axis_format = area
        .children()
        .find(|n| n.has_tag_name((C, "valAx")))
        .and_then(|ax| cnode(ax, "numFmt"))
        .filter(|n| n.attribute("sourceLinked") != Some("1"))
        .and_then(|n| n.attribute("formatCode").map(str::to_string))
        .or_else(|| {
            plots
                .iter()
                .flat_map(|p| p.series.iter())
                .find_map(|s| s.source_format.clone())
        });
    let second_ax = second_axes.as_ref().and_then(|ids| {
        area.children()
            .filter(|n| n.has_tag_name((C, "valAx")))
            .find(|ax| cval(*ax, "axId").is_some_and(|id| ids.contains(&id)))
    });
    let second_bound = |name: &str| {
        second_ax
            .and_then(|ax| cnode(ax, "scaling"))
            .and_then(|s| cval(s, name))
            .and_then(|v| v.parse::<f64>().ok())
            .filter(|v| v.is_finite())
    };
    let secondary_format = second_ax
        .and_then(|ax| cnode(ax, "numFmt"))
        .filter(|n| n.attribute("sourceLinked") != Some("1"))
        .and_then(|n| n.attribute("formatCode").map(str::to_string))
        .or_else(|| {
            secondary_plots
                .iter()
                .flat_map(|p| p.series.iter())
                .find_map(|s| s.source_format.clone())
        });
    let background = cnode(space, "spPr")
        .and_then(|p| child(p, A, "solidFill"))
        .and_then(|f| color_child(f, colors, None));
    Some(Chart {
        title,
        categories,
        plots,
        legend,
        min: bound("min"),
        max: bound("max"),
        axis_format,
        secondary_min: second_bound("min"),
        secondary_max: second_bound("max"),
        secondary_format,
        secondary: secondary_plots,
        background,
        text_size: declared_size(
            area.children()
                .find(|n| n.has_tag_name((C, "valAx")))
                .and_then(|ax| cnode(ax, "txPr")),
        )
        .or_else(|| {
            declared_size(
                area.children()
                    .find(|n| n.has_tag_name((C, "catAx")))
                    .and_then(|ax| cnode(ax, "txPr")),
            )
        })
        .or_else(|| declared_size(cnode(space, "txPr")))
        .unwrap_or(TEXT_SIZE),
        legend_size: declared_size(cnode(chart, "legend").and_then(|l| cnode(l, "txPr")))
            .or_else(|| declared_size(cnode(space, "txPr")))
            .unwrap_or(TEXT_SIZE),
        title_size: declared_size(cnode(chart, "title")).unwrap_or(TITLE_SIZE),
    })
}

/// The lowest and highest value the `plots` reach over `n` categories, stacked ones by their
/// running totals; zero is always inside.
fn extent(plots: &[Plot], n: usize) -> (f64, f64) {
    let (mut lo, mut hi) = (0f64, 0f64);
    for plot in plots {
        for i in 0..n {
            let value = |s: &Series| s.values.get(i).copied().flatten();
            if plot.kind != Kind::Line && plot.grouping != Grouping::Clustered {
                let (pos, neg) =
                    plot.series
                        .iter()
                        .filter_map(value)
                        .fold(
                            (0.0, 0.0),
                            |(p, n), v| {
                                if v >= 0.0 {
                                    (p + v, n)
                                } else {
                                    (p, n + v)
                                }
                            },
                        );
                hi = hi.max(pos);
                lo = lo.min(neg);
            } else {
                for v in plot.series.iter().filter_map(value) {
                    hi = hi.max(v);
                    lo = lo.min(v);
                }
            }
        }
    }
    (lo, hi)
}

/// A value axis from `lo` to `hi`: its ends and the ticks between, on 1, 2 or 5 times a power
/// of ten, honoring a bound the chart fixes.
fn scale(lo: f64, hi: f64, fixed_min: Option<f64>, fixed_max: Option<f64>) -> (f64, f64, Vec<f64>) {
    scale_from(lo, hi, fixed_min, fixed_max, true)
}

/// [`scale`], with the axis starting at the lowest value instead of at zero when `from_zero` is
/// false.
fn scale_from(
    lo: f64,
    hi: f64,
    fixed_min: Option<f64>,
    fixed_max: Option<f64>,
    from_zero: bool,
) -> (f64, f64, Vec<f64>) {
    let mut lo = fixed_min.unwrap_or(if from_zero { lo.min(0.0) } else { lo });
    let mut hi = fixed_max.unwrap_or(if from_zero { hi.max(0.0) } else { hi });
    if hi <= lo {
        hi = lo + 1.0;
    }
    let raw = (hi - lo) / 5.0;
    let magnitude = 10f64.powf(raw.log10().floor());
    let step = [1.0, 2.0, 5.0, 10.0]
        .into_iter()
        .map(|m| m * magnitude)
        .find(|s| *s >= raw)
        .unwrap_or(10.0 * magnitude);
    if fixed_min.is_none() {
        lo = (lo / step).floor() * step;
    }
    if fixed_max.is_none() {
        hi = (hi / step).ceil() * step;
    }
    let mut ticks = Vec::new();
    let mut tick = (lo / step).ceil() * step;
    while tick <= hi + step * 1e-9 && ticks.len() < 50 {
        ticks.push(if tick.abs() < step * 1e-9 { 0.0 } else { tick });
        tick += step;
    }
    (lo, hi, ticks)
}

fn format_value(v: f64, percent: bool) -> String {
    if percent {
        return format!("{}%", (v * 100.0).round());
    }
    if v.fract().abs() < 1e-9 {
        return format!("{}", v.round() as i64);
    }
    let s = format!("{v:.2}");
    s.trim_end_matches('0').trim_end_matches('.').to_string()
}

#[derive(Clone, Copy, PartialEq)]
enum Align {
    Left,
    Center,
    Right,
}

impl<'a> Painter<'a> {
    /// Paint the chart of the graphic frame `frame`.
    pub(super) fn chart(&mut self, frame: Node, parent: Transform) {
        let place = child(frame, P, "xfrm").and_then(|x| {
            let (off, ext) = (child(x, A, "off")?, child(x, A, "ext")?);
            Some((
                num(off, "x")?,
                num(off, "y")?,
                num(ext, "cx")?,
                num(ext, "cy")?,
            ))
        });
        let part = (|| {
            let id = frame
                .descendants()
                .find(|n| n.has_tag_name((C, "chart")))?
                .attribute((R, "id"))?;
            self.rels_of(frame)?.get(id).cloned()
        })();
        let text = part.and_then(|p| self.container.read_xml(&p).ok());
        let chart = text
            .as_deref()
            .and_then(|t| roxmltree::Document::parse(t).ok())
            .and_then(|doc| read_chart(&doc, self.colors));
        let (Some((x, y, w, h)), Some(chart)) = (place, chart) else {
            self.gaps.charts += 1;
            return;
        };
        if w <= 0.0 || h <= 0.0 {
            self.gaps.charts += 1;
            return;
        }
        let at = Transform::from_translate(x as f32, y as f32).post_concat(parent);
        self.draw_chart(&chart, w as f32, h as f32, at);
        self.gaps.approximated_charts += 1;
    }

    fn rect(&mut self, x: f32, y: f32, w: f32, h: f32, color: Rgba, t: Transform) {
        if let Some(rect) = tiny_skia::Rect::from_xywh(x, y, w.max(0.0), h.max(0.0)) {
            if w > 0.0 && h > 0.0 {
                let mut paint = Paint::default();
                paint.set_color(color.to_color());
                paint.anti_alias = true;
                self.pixmap.fill_rect(rect, &paint, t, None);
            }
        }
    }

    fn line(&mut self, from: (f32, f32), to: (f32, f32), width: f32, color: Rgba, t: Transform) {
        let mut pb = PathBuilder::new();
        pb.move_to(from.0, from.1);
        pb.line_to(to.0, to.1);
        if let Some(path) = pb.finish() {
            self.stroke(&path, width, color, t);
        }
    }

    fn stroke(&mut self, path: &tiny_skia::Path, width: f32, color: Rgba, t: Transform) {
        let mut paint = Paint::default();
        paint.set_color(color.to_color());
        paint.anti_alias = true;
        let stroke = Stroke {
            width,
            line_cap: tiny_skia::LineCap::Round,
            line_join: tiny_skia::LineJoin::Round,
            ..Stroke::default()
        };
        self.pixmap.stroke_path(path, &paint, &stroke, t, None);
    }

    fn fill_with(&mut self, path: &tiny_skia::Path, color: Rgba, t: Transform) {
        let mut paint = Paint::default();
        paint.set_color(color.to_color());
        paint.anti_alias = true;
        self.pixmap
            .fill_path(path, &paint, FillRule::Winding, t, None);
    }

    /// The width of `text` at `size` (EMU) in the faces available.
    fn text_width(&mut self, text: &str, size: f32) -> f32 {
        text.chars()
            .map(|ch| match self.fonts.pick(&[], false, false, ch) {
                Some((face, _)) => self.fonts.advance(face, ch, size),
                None => size * 0.5,
            })
            .sum()
    }

    /// Paint `text` with its top at `top`, aligned on `x`; a text no face can draw counts as not
    /// painted.
    #[allow(clippy::too_many_arguments)]
    fn label(
        &mut self,
        text: &str,
        x: f32,
        top: f32,
        size: f32,
        color: Rgba,
        align: Align,
        t: Transform,
    ) {
        let width = self.text_width(text, size);
        let mut pen = match align {
            Align::Left => x,
            Align::Center => x - width / 2.0,
            Align::Right => x - width,
        };
        let mut missing = false;
        let mut baseline = None;
        for ch in text.chars() {
            let Some((face, _)) = self.fonts.pick(&[], false, false, ch) else {
                missing |= !ch.is_whitespace();
                pen += size * 0.5;
                continue;
            };
            let base = *baseline.get_or_insert_with(|| top + self.fonts.ascent(face, size));
            if !ch.is_whitespace() {
                if let Some(path) = self.fonts.glyph(face, ch, size, pen, base) {
                    self.fill_with(&path, color, t);
                }
            }
            pen += self.fonts.advance(face, ch, size);
        }
        if missing {
            self.gaps.text_runs += 1;
        }
    }

    fn draw_chart(&mut self, chart: &Chart, w: f32, h: f32, t: Transform) {
        let pt = EMU_PER_PT;
        let fit = |points: f32| text_points(points, h / pt) * pt;
        let size = fit(chart.text_size);
        let lsize = fit(chart.legend_size);
        let label_size = size * 0.9;
        let title_points = chart
            .title_size
            .clamp(6.0, 60.0)
            .min((h / pt / 3.0).max(TITLE_SIZE));
        let pad = 6.0 * pt;
        if let Some(bg) = chart.background {
            self.rect(0.0, 0.0, w, h, bg, t);
        }
        let (mut left, mut top, mut right, mut bottom) = (pad, pad, w - pad, h - pad);
        if let Some(title) = &chart.title {
            // A title wider than the frame shrinks to fit, down to 60%.
            let wide = self.text_width(title, title_points * pt).max(1.0);
            let title_size = title_points * pt * ((w - 2.0 * pad) / wide).clamp(0.6, 1.0);
            self.label(title, w / 2.0, top, title_size, INK, Align::Center, t);
            top += title_size * 1.6;
        }
        let first = &chart.plots[0];
        let pie = first.kind == Kind::Pie;

        // Legend: the series' colors, or a pie's slices.
        let entries: Vec<(String, Rgba)> = if pie {
            first
                .series
                .first()
                .map(|s| {
                    s.slices
                        .iter()
                        .enumerate()
                        .map(|(i, c)| (chart.categories.get(i).cloned().unwrap_or_default(), *c))
                        .collect()
                })
                .unwrap_or_default()
        } else {
            chart
                .plots
                .iter()
                .chain(&chart.secondary)
                .flat_map(|p| p.series.iter())
                .map(|s| (s.name.clone(), s.color))
                .collect()
        };
        if let (Some(side), false) = (chart.legend, entries.is_empty()) {
            let swatch = lsize * 0.7;
            let gap = lsize * 0.4;
            let widths: Vec<f32> = entries
                .iter()
                .map(|(n, _)| self.text_width(n, lsize))
                .collect();
            if side == 'r' || side == 'l' {
                let rows = entries.len() as f32 * lsize * 1.5;
                let lw = swatch + gap + widths.iter().cloned().fold(0.0, f32::max);
                let x0 = if side == 'r' { right - lw } else { left };
                let y0 = top + ((bottom - top) - rows).max(0.0) / 2.0;
                for (i, (name, color)) in entries.iter().enumerate() {
                    let y = y0 + i as f32 * lsize * 1.5;
                    self.rect(x0, y + lsize * 0.15, swatch, swatch, *color, t);
                    self.label(name, x0 + swatch + gap, y, lsize, INK, Align::Left, t);
                }
                if side == 'r' {
                    right -= lw + pad;
                } else {
                    left += lw + pad;
                }
            } else {
                let total: f32 = widths.iter().map(|w| swatch + gap + w + lsize * 1.2).sum();
                let mut x = ((left + right) - total).max(0.0) / 2.0 + left;
                let y = if side == 'b' {
                    bottom - lsize * 1.3
                } else {
                    top
                };
                for ((name, color), width) in entries.iter().zip(&widths) {
                    self.rect(x, y + lsize * 0.15, swatch, swatch, *color, t);
                    self.label(name, x + swatch + gap, y, lsize, INK, Align::Left, t);
                    x += swatch + gap + width + lsize * 1.2;
                }
                if side == 'b' {
                    bottom -= lsize * 1.8;
                } else {
                    top += lsize * 1.8;
                }
            }
        }

        if pie {
            self.draw_pie(
                first,
                &chart.categories,
                (left, top, right, bottom),
                label_size,
                t,
            );
            return;
        }
        if first.kind == Kind::Scatter {
            self.draw_scatter(chart, (left, top, right, bottom), size, t);
            return;
        }
        if first.kind == Kind::Radar {
            self.draw_radar(chart, (left, top, right, bottom), size, t);
            return;
        }

        // The value range over every plot, stacked ones by their running totals.
        let n = chart
            .categories
            .len()
            .max(
                chart
                    .plots
                    .iter()
                    .chain(&chart.secondary)
                    .flat_map(|p| p.series.iter())
                    .map(|s| s.values.len())
                    .max()
                    .unwrap_or(0),
            )
            .max(1);
        let percent = chart
            .plots
            .iter()
            .any(|p| p.kind != Kind::Line && p.grouping == Grouping::Percent);
        let (lo, hi) = extent(&chart.plots, n);
        let (min, max, ticks) = if percent {
            (0.0, 1.0, vec![0.0, 0.2, 0.4, 0.6, 0.8, 1.0])
        } else {
            scale(lo, hi, chart.min, chart.max)
        };
        let horizontal = first.kind == Kind::Bar && first.horizontal;

        let tick_labels: Vec<String> = ticks
            .iter()
            .map(|v| match (&chart.axis_format, percent) {
                (Some(code), false) => format_code(*v, code),
                _ => format_value(*v, percent),
            })
            .collect();
        let cat_labels: Vec<String> = (0..n)
            .map(|i| chart.categories.get(i).cloned().unwrap_or_default())
            .collect();
        // A second value axis stands on the right of a vertical chart.
        let second = if chart.secondary.is_empty() || horizontal {
            None
        } else {
            let (lo2, hi2) = extent(&chart.secondary, n);
            let (min2, max2, ticks2) = scale(lo2, hi2, chart.secondary_min, chart.secondary_max);
            let labels2: Vec<String> = ticks2
                .iter()
                .map(|v| match &chart.secondary_format {
                    Some(code) => format_code(*v, code),
                    None => format_value(*v, false),
                })
                .collect();
            Some((min2, max2, ticks2, labels2))
        };
        let tick_width = tick_labels
            .iter()
            .map(|s| self.text_width(s, size))
            .fold(0.0, f32::max);
        let cat_width = cat_labels
            .iter()
            .map(|s| self.text_width(s, size))
            .fold(0.0, f32::max);
        if horizontal {
            left += cat_width + size * 0.6;
            bottom -= size * 1.6;
        } else {
            left += tick_width + size * 0.6;
            bottom -= size * 1.6;
        }
        let second_width = second
            .as_ref()
            .map(|(_, _, _, labels)| {
                labels
                    .iter()
                    .map(|s| self.text_width(s, size))
                    .fold(0.0, f32::max)
                    + size * 0.6
            })
            .unwrap_or(0.0);
        right -= second_width;
        if right - left < size || bottom - top < size {
            return;
        }
        let span = (max - min).max(f64::MIN_POSITIVE);
        let (pw, ph) = (right - left, bottom - top);
        // Position along the value axis, and a category position (in categories) along the other.
        let val = move |v: f64| -> f32 {
            let f = ((v - min) / span) as f32;
            if horizontal {
                left + f * pw
            } else {
                bottom - f * ph
            }
        };
        let cat = move |c: f32| -> f32 {
            if horizontal {
                bottom - c / n as f32 * ph
            } else {
                left + c / n as f32 * pw
            }
        };
        let cell = (if horizontal { ph } else { pw }) / n as f32;
        let (min2, span2) = second
            .as_ref()
            .map(|(lo2, hi2, _, _)| (*lo2, (hi2 - lo2).max(f64::MIN_POSITIVE)))
            .unwrap_or((0.0, 1.0));
        let val2 = move |v: f64| -> f32 { bottom - ((v - min2) / span2) as f32 * ph };
        // The second axis's labels, at its own ticks.
        if let Some((_, _, ticks2, labels2)) = &second {
            for (tick, text) in ticks2.iter().zip(labels2) {
                self.label(
                    text,
                    right + size * 0.4,
                    val2(*tick) - size * 0.55,
                    size,
                    INK,
                    Align::Left,
                    t,
                );
            }
        }

        // Gridlines and the value labels.
        for (tick, text) in ticks.iter().zip(&tick_labels) {
            let p = val(*tick);
            if horizontal {
                self.line((p, top), (p, bottom), pt * 0.75, GRID, t);
                self.label(text, p, bottom + size * 0.3, size, INK, Align::Center, t);
            } else {
                self.line((left, p), (right, p), pt * 0.75, GRID, t);
                self.label(
                    text,
                    left - size * 0.4,
                    p - size * 0.55,
                    size,
                    INK,
                    Align::Right,
                    t,
                );
            }
        }
        // Category labels: every k-th when they would run into each other.
        let every = if horizontal {
            1
        } else {
            (cat_width / (cell * 0.95).max(1.0)).ceil().max(1.0) as usize
        };
        for (i, name) in cat_labels.iter().enumerate().step_by(every) {
            let mid = i as f32 + 0.5;
            if horizontal {
                self.label(
                    name,
                    left - size * 0.4,
                    cat(mid) - size * 0.55,
                    size,
                    INK,
                    Align::Right,
                    t,
                );
            } else {
                self.label(
                    name,
                    cat(mid),
                    bottom + size * 0.3,
                    size,
                    INK,
                    Align::Center,
                    t,
                );
            }
        }

        // The plots, then the axis line at zero.
        for plot in &chart.plots {
            match plot.kind {
                Kind::Bar => {
                    self.draw_bars(plot, n, percent, &val, &cat, horizontal, label_size, t)
                }
                Kind::Line | Kind::Area => {
                    self.draw_lines(plot, n, percent, &val, &cat, label_size, t)
                }
                Kind::Pie | Kind::Scatter | Kind::Radar => {}
            }
        }
        for plot in chart.secondary.iter().filter(|_| !horizontal) {
            match plot.kind {
                Kind::Bar => self.draw_bars(plot, n, false, &val2, &cat, false, label_size, t),
                Kind::Line | Kind::Area => {
                    self.draw_lines(plot, n, false, &val2, &cat, label_size, t)
                }
                Kind::Pie | Kind::Scatter | Kind::Radar => {}
            }
        }
        let zero = val(0.0f64.clamp(min, max));
        if horizontal {
            self.line((zero, top), (zero, bottom), pt, INK, t);
        } else {
            self.line((left, zero), (right, zero), pt, INK, t);
        }
    }

    #[allow(clippy::too_many_arguments)]
    fn draw_bars(
        &mut self,
        plot: &Plot,
        n: usize,
        percent: bool,
        val: &dyn Fn(f64) -> f32,
        cat: &dyn Fn(f32) -> f32,
        horizontal: bool,
        size: f32,
        t: Transform,
    ) {
        let stacked = plot.grouping != Grouping::Clustered;
        let slots = (if stacked { 1 } else { plot.series.len().max(1) }) as f32;
        let overlap = if stacked {
            1.0
        } else {
            (plot.overlap / 100.0).clamp(-1.0, 1.0)
        };
        let gap = (plot.gap / 100.0).max(0.0);
        let width = 1.0 / (slots - (slots - 1.0) * overlap + gap);
        let group = width * (slots - (slots - 1.0) * overlap);
        // Labels go on once every bar is down, so no later bar covers one.
        let mut labels: Vec<(String, f32, f32, Align, Rgba)> = Vec::new();
        for i in 0..n {
            let total: f64 = plot
                .series
                .iter()
                .filter_map(|s| s.values.get(i).copied().flatten())
                .map(f64::abs)
                .sum();
            let (mut pos, mut neg) = (0.0f64, 0.0f64);
            for (j, series) in plot.series.iter().enumerate() {
                let Some(raw) = series.values.get(i).copied().flatten() else {
                    continue;
                };
                let mut v = raw;
                if percent && total > 0.0 {
                    v /= total;
                }
                let (from, to) = if stacked {
                    if v >= 0.0 {
                        pos += v;
                        (pos - v, pos)
                    } else {
                        neg += v;
                        (neg - v, neg)
                    }
                } else {
                    (0.0, v)
                };
                let slot = if stacked { 0.0 } else { j as f32 };
                let c0 = i as f32 + (1.0 - group) / 2.0 + slot * width * (1.0 - overlap);
                let c1 = c0 + width;
                let (a, b) = (cat(c0), cat(c1));
                let (p, q) = (val(from), val(to));
                if horizontal {
                    self.rect(
                        p.min(q),
                        a.min(b),
                        (q - p).abs(),
                        (b - a).abs(),
                        series.color,
                        t,
                    );
                } else {
                    self.rect(
                        a.min(b),
                        p.min(q),
                        (b - a).abs(),
                        (q - p).abs(),
                        series.color,
                        t,
                    );
                }
                if series.labels.value {
                    let text = series.labels.render(raw);
                    let (mid_c, lo, hi) = ((a + b) / 2.0, p.min(q), p.max(q));
                    // Beside a bar the label is in ink; inside a stacked segment it contrasts
                    // with the segment.
                    let inside = contrast(series.color);
                    labels.push(match (horizontal, stacked) {
                        (false, false) if v >= 0.0 => {
                            (text, mid_c, lo - size * 1.4, Align::Center, INK)
                        }
                        (false, false) => (text, mid_c, hi + size * 0.2, Align::Center, INK),
                        (true, false) if v >= 0.0 => {
                            (text, hi + size * 0.3, mid_c - size * 0.55, Align::Left, INK)
                        }
                        (true, false) => (
                            text,
                            lo - size * 0.3,
                            mid_c - size * 0.55,
                            Align::Right,
                            INK,
                        ),
                        (false, true) => (
                            text,
                            mid_c,
                            (lo + hi) / 2.0 - size * 0.55,
                            Align::Center,
                            inside,
                        ),
                        (true, true) => (
                            text,
                            (lo + hi) / 2.0,
                            mid_c - size * 0.55,
                            Align::Center,
                            inside,
                        ),
                    });
                }
            }
        }
        for (text, x, y, align, ink) in labels {
            self.label(&text, x, y, size, ink, align, t);
        }
    }

    #[allow(clippy::too_many_arguments)]
    fn draw_lines(
        &mut self,
        plot: &Plot,
        n: usize,
        percent: bool,
        val: &dyn Fn(f64) -> f32,
        cat: &dyn Fn(f32) -> f32,
        size: f32,
        t: Transform,
    ) {
        let area = plot.kind == Kind::Area;
        let stacked = area && plot.grouping != Grouping::Clustered;
        let mut running = vec![0.0f64; n];
        for series in &plot.series {
            let totals: Vec<f64> = (0..n)
                .map(|i| {
                    plot.series
                        .iter()
                        .filter_map(|s| s.values.get(i).copied().flatten())
                        .map(f64::abs)
                        .sum()
                })
                .collect();
            let mut runs: Vec<Vec<(f32, f32, f32, f64)>> = vec![Vec::new()];
            for i in 0..n {
                match series.values.get(i).copied().flatten() {
                    Some(raw) => {
                        let mut v = raw;
                        if percent && totals[i] > 0.0 {
                            v /= totals[i];
                        }
                        let base = if stacked { running[i] } else { 0.0 };
                        if stacked {
                            running[i] += v;
                        }
                        let x = cat(i as f32 + 0.5);
                        runs.last_mut()
                            .unwrap()
                            .push((x, val(base + v), val(base), raw));
                    }
                    None => runs.push(Vec::new()),
                }
            }
            for run in runs.iter().filter(|r| !r.is_empty()) {
                if area {
                    let mut pb = PathBuilder::new();
                    pb.move_to(run[0].0, run[0].2);
                    for p in run {
                        pb.line_to(p.0, p.1);
                    }
                    pb.line_to(run[run.len() - 1].0, run[run.len() - 1].2);
                    pb.close();
                    if let Some(path) = pb.finish() {
                        self.fill_with(&path, series.color, t);
                    }
                } else if run.len() > 1 {
                    let mut pb = PathBuilder::new();
                    pb.move_to(run[0].0, run[0].1);
                    for p in &run[1..] {
                        pb.line_to(p.0, p.1);
                    }
                    if let Some(path) = pb.finish() {
                        self.stroke(&path, 2.25 * EMU_PER_PT, series.color, t);
                    }
                }
                if !area && plot.markers && series.marker {
                    for p in run {
                        if let Some(circle) = PathBuilder::from_circle(p.0, p.1, 3.5 * EMU_PER_PT) {
                            self.fill_with(&circle, series.color, t);
                        }
                    }
                }
                if !area && series.labels.value {
                    for p in run {
                        let text = series.labels.render(p.3);
                        self.label(&text, p.0, p.1 - size * 1.9, size, INK, Align::Center, t);
                    }
                }
            }
        }
    }

    /// Radar series: one spoke per category from twelve o'clock, clockwise; a value is how far
    /// along its spoke the series' polygon reaches, on webs at the value axis's ticks.
    fn draw_radar(&mut self, chart: &Chart, rect: (f32, f32, f32, f32), size: f32, t: Transform) {
        let pt = EMU_PER_PT;
        let series: Vec<(&Plot, &Series)> = chart
            .plots
            .iter()
            .filter(|p| p.kind == Kind::Radar)
            .flat_map(|p| p.series.iter().map(move |s| (p, s)))
            .collect();
        let n = chart.categories.len().max(
            series
                .iter()
                .map(|(_, s)| s.values.len())
                .max()
                .unwrap_or(0),
        );
        if n < 3 {
            return;
        }
        let (lo, hi) = extent(&chart.plots, n);
        let (min, max, ticks) = scale(lo, hi, chart.min, chart.max);
        let span = (max - min).max(f64::MIN_POSITIVE);
        let (cx, cy) = ((rect.0 + rect.2) / 2.0, (rect.1 + rect.3) / 2.0);
        let half = (rect.2 - rect.0).min(rect.3 - rect.1) / 2.0;
        // Room for the category labels around the web, but never less than a sixth of it.
        let radius = (half - size * 1.8).max(half * 0.3);
        if radius <= 0.0 {
            return;
        }
        let angle = |i: usize| i as f32 / n as f32 * std::f32::consts::TAU;
        let at = |i: usize, r: f32| (cx + r * angle(i).sin(), cy - r * angle(i).cos());
        for tick in ticks.iter().filter(|v| **v > min) {
            let r = radius * ((*tick - min) / span) as f32;
            let mut pb = PathBuilder::new();
            let first = at(0, r);
            pb.move_to(first.0, first.1);
            for i in 1..n {
                let p = at(i, r);
                pb.line_to(p.0, p.1);
            }
            pb.close();
            if let Some(path) = pb.finish() {
                self.stroke(&path, pt * 0.75, GRID, t);
            }
        }
        for i in 0..n {
            let end = at(i, radius);
            self.line((cx, cy), end, pt * 0.75, GRID, t);
        }
        // The value labels run up the first spoke.
        for tick in ticks.iter().filter(|v| **v > min) {
            let r = radius * ((*tick - min) / span) as f32;
            let text = match &chart.axis_format {
                Some(code) => format_code(*tick, code),
                None => format_value(*tick, false),
            };
            self.label(
                &text,
                cx - size * 0.3,
                cy - r - size * 0.55,
                size,
                INK,
                Align::Right,
                t,
            );
        }
        for i in 0..n {
            let name = chart.categories.get(i).cloned().unwrap_or_default();
            let (x, y) = at(i, radius + size * 0.7);
            let (s, c) = (angle(i).sin(), angle(i).cos());
            let align = if s > 0.3 {
                Align::Left
            } else if s < -0.3 {
                Align::Right
            } else {
                Align::Center
            };
            let top = if c > 0.5 {
                y - size * 1.1
            } else if c < -0.5 {
                y
            } else {
                y - size * 0.55
            };
            self.label(&name, x, top, size, INK, align, t);
        }
        for (plot, series) in series {
            let points: Vec<(f32, f32)> = (0..n)
                .map(|i| {
                    let v = series.values.get(i).copied().flatten().unwrap_or(min);
                    at(i, radius * ((v - min) / span).clamp(0.0, 1.0) as f32)
                })
                .collect();
            let mut pb = PathBuilder::new();
            pb.move_to(points[0].0, points[0].1);
            for p in &points[1..] {
                pb.line_to(p.0, p.1);
            }
            pb.close();
            if let Some(path) = pb.finish() {
                if plot.filled {
                    let wash = Rgba {
                        a: 0.4,
                        ..series.color
                    };
                    self.fill_with(&path, wash, t);
                }
                self.stroke(&path, 2.25 * pt, series.color, t);
            }
            if plot.markers && series.marker {
                for p in &points {
                    if let Some(dot) = PathBuilder::from_circle(p.0, p.1, 3.5 * pt) {
                        self.fill_with(&dot, series.color, t);
                    }
                }
            }
        }
    }

    /// Scatter series: a numeric x axis under a numeric y axis, markers at the points and, for
    /// the styles that have them, a line through them.
    fn draw_scatter(&mut self, chart: &Chart, rect: (f32, f32, f32, f32), size: f32, t: Transform) {
        let pt = EMU_PER_PT;
        let (mut left, top, right, mut bottom) = rect;
        let series: Vec<&Series> = chart
            .plots
            .iter()
            .filter(|p| p.kind == Kind::Scatter)
            .flat_map(|p| p.series.iter())
            .collect();
        let points: Vec<Vec<(f64, f64)>> = series
            .iter()
            .map(|s| {
                s.values
                    .iter()
                    .enumerate()
                    .filter_map(|(i, y)| {
                        let x = s.xs.get(i).copied().flatten().unwrap_or(i as f64 + 1.0);
                        Some((x, (*y)?))
                    })
                    .collect()
            })
            .collect();
        let (mut xlo, mut xhi, mut ylo, mut yhi) = (f64::MAX, f64::MIN, f64::MAX, f64::MIN);
        for (x, y) in points.iter().flatten() {
            xlo = xlo.min(*x);
            xhi = xhi.max(*x);
            ylo = ylo.min(*y);
            yhi = yhi.max(*y);
        }
        if xlo > xhi {
            return;
        }
        // Excel starts an axis at zero unless the values sit in the top sixth of their range.
        let from_zero = |lo: f64, hi: f64| !(lo > 0.0 && lo >= hi * 5.0 / 6.0);
        let (x0, x1, x_ticks) = scale_from(xlo, xhi, None, None, from_zero(xlo, xhi));
        let (y0, y1, y_ticks) = scale_from(ylo, yhi, None, None, from_zero(ylo, yhi));
        let y_labels: Vec<String> = y_ticks.iter().map(|v| format_value(*v, false)).collect();
        let label_width = y_labels
            .iter()
            .map(|s| self.text_width(s, size))
            .fold(0.0, f32::max);
        left += label_width + size * 0.6;
        bottom -= size * 1.6;
        if right - left < size || bottom - top < size {
            return;
        }
        let (pw, ph) = (right - left, bottom - top);
        let sx = move |x: f64| left + ((x - x0) / (x1 - x0).max(f64::MIN_POSITIVE)) as f32 * pw;
        let sy = move |y: f64| bottom - ((y - y0) / (y1 - y0).max(f64::MIN_POSITIVE)) as f32 * ph;
        for (tick, text) in y_ticks.iter().zip(&y_labels) {
            let y = sy(*tick);
            self.line((left, y), (right, y), pt * 0.75, GRID, t);
            self.label(
                text,
                left - size * 0.4,
                y - size * 0.55,
                size,
                INK,
                Align::Right,
                t,
            );
        }
        for tick in &x_ticks {
            let text = format_value(*tick, false);
            self.label(
                &text,
                sx(*tick),
                bottom + size * 0.3,
                size,
                INK,
                Align::Center,
                t,
            );
        }
        for (series, run) in series.iter().zip(&points) {
            if series.line && run.len() > 1 {
                let mut pb = PathBuilder::new();
                pb.move_to(sx(run[0].0), sy(run[0].1));
                for (x, y) in &run[1..] {
                    pb.line_to(sx(*x), sy(*y));
                }
                if let Some(path) = pb.finish() {
                    self.stroke(&path, 2.25 * pt, series.color, t);
                }
            }
            if series.marker {
                for (x, y) in run {
                    if let Some(dot) = PathBuilder::from_circle(sx(*x), sy(*y), 3.5 * pt) {
                        self.fill_with(&dot, series.color, t);
                    }
                }
            }
        }
        self.line(
            (left, sy(0.0f64.clamp(y0, y1))),
            (right, sy(0.0f64.clamp(y0, y1))),
            pt,
            INK,
            t,
        );
        self.line(
            (sx(0.0f64.clamp(x0, x1)), top),
            (sx(0.0f64.clamp(x0, x1)), bottom),
            pt,
            INK,
            t,
        );
    }

    fn draw_pie(
        &mut self,
        plot: &Plot,
        categories: &[String],
        rect: (f32, f32, f32, f32),
        size: f32,
        t: Transform,
    ) {
        let Some(series) = plot.series.first() else {
            return;
        };
        let values: Vec<f64> = series
            .values
            .iter()
            .map(|v| v.unwrap_or(0.0).max(0.0))
            .collect();
        let total: f64 = values.iter().sum();
        let (cx, cy) = ((rect.0 + rect.2) / 2.0, (rect.1 + rect.3) / 2.0);
        let radius = ((rect.2 - rect.0).min(rect.3 - rect.1) / 2.0 * 0.92).max(0.0);
        if total <= 0.0 || radius <= 0.0 {
            return;
        }
        let mut angle = (plot.first_angle - 90.0).to_radians();
        for (i, v) in values.iter().enumerate() {
            let sweep = (*v / total) as f32 * std::f32::consts::TAU;
            if sweep <= 0.0 {
                continue;
            }
            let steps = (sweep.to_degrees() / 2.0).ceil().max(1.0) as usize;
            let point = |r: f32, a: f32| (cx + r * a.cos(), cy + r * a.sin());
            let mut pb = PathBuilder::new();
            let inner = radius * plot.hole;
            if inner > 0.0 {
                let (x, y) = point(radius, angle);
                pb.move_to(x, y);
            } else {
                pb.move_to(cx, cy);
                let (x, y) = point(radius, angle);
                pb.line_to(x, y);
            }
            for s in 1..=steps {
                let (x, y) = point(radius, angle + sweep * s as f32 / steps as f32);
                pb.line_to(x, y);
            }
            if inner > 0.0 {
                for s in (0..=steps).rev() {
                    let (x, y) = point(inner, angle + sweep * s as f32 / steps as f32);
                    pb.line_to(x, y);
                }
            }
            pb.close();
            if let Some(path) = pb.finish() {
                let color = series.slices.get(i).copied().unwrap_or(series.color);
                self.fill_with(&path, color, t);
                self.stroke(&path, EMU_PER_PT, WHITE, t);
                let labels = &series.labels;
                if labels.value || labels.percent || labels.category {
                    let mut parts = Vec::new();
                    if labels.category {
                        parts.push(categories.get(i).cloned().unwrap_or_default());
                    }
                    if labels.value {
                        parts.push(labels.render(*v));
                    }
                    if labels.percent {
                        let share = *v / total;
                        parts.push(match &labels.format {
                            Some(code) if code.contains('%') => format_code(share, code),
                            _ => format_value(share, true),
                        });
                    }
                    let mid = angle + sweep / 2.0;
                    let at = if inner > 0.0 {
                        (radius + inner) / 2.0
                    } else {
                        radius * 0.65
                    };
                    let ink = contrast(color);
                    self.label(
                        &parts.join(", "),
                        cx + at * mid.cos(),
                        cy + at * mid.sin() - size * 0.55,
                        size,
                        ink,
                        Align::Center,
                        t,
                    );
                }
            }
            angle += sweep;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_value_axis_runs_on_round_numbers() {
        let (lo, hi, ticks) = scale(0.0, 87.0, None, None);
        assert_eq!((lo, hi), (0.0, 100.0));
        assert_eq!(ticks, [0.0, 20.0, 40.0, 60.0, 80.0, 100.0]);
        let (lo, hi, ticks) = scale(-3.0, 7.0, None, None);
        assert_eq!((lo, hi), (-4.0, 8.0));
        assert_eq!(ticks.first(), Some(&-4.0));
        assert_eq!(ticks.last(), Some(&8.0));
        // A fixed bound stays where the chart puts it.
        let (lo, hi, _) = scale(0.0, 4.2, Some(1.0), Some(5.0));
        assert_eq!((lo, hi), (1.0, 5.0));
        // No data range still gives an axis.
        let (lo, hi, ticks) = scale(0.0, 0.0, None, None);
        assert!(hi > lo && !ticks.is_empty());
    }

    #[test]
    fn values_read_the_way_an_axis_shows_them() {
        assert_eq!(format_value(40.0, false), "40");
        assert_eq!(format_value(0.25, false), "0.25");
        assert_eq!(format_value(0.5, true), "50%");
    }

    #[test]
    fn number_formats_are_written_the_way_the_format_asks() {
        assert_eq!(format_code(0.5, "0%"), "50%");
        assert_eq!(format_code(0.1234, "0.0%"), "12.3%");
        assert_eq!(format_code(1234567.0, "#,##0"), "1,234,567");
        assert_eq!(format_code(1.23456, "0.00"), "1.23");
        assert_eq!(format_code(2.5, "General"), "2.5");
        assert_eq!(format_code(1500.0, r##""$"#,##0"##), "$1,500");
        assert_eq!(format_code(2.0, r#"0.0"x""#), "2.0x");
        assert_eq!(format_code(-5.0, "0"), "-5");
        assert_eq!(format_code(7.0, "0;[Red]-0"), "7");
    }

    #[test]
    fn data_labels_and_number_formats_are_read_from_the_chart_part() {
        let with_labels = series(0, "S", &["A", "B"], &[0.25, 0.5]).replace(
            "<c:cat>",
            r#"<c:dLbls><c:numFmt formatCode="0.0%" sourceLinked="0"/><c:showVal val="1"/><c:showPercent val="0"/></c:dLbls><c:cat>"#,
        );
        let part = chart_part(
            &format!(
                r#"<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>{with_labels}<c:axId val="1"/><c:axId val="2"/></c:barChart><c:valAx><c:axId val="2"/><c:scaling/><c:numFmt formatCode="0%" sourceLinked="0"/></c:valAx>"#
            ),
            "",
        );
        let doc = roxmltree::Document::parse(&part).unwrap();
        let colors = Colors {
            scheme: HashMap::new(),
            map: HashMap::new(),
        };
        let chart = read_chart(&doc, &colors).unwrap();
        let labels = &chart.plots[0].series[0].labels;
        assert!(labels.value && !labels.percent && !labels.category);
        assert_eq!(labels.render(0.25), "25.0%");
        assert_eq!(chart.axis_format.as_deref(), Some("0%"));
        // A series without labels shows none.
        let plain = chart_part(
            &format!(
                r#"<c:barChart><c:barDir val="col"/>{}<c:axId val="1"/><c:axId val="2"/></c:barChart>"#,
                series(0, "S", &["A"], &[1.0])
            ),
            "",
        );
        let doc = roxmltree::Document::parse(&plain).unwrap();
        let chart = read_chart(&doc, &colors).unwrap();
        assert!(!chart.plots[0].series[0].labels.value);
    }

    #[test]
    fn declared_text_sizes_are_read_and_kept_within_the_frame() {
        let colors = Colors {
            scheme: HashMap::new(),
            map: HashMap::new(),
        };
        let read = |extra_space: &str, extra_chart: &str, axes: &str| {
            let plot = format!(
                r#"<c:barChart><c:barDir val="col"/>{}<c:axId val="1"/><c:axId val="2"/></c:barChart>{axes}"#,
                series(0, "S", &["A"], &[1.0])
            );
            let part = chart_part(&plot, extra_chart).replace(
                "</c:chart></c:chartSpace>",
                &format!("</c:chart>{extra_space}</c:chartSpace>"),
            );
            let doc = roxmltree::Document::parse(&part).unwrap();
            read_chart(&doc, &colors).unwrap()
        };
        let sz = |points: u32| {
            format!(
                r#"<c:txPr><a:p><a:pPr><a:defRPr sz="{}"/></a:pPr></a:p></c:txPr>"#,
                points * 100
            )
        };
        // Nothing declared: Office's defaults.
        let chart = read("", "", "");
        assert_eq!(
            (chart.text_size, chart.legend_size, chart.title_size),
            (10.0, 10.0, 14.0)
        );
        // The value axis sets the axes' size, the legend its own, the title its rich text's.
        let chart = read(
            "",
            &format!(
                r#"<c:title><c:tx><c:rich><a:p><a:r><a:rPr sz="2400"/><a:t>T</a:t></a:r></a:p></c:rich></c:tx></c:title><c:legend><c:legendPos val="b"/>{}</c:legend>"#,
                sz(12)
            ),
            &format!(
                r#"<c:valAx><c:axId val="2"/><c:scaling/>{}</c:valAx>"#,
                sz(18)
            ),
        );
        assert_eq!(
            (chart.text_size, chart.legend_size, chart.title_size),
            (18.0, 12.0, 24.0)
        );
        // The chart space's size stands for what no element declares.
        let chart = read(&sz(14), "<c:legend><c:legendPos val=\"r\"/></c:legend>", "");
        assert_eq!((chart.text_size, chart.legend_size), (14.0, 14.0));
        // A frame that is too short for a large size keeps the text readable and the plot alive.
        assert_eq!(text_points(18.0, 300.0), 18.0);
        assert_eq!(text_points(18.0, 50.0), 10.0);
        assert_eq!(text_points(18.0, 108.0), 18.0);
        assert_eq!(text_points(2.0, 300.0), 6.0);
        assert_eq!(text_points(90.0, 3000.0), 40.0);
    }

    use crate::pptx::raster::tests::{deck_full, pixel, render};

    const WHITE: [u8; 3] = [255, 255, 255];
    const ACCENT1: [u8; 3] = [0x44, 0x72, 0xC4];
    const ACCENT2: [u8; 3] = [0xED, 0x7D, 0x31];

    /// A 100 × 50 pt slide holding one chart frame over the whole slide, and the chart part.
    fn chart_slide(chart_xml: &str) -> crate::raster::RasteredSlide {
        let frame = r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="Ch"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="635000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdC"/></a:graphicData></a:graphic></p:graphicFrame>"#;
        render(deck_full(
            frame,
            "",
            "",
            &[("rIdC", "chart", "../charts/chart1.xml")],
            &[("ppt/charts/chart1.xml", chart_xml.as_bytes())],
        ))
    }

    fn series(index: usize, name: &str, cats: &[&str], values: &[f64]) -> String {
        let cat: String = cats
            .iter()
            .enumerate()
            .map(|(i, c)| format!(r#"<c:pt idx="{i}"><c:v>{c}</c:v></c:pt>"#))
            .collect();
        let val: String = values
            .iter()
            .enumerate()
            .map(|(i, v)| format!(r#"<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>"#))
            .collect();
        format!(
            r#"<c:ser><c:idx val="{index}"/><c:order val="{index}"/><c:tx><c:strRef><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{name}</c:v></c:pt></c:strCache></c:strRef></c:tx><c:cat><c:strRef><c:strCache><c:ptCount val="{}"/>{cat}</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:ptCount val="{}"/>{val}</c:numCache></c:numRef></c:val></c:ser>"#,
            cats.len(),
            values.len()
        )
    }

    fn chart_part(plot: &str, extra: &str) -> String {
        format!(
            r#"<?xml version="1.0"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea>{plot}</c:plotArea>{extra}</c:chart></c:chartSpace>"#
        )
    }

    #[test]
    fn a_column_chart_draws_a_column_per_value_in_the_series_color() {
        // Categories A and B with 10 and 20 on an axis that runs 0 to 20: the second column
        // is twice as tall as the first.
        let part = chart_part(
            &format!(
                r#"<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>{}<c:axId val="1"/><c:axId val="2"/></c:barChart>"#,
                series(0, "S", &["A", "B"], &[10.0, 20.0])
            ),
            "",
        );
        let slide = chart_slide(&part);
        assert_eq!(pixel(&slide, 40, 24), ACCENT1, "the first column");
        assert_eq!(pixel(&slide, 40, 14), WHITE, "above the first column");
        assert_eq!(
            pixel(&slide, 76, 10),
            ACCENT1,
            "the second column reaches the top"
        );
        assert_eq!(pixel(&slide, 76, 3), WHITE, "above the chart's plot");
        assert_eq!(slide.gaps.approximated_charts, 1);
        assert_eq!(slide.gaps.charts, 0);
    }

    #[test]
    fn stacked_columns_pile_up_and_the_series_take_their_own_colors() {
        let part = chart_part(
            &format!(
                r#"<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/>{}{}<c:overlap val="100"/><c:axId val="1"/><c:axId val="2"/></c:barChart>"#,
                series(0, "S1", &["A"], &[10.0]),
                series(1, "S2", &["A"], &[10.0])
            ),
            "",
        );
        let slide = chart_slide(&part);
        // One category, 0 to 20: the lower half of the plot is S1, the upper half S2.
        assert_eq!(pixel(&slide, 50, 24), ACCENT1);
        assert_eq!(pixel(&slide, 50, 9), ACCENT2);
    }

    #[test]
    fn a_pie_draws_slices_clockwise_from_twelve_oclock() {
        let part = chart_part(
            &format!(
                r#"<c:pieChart><c:varyColors val="1"/>{}</c:pieChart>"#,
                series(0, "S", &["A", "B"], &[1.0, 1.0])
            ),
            "",
        );
        let slide = chart_slide(&part);
        assert_eq!(
            pixel(&slide, 60, 25),
            ACCENT1,
            "the first slice is the right half"
        );
        assert_eq!(
            pixel(&slide, 40, 25),
            ACCENT2,
            "the second slice is the left half"
        );
        assert_eq!(pixel(&slide, 5, 25), WHITE);
        assert_eq!(slide.gaps.approximated_charts, 1);
    }

    #[test]
    fn a_line_chart_connects_its_points() {
        let part = chart_part(
            &format!(
                r#"<c:lineChart><c:grouping val="standard"/>{}<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>"#,
                series(0, "S", &["A", "B"], &[10.0, 10.0])
            ),
            "",
        );
        let slide = chart_slide(&part);
        // A flat line at 10 of 0 to 10 sits at the top of the plot, between the two points.
        let top_row: Vec<[u8; 3]> = (0..8).map(|dy| pixel(&slide, 58, 4 + dy)).collect();
        assert!(top_row.contains(&ACCENT1), "{top_row:?}");
    }

    #[test]
    fn a_line_on_a_second_axis_is_drawn_on_its_own_scale() {
        // Columns of 10 and 20 on the first axis; a line of 1000 and 2000 on a second one, whose
        // own scale puts the line's last point at the top of the plot, where the second column
        // also ends.
        let part = chart_part(
            &format!(
                r#"<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>{}<c:axId val="1"/><c:axId val="2"/></c:barChart><c:lineChart><c:grouping val="standard"/>{}<c:marker val="1"/><c:axId val="3"/><c:axId val="4"/></c:lineChart><c:valAx><c:axId val="2"/><c:scaling/></c:valAx><c:valAx><c:axId val="4"/><c:scaling><c:max val="2000"/></c:scaling><c:crosses val="max"/></c:valAx>"#,
                series(0, "Columns", &["A", "B"], &[10.0, 20.0]),
                series(1, "Line", &["A", "B"], &[1000.0, 2000.0])
            ),
            "",
        );
        let doc = roxmltree::Document::parse(&part).unwrap();
        let colors = Colors {
            scheme: HashMap::new(),
            map: HashMap::new(),
        };
        let chart = read_chart(&doc, &colors).unwrap();
        assert_eq!((chart.plots.len(), chart.secondary.len()), (1, 1));
        assert_eq!(chart.secondary_max, Some(2000.0));
        let slide = chart_slide(&part);
        assert_eq!(pixel(&slide, 56, 6), ACCENT2, "the line's last point");
        assert_eq!(slide.gaps.approximated_charts, 1);
    }

    #[test]
    fn a_scatter_chart_marks_its_points_and_joins_them() {
        let part = chart_part(
            r#"<c:scatterChart><c:scatterStyle val="lineMarker"/><c:ser><c:idx val="0"/><c:order val="0"/><c:xVal><c:numRef><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>10</c:v></c:pt></c:numCache></c:numRef></c:xVal><c:yVal><c:numRef><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>10</c:v></c:pt></c:numCache></c:numRef></c:yVal></c:ser><c:axId val="1"/><c:axId val="2"/></c:scatterChart>"#,
            "",
        );
        let slide = chart_slide(&part);
        assert_eq!(pixel(&slide, 92, 7), ACCENT1, "the point at (10, 10)");
        assert_eq!(
            pixel(&slide, 58, 17),
            ACCENT1,
            "the line between the points"
        );
        assert_eq!(pixel(&slide, 40, 8), WHITE);
        assert_eq!((slide.gaps.charts, slide.gaps.approximated_charts), (0, 1));
    }

    #[test]
    fn a_radar_chart_puts_each_value_on_its_spoke() {
        // Three categories of 10 on an axis that ends at 10: the first vertex is straight above
        // the center, and its marker is drawn there.
        let part = chart_part(
            &format!(
                r#"<c:radarChart><c:radarStyle val="marker"/>{}<c:axId val="1"/><c:axId val="2"/></c:radarChart>"#,
                series(0, "S", &["A", "B", "C"], &[10.0, 10.0, 10.0])
            ),
            "",
        );
        let slide = chart_slide(&part);
        let column: Vec<[u8; 3]> = (12..26).map(|y| pixel(&slide, 50, y)).collect();
        assert!(column.contains(&ACCENT1), "{column:?}");
        assert_eq!(pixel(&slide, 4, 25), WHITE);
        assert_eq!((slide.gaps.charts, slide.gaps.approximated_charts), (0, 1));
    }

    #[test]
    fn a_chart_of_a_kind_that_is_not_painted_counts_as_not_drawn() {
        let part = chart_part(
            r#"<c:bubbleChart><c:axId val="1"/><c:axId val="2"/></c:bubbleChart>"#,
            "",
        );
        let slide = chart_slide(&part);
        assert_eq!((slide.gaps.charts, slide.gaps.approximated_charts), (1, 0));
    }

    #[test]
    fn a_chart_part_that_is_missing_counts_as_not_drawn() {
        let frame = r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="Ch"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="635000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdC"/></a:graphicData></a:graphic></p:graphicFrame>"#;
        let slide = render(deck_full(
            frame,
            "",
            "",
            &[("rIdC", "chart", "../charts/missing.xml")],
            &[],
        ));
        assert_eq!((slide.gaps.charts, slide.gaps.approximated_charts), (1, 0));
    }
}
