//! A slide painted to pixels (feature `raster`).
//!
//! The slide is read as a scene — background, then the master's and layout's own shapes, then
//! the slide's shape tree — and each shape's outline comes from [`crate::geometry`], filled and
//! stroked with tiny-skia. Colors resolve through the master's color map and the theme.
//!
//! Text is laid out in each shape's text rectangle with the fonts [`super::raster_text`] finds;
//! placeholders take their position, body and list styles from the layout and master.
//!
//! Pictures (PNG, JPEG) fill their shape's outline, cropped by `srcRect` and stretched.
//!
//! Tables are drawn cell by cell — fills, borders and text — in their table style when the
//! presentation defines it.
//!
//! SmartArt is drawn from the shapes PowerPoint drew for it (`ppt/diagrams/drawingN.xml`).
//!
//! Not painted yet, and counted in the gaps instead: pictures in other formats, charts,
//! embedded objects and other graphic frames, custom geometry, vertical text and scripts that
//! need shaping.

use std::collections::HashMap;

use roxmltree::Node;
use tiny_skia::{
    Color, FillRule, FilterQuality, GradientStop, IntSize, LinearGradient, Paint, PathBuilder,
    Pattern, Pixmap, RadialGradient, Shader, SpreadMode, Stroke, StrokeDash, Transform,
};

use super::raster_text::{breaks_anywhere, is_east_asian, needs_shaping, FontBook};
use super::PptxParser;
use crate::container::OoxmlContainer;
use crate::error::{Error, Result};
use crate::geometry::{self, PathFill, Segment};
use crate::raster::{RasteredSlide, SlideRasterGaps, SlideRasterOptions};

const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
/// PowerPoint's pre-drawn SmartArt (`ppt/diagrams/drawingN.xml`).
const DSP: &str = "http://schemas.microsoft.com/office/drawing/2008/diagram";

const REL_LAYOUT: &str = "/slideLayout";
const REL_MASTER: &str = "/slideMaster";
const REL_THEME: &str = "/theme";
const REL_TABLE_STYLES: &str = "/tableStyles";

/// EMU per point.
const EMU_PER_PT: f32 = 12_700.0;
/// Slides larger than this many pixels are refused rather than allocated.
const MAX_PIXELS: u64 = 200_000_000;

impl PptxParser {
    /// Paint slide `index` (0-based, in presentation order).
    pub fn render_slide(
        &self,
        index: usize,
        options: &SlideRasterOptions,
    ) -> Result<RasteredSlide> {
        let info = self.slides.get(index).ok_or(Error::SectionOutOfRange {
            index,
            count: self.slides.len(),
        })?;
        let target = self.relationships.get(&info.rel_id).ok_or_else(|| {
            Error::MissingComponent(format!("slide relationship {}", info.rel_id))
        })?;
        let slide_path = OoxmlContainer::resolve_path("ppt/presentation.xml", target);
        // The resolution is a rendering request, so a bad one is a rendering failure — the
        // classification unpdf's page renderer gives the same two cases.
        if !(options.dpi.is_finite() && options.dpi > 0.0) {
            return Err(Error::Render(format!(
                "dpi must be positive, got {}",
                options.dpi
            )));
        }

        let (cx, cy) = slide_size(&self.container)?;
        let scale = options.dpi / 72.0 / EMU_PER_PT;
        let width = (cx as f32 * scale).round().max(1.0) as u32;
        let height = (cy as f32 * scale).round().max(1.0) as u32;
        if u64::from(width) * u64::from(height) > MAX_PIXELS {
            return Err(Error::Render(format!(
                "slide {index} at {} dpi would be {width} x {height} pixels",
                options.dpi
            )));
        }
        let mut pixmap = Pixmap::new(width, height)
            .ok_or_else(|| Error::InvalidData("the slide has no area".into()))?;

        let texts = PartTexts::load(&self.container, &slide_path)?;
        let parts = texts.parse()?;
        let colors = parts.colors();
        let mut painter = Painter {
            pixmap: &mut pixmap,
            colors: &colors,
            line_widths: parts.theme_line_widths(),
            gaps: SlideRasterGaps::default(),
            fonts: FontBook::new(&options.fonts, &options.font_dirs, options.system_fonts),
            layout: parts.layout.as_ref().map(|d| d.root_element()),
            master: parts.master.as_ref().map(|d| d.root_element()),
            theme: parts.theme.as_ref().map(|d| d.root_element()),
            table_styles: parts.table_styles.as_ref().map(|d| d.root_element()),
            theme_fonts: ThemeFonts::read(parts.theme.as_ref()),
            substituted: 0,
            container: &self.container,
            rels: &texts.rels,
            local_rels: Vec::new(),
            images: HashMap::new(),
        };

        // Background: the slide's own, else the layout's, else the master's; white without.
        let to_pixels = Transform::from_scale(scale, scale);
        let background = parts
            .chain()
            .find_map(|doc| painter.background(doc.root_element()));
        painter.pixmap.fill(Color::WHITE);
        if let Some(fill) = background {
            let slide_box = tiny_skia::Rect::from_xywh(0.0, 0.0, cx as f32, cy as f32);
            let paint = area_paint(&fill, &painter.images, cx as f64, cy as f64, 0.0);
            if let (Some(rect), Some(paint)) = (slide_box, paint) {
                painter.pixmap.fill_rect(rect, &paint, to_pixels, None);
            }
        }

        let slide_root = parts.slide.root_element();
        let show_master = |root: Node| root.attribute("showMasterSp") != Some("0");
        if let (Some(master), true) = (&parts.master, show_master(slide_root)) {
            if parts
                .layout
                .as_ref()
                .is_none_or(|l| show_master(l.root_element()))
            {
                painter.tree(master.root_element(), to_pixels, false);
            }
        }
        if let (Some(layout), true) = (&parts.layout, show_master(slide_root)) {
            painter.tree(layout.root_element(), to_pixels, false);
        }
        painter.tree(slide_root, to_pixels, true);

        let (gaps, substituted_text_runs) = (painter.gaps, painter.substituted);
        Ok(RasteredSlide {
            width,
            height,
            rgba: pixmap.take(),
            gaps,
            substituted_text_runs,
        })
    }
}

fn slide_size(container: &OoxmlContainer) -> Result<(u64, u64)> {
    let xml = container.read_xml("ppt/presentation.xml")?;
    let doc = parse(&xml)?;
    let size = doc
        .descendants()
        .find(|n| n.has_tag_name((P, "sldSz")))
        .and_then(|n| Some((num(n, "cx")?, num(n, "cy")?)))
        // ECMA-376's default slide: 10 × 7.5 inches.
        .unwrap_or((9_144_000.0, 6_858_000.0));
    Ok((size.0.max(0.0) as u64, size.1.max(0.0) as u64))
}

fn parse(xml: &str) -> Result<roxmltree::Document<'_>> {
    roxmltree::Document::parse(xml)
        .map_err(|e| Error::InvalidData(format!("presentation part: {e}")))
}

fn num(node: Node, attr: &str) -> Option<f64> {
    node.attribute(attr)?.parse().ok()
}

fn child<'a, 'i>(node: Node<'a, 'i>, ns: &str, name: &str) -> Option<Node<'a, 'i>> {
    node.children().find(|n| n.has_tag_name((ns, name)))
}

/// A shape's part (`spPr`, `txBody`, `style`, …): in PresentationML, or in a pre-drawn
/// SmartArt shape's namespace.
fn sp_child<'a, 'i>(node: Node<'a, 'i>, name: &str) -> Option<Node<'a, 'i>> {
    child(node, P, name).or_else(|| child(node, DSP, name))
}

/// The text of the slide and of the parts it inherits from.
struct PartTexts {
    slide: String,
    layout: Option<String>,
    master: Option<String>,
    theme: Option<String>,
    /// The presentation's table styles (`ppt/tableStyles.xml`).
    table_styles: Option<String>,
    /// Each part's internal relationships (id -> part path): slide, layout, master.
    rels: [HashMap<String, String>; 3],
}

impl PartTexts {
    fn load(container: &OoxmlContainer, slide_path: &str) -> Result<Self> {
        let related = |part: &str, kind: &str| -> Result<Option<String>> {
            let rels = container.read_optional_relationships_for_part(part)?;
            Ok(rels
                .by_id
                .values()
                .find(|r| !r.external && r.rel_type.ends_with(kind))
                .map(|r| OoxmlContainer::resolve_path(part, &r.target)))
        };
        let read = |path: &Option<String>| -> Result<Option<String>> {
            match path {
                Some(p) => container.read_xml_optional(p),
                None => Ok(None),
            }
        };
        let slide = container.read_xml(slide_path)?;
        let layout_path = related(slide_path, REL_LAYOUT)?;
        let master_path = match &layout_path {
            Some(p) => related(p, REL_MASTER)?,
            None => None,
        };
        let theme_path = match &master_path {
            Some(p) => related(p, REL_THEME)?,
            None => None,
        };
        let table_styles_path = related("ppt/presentation.xml", REL_TABLE_STYLES)?;
        let targets = |part: &Option<String>| -> Result<HashMap<String, String>> {
            let Some(part) = part else {
                return Ok(HashMap::new());
            };
            Ok(container
                .read_optional_relationships_for_part(part)?
                .by_id
                .into_values()
                .filter(|r| !r.external)
                .map(|r| (r.id, OoxmlContainer::resolve_path(part, &r.target)))
                .collect())
        };
        let rels = [
            targets(&Some(slide_path.to_string()))?,
            targets(&layout_path)?,
            targets(&master_path)?,
        ];
        Ok(Self {
            slide,
            layout: read(&layout_path)?,
            master: read(&master_path)?,
            theme: read(&theme_path)?,
            table_styles: read(&table_styles_path)?,
            rels,
        })
    }

    fn parse(&self) -> Result<Parts<'_>> {
        Ok(Parts {
            slide: parse(&self.slide)?,
            layout: parse_optional(&self.layout)?,
            master: parse_optional(&self.master)?,
            theme: parse_optional(&self.theme)?,
            table_styles: parse_optional(&self.table_styles)?,
        })
    }
}

fn parse_optional(text: &Option<String>) -> Result<Option<roxmltree::Document<'_>>> {
    text.as_deref().map(parse).transpose()
}

/// The slide and the parts it inherits from, parsed.
struct Parts<'a> {
    slide: roxmltree::Document<'a>,
    layout: Option<roxmltree::Document<'a>>,
    master: Option<roxmltree::Document<'a>>,
    theme: Option<roxmltree::Document<'a>>,
    table_styles: Option<roxmltree::Document<'a>>,
}

impl<'a> Parts<'a> {
    /// Slide, layout, master — the order a property is looked up in.
    fn chain(&self) -> impl Iterator<Item = &roxmltree::Document<'a>> {
        std::iter::once(&self.slide)
            .chain(self.layout.as_ref())
            .chain(self.master.as_ref())
    }

    /// The colors scheme names resolve to: the theme's scheme, reached through the master's
    /// color map (`bg1` → `lt1`, …) and the slide's override of it.
    fn colors(&self) -> Colors {
        let mut scheme = HashMap::new();
        if let Some(theme) = &self.theme {
            if let Some(clr) = theme
                .descendants()
                .find(|n| n.has_tag_name((A, "clrScheme")))
            {
                for entry in clr.children().filter(|n| n.is_element()) {
                    if let Some(c) = entry
                        .children()
                        .find(|n| n.is_element())
                        .and_then(|c| base_color(c, &HashMap::new()))
                    {
                        scheme.insert(entry.tag_name().name().to_string(), c);
                    }
                }
            }
        }
        let mut map: HashMap<String, String> = HashMap::new();
        let mut apply = |node: Option<Node>| {
            if let Some(n) = node {
                for a in n.attributes() {
                    map.insert(a.name().to_string(), a.value().to_string());
                }
            }
        };
        if let Some(master) = &self.master {
            apply(child(master.root_element(), P, "clrMap"));
        }
        apply(
            child(self.slide.root_element(), P, "clrMapOvr")
                .and_then(|o| child(o, A, "overrideClrMapping")),
        );
        Colors { scheme, map }
    }

    /// The theme's line widths by style index (`lnRef idx` 1, 2, 3 …), in EMU.
    fn theme_line_widths(&self) -> Vec<f32> {
        self.theme
            .as_ref()
            .and_then(|t| t.descendants().find(|n| n.has_tag_name((A, "lnStyleLst"))))
            .map(|l| {
                l.children()
                    .filter(|n| n.has_tag_name((A, "ln")))
                    .map(|n| num(n, "w").unwrap_or(9_525.0) as f32)
                    .collect()
            })
            .unwrap_or_default()
    }
}

/// Theme colors by scheme name, and the color map from the names shapes use to them.
struct Colors {
    scheme: HashMap<String, Rgba>,
    map: HashMap<String, String>,
}

impl Colors {
    fn scheme(&self, name: &str) -> Option<Rgba> {
        let mapped = self.map.get(name).map(String::as_str).unwrap_or(name);
        self.scheme.get(mapped).copied()
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
struct Rgba {
    r: f32,
    g: f32,
    b: f32,
    a: f32,
}

impl Rgba {
    const BLACK: Rgba = Rgba {
        r: 0.0,
        g: 0.0,
        b: 0.0,
        a: 1.0,
    };

    fn to_color(self) -> Color {
        Color::from_rgba(
            self.r.clamp(0.0, 1.0),
            self.g.clamp(0.0, 1.0),
            self.b.clamp(0.0, 1.0),
            self.a.clamp(0.0, 1.0),
        )
        .unwrap_or(Color::BLACK)
    }

    fn hsl(self) -> (f32, f32, f32) {
        let (max, min) = (
            self.r.max(self.g).max(self.b),
            self.r.min(self.g).min(self.b),
        );
        let l = (max + min) / 2.0;
        if max == min {
            return (0.0, 0.0, l);
        }
        let d = max - min;
        let s = if l > 0.5 {
            d / (2.0 - max - min)
        } else {
            d / (max + min)
        };
        let h = if max == self.r {
            (self.g - self.b) / d + if self.g < self.b { 6.0 } else { 0.0 }
        } else if max == self.g {
            (self.b - self.r) / d + 2.0
        } else {
            (self.r - self.g) / d + 4.0
        } / 6.0;
        (h, s, l)
    }

    fn from_hsl(h: f32, s: f32, l: f32, a: f32) -> Rgba {
        if s == 0.0 {
            return Rgba {
                r: l,
                g: l,
                b: l,
                a,
            };
        }
        let q = if l < 0.5 {
            l * (1.0 + s)
        } else {
            l + s - l * s
        };
        let p = 2.0 * l - q;
        let hue = |mut t: f32| {
            t = t.rem_euclid(1.0);
            if t < 1.0 / 6.0 {
                p + (q - p) * 6.0 * t
            } else if t < 0.5 {
                q
            } else if t < 2.0 / 3.0 {
                p + (q - p) * (2.0 / 3.0 - t) * 6.0
            } else {
                p
            }
        };
        Rgba {
            r: hue(h + 1.0 / 3.0),
            g: hue(h),
            b: hue(h - 1.0 / 3.0),
            a,
        }
    }
}

/// A color element's base color, before its modifiers: `srgbClr`, `sysClr` (its last
/// computed value), `prstClr` for the few names in use, or `schemeClr` through `scheme`.
fn base_color(node: Node, scheme: &HashMap<String, Rgba>) -> Option<Rgba> {
    let hex = |v: &str| -> Option<Rgba> {
        let n = u32::from_str_radix(v, 16).ok()?;
        Some(Rgba {
            r: ((n >> 16) & 0xFF) as f32 / 255.0,
            g: ((n >> 8) & 0xFF) as f32 / 255.0,
            b: (n & 0xFF) as f32 / 255.0,
            a: 1.0,
        })
    };
    match node.tag_name().name() {
        "srgbClr" => hex(node.attribute("val")?),
        "sysClr" => hex(node.attribute("lastClr")?),
        "prstClr" => match node.attribute("val")? {
            "black" => hex("000000"),
            "white" => hex("FFFFFF"),
            "red" => hex("FF0000"),
            "green" => hex("008000"),
            "blue" => hex("0000FF"),
            "yellow" => hex("FFFF00"),
            "gray" => hex("808080"),
            _ => None,
        },
        "schemeClr" => scheme.get(node.attribute("val")?).copied(),
        _ => None,
    }
}

/// A color element with its modifiers applied (`lumMod`, `lumOff`, `tint`, `shade`, `alpha`,
/// `satMod`), the scheme name `phClr` standing for `placeholder`.
fn color(node: Node, colors: &Colors, placeholder: Option<Rgba>) -> Option<Rgba> {
    let mut c = match node.tag_name().name() {
        "schemeClr" => match node.attribute("val")? {
            "phClr" => placeholder?,
            name => colors.scheme(name)?,
        },
        _ => base_color(node, &colors.scheme)?,
    };
    for m in node.children().filter(|n| n.is_element()) {
        let v = num(m, "val").unwrap_or(0.0) as f32 / 100_000.0;
        match m.tag_name().name() {
            "alpha" => c.a = v,
            "lumMod" | "lumOff" | "satMod" => {
                let (h, mut s, mut l) = c.hsl();
                match m.tag_name().name() {
                    "lumMod" => l *= v,
                    "lumOff" => l += v,
                    _ => s *= v,
                }
                c = Rgba::from_hsl(h, s.clamp(0.0, 1.0), l.clamp(0.0, 1.0), c.a);
            }
            "shade" => {
                c.r *= v;
                c.g *= v;
                c.b *= v;
            }
            "tint" => {
                c.r += (1.0 - c.r) * (1.0 - v);
                c.g += (1.0 - c.g) * (1.0 - v);
                c.b += (1.0 - c.b) * (1.0 - v);
            }
            _ => {}
        }
    }
    Some(c)
}

/// The first color element among `node`'s children.
fn color_child(node: Node, colors: &Colors, placeholder: Option<Rgba>) -> Option<Rgba> {
    node.children()
        .filter(|n| n.is_element())
        .find_map(|c| color(c, colors, placeholder))
}

/// What a fill paints.
#[derive(Debug, Clone)]
enum Fill {
    Solid(Rgba),
    Gradient(Gradient),
    /// A decoded picture, by its part's path, stretched over the area it fills.
    Picture(String),
}

impl Fill {
    /// The one color a line draws this fill in: a gradient's first stop, counted as
    /// approximated; a picture draws no line.
    fn line_color(&self, gaps: &mut SlideRasterGaps) -> Option<Rgba> {
        match self {
            Fill::Solid(c) => Some(*c),
            Fill::Gradient(g) => {
                gaps.approximated_fills += 1;
                g.stops.first().map(|s| s.1)
            }
            Fill::Picture(_) => {
                gaps.approximated_fills += 1;
                None
            }
        }
    }
}

/// A gradient: its stops (position 0–1, color) in order, and how they lie over the area.
#[derive(Debug, Clone)]
struct Gradient {
    stops: Vec<(f32, Rgba)>,
    kind: GradientKind,
}

#[derive(Debug, Clone, Copy)]
enum GradientKind {
    /// Along a line at `angle` degrees clockwise from the x axis; `scaled` turns the angle
    /// with the area's proportions, so a 45° gradient runs corner to corner.
    Linear { angle: f32, scaled: bool },
    /// Outward from the focus rectangle (`l`, `t`, `r`, `b` insets as fractions of the area),
    /// the first stop at the focus and the last at the farthest corner.
    Radial { focus: [f32; 4] },
}

/// The gradient an `a:gradFill` describes, and whether drawing it is an approximation (a
/// rectangular or shape-following path is drawn as a radial one). `placeholder` stands for
/// `phClr`.
fn gradient(node: Node, colors: &Colors, placeholder: Option<Rgba>) -> Option<(Gradient, bool)> {
    let mut stops: Vec<(f32, Rgba)> = child(node, A, "gsLst")?
        .children()
        .filter(|n| n.has_tag_name((A, "gs")))
        .filter_map(|gs| {
            let pos = (num(gs, "pos").unwrap_or(0.0) / 100_000.0).clamp(0.0, 1.0) as f32;
            Some((pos, color_child(gs, colors, placeholder)?))
        })
        .collect();
    if stops.is_empty() {
        return None;
    }
    stops.sort_by(|a, b| a.0.total_cmp(&b.0));
    if let Some(path) = child(node, A, "path") {
        let inset = |name: &str| {
            child(path, A, "fillToRect")
                .and_then(|r| num(r, name))
                .map_or(0.0, |v| (v / 100_000.0) as f32)
        };
        let focus = [inset("l"), inset("t"), inset("r"), inset("b")];
        let approximated = path.attribute("path") != Some("circle");
        return Some((
            Gradient {
                stops,
                kind: GradientKind::Radial { focus },
            },
            approximated,
        ));
    }
    let lin = child(node, A, "lin");
    let angle = lin
        .and_then(|l| num(l, "ang"))
        .map_or(0.0, |a| (a / 60_000.0) as f32);
    let scaled = lin
        .and_then(|l| l.attribute("scaled"))
        .is_some_and(|v| v == "1" || v == "true");
    Some((
        Gradient {
            stops,
            kind: GradientKind::Linear { angle, scaled },
        },
        false,
    ))
}

/// A gradient shader over a `w` × `h` area (in its own coordinates), its colors made lighter
/// (`shade` > 0) or darker.
fn gradient_shader(g: &Gradient, w: f64, h: f64, shade: f32) -> Option<Shader<'static>> {
    let stops: Vec<GradientStop> = g
        .stops
        .iter()
        .map(|&(pos, c)| GradientStop::new(pos, shaded(c, shade).to_color()))
        .collect();
    let (w, h) = (w as f32, h as f32);
    match g.kind {
        GradientKind::Linear { angle, scaled } => {
            let (sin, cos) = angle.to_radians().sin_cos();
            // A scaled angle is laid out in a unit square and stretched with the area: the
            // direction keeps its isolines on the stretched ones.
            let (dx, dy) = if scaled {
                (cos * h, sin * w)
            } else {
                (cos, sin)
            };
            let norm = dx.hypot(dy);
            if norm == 0.0 {
                return None;
            }
            let (dx, dy) = (dx / norm, dy / norm);
            let half = (w * dx.abs() + h * dy.abs()) / 2.0;
            let (cx, cy) = (w / 2.0, h / 2.0);
            LinearGradient::new(
                tiny_skia::Point::from_xy(cx - dx * half, cy - dy * half),
                tiny_skia::Point::from_xy(cx + dx * half, cy + dy * half),
                stops,
                SpreadMode::Pad,
                Transform::identity(),
            )
        }
        GradientKind::Radial {
            focus: [l, t, r, b],
        } => {
            let cx = w * (l + (1.0 - l - r) / 2.0);
            let cy = h * (t + (1.0 - t - b) / 2.0);
            let radius = [(0.0, 0.0), (w, 0.0), (0.0, h), (w, h)]
                .iter()
                .map(|&(x, y)| (x - cx).hypot(y - cy))
                .fold(0.0, f32::max);
            let center = tiny_skia::Point::from_xy(cx, cy);
            RadialGradient::new(
                center,
                0.0,
                center,
                radius.max(1.0),
                stops,
                SpreadMode::Pad,
                Transform::identity(),
            )
        }
    }
}

/// How `fill` paints a `w` × `h` area in its own coordinates, lighter (`shade` > 0) or darker;
/// `None` when it paints nothing (a picture that did not decode, a degenerate gradient).
fn area_paint<'p>(
    fill: &Fill,
    images: &'p HashMap<String, Option<Pixmap>>,
    w: f64,
    h: f64,
    shade: f32,
) -> Option<Paint<'p>> {
    let shader = match fill {
        Fill::Solid(c) => Shader::SolidColor(shaded(*c, shade).to_color()),
        Fill::Gradient(g) => gradient_shader(g, w, h, shade)?,
        Fill::Picture(path) => {
            let image = images.get(path)?.as_ref()?;
            let (iw, ih) = (image.width() as f32, image.height() as f32);
            Pattern::new(
                image.as_ref(),
                SpreadMode::Pad,
                FilterQuality::Bilinear,
                1.0,
                Transform::from_scale(w as f32 / iw, h as f32 / ih),
            )
        }
    };
    Some(Paint {
        shader,
        anti_alias: true,
        ..Paint::default()
    })
}

struct Painter<'a> {
    pixmap: &'a mut Pixmap,
    colors: &'a Colors,
    /// The theme's line widths by style index, in EMU.
    line_widths: Vec<f32>,
    gaps: SlideRasterGaps,
    fonts: FontBook,
    /// The layout's and master's root elements, for what placeholders inherit.
    layout: Option<Node<'a, 'a>>,
    master: Option<Node<'a, 'a>>,
    /// The theme's root element, for the fill styles shapes and backgrounds name.
    theme: Option<Node<'a, 'a>>,
    /// The presentation's table styles (`a:tblStyleLst`).
    table_styles: Option<Node<'a, 'a>>,
    theme_fonts: ThemeFonts,
    /// Text runs drawn in a face standing in for the one they ask for.
    substituted: u32,
    container: &'a OoxmlContainer,
    /// The relationships of the slide, its layout and its master, in that order.
    rels: &'a [HashMap<String, String>; 3],
    /// The relationships of parts read while painting (a SmartArt drawing), by the address of
    /// their parsed document.
    local_rels: Vec<(usize, HashMap<String, String>)>,
    /// Decoded pictures by part path; `None` for one that did not decode.
    images: HashMap<String, Option<Pixmap>>,
}

impl<'a> Painter<'a> {
    /// The relationships a reference in `node` resolves against: those of the part `node` is
    /// in. A placeholder's picture fill or a list style's picture bullet can come from the
    /// layout or the master, and its `r:embed` names a relationship of that part.
    fn rels_of(&self, node: Node) -> Option<&HashMap<String, String>> {
        let address = node.document() as *const _ as usize;
        if let Some((_, rels)) = self.local_rels.iter().find(|(a, _)| *a == address) {
            return Some(rels);
        }
        let same = |root: Option<Node>| {
            root.is_some_and(|r| {
                std::ptr::eq(
                    r.document() as *const _ as *const (),
                    node.document() as *const _ as *const (),
                )
            })
        };
        // A theme's own relationships are not read: a picture in a theme fill style is not
        // found, and counts as one not painted.
        if same(self.theme) {
            None
        } else if same(self.master) {
            Some(&self.rels[2])
        } else if same(self.layout) {
            Some(&self.rels[1])
        } else {
            Some(&self.rels[0])
        }
    }

    /// What the fill element `node` paints: `Some(None)` for `noFill`, `Some(Some(_))` for a
    /// fill, `None` for anything that is not a fill. `placeholder` stands for `phClr`, the color
    /// of the style reference that named a theme fill. Patterns are painted in their foreground
    /// color, rectangular and shape-following gradients as radial ones, and tiled pictures
    /// stretched, each counted as approximated; a picture that cannot be shown counts as one
    /// not painted.
    fn fill(&mut self, node: Node, placeholder: Option<Rgba>) -> Option<Option<Fill>> {
        match node.tag_name().name() {
            "noFill" => Some(None),
            "solidFill" => Some(color_child(node, self.colors, placeholder).map(Fill::Solid)),
            "gradFill" => Some(gradient(node, self.colors, placeholder).map(
                |(g, approximated)| {
                    if approximated {
                        self.gaps.approximated_fills += 1;
                    }
                    Fill::Gradient(g)
                },
            )),
            "pattFill" => {
                self.gaps.approximated_fills += 1;
                Some(
                    child(node, A, "fgClr")
                        .and_then(|c| color_child(c, self.colors, placeholder))
                        .map(Fill::Solid),
                )
            }
            "blipFill" => match self.blip_image(node) {
                Some(path) => {
                    if child(node, A, "tile").is_some() {
                        self.gaps.approximated_fills += 1;
                    }
                    Some(Some(Fill::Picture(path)))
                }
                None => {
                    self.gaps.images += 1;
                    Some(None)
                }
            },
            _ => None,
        }
    }

    /// The theme fill style a style reference (`a:fillRef`, `p:bgRef`) names: `idx` 1, 2, …
    /// in the theme's `fillStyleLst`, 1001, 1002, … in its `bgFillStyleLst`, 0 for none. Its
    /// `phClr` is the reference's own color, which also stands in when the theme has no such
    /// style.
    fn style_fill(&mut self, reference: Node) -> Option<Fill> {
        let idx = num(reference, "idx").unwrap_or(0.0) as usize;
        if idx == 0 {
            return None;
        }
        let color = color_child(reference, self.colors, None);
        let (list, i) = if idx >= 1001 {
            ("bgFillStyleLst", idx - 1001)
        } else {
            ("fillStyleLst", idx - 1)
        };
        let style = self
            .theme
            .and_then(|t| t.descendants().find(|n| n.has_tag_name((A, list))))
            .and_then(|l| l.children().filter(|n| n.is_element()).nth(i));
        match style.and_then(|s| self.fill(s, color)) {
            Some(fill) => fill,
            None => color.map(Fill::Solid),
        }
    }

    /// What a slide, layout or master paints behind its shapes: its `p:bgPr` fill, or the theme
    /// style its `p:bgRef` names.
    fn background(&mut self, root: Node) -> Option<Fill> {
        let bg = child(child(root, P, "cSld")?, P, "bg")?;
        if let Some(pr) = child(bg, P, "bgPr") {
            return pr
                .children()
                .filter(|n| n.is_element())
                .find_map(|n| self.fill(n, None))
                .flatten();
        }
        child(bg, P, "bgRef").and_then(|r| self.style_fill(r))
    }

    /// The picture a `blip`-holding element (`blipFill`, `buBlip`) names, decoded and cached;
    /// `None` when it is missing or not PNG or JPEG.
    fn blip_image(&mut self, holder: Node) -> Option<String> {
        let path = child(holder, A, "blip")
            .and_then(|b| b.attribute((R, "embed")))
            .and_then(|id| self.rels_of(holder)?.get(id))
            .cloned()?;
        if !self.images.contains_key(&path) {
            let decoded = self
                .container
                .read_binary(&path)
                .ok()
                .and_then(|b| decode_image(&b));
            self.images.insert(path.clone(), decoded);
        }
        matches!(self.images.get(&path), Some(Some(_))).then_some(path)
    }

    /// Paint the shape tree of a slide, layout or master. Placeholders on a layout or master
    /// are prompts for the slide's content, not content of their own, and are passed over.
    fn tree(&mut self, root: Node, parent: Transform, placeholders: bool) {
        let Some(tree) = child(root, P, "cSld").and_then(|c| child(c, P, "spTree")) else {
            return;
        };
        self.children(tree, parent, placeholders);
    }

    fn children(&mut self, group: Node, parent: Transform, placeholders: bool) {
        for node in group.children().filter(|n| n.is_element()) {
            if !placeholders && is_placeholder(node) {
                continue;
            }
            match node.tag_name().name() {
                "sp" | "cxnSp" => self.shape(node, parent),
                "grpSp" => {
                    let transform = sp_child(node, "grpSpPr")
                        .and_then(|pr| child(pr, A, "xfrm"))
                        .map_or(parent, |x| group_transform(x).post_concat(parent));
                    self.children(node, transform, placeholders);
                }
                "pic" => self.shape(node, parent),
                "graphicFrame" => {
                    let uri = node
                        .descendants()
                        .find(|n| n.has_tag_name((A, "graphicData")))
                        .and_then(|g| g.attribute("uri"))
                        .unwrap_or("");
                    if uri.ends_with("/chart") {
                        self.gaps.charts += 1;
                    } else if uri.ends_with("/table") {
                        self.table(node, parent);
                    } else if uri.ends_with("/diagram") {
                        self.diagram(node, parent);
                    } else {
                        self.gaps.graphic_frames += 1;
                    }
                }
                "AlternateContent" => {
                    // The fallback is what a reader without the extension draws.
                    if let Some(fallback) =
                        node.children().find(|n| n.tag_name().name() == "Fallback")
                    {
                        self.children(fallback, parent, placeholders);
                    }
                }
                _ => {}
            }
        }
    }

    fn shape(&mut self, node: Node, parent: Transform) {
        let ph = placeholder(node);
        let inherited = self.inherited(ph.as_ref());
        let sp_pr = sp_child(node, "spPr");
        // A placeholder without a position of its own takes its layout's, else its master's.
        let xfrm = sp_pr.and_then(|s| child(s, A, "xfrm")).or_else(|| {
            inherited
                .iter()
                .find_map(|n| child(*n, P, "spPr").and_then(|s| child(s, A, "xfrm")))
        });
        let place = xfrm.and_then(|x| {
            let (off, ext) = (child(x, A, "off")?, child(x, A, "ext")?);
            Some((
                x,
                num(off, "x").unwrap_or(0.0),
                num(off, "y").unwrap_or(0.0),
                num(ext, "cx").unwrap_or(0.0),
                num(ext, "cy").unwrap_or(0.0),
            ))
        });
        let Some((xfrm, x, y, w, h)) = place else {
            // Nowhere to draw it: whatever it holds is missing.
            self.gaps.text_runs += count_runs(node);
            if sp_child(node, "blipFill").is_some() {
                self.gaps.images += 1;
            }
            return;
        };
        let transform = shape_transform(xfrm, x, y, w, h).post_concat(parent);
        // Shape properties a placeholder leaves out come from its layout's, then its master's.
        let sp_prs: Vec<Node> = sp_pr
            .into_iter()
            .chain(inherited.iter().filter_map(|n| child(*n, P, "spPr")))
            .collect();
        let has_geometry =
            |s: &Node| child(*s, A, "prstGeom").is_some() || child(*s, A, "custGeom").is_some();
        let geometry = match sp_prs.iter().find(|s| has_geometry(s)) {
            Some(s) => shape_geometry(*s, w, h),
            // No geometry anywhere: the rectangle a shape's box is.
            None => geometry::preset("rect").map(|def| geometry::evaluate(def, w, h, &[])),
        };
        // A picture's fill is its own `blipFill`; a shape's is in its properties.
        let fill_node =
            sp_child(node, "blipFill").or_else(|| sp_prs.iter().find_map(|s| fill_element(*s)));
        let picture = fill_node.filter(|n| n.tag_name().name() == "blipFill");

        let style = std::iter::once(node)
            .chain(inherited.iter().copied())
            .find_map(|n| sp_child(n, "style"));
        let colors = self.colors;
        let style_color = |name: &str| {
            style
                .and_then(|s| child(s, A, name))
                .filter(|r| r.attribute("idx") != Some("0"))
                .and_then(|r| color_child(r, colors, None))
        };
        let fill_ref = style
            .and_then(|s| child(s, A, "fillRef"))
            .filter(|r| r.attribute("idx") != Some("0"));
        let shape_fill = match picture {
            Some(_) => None,
            None => match fill_node.and_then(|n| self.fill(n, None)) {
                Some(fill) => fill,
                None => fill_ref.and_then(|r| self.style_fill(r)),
            },
        };
        let ln = sp_prs.iter().find_map(|s| child(*s, A, "ln"));
        let ln_fill = ln.and_then(|l| {
            l.children()
                .filter(|n| n.is_element())
                .find_map(|n| self.fill(n, None))
        });
        let line_color = match ln_fill {
            Some(fill) => fill.and_then(|f| f.line_color(&mut self.gaps)),
            None => style_color("lnRef"),
        };
        let line = LineStyle {
            color: line_color,
            width: ln.and_then(|l| num(l, "w")).map(|w| w as f32).or_else(|| {
                style
                    .and_then(|s| child(s, A, "lnRef"))
                    .and_then(|r| num(r, "idx"))
                    .and_then(|i| self.line_widths.get((i as usize).checked_sub(1)?).copied())
            }),
            dashed: ln
                .and_then(|l| child(l, A, "prstDash"))
                .and_then(|d| d.attribute("val"))
                .is_some_and(|v| v != "solid"),
            ends: LineEnds {
                head: ln.and_then(|l| line_end(l, "headEnd")),
                tail: ln.and_then(|l| line_end(l, "tailEnd")),
            },
        };

        if let Some(blip_fill) = picture {
            let painted = geometry
                .as_ref()
                .is_some_and(|g| self.picture(blip_fill, g, w, h, transform));
            if !painted {
                self.gaps.images += 1;
            }
        }
        if shape_fill.is_some() || line.color.is_some() {
            match &geometry {
                Some(geometry) => {
                    self.outlines(geometry, shape_fill.as_ref(), &line, (w, h), transform)
                }
                None => self.gaps.shapes += 1,
            }
        }

        if let Some(body) = sp_child(node, "txBody") {
            // A pre-drawn SmartArt shape places its text in a rectangle of its own
            // (`dsp:txXfrm`, in the same space as the shape's).
            let own_rect = child(node, DSP, "txXfrm").and_then(|t| {
                let (off, ext) = (child(t, A, "off")?, child(t, A, "ext")?);
                let (tx, ty) = (num(off, "x")? - x, num(off, "y")? - y);
                Some((tx, ty, tx + num(ext, "cx")?, ty + num(ext, "cy")?))
            });
            let rect = own_rect
                .or_else(|| geometry.as_ref().map(|g| g.text_rect))
                .unwrap_or((0.0, 0.0, w, h));
            let font_color = style_color("fontRef");
            let frame = TextFrame {
                rect,
                anchor: None,
                color: None,
                bold: None,
            };
            self.text(body, ph.as_ref(), &inherited, frame, font_color, transform);
        }
    }

    /// Fill `geometry` with the picture `blip_fill` names, cropped by its `srcRect` and stretched
    /// over the shape's `w` × `h`. False when the picture is missing or not PNG or JPEG.
    fn picture(
        &mut self,
        blip_fill: Node,
        geometry: &geometry::Geometry,
        w: f64,
        h: f64,
        transform: Transform,
    ) -> bool {
        let Some(path) = self.blip_image(blip_fill) else {
            return false;
        };
        let Some(Some(image)) = self.images.get(&path) else {
            return false;
        };
        let (iw, ih) = (image.width() as f32, image.height() as f32);
        let crop = |name: &str| {
            child(blip_fill, A, "srcRect")
                .and_then(|r| num(r, name))
                .map_or(0.0, |v| v as f32 / 100_000.0)
        };
        let (x0, y0) = (crop("l") * iw, crop("t") * ih);
        let (cw, ch) = (
            iw * (1.0 - crop("l") - crop("r")),
            ih * (1.0 - crop("t") - crop("b")),
        );
        if cw <= 0.0 || ch <= 0.0 {
            return false;
        }
        let to_box = Transform::from_translate(-x0, -y0).post_scale(w as f32 / cw, h as f32 / ch);
        let pattern = Pattern::new(
            image.as_ref(),
            SpreadMode::Pad,
            FilterQuality::Bilinear,
            1.0,
            to_box,
        );
        let paint = Paint {
            shader: pattern,
            anti_alias: true,
            ..Paint::default()
        };
        for outline in geometry
            .outlines
            .iter()
            .filter(|o| o.fill != PathFill::None)
        {
            if let Some(path) = to_path(&outline.segments) {
                self.pixmap
                    .fill_path(&path, &paint, FillRule::EvenOdd, transform, None);
            }
        }
        true
    }

    fn outlines(
        &mut self,
        geometry: &geometry::Geometry,
        shape_fill: Option<&Fill>,
        line: &LineStyle,
        (w, h): (f64, f64),
        transform: Transform,
    ) {
        let width = line.width.unwrap_or(9_525.0).max(1.0);
        for outline in &geometry.outlines {
            let Some(path) = to_path(&outline.segments) else {
                continue;
            };
            // An open, stroked path carries the line's ends; its line stops short of a filled
            // arrowhead so the line's square end does not show through the tip.
            let ended = (outline.stroke && line.color.is_some())
                .then(|| arrowheads(&outline.segments, line.ends, width))
                .flatten();
            if let Some(fill) = shape_fill {
                let shade = match outline.fill {
                    PathFill::None => None,
                    PathFill::Norm => Some(0.0),
                    PathFill::Lighten => Some(0.4),
                    PathFill::LightenLess => Some(0.2),
                    PathFill::Darken => Some(-0.4),
                    PathFill::DarkenLess => Some(-0.2),
                };
                if let Some(paint) = shade.and_then(|s| area_paint(fill, &self.images, w, h, s)) {
                    self.pixmap
                        .fill_path(&path, &paint, FillRule::EvenOdd, transform, None);
                }
            }
            if let (true, Some(color)) = (outline.stroke, line.color) {
                let mut paint = Paint::default();
                paint.set_color(color.to_color());
                paint.anti_alias = true;
                let stroke = Stroke {
                    width,
                    dash: if line.dashed {
                        StrokeDash::new(vec![width * 4.0, width * 3.0], 0.0)
                    } else {
                        None
                    },
                    ..Stroke::default()
                };
                let stroked = ended
                    .as_ref()
                    .and_then(|e| to_path(&e.line))
                    .unwrap_or(path);
                self.pixmap
                    .stroke_path(&stroked, &paint, &stroke, transform, None);
                for head in ended.iter().flat_map(|e| &e.heads) {
                    match head {
                        Arrowhead::Fill(shape) => {
                            self.pixmap
                                .fill_path(shape, &paint, FillRule::Winding, transform, None)
                        }
                        Arrowhead::Stroke(shape) => {
                            let open = Stroke {
                                width,
                                line_join: tiny_skia::LineJoin::Miter,
                                ..Stroke::default()
                            };
                            self.pixmap
                                .stroke_path(shape, &paint, &open, transform, None)
                        }
                    }
                }
            }
        }
    }

    /// Draw a table (`a:tbl` in a graphic frame): its grid of columns and rows, each cell's
    /// fill, borders and text. A cell that spans columns or rows covers them; the cells it
    /// covers (`hMerge`, `vMerge`) draw nothing of their own. The table's style
    /// (`tableStyleId`, defined in `ppt/tableStyles.xml`) gives each cell the fill, borders and
    /// text color and weight of the parts it falls in; the cell's own properties override it. A
    /// style the presentation does not define (PowerPoint's built-in ones are not in the file
    /// unless used from it) is not applied, and the table counts as an approximated fill.
    fn table(&mut self, frame: Node, parent: Transform) {
        let tbl = frame.descendants().find(|n| n.has_tag_name((A, "tbl")));
        let place = child(frame, P, "xfrm").and_then(|x| {
            let (off, ext) = (child(x, A, "off")?, child(x, A, "ext")?);
            Some((
                x,
                num(off, "x").unwrap_or(0.0),
                num(off, "y").unwrap_or(0.0),
                num(ext, "cx").unwrap_or(0.0),
                num(ext, "cy").unwrap_or(0.0),
            ))
        });
        let (Some(tbl), Some((xfrm, x, y, w, h))) = (tbl, place) else {
            self.gaps.graphic_frames += 1;
            return;
        };
        let transform = shape_transform(xfrm, x, y, w, h).post_concat(parent);
        let edges = |widths: Vec<f64>| {
            std::iter::once(0.0)
                .chain(widths.iter().scan(0.0, |at, w| {
                    *at += w;
                    Some(*at)
                }))
                .collect::<Vec<f64>>()
        };
        let cols = edges(
            child(tbl, A, "tblGrid")
                .into_iter()
                .flat_map(|g| g.children().filter(|n| n.has_tag_name((A, "gridCol"))))
                .map(|c| num(c, "w").unwrap_or(0.0))
                .collect(),
        );
        let rows: Vec<Node> = tbl
            .children()
            .filter(|n| n.has_tag_name((A, "tr")))
            .collect();
        let row_edges = edges(rows.iter().map(|r| num(*r, "h").unwrap_or(0.0)).collect());
        if cols.len() < 2 || rows.is_empty() {
            self.gaps.graphic_frames += 1;
            return;
        }
        let tbl_pr = child(tbl, A, "tblPr");
        let style_id = tbl_pr
            .and_then(|p| child(p, A, "tableStyleId"))
            .and_then(|i| i.text())
            .map(str::trim);
        let style = style_id.and_then(|id| {
            self.table_styles?
                .children()
                .find(|s| s.has_tag_name((A, "tblStyle")) && s.attribute("styleId") == Some(id))
        });
        if style_id.is_some() && style.is_none() {
            self.gaps.approximated_fills += 1;
        }
        let flag = |name: &str| {
            tbl_pr
                .and_then(|p| p.attribute(name))
                .is_some_and(|v| v == "1" || v == "true")
        };
        let flags = TableFlags {
            first_row: flag("firstRow"),
            last_row: flag("lastRow"),
            first_col: flag("firstCol"),
            last_col: flag("lastCol"),
            band_row: flag("bandRow"),
            band_col: flag("bandCol"),
        };
        let (n_cols, n_rows) = (cols.len() - 1, row_edges.len() - 1);

        // Each drawn cell with the rectangle it covers and its grid span (rows, columns).
        let mut cells: Vec<PlacedCell> = Vec::new();
        for (ri, row) in rows.iter().enumerate() {
            let tcs = row.children().filter(|n| n.has_tag_name((A, "tc")));
            for (ci, tc) in tcs.enumerate() {
                let merged = ["hMerge", "vMerge"]
                    .iter()
                    .any(|a| matches!(tc.attribute(*a), Some("1" | "true")));
                if merged || ci + 1 >= cols.len() {
                    continue;
                }
                let span = |attr: &str| num(tc, attr).map_or(1, |v| v.max(1.0) as usize);
                let c1 = (ci + span("gridSpan")).min(cols.len() - 1);
                let r1 = (ri + span("rowSpan")).min(row_edges.len() - 1);
                cells.push(PlacedCell {
                    tc,
                    rect: (cols[ci], row_edges[ri], cols[c1], row_edges[r1]),
                    span: Span {
                        rows: (ri, r1),
                        cols: (ci, c1),
                    },
                });
            }
        }

        // Fills, then borders over them, then text over both.
        let grid = (n_rows, n_cols);
        for &PlacedCell {
            tc,
            rect: (x0, y0, x1, y1),
            span,
        } in &cells
        {
            let own = child(tc, A, "tcPr").and_then(|pr| {
                pr.children()
                    .filter(|n| n.is_element())
                    .find_map(|n| self.fill(n, None))
            });
            let fill = match own {
                Some(fill) => fill,
                None => {
                    // The last part of the style that fills the cell.
                    let parts = style_parts(style, flags, span, grid);
                    let filled = parts.iter().rev().find_map(|(part, _)| {
                        let tc_style = child(*part, A, "tcStyle")?;
                        child(tc_style, A, "fill").or_else(|| child(tc_style, A, "fillRef"))
                    });
                    match filled {
                        Some(f) if f.tag_name().name() == "fillRef" => self.style_fill(f),
                        Some(f) => f
                            .children()
                            .filter(|n| n.is_element())
                            .find_map(|n| self.fill(n, None))
                            .flatten(),
                        None => None,
                    }
                }
            };
            let (Some(fill), Some(rect)) = (
                fill,
                tiny_skia::Rect::from_xywh(0.0, 0.0, (x1 - x0) as f32, (y1 - y0) as f32),
            ) else {
                continue;
            };
            let at = Transform::from_translate(x0 as f32, y0 as f32).post_concat(transform);
            if let Some(paint) = area_paint(&fill, &self.images, x1 - x0, y1 - y0, 0.0) {
                self.pixmap.fill_rect(rect, &paint, at, None);
            }
        }
        for &PlacedCell {
            tc,
            rect: (x0, y0, x1, y1),
            span,
        } in &cells
        {
            let pr = child(tc, A, "tcPr");
            let parts = style_parts(style, flags, span, grid);
            for (own, edge, from, to) in [
                ("lnL", Edge::Left, (x0, y0), (x0, y1)),
                ("lnR", Edge::Right, (x1, y0), (x1, y1)),
                ("lnT", Edge::Top, (x0, y0), (x1, y0)),
                ("lnB", Edge::Bottom, (x0, y1), (x1, y1)),
                ("lnTlToBr", Edge::TopLeftToBottomRight, (x0, y0), (x1, y1)),
                ("lnBlToTr", Edge::BottomLeftToTopRight, (x0, y1), (x1, y0)),
            ] {
                // The cell's own border, else the last style part that draws this edge.
                let border = pr.and_then(|p| child(p, A, own)).or_else(|| {
                    parts.iter().rev().find_map(|(part, region)| {
                        let name = edge.style_name(span, *region);
                        child(*part, A, "tcStyle")
                            .and_then(|s| child(s, A, "tcBdr"))
                            .and_then(|b| child(b, A, name))
                    })
                });
                let Some(border) = border else {
                    continue;
                };
                let Some((color, width)) = self.border(border) else {
                    continue;
                };
                let mut pb = PathBuilder::new();
                pb.move_to(from.0 as f32, from.1 as f32);
                pb.line_to(to.0 as f32, to.1 as f32);
                let Some(path) = pb.finish() else {
                    continue;
                };
                let mut paint = Paint::default();
                paint.set_color(color.to_color());
                paint.anti_alias = true;
                let stroke = Stroke {
                    width: width.max(1.0),
                    ..Stroke::default()
                };
                self.pixmap
                    .stroke_path(&path, &paint, &stroke, transform, None);
            }
        }
        for &PlacedCell {
            tc,
            rect: (x0, y0, x1, y1),
            span,
        } in &cells
        {
            let Some(body) = child(tc, A, "txBody") else {
                continue;
            };
            // The last style parts that set the text's color and weight.
            let parts = style_parts(style, flags, span, grid);
            let tx_styles: Vec<Node> = parts
                .iter()
                .filter_map(|(p, _)| child(*p, A, "tcTxStyle"))
                .collect();
            let text_color = tx_styles
                .iter()
                .rev()
                .find_map(|t| color_child(*t, self.colors, None));
            let text_bold = tx_styles
                .iter()
                .rev()
                .find_map(|t| t.attribute("b"))
                .map(|b| b == "on");
            // Cell margins stand where a text box's insets would; the defaults are the same.
            let pr = child(tc, A, "tcPr");
            let margin = |name: &str, default: f64| {
                pr.and_then(|p| num(p, name)).unwrap_or(default) - default
            };
            let frame = TextFrame {
                rect: (
                    x0 + margin("marL", 91_440.0),
                    y0 + margin("marT", 45_720.0),
                    x1 - margin("marR", 91_440.0),
                    y1 - margin("marB", 45_720.0),
                ),
                anchor: pr.and_then(|p| p.attribute("anchor")),
                color: text_color,
                bold: text_bold,
            };
            self.text(body, None, &[], frame, None, transform);
        }
    }

    /// Draw a SmartArt diagram from the shapes PowerPoint drew for it (`dsp:drawing`, reached
    /// through the data part's `dsp:dataModelExt`), placed at the graphic frame. A diagram
    /// with no drawing — its layout would have to be computed — counts as a graphic frame not
    /// drawn.
    fn diagram(&mut self, frame: Node, parent: Transform) {
        let drawing_path = (|| {
            let rel_ids = frame
                .descendants()
                .find(|n| n.tag_name().name() == "relIds")?;
            let rels = self.rels_of(frame)?;
            let data_path = rels.get(rel_ids.attribute((R, "dm"))?)?;
            let data = self.container.read_xml(data_path).ok()?;
            let data = roxmltree::Document::parse(&data).ok()?;
            let ext = data
                .descendants()
                .find(|n| n.has_tag_name((DSP, "dataModelExt")))?;
            rels.get(ext.attribute("relId")?).cloned()
        })();
        let offset = child(frame, P, "xfrm")
            .and_then(|x| child(x, A, "off"))
            .map(|o| (num(o, "x").unwrap_or(0.0), num(o, "y").unwrap_or(0.0)));
        let text = drawing_path
            .as_deref()
            .and_then(|p| self.container.read_xml(p).ok());
        let (Some(path), Some((x, y)), Some(text)) = (drawing_path, offset, text) else {
            self.gaps.graphic_frames += 1;
            return;
        };
        let Ok(drawing) = roxmltree::Document::parse(&text) else {
            self.gaps.graphic_frames += 1;
            return;
        };
        let Some(tree) = child(drawing.root_element(), DSP, "spTree") else {
            self.gaps.graphic_frames += 1;
            return;
        };
        let rels = self
            .container
            .read_optional_relationships_for_part(&path)
            .map(|r| {
                r.by_id
                    .into_values()
                    .filter(|r| !r.external)
                    .map(|r| (r.id, OoxmlContainer::resolve_path(&path, &r.target)))
                    .collect()
            })
            .unwrap_or_default();
        let address = &drawing as *const _ as usize;
        self.local_rels.push((address, rels));
        let at = Transform::from_translate(x as f32, y as f32).post_concat(parent);
        self.children(tree, at, true);
        self.local_rels.pop();
    }

    /// The color and width (EMU) a table border draws in: from the `a:ln` it holds, or from
    /// the theme line style its `a:lnRef` names, in the reference's color. `None` draws nothing.
    fn border(&mut self, border: Node) -> Option<(Rgba, f32)> {
        let name = border.tag_name().name();
        let (ln, reference) = if name.starts_with("ln") && name != "lnRef" {
            (Some(border), None)
        } else {
            (child(border, A, "ln"), child(border, A, "lnRef"))
        };
        if let Some(ln) = ln {
            let color = ln
                .children()
                .filter(|n| n.is_element())
                .find_map(|n| self.fill(n, None))
                .flatten()
                .and_then(|f| f.line_color(&mut self.gaps))?;
            return Some((color, num(ln, "w").map_or(12_700.0, |w| w as f32)));
        }
        let reference = reference?;
        let idx = num(reference, "idx")? as usize;
        let width = *self.line_widths.get(idx.checked_sub(1)?)?;
        Some((color_child(reference, self.colors, None)?, width))
    }

    /// The layout's and the master's counterparts of a placeholder, nearest first.
    fn inherited(&self, ph: Option<&Placeholder>) -> Vec<Node<'a, 'a>> {
        let Some(ph) = ph else {
            return Vec::new();
        };
        let mut found = Vec::new();
        if let Some(layout) = self.layout {
            let candidates: Vec<(Node, Placeholder)> = placeholders_in(layout).collect();
            let by_idx = ph
                .idx
                .as_ref()
                .and_then(|i| candidates.iter().find(|(_, c)| c.idx.as_ref() == Some(i)));
            let by_kind = || candidates.iter().find(|(_, c)| c.kind() == ph.kind());
            if let Some((n, _)) = by_idx.or_else(by_kind) {
                found.push(*n);
            }
        }
        if let Some(master) = self.master {
            if let Some((n, _)) = placeholders_in(master).find(|(_, c)| c.kind() == ph.kind()) {
                found.push(n);
            }
        }
        found
    }

    /// Lay out and paint a text body in the shape's text rectangle `(l, t, r, b)`.
    fn text(
        &mut self,
        body: Node,
        ph: Option<&Placeholder>,
        inherited: &[Node<'a, 'a>],
        frame: TextFrame,
        font_color: Option<Rgba>,
        transform: Transform,
    ) {
        let rect = frame.rect;
        // Body properties: the master's, overridden by the layout's, then the shape's own.
        let body_prs: Vec<Node> = inherited
            .iter()
            .rev()
            .filter_map(|n| child(*n, P, "txBody").and_then(|b| child(b, A, "bodyPr")))
            .chain(child(body, A, "bodyPr"))
            .collect();
        let attr = |name: &str| body_prs.iter().rev().find_map(|b| b.attribute(name));
        let inset =
            |name: &str, default: f64| attr(name).and_then(|v| v.parse().ok()).unwrap_or(default);
        if attr("vert").is_some_and(|v| v != "horz") {
            self.gaps.text_runs += count_runs(body);
            return;
        }
        let (l, t) = (
            rect.0 + inset("lIns", 91_440.0),
            rect.1 + inset("tIns", 45_720.0),
        );
        let (r, b) = (
            rect.2 - inset("rIns", 91_440.0),
            rect.3 - inset("bIns", 45_720.0),
        );
        let wrap = attr("wrap") != Some("none");
        let anchor = frame.anchor.or_else(|| attr("anchor")).unwrap_or("t");
        let font_scale = child(body, A, "bodyPr")
            .and_then(|b| child(b, A, "normAutofit"))
            .and_then(|a| num(a, "fontScale"))
            .map_or(1.0, |v| v / 100_000.0) as f32;

        // List styles from the lowest priority up: the master's text styles, the master's and
        // layout's placeholder, the shape's own.
        let kind = ph.map_or("other", |p| match p.kind() {
            "title" => "title",
            _ => "body",
        });
        let mut list_styles: Vec<Node> = Vec::new();
        if let Some(master) = self.master {
            let name = match kind {
                "title" => "titleStyle",
                "body" => "bodyStyle",
                _ => "otherStyle",
            };
            if let Some(s) = child(master, P, "txStyles").and_then(|t| child(t, P, name)) {
                list_styles.push(s);
            }
        }
        for n in inherited.iter().rev() {
            if let Some(s) = child(*n, P, "txBody").and_then(|b| child(b, A, "lstStyle")) {
                list_styles.push(s);
            }
        }
        if let Some(s) = child(body, A, "lstStyle") {
            list_styles.push(s);
        }

        let box_width = (r - l).max(0.0) as f32;
        let mut lines: Vec<Line> = Vec::new();
        let mut run_index = 0u32;
        let mut missing_runs: Vec<u32> = Vec::new();
        let mut substituted_runs: Vec<u32> = Vec::new();
        let mut numbering = Numbering::default();
        for para in body.children().filter(|n| n.has_tag_name((A, "p"))) {
            let p_pr = child(para, A, "pPr");
            let level = p_pr.and_then(|p| num(p, "lvl")).unwrap_or(0.0) as usize;
            let level_name = format!("lvl{}pPr", level.min(8) + 1);
            let levels: Vec<Node> = list_styles
                .iter()
                .filter_map(|s| child(*s, A, &level_name))
                .chain(p_pr)
                .collect();
            let para_attr = |name: &str| levels.iter().rev().find_map(|n| n.attribute(name));
            let align = para_attr("algn").unwrap_or("l").to_string();
            let mar_l = para_attr("marL")
                .and_then(|v| v.parse::<f32>().ok())
                .unwrap_or(0.0);
            let indent = para_attr("indent")
                .and_then(|v| v.parse::<f32>().ok())
                .unwrap_or(0.0);
            let line_spacing = levels
                .iter()
                .rev()
                .find_map(|n| child(*n, A, "lnSpc").and_then(|s| child(s, A, "spcPct")))
                .and_then(|p| num(p, "val"))
                .map_or(1.0, |v| v / 100_000.0) as f32;

            let mut base = RunStyle::default_with(font_color.or_else(|| self.colors.scheme("tx1")));
            for n in &levels {
                if let Some(d) = child(*n, A, "defRPr") {
                    base.apply(d, self.colors);
                }
            }
            if let Some(color) = frame.color {
                base.color = Some(color);
            }
            if let Some(bold) = frame.bold {
                base.bold = bold;
            }

            let mut items: Vec<Item> = Vec::new();
            let mut end_size = base.size;
            let mut first_style: Option<RunStyle> = None;
            for run in para.children().filter(|n| n.is_element()) {
                let name = run.tag_name().name();
                if name == "endParaRPr" {
                    let mut s = base.clone();
                    s.apply(run, self.colors);
                    end_size = s.size;
                    continue;
                }
                if name == "br" {
                    items.push(Item::Break);
                    continue;
                }
                if name != "r" && name != "fld" {
                    continue;
                }
                let mut style = base.clone();
                if let Some(rpr) = child(run, A, "rPr") {
                    style.apply(rpr, self.colors);
                }
                if style.approximated_fill && child(run, A, "t").and_then(|t| t.text()).is_some() {
                    self.gaps.approximated_fills += 1;
                }
                let text: String = child(run, A, "t")
                    .and_then(|t| t.text())
                    .unwrap_or("")
                    .to_string();
                if text.is_empty() {
                    continue;
                }
                if first_style.is_none() {
                    first_style = Some(style.clone());
                }
                let size = style.size * font_scale * EMU_PER_PT;
                for ch in text.chars() {
                    if ch == '\t' || ch == '\n' || ch == '\r' {
                        items.push(Item::Glyph(Glyph::new(
                            ' ',
                            None,
                            size,
                            size * 0.25,
                            style.color,
                        )));
                        continue;
                    }
                    if needs_shaping(ch) {
                        missing_runs.push(run_index);
                        continue;
                    }
                    let families = self.families(&style, ch);
                    let refs: Vec<&str> = families.iter().map(String::as_str).collect();
                    let picked = self.fonts.pick(&refs, style.bold, style.italic, ch);
                    let (face, advance) = match picked {
                        Some((face, substituted)) => {
                            if substituted {
                                substituted_runs.push(run_index);
                            }
                            (Some(face), self.fonts.advance(face, ch, size))
                        }
                        None => {
                            if !ch.is_whitespace() {
                                missing_runs.push(run_index);
                            }
                            (None, size * 0.3)
                        }
                    };
                    items.push(Item::Glyph(Glyph::new(
                        ch,
                        face,
                        size,
                        advance,
                        style.color,
                    )));
                }
                run_index += 1;
            }
            // A bullet hangs in the first line's indent, sized and colored after the
            // paragraph's first run unless the list says otherwise. A paragraph with no text
            // shows none, but still takes its place in a numbered list.
            let first_glyph = items.iter().find_map(|i| match i {
                Item::Glyph(g) => Some(g.clone()),
                Item::Break => None,
            });
            let marker = bullet(&levels);
            let number = numbering.next(level, marker.as_ref());
            if let (Some(marker), Some(first_glyph)) = (marker, first_glyph) {
                let size_pct = levels
                    .iter()
                    .rev()
                    .find_map(|n| child(*n, A, "buSzPct"))
                    .and_then(|n| num(n, "val"))
                    .map_or(1.0, |v| v / 100_000.0) as f32;
                let size = first_glyph.size * size_pct;
                let color = levels
                    .iter()
                    .rev()
                    .find_map(|n| child(*n, A, "buClr"))
                    .and_then(|c| color_child(c, self.colors, None))
                    .or(first_glyph.color);
                let font = levels
                    .iter()
                    .rev()
                    .find_map(|n| child(*n, A, "buFont"))
                    .and_then(|f| f.attribute("typeface"))
                    .map(str::to_string);
                let wanted: Vec<&str> = font.iter().map(String::as_str).collect();
                let mut glyphs: Vec<Glyph> = Vec::new();
                match marker {
                    Bullet::Char(bullet) => {
                        // A symbol face's private-use codes mean nothing in another face: a
                        // bullet the named face cannot draw is drawn as a plain one.
                        let (ch, face) = match self.fonts.pick(&wanted, false, false, bullet) {
                            Some((face, substituted)) if !substituted || wanted.is_empty() => {
                                (bullet, Some(face))
                            }
                            _ => (
                                '\u{2022}',
                                self.fonts.pick(&[], false, false, '\u{2022}').map(|p| p.0),
                            ),
                        };
                        let advance = face.map_or(size * 0.5, |f| self.fonts.advance(f, ch, size));
                        glyphs.push(Glyph::new(ch, face, size, advance, color));
                    }
                    Bullet::Number { scheme, .. } => {
                        // Numbers are drawn in the bullet font when the list names one, else
                        // in the paragraph's first run's face.
                        let style = first_style.clone().unwrap_or_else(|| base.clone());
                        for ch in autonumber(&scheme, number.unwrap_or(1)).chars() {
                            let families = if wanted.is_empty() {
                                self.families(&style, ch)
                            } else {
                                font.iter().cloned().collect()
                            };
                            let refs: Vec<&str> = families.iter().map(String::as_str).collect();
                            let face = self
                                .fonts
                                .pick(&refs, style.bold, style.italic, ch)
                                .map(|p| p.0);
                            let advance =
                                face.map_or(size * 0.5, |f| self.fonts.advance(f, ch, size));
                            glyphs.push(Glyph::new(ch, face, size, advance, color));
                        }
                    }
                    Bullet::Picture(holder) => {
                        // A picture that cannot be shown still holds its place: the text
                        // starts where it would beside the picture.
                        let mut g = Glyph::new('\u{FFFC}', None, size, size, color);
                        g.picture = self.blip_image(holder);
                        if g.picture.is_none() {
                            self.gaps.images += 1;
                        }
                        glyphs.push(g);
                    }
                }
                // The text starts at the indent's end, or a quarter em after the bullet when
                // the bullet is wider than the indent leaves.
                let own: f32 = glyphs.iter().map(|g| g.advance).sum();
                let hang = if indent < 0.0 {
                    -indent
                } else {
                    own + size * 0.25
                };
                if let Some(last) = glyphs.last_mut() {
                    last.advance += hang.max(own) - own;
                }
                for (k, g) in glyphs.into_iter().enumerate() {
                    items.insert(k, Item::Glyph(g));
                }
            }
            let first = box_width - mar_l - indent;
            let rest = box_width - mar_l;
            let start = lines.len();
            break_lines(
                &items,
                if wrap { Some((first, rest)) } else { None },
                &mut lines,
            );
            if lines.len() == start {
                lines.push(Line::empty(end_size * font_scale * EMU_PER_PT));
            }
            for (i, line) in lines[start..].iter_mut().enumerate() {
                line.align = align.clone();
                line.left = mar_l + if i == 0 { indent } else { 0.0 };
                line.spacing = line_spacing;
            }
        }
        missing_runs.dedup();
        substituted_runs.dedup();
        self.gaps.text_runs += missing_runs.len() as u32;
        self.substituted += substituted_runs
            .iter()
            .filter(|r| !missing_runs.contains(r))
            .count() as u32;

        // Vertical placement.
        let heights: Vec<f32> = lines.iter().map(|l| l.height() * l.spacing).collect();
        let total: f32 = heights.iter().sum();
        let box_height = (b - t) as f32;
        let mut y = t as f32
            + match anchor {
                "ctr" => (box_height - total) / 2.0,
                "b" => box_height - total,
                _ => 0.0,
            };
        for (line, height) in lines.iter().zip(heights) {
            let width = line.width();
            let x0 = l as f32
                + line.left
                + match line.align.as_str() {
                    "ctr" => (box_width - line.left - width) / 2.0,
                    "r" => box_width - line.left - width,
                    _ => 0.0,
                };
            let ascent = line
                .glyphs
                .iter()
                .filter_map(|g| g.face.map(|f| self.fonts.ascent(f, g.size)))
                .fold(line.height() * 0.8, f32::max);
            let baseline = y + ascent + (height - line.height()) / 2.0;
            let mut x = x0;
            for g in &line.glyphs {
                if let Some(Some(image)) = g.picture.as_ref().and_then(|p| self.images.get(p)) {
                    let side = g.size * 0.8;
                    let (iw, ih) = (image.width() as f32, image.height() as f32);
                    let place = Transform::from_scale(side / iw, side / ih)
                        .post_translate(x, baseline - side);
                    let paint = Paint {
                        shader: Pattern::new(
                            image.as_ref(),
                            SpreadMode::Pad,
                            FilterQuality::Bilinear,
                            1.0,
                            place,
                        ),
                        anti_alias: true,
                        ..Paint::default()
                    };
                    if let Some(rect) = tiny_skia::Rect::from_xywh(x, baseline - side, side, side) {
                        self.pixmap.fill_rect(rect, &paint, transform, None);
                    }
                    x += g.advance;
                    continue;
                }
                if let (Some(face), false) = (g.face, g.ch.is_whitespace()) {
                    if let Some(path) = self.fonts.glyph(face, g.ch, g.size, x, baseline) {
                        let mut paint = Paint::default();
                        paint.set_color(g.color.unwrap_or(Rgba::BLACK).to_color());
                        paint.anti_alias = true;
                        self.pixmap
                            .fill_path(&path, &paint, FillRule::Winding, transform, None);
                    }
                }
                x += g.advance;
            }
            y += height;
        }
    }

    /// The families to draw `ch` in for a run styled `style`: its East Asian face for an East
    /// Asian character, else its Latin face, theme references resolved.
    fn families(&self, style: &RunStyle, ch: char) -> Vec<String> {
        let mut out = Vec::new();
        let latin = style.latin.as_deref().unwrap_or("+mn-lt");
        let ea = style.ea.as_deref().unwrap_or("+mn-ea");
        if is_east_asian(ch) {
            out.extend(self.theme_fonts.resolve(ea, ch));
        }
        out.extend(self.theme_fonts.resolve(latin, ch));
        out.retain(|f| !f.is_empty());
        out.dedup();
        out
    }
}

/// Which of a table style's parts a table turns on (`a:tblPr` attributes).
#[derive(Debug, Clone, Copy, Default)]
struct TableFlags {
    first_row: bool,
    last_row: bool,
    first_col: bool,
    last_col: bool,
    band_row: bool,
    band_col: bool,
}

/// A table cell that draws: its element, the rectangle it covers and its place in the grid.
#[derive(Debug, Clone, Copy)]
struct PlacedCell<'a, 'i> {
    tc: Node<'a, 'i>,
    rect: (f64, f64, f64, f64),
    span: Span,
}

/// A cell's place in the grid: the rows and columns it covers, end exclusive.
#[derive(Debug, Clone, Copy, PartialEq)]
struct Span {
    rows: (usize, usize),
    cols: (usize, usize),
}

/// A cell edge, and which border of a style part draws it.
#[derive(Debug, Clone, Copy)]
enum Edge {
    Left,
    Right,
    Top,
    Bottom,
    TopLeftToBottomRight,
    BottomLeftToTopRight,
}

impl Edge {
    /// The part's border element for this edge of a cell in `region`: the region's own edge
    /// (`left`, `top`, …) where the cell meets it, the inside border (`insideV`, `insideH`)
    /// where it does not.
    fn style_name(self, cell: Span, region: Span) -> &'static str {
        match self {
            Edge::Left if cell.cols.0 == region.cols.0 => "left",
            Edge::Right if cell.cols.1 == region.cols.1 => "right",
            Edge::Left | Edge::Right => "insideV",
            Edge::Top if cell.rows.0 == region.rows.0 => "top",
            Edge::Bottom if cell.rows.1 == region.rows.1 => "bottom",
            Edge::Top | Edge::Bottom => "insideH",
            Edge::TopLeftToBottomRight => "tl2br",
            Edge::BottomLeftToTopRight => "tr2bl",
        }
    }
}

/// The parts of table style `style` a cell in `span` of a `rows` × `cols` grid falls in, in the
/// order they apply — later ones win — each with the region it covers: the whole table, the
/// column band, the row band, the last and first column, the last and first row, then the
/// corner cells.
fn style_parts<'a, 'i>(
    style: Option<Node<'a, 'i>>,
    flags: TableFlags,
    span: Span,
    (rows, cols): (usize, usize),
) -> Vec<(Node<'a, 'i>, Span)> {
    let Some(style) = style else {
        return Vec::new();
    };
    let whole = Span {
        rows: (0, rows),
        cols: (0, cols),
    };
    let row_strip = Span {
        rows: span.rows,
        cols: (0, cols),
    };
    let col_strip = Span {
        rows: (0, rows),
        cols: span.cols,
    };
    let first_row = flags.first_row && span.rows.0 == 0;
    let last_row = flags.last_row && span.rows.1 == rows;
    let first_col = flags.first_col && span.cols.0 == 0;
    let last_col = flags.last_col && span.cols.1 == cols;
    // Bands count from the first row or column that is not a header.
    let band = |at: usize, header: bool, edge: bool| {
        (!edge).then(|| (at - usize::from(header)).is_multiple_of(2))
    };
    let mut wanted: Vec<(&str, Span)> = vec![("wholeTbl", whole)];
    if flags.band_col {
        if let Some(odd) = band(span.cols.0, flags.first_col, first_col || last_col) {
            wanted.push((if odd { "band1V" } else { "band2V" }, col_strip));
        }
    }
    if flags.band_row {
        if let Some(odd) = band(span.rows.0, flags.first_row, first_row || last_row) {
            wanted.push((if odd { "band1H" } else { "band2H" }, row_strip));
        }
    }
    if last_col {
        wanted.push(("lastCol", col_strip));
    }
    if first_col {
        wanted.push(("firstCol", col_strip));
    }
    if last_row {
        wanted.push(("lastRow", row_strip));
    }
    if first_row {
        wanted.push(("firstRow", row_strip));
    }
    for (corner, on) in [
        ("seCell", last_row && last_col),
        ("swCell", last_row && first_col),
        ("neCell", first_row && last_col),
        ("nwCell", first_row && first_col),
    ] {
        if on {
            wanted.push((corner, span));
        }
    }
    wanted
        .into_iter()
        .filter_map(|(name, region)| Some((child(style, A, name)?, region)))
        .collect()
}

/// Where a text body is laid out: its rectangle (before the body's insets) and, for a table
/// cell, the vertical anchor its cell properties give.
#[derive(Debug, Clone, Copy)]
struct TextFrame<'s> {
    rect: (f64, f64, f64, f64),
    anchor: Option<&'s str>,
    /// A table style's text color and weight, over the list styles' defaults and under the
    /// runs' own.
    color: Option<Rgba>,
    bold: Option<bool>,
}

/// A placeholder's type and index.
#[derive(Debug, Clone)]
struct Placeholder {
    kind: Option<String>,
    idx: Option<String>,
}

impl Placeholder {
    /// The kind a layout or master placeholder answers to: a centered title is a title, a
    /// subtitle or an untyped (object) placeholder is a body.
    fn kind(&self) -> &str {
        match self.kind.as_deref() {
            Some("title" | "ctrTitle") => "title",
            None | Some("body" | "subTitle" | "obj") => "body",
            Some(other) => other,
        }
    }
}

fn placeholder(node: Node) -> Option<Placeholder> {
    let ph = node
        .children()
        .find(|n| n.tag_name().name().starts_with("nv"))?
        .children()
        .find(|n| n.has_tag_name((P, "nvPr")))
        .and_then(|n| child(n, P, "ph"))?;
    Some(Placeholder {
        kind: ph.attribute("type").map(str::to_string),
        idx: ph.attribute("idx").map(str::to_string),
    })
}

fn placeholders_in<'a, 'i>(
    root: Node<'a, 'i>,
) -> impl Iterator<Item = (Node<'a, 'i>, Placeholder)> {
    child(root, P, "cSld")
        .and_then(|c| child(c, P, "spTree"))
        .into_iter()
        .flat_map(|t| t.children())
        .filter(|n| n.has_tag_name((P, "sp")))
        .filter_map(|n| Some((n, placeholder(n)?)))
}

fn count_runs(node: Node) -> u32 {
    node.descendants()
        .filter(|n| n.has_tag_name((A, "r")))
        .filter(|r| {
            r.descendants()
                .any(|t| t.has_tag_name((A, "t")) && t.text().is_some_and(|s| !s.trim().is_empty()))
        })
        .count() as u32
}

/// A run's character properties, as far as painting needs them.
#[derive(Debug, Clone)]
struct RunStyle {
    /// Points.
    size: f32,
    bold: bool,
    italic: bool,
    color: Option<Rgba>,
    /// The text's fill is drawn as a stand-in: a gradient in its first color, a pattern in its
    /// foreground color.
    approximated_fill: bool,
    latin: Option<String>,
    ea: Option<String>,
}

impl RunStyle {
    fn default_with(color: Option<Rgba>) -> Self {
        Self {
            size: 18.0,
            bold: false,
            italic: false,
            color,
            approximated_fill: false,
            latin: None,
            ea: None,
        }
    }

    /// Apply a run property element (`rPr`, `defRPr`, `endParaRPr`).
    fn apply(&mut self, rpr: Node, colors: &Colors) {
        if let Some(sz) = num(rpr, "sz") {
            self.size = (sz / 100.0) as f32;
        }
        if let Some(b) = rpr.attribute("b") {
            self.bold = b == "1" || b == "true";
        }
        if let Some(i) = rpr.attribute("i") {
            self.italic = i == "1" || i == "true";
        }
        // The text's fill: a color, none (the text is not seen), or a gradient or pattern
        // drawn in one of its colors.
        for fill in rpr.children().filter(|n| n.is_element()) {
            let (color, approximated) = match fill.tag_name().name() {
                "solidFill" => (color_child(fill, colors, None), false),
                "noFill" => (
                    Some(Rgba {
                        a: 0.0,
                        ..Rgba::BLACK
                    }),
                    false,
                ),
                "gradFill" => (
                    gradient(fill, colors, None).and_then(|(g, _)| g.stops.first().map(|s| s.1)),
                    true,
                ),
                "pattFill" => (
                    child(fill, A, "fgClr").and_then(|c| color_child(c, colors, None)),
                    true,
                ),
                _ => continue,
            };
            if let Some(color) = color {
                self.color = Some(color);
                self.approximated_fill = approximated;
            }
            break;
        }
        if let Some(f) = child(rpr, A, "latin").and_then(|n| n.attribute("typeface")) {
            self.latin = Some(f.to_string());
        }
        if let Some(f) = child(rpr, A, "ea").and_then(|n| n.attribute("typeface")) {
            self.ea = Some(f.to_string());
        }
    }
}

/// The theme's major (headings) and minor (body) fonts.
#[derive(Debug, Default)]
struct ThemeFonts {
    major: FontSet,
    minor: FontSet,
}

#[derive(Debug, Default)]
struct FontSet {
    latin: String,
    ea: String,
    /// Per-script faces (`Hang`, `Jpan`, `Hans`, …).
    scripts: HashMap<String, String>,
}

impl ThemeFonts {
    fn read(theme: Option<&roxmltree::Document>) -> Self {
        let set = |name: &str| -> FontSet {
            let Some(n) = theme.and_then(|t| t.descendants().find(|n| n.has_tag_name((A, name))))
            else {
                return FontSet::default();
            };
            let face = |tag: &str| {
                child(n, A, tag)
                    .and_then(|c| c.attribute("typeface"))
                    .unwrap_or("")
                    .to_string()
            };
            FontSet {
                latin: face("latin"),
                ea: face("ea"),
                scripts: n
                    .children()
                    .filter(|c| c.has_tag_name((A, "font")))
                    .filter_map(|c| {
                        Some((
                            c.attribute("script")?.to_string(),
                            c.attribute("typeface")?.to_string(),
                        ))
                    })
                    .collect(),
            }
        };
        Self {
            major: set("majorFont"),
            minor: set("minorFont"),
        }
    }

    /// A typeface, with a theme reference (`+mj-lt`, `+mn-ea`, …) resolved for `ch`.
    fn resolve(&self, typeface: &str, ch: char) -> Option<String> {
        let (set, slot) = match typeface {
            t if t.starts_with("+mj-") => (&self.major, &t[4..]),
            t if t.starts_with("+mn-") => (&self.minor, &t[4..]),
            t => return Some(t.to_string()),
        };
        match slot {
            "ea" => {
                if !set.ea.is_empty() {
                    return Some(set.ea.clone());
                }
                let scripts: &[&str] = match ch as u32 {
                    0x1100..=0x11FF | 0x3130..=0x318F | 0xAC00..=0xD7AF => &["Hang"],
                    0x3040..=0x30FF => &["Jpan"],
                    _ => &["Hang", "Jpan", "Hans", "Hant"],
                };
                scripts.iter().find_map(|s| set.scripts.get(*s).cloned())
            }
            _ => Some(set.latin.clone()),
        }
    }
}

/// A character placed on a line: its face (none when no face has it), size and advance (EMU).
/// A picture bullet is a glyph that draws a picture (a decoded part's path) instead.
#[derive(Debug, Clone)]
struct Glyph {
    ch: char,
    face: Option<usize>,
    size: f32,
    advance: f32,
    color: Option<Rgba>,
    picture: Option<String>,
}

impl Glyph {
    fn new(ch: char, face: Option<usize>, size: f32, advance: f32, color: Option<Rgba>) -> Self {
        Self {
            ch,
            face,
            size,
            advance,
            color,
            picture: None,
        }
    }
}

enum Item {
    Glyph(Glyph),
    Break,
}

#[derive(Debug, Clone)]
struct Line {
    glyphs: Vec<Glyph>,
    /// The height of an empty line (the paragraph-end size).
    empty_size: f32,
    align: String,
    left: f32,
    spacing: f32,
}

impl Line {
    fn empty(size: f32) -> Self {
        Self {
            glyphs: Vec::new(),
            empty_size: size,
            align: String::new(),
            left: 0.0,
            spacing: 1.0,
        }
    }

    fn width(&self) -> f32 {
        let trailing = self
            .glyphs
            .iter()
            .rev()
            .take_while(|g| g.ch == ' ')
            .map(|g| g.advance)
            .sum::<f32>();
        self.glyphs.iter().map(|g| g.advance).sum::<f32>() - trailing
    }

    fn height(&self) -> f32 {
        self.glyphs
            .iter()
            .map(|g| g.size)
            .fold(0.0, f32::max)
            .max(if self.glyphs.is_empty() {
                self.empty_size
            } else {
                0.0
            })
            * 1.2
    }
}

/// Break `items` into lines no wider than `widths` (first line, the rest) allow — at spaces,
/// and between CJK characters; a word wider than a whole line is broken where it overflows.
fn break_lines(items: &[Item], widths: Option<(f32, f32)>, lines: &mut Vec<Line>) {
    let mut line: Vec<Glyph> = Vec::new();
    let mut width = 0.0f32;
    let mut last_break: Option<usize> = None;
    let push = |lines: &mut Vec<Line>, glyphs: Vec<Glyph>| {
        let size = glyphs.iter().map(|g| g.size).fold(0.0, f32::max);
        let mut l = Line::empty(size);
        l.glyphs = glyphs;
        lines.push(l);
    };
    let limit = |lines_so_far: usize, first_index: usize| {
        widths.map(|(first, rest)| {
            if lines_so_far == first_index {
                first
            } else {
                rest
            }
        })
    };
    let first_index = lines.len();
    for item in items {
        match item {
            Item::Break => {
                push(lines, std::mem::take(&mut line));
                width = 0.0;
                last_break = None;
            }
            Item::Glyph(g) => {
                let max = limit(lines.len(), first_index);
                let overflow =
                    max.is_some_and(|m| width + g.advance > m) && g.ch != ' ' && !line.is_empty();
                if overflow {
                    let at = last_break.unwrap_or(line.len());
                    let rest: Vec<Glyph> = line.split_off(at);
                    let rest: Vec<Glyph> = rest.into_iter().skip_while(|g| g.ch == ' ').collect();
                    push(lines, std::mem::take(&mut line));
                    line = rest;
                    width = line.iter().map(|g| g.advance).sum();
                    last_break = None;
                }
                if breaks_anywhere(g.ch) && !line.is_empty() {
                    last_break = Some(line.len());
                }
                width += g.advance;
                line.push(g.clone());
                if g.ch == ' ' || breaks_anywhere(g.ch) {
                    last_break = Some(line.len());
                }
            }
        }
    }
    if !line.is_empty() {
        push(lines, line);
    }
}

/// A paragraph's bullet.
#[derive(Debug, Clone)]
enum Bullet<'a, 'i> {
    Char(char),
    /// A number in `scheme` (`arabicPeriod`, `alphaLcParenR`, …), counting from `start`.
    Number {
        scheme: String,
        start: u32,
    },
    /// The picture the `a:buBlip` element names.
    Picture(Node<'a, 'i>),
}

/// The bullet a paragraph draws, from its list levels (nearest last); `None` with no bullet.
fn bullet<'a, 'i>(levels: &[Node<'a, 'i>]) -> Option<Bullet<'a, 'i>> {
    for n in levels.iter().rev() {
        for c in n.children().filter(|c| c.is_element()) {
            match c.tag_name().name() {
                "buNone" => return None,
                "buChar" => return c.attribute("char")?.chars().next().map(Bullet::Char),
                "buAutoNum" => {
                    return Some(Bullet::Number {
                        scheme: c.attribute("type").unwrap_or("arabicPeriod").to_string(),
                        start: num(c, "startAt").map_or(1, |v| v.max(1.0) as u32),
                    })
                }
                "buBlip" => return Some(Bullet::Picture(c)),
                _ => {}
            }
        }
    }
    None
}

/// The running numbers of a text body's numbered paragraphs, one per list level.
///
/// A number continues from the paragraph before at the same level when that one was numbered
/// in the same scheme; a paragraph at an outer level ends the deeper lists, and one at the same
/// level without a number ends its list. A paragraph with no text is not drawn but counts.
#[derive(Default)]
struct Numbering {
    /// For each level: the scheme and the last number given.
    levels: [Option<(String, u32)>; 9],
}

impl Numbering {
    fn next(&mut self, level: usize, bullet: Option<&Bullet>) -> Option<u32> {
        let level = level.min(8);
        for deeper in &mut self.levels[level + 1..] {
            *deeper = None;
        }
        match bullet {
            Some(Bullet::Number { scheme, start }) => {
                let n = match &self.levels[level] {
                    Some((s, n)) if s == scheme => n + 1,
                    _ => *start,
                };
                self.levels[level] = Some((scheme.clone(), n));
                Some(n)
            }
            _ => {
                self.levels[level] = None;
                None
            }
        }
    }
}

/// `n` written in a DrawingML autonumber `scheme`. Schemes this does not know are written as
/// `arabicPeriod`.
fn autonumber(scheme: &str, n: u32) -> String {
    const SUFFIXES: [(&str, &str, &str); 5] = [
        ("ParenBoth", "(", ")"),
        ("ParenR", "", ")"),
        ("Period", "", "."),
        ("Plain", "", ""),
        ("Minus", "", "-"),
    ];
    let digits = |n: u32| n.to_string();
    let fullwidth = |n: u32| {
        n.to_string()
            .chars()
            .map(|c| char::from_u32(0xFF10 + (c as u32 - '0' as u32)).unwrap_or(c))
            .collect::<String>()
    };
    let alpha = |n: u32, upper: bool| {
        let letter = (b'a' + ((n - 1) % 26) as u8) as char;
        let letter = if upper {
            letter.to_ascii_uppercase()
        } else {
            letter
        };
        std::iter::repeat_n(letter, ((n - 1) / 26 + 1) as usize).collect::<String>()
    };
    let roman = |n: u32, upper: bool| {
        const NUMERALS: [(u32, &str); 13] = [
            (1000, "m"),
            (900, "cm"),
            (500, "d"),
            (400, "cd"),
            (100, "c"),
            (90, "xc"),
            (50, "l"),
            (40, "xl"),
            (10, "x"),
            (9, "ix"),
            (5, "v"),
            (4, "iv"),
            (1, "i"),
        ];
        let mut rest = n;
        let mut out = String::new();
        for (value, numeral) in NUMERALS {
            while rest >= value {
                out.push_str(numeral);
                rest -= value;
            }
        }
        if upper {
            out.to_ascii_uppercase()
        } else {
            out
        }
    };
    let cjk = |n: u32| {
        const DIGITS: [char; 10] = ['〇', '一', '二', '三', '四', '五', '六', '七', '八', '九'];
        match n {
            0..=9 => DIGITS[n as usize].to_string(),
            10..=99 => {
                let (tens, ones) = (n / 10, n % 10);
                let mut out = String::new();
                if tens > 1 {
                    out.push(DIGITS[tens as usize]);
                }
                out.push('十');
                if ones > 0 {
                    out.push(DIGITS[ones as usize]);
                }
                out
            }
            _ => n.to_string(),
        }
    };
    let circled = |n: u32, black: bool| {
        match (n, black) {
            (1..=20, false) => char::from_u32(0x2460 + n - 1).map(String::from),
            (1..=10, true) => char::from_u32(0x2776 + n - 1).map(String::from),
            (11..=20, true) => char::from_u32(0x24EB + n - 11).map(String::from),
            _ => None,
        }
        .unwrap_or_else(|| n.to_string())
    };

    match scheme {
        "circleNumDbPlain" | "circleNumWdWhitePlain" => return circled(n, false),
        "circleNumWdBlackPlain" => return circled(n, true),
        "arabicDbPlain" => return fullwidth(n),
        "arabicDbPeriod" => return format!("{}．", fullwidth(n)),
        "ea1JpnKorPlain" | "ea1ChsPlain" | "ea1ChtPlain" => return cjk(n),
        "ea1JpnKorPeriod" => return format!("{}．", cjk(n)),
        "ea1ChsPeriod" | "ea1ChtPeriod" => return format!("{}、", cjk(n)),
        _ => {}
    }
    let (kind, rest) = ["alphaLc", "alphaUc", "romanLc", "romanUc", "arabic"]
        .iter()
        .find_map(|k| scheme.strip_prefix(k).map(|r| (*k, r)))
        .unwrap_or(("arabic", "Period"));
    let body = match kind {
        "alphaLc" => alpha(n, false),
        "alphaUc" => alpha(n, true),
        "romanLc" => roman(n, false),
        "romanUc" => roman(n, true),
        _ => digits(n),
    };
    let (open, close) = SUFFIXES
        .iter()
        .find(|(name, _, _)| rest.ends_with(name))
        .map_or(("", "."), |(_, o, c)| (*o, *c));
    format!("{open}{body}{close}")
}

/// A picture as pixels: PNG and JPEG; anything else (EMF, WMF, TIFF, …) is not decoded.
fn decode_image(bytes: &[u8]) -> Option<Pixmap> {
    if bytes.starts_with(b"\x89PNG") {
        return Pixmap::decode_png(bytes).ok();
    }
    if bytes.starts_with(&[0xFF, 0xD8]) {
        use zune_jpeg::zune_core::bytestream::ZCursor;
        use zune_jpeg::zune_core::colorspace::ColorSpace;
        use zune_jpeg::zune_core::options::DecoderOptions;
        let options = DecoderOptions::default().jpeg_set_out_colorspace(ColorSpace::RGB);
        let mut decoder = zune_jpeg::JpegDecoder::new_with_options(ZCursor::new(bytes), options);
        let rgb = decoder.decode().ok()?;
        let info = decoder.info()?;
        let (w, h) = (u32::from(info.width), u32::from(info.height));
        if rgb.len() != w as usize * h as usize * 3 {
            return None;
        }
        let rgba: Vec<u8> = rgb
            .as_chunks::<3>()
            .0
            .iter()
            .flat_map(|p| [p[0], p[1], p[2], 255])
            .collect();
        return Pixmap::from_vec(rgba, IntSize::from_wh(w, h)?);
    }
    None
}

fn is_placeholder(node: Node) -> bool {
    node.descendants()
        .take(8)
        .any(|n| n.has_tag_name((P, "ph")))
}

/// The outline of a shape `w` × `h` (EMU): its preset geometry with the shape's adjust values.
/// The fill elements of a shape's properties.
const FILLS: &[&str] = &[
    "noFill",
    "solidFill",
    "gradFill",
    "pattFill",
    "blipFill",
    "grpFill",
];

/// A shape-properties element's fill, if it has one.
fn fill_element<'a, 'i>(sp_pr: Node<'a, 'i>) -> Option<Node<'a, 'i>> {
    sp_pr
        .children()
        .filter(|n| n.is_element())
        .find(|n| FILLS.contains(&n.tag_name().name()))
}

fn shape_geometry(sp_pr: Node, w: f64, h: f64) -> Option<geometry::Geometry> {
    let prst = child(sp_pr, A, "prstGeom")?;
    let def = geometry::preset(prst.attribute("prst")?)?;
    let adjust: Vec<(&str, f64)> = child(prst, A, "avLst")
        .into_iter()
        .flat_map(|av| av.children().filter(|n| n.has_tag_name((A, "gd"))))
        .filter_map(|gd| {
            let mut f = gd.attribute("fmla")?.split_whitespace();
            match (f.next(), f.next()) {
                (Some("val"), Some(v)) => Some((gd.attribute("name")?, v.parse().ok()?)),
                _ => None,
            }
        })
        .collect();
    Some(geometry::evaluate(def, w, h, &adjust))
}

/// A shape's placement: flipped and rotated about its centre, then moved to its offset.
fn shape_transform(xfrm: Node, x: f64, y: f64, w: f64, h: f64) -> Transform {
    let (w, h) = (w as f32, h as f32);
    let mut t = Transform::identity();
    if xfrm.attribute("flipH") == Some("1") {
        t = t.post_scale(-1.0, 1.0).post_translate(w, 0.0);
    }
    if xfrm.attribute("flipV") == Some("1") {
        t = t.post_scale(1.0, -1.0).post_translate(0.0, h);
    }
    if let Some(rot) = num(xfrm, "rot") {
        t = t.post_concat(Transform::from_rotate_at(
            (rot / 60_000.0) as f32,
            w / 2.0,
            h / 2.0,
        ));
    }
    t.post_translate(x as f32, y as f32)
}

/// A group's mapping from its children's coordinate space (`chOff`/`chExt`) to its own place.
fn group_transform(xfrm: Node) -> Transform {
    let pair = |name: &str, a: &str, b: &str| {
        child(xfrm, A, name).map(|n| {
            (
                num(n, a).unwrap_or(0.0) as f32,
                num(n, b).unwrap_or(0.0) as f32,
            )
        })
    };
    let (ox, oy) = pair("off", "x", "y").unwrap_or((0.0, 0.0));
    let (ex, ey) = pair("ext", "cx", "cy").unwrap_or((0.0, 0.0));
    let (cox, coy) = pair("chOff", "x", "y").unwrap_or((ox, oy));
    let (cex, cey) = pair("chExt", "cx", "cy").unwrap_or((ex, ey));
    let sx = if cex != 0.0 { ex / cex } else { 1.0 };
    let sy = if cey != 0.0 { ey / cey } else { 1.0 };
    let mut t = Transform::from_translate(-cox, -coy).post_scale(sx, sy);
    if xfrm.attribute("flipH") == Some("1") {
        t = t.post_scale(-1.0, 1.0).post_translate(ex, 0.0);
    }
    if xfrm.attribute("flipV") == Some("1") {
        t = t.post_scale(1.0, -1.0).post_translate(0.0, ey);
    }
    if let Some(rot) = num(xfrm, "rot") {
        t = t.post_concat(Transform::from_rotate_at(
            (rot / 60_000.0) as f32,
            ex / 2.0,
            ey / 2.0,
        ));
    }
    t.post_translate(ox, oy)
}

/// Lighter (`amount` > 0) or darker toward white or black; unchanged at 0.
fn shaded(c: Rgba, amount: f32) -> Rgba {
    let mix = |v: f32| {
        if amount > 0.0 {
            v + (1.0 - v) * amount
        } else {
            v * (1.0 + amount)
        }
    };
    Rgba {
        r: mix(c.r),
        g: mix(c.g),
        b: mix(c.b),
        a: c.a,
    }
}

/// How a shape's outline is stroked: its color (none for no line), width in EMU (the default
/// hairline when unset), dash, and end decorations.
#[derive(Debug, Clone, Copy)]
struct LineStyle {
    color: Option<Rgba>,
    width: Option<f32>,
    dashed: bool,
    ends: LineEnds,
}

/// The decorations at the two ends of a line (`a:headEnd` at its start, `a:tailEnd` at its end).
#[derive(Debug, Clone, Copy, Default)]
struct LineEnds {
    head: Option<LineEnd>,
    tail: Option<LineEnd>,
}

/// One line end: its kind, and its width and length as multiples of the line's width.
#[derive(Debug, Clone, Copy, PartialEq)]
struct LineEnd {
    kind: LineEndKind,
    width: f32,
    length: f32,
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum LineEndKind {
    Triangle,
    Stealth,
    Diamond,
    Oval,
    /// An open arrow: two strokes meeting at the tip.
    Arrow,
}

/// The `a:headEnd` or `a:tailEnd` of a line, unless it is absent or `none`. Sizes `sm`, `med`
/// (the default) and `lg` are 2, 3 and 5 times the line's width, as PowerPoint draws them.
fn line_end(ln: Node, name: &str) -> Option<LineEnd> {
    let end = child(ln, A, name)?;
    let kind = match end.attribute("type")? {
        "triangle" => LineEndKind::Triangle,
        "stealth" => LineEndKind::Stealth,
        "diamond" => LineEndKind::Diamond,
        "oval" => LineEndKind::Oval,
        "arrow" => LineEndKind::Arrow,
        _ => return None,
    };
    let size = |attr: &str| match end.attribute(attr) {
        Some("sm") => 2.0,
        Some("lg") => 5.0,
        _ => 3.0,
    };
    Some(LineEnd {
        kind,
        width: size("w"),
        length: size("len"),
    })
}

/// A line's arrowheads, and the line itself, shortened where a filled head covers its end.
struct EndedLine {
    line: Vec<Segment>,
    heads: Vec<Arrowhead>,
}

enum Arrowhead {
    Fill(tiny_skia::Path),
    Stroke(tiny_skia::Path),
}

/// Arrowheads for the ends of `segments`, if it is an open path with any. Sizes scale with the
/// line `width`, which is never taken below 2 pt so a hairline still gets a visible head.
fn arrowheads(segments: &[Segment], ends: LineEnds, width: f32) -> Option<EndedLine> {
    if ends.head.is_none() && ends.tail.is_none() || segments.contains(&Segment::Close) {
        return None;
    }
    let base = f64::from(width.max(25_400.0));
    let mut line = segments.to_vec();
    let mut heads = Vec::new();
    let mut place = |end: LineEnd, at: usize, tip: geometry::Point, toward: geometry::Point| {
        let (dx, dy) = (tip.x - toward.x, tip.y - toward.y);
        let norm = dx.hypot(dy);
        if norm == 0.0 {
            return;
        }
        let u = (dx / norm, dy / norm);
        let (half_w, len) = (
            base * f64::from(end.width) / 2.0,
            base * f64::from(end.length),
        );
        // Local coordinates: x along the line toward the tip, y across it; the tip at 0.
        let local = |x: f64, y: f64| {
            (
                (tip.x + u.0 * x - u.1 * y) as f32,
                (tip.y + u.1 * x + u.0 * y) as f32,
            )
        };
        let polygon = |points: &[(f64, f64)]| {
            let mut pb = PathBuilder::new();
            for (i, &(x, y)) in points.iter().enumerate() {
                let (px, py) = local(x, y);
                if i == 0 {
                    pb.move_to(px, py);
                } else {
                    pb.line_to(px, py);
                }
            }
            pb
        };
        let (head, setback) = match end.kind {
            LineEndKind::Triangle => {
                let mut pb = polygon(&[(0.0, 0.0), (-len, half_w), (-len, -half_w)]);
                pb.close();
                (pb.finish().map(Arrowhead::Fill), len)
            }
            LineEndKind::Stealth => {
                let mut pb = polygon(&[
                    (0.0, 0.0),
                    (-len, half_w),
                    (-len / 2.0, 0.0),
                    (-len, -half_w),
                ]);
                pb.close();
                (pb.finish().map(Arrowhead::Fill), len / 2.0)
            }
            LineEndKind::Diamond => {
                let mut pb = polygon(&[
                    (len / 2.0, 0.0),
                    (0.0, half_w),
                    (-len / 2.0, 0.0),
                    (0.0, -half_w),
                ]);
                pb.close();
                (pb.finish().map(Arrowhead::Fill), 0.0)
            }
            LineEndKind::Oval => {
                let oval = tiny_skia::Rect::from_xywh(
                    (-len / 2.0) as f32,
                    -half_w as f32,
                    len as f32,
                    (2.0 * half_w) as f32,
                )
                .and_then(PathBuilder::from_oval)
                .and_then(|p| {
                    p.transform(Transform::from_row(
                        u.0 as f32,
                        u.1 as f32,
                        -u.1 as f32,
                        u.0 as f32,
                        tip.x as f32,
                        tip.y as f32,
                    ))
                });
                (oval.map(Arrowhead::Fill), 0.0)
            }
            LineEndKind::Arrow => {
                let pb = polygon(&[(-len, half_w), (0.0, 0.0), (-len, -half_w)]);
                (pb.finish().map(Arrowhead::Stroke), 0.0)
            }
        };
        heads.extend(head);
        // Pull a straight end segment back to the head's base, when it is long enough to keep.
        if setback > 0.0 {
            let shortened = geometry::Point {
                x: tip.x - u.0 * setback,
                y: tip.y - u.1 * setback,
            };
            if norm > setback {
                match line.get_mut(at) {
                    Some(Segment::Move(p)) | Some(Segment::Line(p)) => *p = shortened,
                    _ => {}
                }
            }
        }
    };

    if let Some(end) = ends.head {
        // The start: the first point, and the first point after it that differs from it.
        if let Some((tip, toward, straight)) = path_start(segments) {
            let at = if straight { 0 } else { usize::MAX };
            place(end, at, tip, toward);
        }
    }
    if let Some(end) = ends.tail {
        if let Some((tip, toward, straight)) = path_end(segments) {
            let at = if straight {
                segments.len() - 1
            } else {
                usize::MAX
            };
            place(end, at, tip, toward);
        }
    }
    Some(EndedLine { line, heads })
}

/// Where a path starts, the point its first segment heads for, and whether that segment is a
/// straight line from it.
fn path_start(segments: &[Segment]) -> Option<(geometry::Point, geometry::Point, bool)> {
    let Some(Segment::Move(start)) = segments.first() else {
        return None;
    };
    let next = segments.get(1).and_then(|s| match *s {
        Segment::Line(a) => Some((a, true)),
        Segment::Quad(c, a) => Some((if c == *start { a } else { c }, false)),
        Segment::Cubic(c1, c2, a) => Some((
            if c1 != *start {
                c1
            } else if c2 != *start {
                c2
            } else {
                a
            },
            false,
        )),
        _ => None,
    })?;
    Some((*start, next.0, next.1))
}

/// Where a path ends, the point its last segment comes from, and whether that segment is a
/// straight line to it.
fn path_end(segments: &[Segment]) -> Option<(geometry::Point, geometry::Point, bool)> {
    let last = segments.len().checked_sub(1)?;
    let anchor = |s: &Segment| match *s {
        Segment::Move(a) | Segment::Line(a) | Segment::Quad(_, a) | Segment::Cubic(_, _, a) => {
            Some(a)
        }
        Segment::Close => None,
    };
    let end = anchor(&segments[last])?;
    let previous = segments[..last].iter().rev().find_map(anchor)?;
    match segments[last] {
        Segment::Line(_) => Some((end, previous, true)),
        Segment::Quad(c, _) => Some((end, if c == end { previous } else { c }, false)),
        Segment::Cubic(c1, c2, _) => Some((
            end,
            if c2 != end {
                c2
            } else if c1 != end {
                c1
            } else {
                previous
            },
            false,
        )),
        _ => None,
    }
}

fn to_path(segments: &[Segment]) -> Option<tiny_skia::Path> {
    let mut pb = PathBuilder::new();
    let p = |pt: geometry::Point| (pt.x as f32, pt.y as f32);
    for s in segments {
        match *s {
            Segment::Move(a) => {
                let (x, y) = p(a);
                pb.move_to(x, y);
            }
            Segment::Line(a) => {
                let (x, y) = p(a);
                pb.line_to(x, y);
            }
            Segment::Quad(c, a) => {
                let ((cx, cy), (x, y)) = (p(c), p(a));
                pb.quad_to(cx, cy, x, y);
            }
            Segment::Cubic(c1, c2, a) => {
                let ((x1, y1), (x2, y2), (x, y)) = (p(c1), p(c2), p(a));
                pb.cubic_to(x1, y1, x2, y2, x, y);
            }
            Segment::Close => pb.close(),
        }
    }
    pb.finish()
}

#[cfg(test)]
pub(crate) mod tests;
