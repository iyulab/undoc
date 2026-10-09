//! A slide painted to pixels (feature `raster`).
//!
//! The slide is read as a scene — background, then the master's and layout's own shapes, then
//! the slide's shape tree — and each shape's outline comes from [`crate::geometry`], filled and
//! stroked with tiny-skia. Colors resolve through the master's color map and the theme.
//!
//! Text is laid out in each shape's text rectangle with the fonts [`super::raster_text`] finds;
//! placeholders take their position, body and list styles from the layout and master.
//!
//! Not painted yet, and counted in the gaps instead: pictures, charts, tables and other graphic
//! frames, custom geometry, vertical text and scripts that need shaping.

use std::collections::HashMap;

use roxmltree::Node;
use tiny_skia::{Color, FillRule, Paint, PathBuilder, Pixmap, Stroke, StrokeDash, Transform};

use super::raster_text::{breaks_anywhere, is_east_asian, needs_shaping, FontBook};
use super::PptxParser;
use crate::container::OoxmlContainer;
use crate::error::{Error, Result};
use crate::geometry::{self, PathFill, Segment};
use crate::raster::{RasteredSlide, SlideRasterGaps, SlideRasterOptions};

const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";

const REL_LAYOUT: &str = "/slideLayout";
const REL_MASTER: &str = "/slideMaster";
const REL_THEME: &str = "/theme";

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
        let info = self.slides.get(index).ok_or_else(|| {
            Error::InvalidData(format!(
                "slide index {index} is out of range: the presentation has {} slides",
                self.slides.len()
            ))
        })?;
        let target = self.relationships.get(&info.rel_id).ok_or_else(|| {
            Error::MissingComponent(format!("slide relationship {}", info.rel_id))
        })?;
        let slide_path = OoxmlContainer::resolve_path("ppt/presentation.xml", target);

        let (cx, cy) = slide_size(&self.container)?;
        let scale = options.dpi / 72.0 / EMU_PER_PT;
        let width = (cx as f32 * scale).round().max(1.0) as u32;
        let height = (cy as f32 * scale).round().max(1.0) as u32;
        if u64::from(width) * u64::from(height) > MAX_PIXELS {
            return Err(Error::InvalidData(format!(
                "a slide of {width} x {height} pixels is too large to rasterize"
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
            theme_fonts: ThemeFonts::read(parts.theme.as_ref()),
            substituted: 0,
        };

        // Background: the slide's own, else the layout's, else the master's; white without.
        let background = parts
            .chain()
            .find_map(|doc| background(doc.root_element(), &colors, &mut painter.gaps))
            .unwrap_or(Rgba::WHITE);
        painter.pixmap.fill(background.to_color());

        let to_pixels = Transform::from_scale(scale, scale);
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

/// The text of the slide and of the parts it inherits from.
struct PartTexts {
    slide: String,
    layout: Option<String>,
    master: Option<String>,
    theme: Option<String>,
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
        Ok(Self {
            slide,
            layout: read(&layout_path)?,
            master: read(&master_path)?,
            theme: read(&theme_path)?,
        })
    }

    fn parse(&self) -> Result<Parts<'_>> {
        Ok(Parts {
            slide: parse(&self.slide)?,
            layout: parse_optional(&self.layout)?,
            master: parse_optional(&self.master)?,
            theme: parse_optional(&self.theme)?,
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
    const WHITE: Rgba = Rgba {
        r: 1.0,
        g: 1.0,
        b: 1.0,
        a: 1.0,
    };
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

/// What a fill element paints: `Some(None)` for `noFill`, `Some(Some(c))` for a color, `None`
/// for anything that is not a fill. Gradients and patterns are painted in one of their colors
/// and counted as approximated.
fn fill(node: Node, colors: &Colors, gaps: &mut SlideRasterGaps) -> Option<Option<Rgba>> {
    match node.tag_name().name() {
        "noFill" => Some(None),
        "solidFill" => Some(color_child(node, colors, None)),
        "gradFill" => {
            gaps.approximated_fills += 1;
            let stop = node
                .descendants()
                .find(|n| n.has_tag_name((A, "gs")))
                .and_then(|gs| color_child(gs, colors, None));
            Some(stop)
        }
        "pattFill" => {
            gaps.approximated_fills += 1;
            Some(child(node, A, "fgClr").and_then(|c| color_child(c, colors, None)))
        }
        "blipFill" => {
            gaps.images += 1;
            Some(None)
        }
        _ => None,
    }
}

fn background(root: Node, colors: &Colors, gaps: &mut SlideRasterGaps) -> Option<Rgba> {
    let bg = child(child(root, P, "cSld")?, P, "bg")?;
    if let Some(pr) = child(bg, P, "bgPr") {
        return pr
            .children()
            .filter(|n| n.is_element())
            .find_map(|n| fill(n, colors, gaps))
            .flatten();
    }
    // A theme background reference; its own color stands in for the theme fill it names.
    child(bg, P, "bgRef").and_then(|r| color_child(r, colors, None))
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
    theme_fonts: ThemeFonts,
    /// Text runs drawn in a face standing in for the one they ask for.
    substituted: u32,
}

impl<'a> Painter<'a> {
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
                    let transform = child(node, P, "grpSpPr")
                        .and_then(|pr| child(pr, A, "xfrm"))
                        .map_or(parent, |x| group_transform(x).post_concat(parent));
                    self.children(node, transform, placeholders);
                }
                "pic" => self.gaps.images += 1,
                "graphicFrame" => {
                    let uri = node
                        .descendants()
                        .find(|n| n.has_tag_name((A, "graphicData")))
                        .and_then(|g| g.attribute("uri"))
                        .unwrap_or("");
                    if uri.ends_with("/chart") {
                        self.gaps.charts += 1;
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
        let sp_pr = child(node, P, "spPr");
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
            self.gaps.text_runs += count_runs(node);
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
        let fill_node = sp_prs.iter().find_map(|s| fill_element(*s));

        let style = std::iter::once(node)
            .chain(inherited.iter().copied())
            .find_map(|n| child(n, P, "style"));
        let style_color = |name: &str| {
            style
                .and_then(|s| child(s, A, name))
                .filter(|r| r.attribute("idx") != Some("0"))
                .and_then(|r| color_child(r, self.colors, None))
        };
        let shape_fill = fill_node
            .and_then(|n| fill(n, self.colors, &mut self.gaps))
            .unwrap_or_else(|| style_color("fillRef"));
        let ln = sp_prs.iter().find_map(|s| child(*s, A, "ln"));
        let line_fill = ln
            .and_then(|l| {
                l.children()
                    .filter(|n| n.is_element())
                    .find_map(|n| fill(n, self.colors, &mut self.gaps))
            })
            .unwrap_or_else(|| style_color("lnRef"));
        let line_width = ln.and_then(|l| num(l, "w")).map(|w| w as f32).or_else(|| {
            style
                .and_then(|s| child(s, A, "lnRef"))
                .and_then(|r| num(r, "idx"))
                .and_then(|i| self.line_widths.get((i as usize).checked_sub(1)?).copied())
        });
        let dashed = ln
            .and_then(|l| child(l, A, "prstDash"))
            .and_then(|d| d.attribute("val"))
            .is_some_and(|v| v != "solid");

        if shape_fill.is_some() || line_fill.is_some() {
            match &geometry {
                Some(geometry) => self.outlines(
                    geometry, shape_fill, line_fill, line_width, dashed, transform,
                ),
                None => self.gaps.shapes += 1,
            }
        }

        if let Some(body) = child(node, P, "txBody") {
            let rect = geometry.as_ref().map_or((0.0, 0.0, w, h), |g| g.text_rect);
            let font_color = style_color("fontRef");
            self.text(body, ph.as_ref(), &inherited, rect, font_color, transform);
        }
    }

    fn outlines(
        &mut self,
        geometry: &geometry::Geometry,
        shape_fill: Option<Rgba>,
        line_fill: Option<Rgba>,
        line_width: Option<f32>,
        dashed: bool,
        transform: Transform,
    ) {
        for outline in &geometry.outlines {
            let Some(path) = to_path(&outline.segments) else {
                continue;
            };
            if let Some(color) = shape_fill {
                let color = match outline.fill {
                    PathFill::None => None,
                    PathFill::Norm => Some(color),
                    PathFill::Lighten => Some(shaded(color, 0.4)),
                    PathFill::LightenLess => Some(shaded(color, 0.2)),
                    PathFill::Darken => Some(shaded(color, -0.4)),
                    PathFill::DarkenLess => Some(shaded(color, -0.2)),
                };
                if let Some(color) = color {
                    let mut paint = Paint::default();
                    paint.set_color(color.to_color());
                    paint.anti_alias = true;
                    self.pixmap
                        .fill_path(&path, &paint, FillRule::EvenOdd, transform, None);
                }
            }
            if let (true, Some(color)) = (outline.stroke, line_fill) {
                let mut paint = Paint::default();
                paint.set_color(color.to_color());
                paint.anti_alias = true;
                let width = line_width.unwrap_or(9_525.0).max(1.0);
                let stroke = Stroke {
                    width,
                    dash: if dashed {
                        StrokeDash::new(vec![width * 4.0, width * 3.0], 0.0)
                    } else {
                        None
                    },
                    ..Stroke::default()
                };
                self.pixmap
                    .stroke_path(&path, &paint, &stroke, transform, None);
            }
        }
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
        rect: (f64, f64, f64, f64),
        font_color: Option<Rgba>,
        transform: Transform,
    ) {
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
        let anchor = attr("anchor").unwrap_or("t");
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

            let mut items: Vec<Item> = Vec::new();
            let mut end_size = base.size;
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
                let text: String = child(run, A, "t")
                    .and_then(|t| t.text())
                    .unwrap_or("")
                    .to_string();
                if text.is_empty() {
                    continue;
                }
                let size = style.size * font_scale * EMU_PER_PT;
                for ch in text.chars() {
                    if ch == '\t' || ch == '\n' || ch == '\r' {
                        items.push(Item::Glyph(Glyph {
                            ch: ' ',
                            face: None,
                            size,
                            advance: size * 0.25,
                            color: style.color,
                        }));
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
                    items.push(Item::Glyph(Glyph {
                        ch,
                        face,
                        size,
                        advance,
                        color: style.color,
                    }));
                }
                run_index += 1;
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
        if let Some(c) = child(rpr, A, "solidFill").and_then(|f| color_child(f, colors, None)) {
            self.color = Some(c);
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
#[derive(Debug, Clone)]
struct Glyph {
    ch: char,
    face: Option<usize>,
    size: f32,
    advance: f32,
    color: Option<Rgba>,
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

/// Lighter (`amount` > 0) or darker toward white or black.
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
mod tests;
