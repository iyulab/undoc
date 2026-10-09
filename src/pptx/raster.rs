//! A slide painted to pixels (feature `raster`).
//!
//! The slide is read as a scene — background, then the master's and layout's own shapes, then
//! the slide's shape tree — and each shape's outline comes from [`crate::geometry`], filled and
//! stroked with tiny-skia. Colors resolve through the master's color map and the theme.
//!
//! Not painted yet, and counted in the gaps instead: pictures, text, charts, tables and other
//! graphic frames, custom geometry.

use std::collections::HashMap;

use roxmltree::Node;
use tiny_skia::{Color, FillRule, Paint, PathBuilder, Pixmap, Stroke, StrokeDash, Transform};

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

        let gaps = painter.gaps;
        Ok(RasteredSlide {
            width,
            height,
            rgba: pixmap.take(),
            gaps,
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
}

impl Painter<'_> {
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
        // Text is not painted yet; say how much of it is missing.
        let runs = node
            .descendants()
            .filter(|n| n.has_tag_name((A, "r")))
            .filter(|r| {
                r.descendants()
                    .any(|t| t.has_tag_name((A, "t")) && t.text().is_some_and(|s| !s.is_empty()))
            })
            .count();
        self.gaps.text_runs += runs as u32;

        let Some(sp_pr) = child(node, P, "spPr") else {
            return;
        };
        let Some(xfrm) = child(sp_pr, A, "xfrm") else {
            // A placeholder without its own position takes its layout's; not read yet.
            return;
        };
        let (Some(off), Some(ext)) = (child(xfrm, A, "off"), child(xfrm, A, "ext")) else {
            return;
        };
        let (x, y) = (num(off, "x").unwrap_or(0.0), num(off, "y").unwrap_or(0.0));
        let (w, h) = (num(ext, "cx").unwrap_or(0.0), num(ext, "cy").unwrap_or(0.0));

        let style = child(node, P, "style");
        let style_color = |name: &str| {
            style
                .and_then(|s| child(s, A, name))
                .filter(|r| r.attribute("idx") != Some("0"))
                .and_then(|r| color_child(r, self.colors, None))
        };
        let shape_fill = sp_pr
            .children()
            .filter(|n| n.is_element())
            .find_map(|n| fill(n, self.colors, &mut self.gaps))
            .unwrap_or_else(|| style_color("fillRef"));

        let ln = child(sp_pr, A, "ln");
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

        if shape_fill.is_none() && line_fill.is_none() {
            return;
        }
        let Some(geometry) = shape_geometry(sp_pr, w, h) else {
            self.gaps.shapes += 1;
            return;
        };

        let transform = shape_transform(xfrm, x, y, w, h).post_concat(parent);
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
}

fn is_placeholder(node: Node) -> bool {
    node.descendants()
        .take(8)
        .any(|n| n.has_tag_name((P, "ph")))
}

/// The outline of a shape `w` × `h` (EMU): its preset geometry with the shape's adjust values.
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
