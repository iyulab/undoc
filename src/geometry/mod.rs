//! DrawingML shape geometry: the outline a shape draws, from its preset name or custom
//! geometry and its size.
//!
//! A DrawingML shape does not carry its outline. `<a:prstGeom prst="rightArrow">` names one of
//! the preset shapes ECMA-376 defines as data ([`presets`]): adjust values a shape may
//! override (`<a:avLst>`), guide formulas evaluated against the shape's width and height, and
//! paths whose coordinates are guide names or numbers. `<a:custGeom>` carries the same kind of
//! definition inline. Both are evaluated here into plain paths in the shape's own coordinate
//! space (origin at its top-left, the units its width and height are given in).
//!
//! Angles are in 60,000ths of a degree, as everywhere in DrawingML.

// Generated: rustfmt would rewrap the one-line definitions and the file would stop matching
// what `scripts/gen_preset_geometry.py` writes.
#[rustfmt::skip]
mod presets;

use std::collections::HashMap;
use std::f64::consts::PI;

/// A named guide: `name = fmla`, where `fmla` is an operator and its arguments
/// (ECMA-376 Part 1, 20.1.9.11).
#[derive(Debug, Clone, Copy)]
pub(crate) struct Guide<'a> {
    pub name: &'a str,
    pub fmla: &'a str,
}

/// How a path is filled. The lighter and darker variants shade the shape's fill; a path with
/// `None` is outline only.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum PathFill {
    Norm,
    None,
    Lighten,
    LightenLess,
    Darken,
    DarkenLess,
}

/// A path command; its coordinates and angles are guide names or numbers.
#[derive(Debug, Clone, Copy)]
pub(crate) enum Cmd<'a> {
    Move(&'a str, &'a str),
    Line(&'a str, &'a str),
    Arc {
        wr: &'a str,
        hr: &'a str,
        st: &'a str,
        sw: &'a str,
    },
    Quad([&'a str; 4]),
    Cubic([&'a str; 6]),
    Close,
}

/// One path of a shape. With `w`/`h`, its coordinates are in a coordinate space of that size,
/// stretched over the shape.
#[derive(Debug, Clone, Copy)]
pub(crate) struct PathDef<'a> {
    pub w: Option<i64>,
    pub h: Option<i64>,
    pub fill: PathFill,
    pub stroke: bool,
    pub cmds: &'a [Cmd<'a>],
}

/// A shape's definition: adjust values, guides, text rectangle (`l t r b`) and paths.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Preset<'a> {
    pub name: &'a str,
    pub av: &'a [Guide<'a>],
    pub gd: &'a [Guide<'a>],
    pub text_rect: Option<[&'a str; 4]>,
    pub paths: &'a [PathDef<'a>],
}

/// The definition of a preset shape, by its `prst` name.
pub(crate) fn preset(name: &str) -> Option<&'static Preset<'static>> {
    presets::PRESETS
        .binary_search_by(|p| p.name.cmp(name))
        .ok()
        .map(|i| &presets::PRESETS[i])
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub(crate) struct Point {
    pub x: f64,
    pub y: f64,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub(crate) enum Segment {
    Move(Point),
    Line(Point),
    Cubic(Point, Point, Point),
    Quad(Point, Point),
    Close,
}

/// An evaluated path.
#[derive(Debug, Clone, PartialEq)]
pub(crate) struct Outline {
    pub fill: PathFill,
    pub stroke: bool,
    pub segments: Vec<Segment>,
}

/// An evaluated shape: its paths and the rectangle its text is laid out in.
#[derive(Debug, Clone, PartialEq)]
pub(crate) struct Geometry {
    pub outlines: Vec<Outline>,
    /// `(left, top, right, bottom)`; the whole shape when the definition names none.
    pub text_rect: (f64, f64, f64, f64),
}

/// Evaluate `def` for a shape `w` × `h`, with `adjust` overriding its adjust values by name
/// (a shape's own `<a:avLst>`).
pub(crate) fn evaluate(def: &Preset<'_>, w: f64, h: f64, adjust: &[(&str, f64)]) -> Geometry {
    let mut env = Env::new(w, h);
    for g in def.av {
        let value = adjust
            .iter()
            .find(|(name, _)| *name == g.name)
            .map(|&(_, v)| v)
            .unwrap_or_else(|| env.formula(g.fmla));
        env.set(g.name, value);
    }
    for g in def.gd {
        let value = env.formula(g.fmla);
        env.set(g.name, value);
    }
    let text_rect = match def.text_rect {
        Some([l, t, r, b]) => (env.arg(l), env.arg(t), env.arg(r), env.arg(b)),
        None => (0.0, 0.0, w, h),
    };
    let outlines = def
        .paths
        .iter()
        .map(|path| outline(path, &env, w, h))
        .collect();
    Geometry {
        outlines,
        text_rect,
    }
}

fn outline(path: &PathDef<'_>, env: &Env<'_>, w: f64, h: f64) -> Outline {
    // A path with its own size is stretched over the shape.
    let sx = path.w.filter(|&pw| pw != 0).map_or(1.0, |pw| w / pw as f64);
    let sy = path.h.filter(|&ph| ph != 0).map_or(1.0, |ph| h / ph as f64);
    let pt = |x: &str, y: &str| Point {
        x: env.arg(x) * sx,
        y: env.arg(y) * sy,
    };
    let mut segments = Vec::with_capacity(path.cmds.len());
    let mut current = Point { x: 0.0, y: 0.0 };
    let mut start = current;
    for cmd in path.cmds {
        match *cmd {
            Cmd::Move(x, y) => {
                current = pt(x, y);
                start = current;
                segments.push(Segment::Move(current));
            }
            Cmd::Line(x, y) => {
                current = pt(x, y);
                segments.push(Segment::Line(current));
            }
            Cmd::Quad([x1, y1, x2, y2]) => {
                let (c, p) = (pt(x1, y1), pt(x2, y2));
                current = p;
                segments.push(Segment::Quad(c, p));
            }
            Cmd::Cubic([x1, y1, x2, y2, x3, y3]) => {
                let (c1, c2, p) = (pt(x1, y1), pt(x2, y2), pt(x3, y3));
                current = p;
                segments.push(Segment::Cubic(c1, c2, p));
            }
            Cmd::Arc { wr, hr, st, sw } => {
                let (wr, hr) = (env.arg(wr) * sx, env.arg(hr) * sy);
                current = arc(&mut segments, current, wr, hr, env.arg(st), env.arg(sw));
            }
            Cmd::Close => {
                current = start;
                segments.push(Segment::Close);
            }
        }
    }
    Outline {
        fill: path.fill,
        stroke: path.stroke,
        segments,
    }
}

/// Append an elliptical arc from `from` as cubic Béziers, and return where it ends.
///
/// `st` and `sw` are the start and swing angles (60,000ths of a degree) of the ellipse with
/// radii `wr` × `hr` that passes through `from` at `st`. DrawingML's angles are the visual
/// angle from the ellipse's centre, so each is mapped to the ellipse's parameter angle
/// (`atan2(wr·sin a, hr·cos a)`) before the curve is drawn.
fn arc(out: &mut Vec<Segment>, from: Point, wr: f64, hr: f64, st: f64, sw: f64) -> Point {
    if wr == 0.0 || hr == 0.0 || sw == 0.0 {
        return from;
    }
    let st = st / 60_000.0 * PI / 180.0;
    let sw = sw / 60_000.0 * PI / 180.0;
    let param = |a: f64| (wr * a.sin()).atan2(hr * a.cos());
    let p0 = param(st);
    let centre = Point {
        x: from.x - wr * p0.cos(),
        y: from.y - hr * p0.sin(),
    };
    // The parameter swing has the visual swing's direction; whole turns carry over as they are.
    let turns = (sw / (2.0 * PI)).trunc();
    let mut delta = param(st + sw) - p0;
    let rest = sw - turns * 2.0 * PI;
    if rest > 0.0 && delta <= 0.0 {
        delta += 2.0 * PI;
    } else if rest < 0.0 && delta >= 0.0 {
        delta -= 2.0 * PI;
    } else if rest == 0.0 {
        delta = 0.0;
    }
    delta += turns * 2.0 * PI;
    let pieces = (delta.abs() / (PI / 2.0)).ceil().max(1.0) as usize;
    let step = delta / pieces as f64;
    let k = 4.0 / 3.0 * (step / 4.0).tan();
    let at = |t: f64| Point {
        x: centre.x + wr * t.cos(),
        y: centre.y + hr * t.sin(),
    };
    let mut t = p0;
    let mut end = from;
    for _ in 0..pieces {
        let t1 = t + step;
        let (a, b) = (at(t), at(t1));
        let c1 = Point {
            x: a.x - k * wr * t.sin(),
            y: a.y + k * hr * t.cos(),
        };
        let c2 = Point {
            x: b.x + k * wr * t1.sin(),
            y: b.y - k * hr * t1.cos(),
        };
        out.push(Segment::Cubic(c1, c2, b));
        end = b;
        t = t1;
    }
    end
}

/// The guide values of one shape: the built-in ones its size gives, then its own.
struct Env<'a> {
    values: HashMap<&'a str, f64>,
}

impl<'a> Env<'a> {
    fn new(w: f64, h: f64) -> Self {
        let (ss, ls) = (w.min(h), w.max(h));
        let mut values = HashMap::new();
        for (name, value) in [
            ("w", w),
            ("h", h),
            ("l", 0.0),
            ("t", 0.0),
            ("r", w),
            ("b", h),
            ("hc", w / 2.0),
            ("vc", h / 2.0),
            ("ss", ss),
            ("ls", ls),
            ("wd2", w / 2.0),
            ("wd3", w / 3.0),
            ("wd4", w / 4.0),
            ("wd5", w / 5.0),
            ("wd6", w / 6.0),
            ("wd8", w / 8.0),
            ("wd10", w / 10.0),
            ("wd12", w / 12.0),
            ("wd32", w / 32.0),
            ("hd2", h / 2.0),
            ("hd3", h / 3.0),
            ("hd4", h / 4.0),
            ("hd5", h / 5.0),
            ("hd6", h / 6.0),
            ("hd8", h / 8.0),
            ("hd10", h / 10.0),
            ("hd12", h / 12.0),
            ("hd32", h / 32.0),
            ("ssd2", ss / 2.0),
            ("ssd4", ss / 4.0),
            ("ssd6", ss / 6.0),
            ("ssd8", ss / 8.0),
            ("ssd16", ss / 16.0),
            ("ssd32", ss / 32.0),
            ("cd2", 10_800_000.0),
            ("cd4", 5_400_000.0),
            ("cd8", 2_700_000.0),
            ("3cd4", 16_200_000.0),
            ("3cd8", 8_100_000.0),
            ("5cd8", 13_500_000.0),
            ("7cd8", 18_900_000.0),
        ] {
            values.insert(name, value);
        }
        Self { values }
    }

    fn set(&mut self, name: &'a str, value: f64) {
        self.values.insert(name, value);
    }

    /// A formula argument: a number, or the name of a guide defined before it. An unknown
    /// name reads as 0, as a reader of a malformed definition must read something.
    fn arg(&self, token: &str) -> f64 {
        token
            .parse::<f64>()
            .ok()
            .or_else(|| self.values.get(token).copied())
            .unwrap_or(0.0)
    }

    fn formula(&self, fmla: &str) -> f64 {
        let mut parts = fmla.split_whitespace();
        let op = parts.next().unwrap_or("");
        let a: Vec<f64> = parts.map(|t| self.arg(t)).collect();
        let x = |i: usize| a.get(i).copied().unwrap_or(0.0);
        let angle = |v: f64| v / 60_000.0 * PI / 180.0;
        let to_angle = |rad: f64| rad * 180.0 / PI * 60_000.0;
        match op {
            "val" => x(0),
            "*/" => {
                if x(2) == 0.0 {
                    0.0
                } else {
                    x(0) * x(1) / x(2)
                }
            }
            "+-" => x(0) + x(1) - x(2),
            "+/" => {
                if x(2) == 0.0 {
                    0.0
                } else {
                    (x(0) + x(1)) / x(2)
                }
            }
            "?:" => {
                if x(0) > 0.0 {
                    x(1)
                } else {
                    x(2)
                }
            }
            "abs" => x(0).abs(),
            "at2" => to_angle(x(1).atan2(x(0))),
            "cat2" => x(0) * x(2).atan2(x(1)).cos(),
            "sat2" => x(0) * x(2).atan2(x(1)).sin(),
            "cos" => x(0) * angle(x(1)).cos(),
            "sin" => x(0) * angle(x(1)).sin(),
            "tan" => x(0) * angle(x(1)).tan(),
            "max" => x(0).max(x(1)),
            "min" => x(0).min(x(1)),
            "mod" => (x(0) * x(0) + x(1) * x(1) + x(2) * x(2)).sqrt(),
            "pin" => {
                if x(1) < x(0) {
                    x(0)
                } else if x(1) > x(2) {
                    x(2)
                } else {
                    x(1)
                }
            }
            "sqrt" => x(0).max(0.0).sqrt(),
            _ => 0.0,
        }
    }
}

#[cfg(test)]
mod tests;
