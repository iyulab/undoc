//! Picture bullets a slide inherits from its layout and master.
//!
//! A placeholder on a slide takes its paragraph properties from the layout's placeholder
//! with the same `idx` (or type), which takes them from the master's placeholder of that
//! type, which takes them from the master's `p:txStyles` (`titleStyle` / `bodyStyle` /
//! `otherStyle`). Only the bullet is resolved here: for each list level, either the
//! picture a `a:buBlip` names or "no picture" (`a:buNone`, `a:buChar`, `a:buAutoNum`
//! replace an inherited picture bullet).

use std::collections::HashMap;

use quick_xml::events::{BytesStart, Event};

use crate::container::OoxmlContainer;

/// The bullet of each list level (0-based): `Some(package path)` of a picture, `None`
/// where another bullet kind or none replaces an inherited picture.
pub(super) type BulletLevels = HashMap<u8, Option<String>>;

/// The bullets one layout or master part sets.
#[derive(Debug, Default)]
struct PartBullets {
    /// `titleStyle`, `bodyStyle`, `otherStyle` of `p:txStyles`.
    text_styles: [BulletLevels; 3],
    /// Placeholder shapes: (type, idx, bullets of their list style).
    placeholders: Vec<(String, Option<String>, BulletLevels)>,
}

/// The picture bullets a slide's placeholders inherit, already merged
/// master `txStyles` < master placeholder < layout placeholder.
#[derive(Debug, Default)]
pub(super) struct InheritedBullets {
    by_idx: HashMap<String, BulletLevels>,
    by_type: HashMap<String, BulletLevels>,
}

impl InheritedBullets {
    /// Merge the bullets of a layout over those of its master.
    pub(super) fn new(
        layout: Option<(&str, &HashMap<String, String>)>,
        layout_xml: Option<&str>,
        master: Option<(&str, &HashMap<String, String>)>,
        master_xml: Option<&str>,
    ) -> Self {
        let master_bullets = match (master, master_xml) {
            (Some((path, rels)), Some(xml)) => scan_part(xml, path, rels),
            _ => PartBullets::default(),
        };
        let layout_bullets = match (layout, layout_xml) {
            (Some((path, rels)), Some(xml)) => scan_part(xml, path, rels),
            _ => PartBullets::default(),
        };

        // What the master gives a placeholder of `kind`: its text style, then its own
        // placeholder of that type.
        let from_master = |kind: &str| -> BulletLevels {
            let mut levels = master_bullets.text_styles[text_style_of(kind)].clone();
            let master_kind = master_placeholder_type(kind);
            for (ty, _, own) in &master_bullets.placeholders {
                if ty == master_kind || (ty.is_empty() && master_kind == "body") {
                    levels.extend(own.clone());
                }
            }
            levels
        };

        let mut inherited = Self::default();
        for (ty, idx, own) in &layout_bullets.placeholders {
            let mut levels = from_master(ty);
            levels.extend(own.clone());
            if let Some(idx) = idx {
                inherited.by_idx.insert(idx.clone(), levels.clone());
            }
            if !ty.is_empty() {
                inherited.by_type.insert(ty.clone(), levels);
            }
        }
        for (ty, _, _) in &master_bullets.placeholders {
            if !ty.is_empty() && !inherited.by_type.contains_key(ty) {
                inherited.by_type.insert(ty.clone(), from_master(ty));
            }
        }
        inherited
    }

    /// The bullets a slide placeholder with this type and index inherits.
    pub(super) fn levels(&self, ty: Option<&str>, idx: Option<&str>) -> Option<&BulletLevels> {
        idx.and_then(|i| self.by_idx.get(i))
            .or_else(|| ty.and_then(|t| self.by_type.get(t)))
    }

    /// The package path of an inherited picture bullet with this file name.
    pub(super) fn path_of(&self, file: &str) -> Option<&str> {
        self.by_idx
            .values()
            .chain(self.by_type.values())
            .flat_map(|levels| levels.values().flatten())
            .map(String::as_str)
            .find(|path| path.rsplit('/').next() == Some(file))
    }
}

/// Which `txStyles` entry formats a placeholder of this type.
fn text_style_of(ty: &str) -> usize {
    match ty {
        "title" | "ctrTitle" => 0,
        "dt" | "ftr" | "sldNum" | "hdr" => 2,
        _ => 1,
    }
}

/// The master placeholder type a layout placeholder type takes its properties from.
fn master_placeholder_type(ty: &str) -> &str {
    match ty {
        "ctrTitle" => "title",
        "title" | "dt" | "ftr" | "sldNum" => ty,
        _ => "body",
    }
}

fn attr(e: &BytesStart<'_>, local: &str) -> Option<String> {
    e.attributes()
        .flatten()
        .find(|a| a.key.local_name().as_ref() == local)
        .map(|a| crate::decode::attr_value(&a))
}

/// The 0-based level of `a:lvl1pPr` … `a:lvl9pPr`.
fn level_of(name: &str) -> Option<u8> {
    let digit = name.strip_prefix("lvl")?.strip_suffix("pPr")?;
    digit
        .parse::<u8>()
        .ok()
        .filter(|n| (1..=9).contains(n))
        .map(|n| n - 1)
}

fn scan_part(xml: &str, part_path: &str, rels: &HashMap<String, String>) -> PartBullets {
    let mut out = PartBullets::default();
    let mut reader = crate::decode::reader_for(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut text_style: Option<usize> = None;
    let mut in_shape = false;
    let mut ph: Option<(String, Option<String>)> = None;
    let mut shape_levels = BulletLevels::new();
    let mut level: Option<u8> = None;
    let mut in_bu_blip = false;

    loop {
        let (e, is_start) = match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => (e, true),
            Ok(Event::Empty(e)) => (e, false),
            Ok(Event::End(e)) => {
                let name = e.name();
                let local = name.local_name();
                match local.as_ref() {
                    "titleStyle" | "bodyStyle" | "otherStyle" => text_style = None,
                    "buBlip" => in_bu_blip = false,
                    "sp" => {
                        if let Some((ty, idx)) = ph.take() {
                            out.placeholders
                                .push((ty, idx, std::mem::take(&mut shape_levels)));
                        }
                        shape_levels.clear();
                        in_shape = false;
                    }
                    n if level_of(n).is_some() => level = None,
                    _ => {}
                }
                buf.clear();
                continue;
            }
            Ok(Event::Eof) | Err(_) => break,
            Ok(_) => {
                buf.clear();
                continue;
            }
        };
        let name = e.name();
        let local = name.local_name();
        match local.as_ref() {
            "titleStyle" => text_style = Some(0),
            "bodyStyle" => text_style = Some(1),
            "otherStyle" => text_style = Some(2),
            "sp" if is_start => {
                in_shape = true;
                ph = None;
                shape_levels.clear();
            }
            "ph" if in_shape => {
                ph = Some((attr(&e, "type").unwrap_or_default(), attr(&e, "idx")));
            }
            n if level_of(n).is_some() => {
                level = if is_start { level_of(n) } else { None };
            }
            "buNone" | "buChar" | "buAutoNum" => {
                if let Some(l) = level {
                    target(&mut out, text_style, &mut shape_levels).insert(l, None);
                }
            }
            "buBlip" => in_bu_blip = is_start,
            "blip" if in_bu_blip => {
                if let (Some(l), Some(rel)) = (level, attr(&e, "embed")) {
                    if let Some(rel_target) = rels.get(&rel) {
                        let path = OoxmlContainer::resolve_path(part_path, rel_target);
                        target(&mut out, text_style, &mut shape_levels).insert(l, Some(path));
                    }
                }
            }
            _ => {}
        }
        buf.clear();
    }
    out
}

/// Where the level under the cursor records its bullet: the `txStyles` entry being read,
/// else the shape's own list style.
fn target<'a>(
    out: &'a mut PartBullets,
    text_style: Option<usize>,
    shape_levels: &'a mut BulletLevels,
) -> &'a mut BulletLevels {
    match text_style {
        Some(i) => &mut out.text_styles[i],
        None => shape_levels,
    }
}
