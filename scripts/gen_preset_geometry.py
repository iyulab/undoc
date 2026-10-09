#!/usr/bin/env python3
"""Generate src/geometry/presets.rs from the DrawingML preset shape definitions.

DrawingML names its built-in shapes (`<a:prstGeom prst="rightArrow"/>`) instead of drawing
them, and ECMA-376 defines each one as data: adjust values, guide formulas evaluated against
the shape's size, a text rectangle and the paths. This script turns that file into Rust
constants, so every preset shape is drawn by one formula engine rather than by hand.

Usage:
    python scripts/gen_preset_geometry.py <presetShapeDefinitions.xml>

Source (public):
    ECMA-376 Part 1, 5th edition (2016), electronic annex
    https://ecma-international.org/wp-content/uploads/ECMA-376-1_5th_edition_december_2016.zip
    -> OfficeOpenXML-DrawingMLGeometries.zip -> presetShapeDefinitions.xml

The definitions are used under the Ecma copyright license (an implementation of the
specification's functionality in a conformant product), whose notice the generated file
carries in full.
"""

import pathlib
import sys
import xml.etree.ElementTree as ET

NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
OUT = pathlib.Path(__file__).resolve().parent.parent / "src" / "geometry" / "presets.rs"

ECMA_NOTICE = """\
COPYRIGHT NOTICE
© Ecma International
By obtaining and/or copying this work, you (the licensee) agree that you have read, understood, and will comply
with the following terms and conditions.
This document may be copied, published and distributed to others, and certain derivative works of it may be
prepared, copied, published, and distributed, in whole or in part, provided that the above copyright notice and
this Copyright License and Disclaimer are included on all such copies and derivative works. The only derivative
works that are permissible under this Copyright License and Disclaimer are:
(i) works which incorporate all or portion of this document for the purpose of providing commentary or
 explanation (such as an annotated version of the document),
(ii) works which incorporate all or portion of this document for the purpose of incorporating features that
 provide accessibility,
(iii) translations of this document into languages other than English and into different formats and
(iv) works by making use of this specification in standard conformant products by implementing (e.g. by
 copy and paste wholly or partly) the functionality therein.
However, the content of this document itself may not be modified in any way, including by removing the
copyright notice or references to Ecma International, except as required to translate it into languages other
than English or into a different format.
The official version of an Ecma International document is the English language version on the Ecma
International website. In the event of discrepancies between a translated version and the official version, the
official version shall govern.
The limited permissions granted above are perpetual and will not be revoked by Ecma International or its
successors or assigns.
This document and the information contained herein is provided on an “AS IS” basis and ECMA
INTERNATIONAL DISCLAIMS ALL WARRANTIES, EXPRESS OR IMPLIED, INCLUDING BUT NOT
LIMITED TO ANY WARRANTY THAT THE USE OF THE INFORMATION HEREIN WILL NOT INFRINGE
ANY OWNERSHIP RIGHTS OR ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR FITNESS FOR
A PARTICULAR PURPOSE."""

FILLS = {
    None: "Norm",
    "norm": "Norm",
    "none": "None",
    "lighten": "Lighten",
    "lightenLess": "LightenLess",
    "darken": "Darken",
    "darkenLess": "DarkenLess",
}


def rust_str(s: str) -> str:
    return '"' + s.replace("\\", "\\\\").replace('"', '\\"') + '"'


def guides(parent) -> str:
    if parent is None:
        return "&[]"
    items = [
        f"Guide {{ name: {rust_str(g.get('name'))}, fmla: {rust_str(g.get('fmla'))} }}"
        for g in parent.findall(f"{NS}gd")
    ]
    return "&[" + ", ".join(items) + "]"


def points(cmd) -> list:
    out = []
    for pt in cmd.findall(f"{NS}pt"):
        out += [rust_str(pt.get("x")), rust_str(pt.get("y"))]
    return out


def command(cmd) -> str:
    tag = cmd.tag[len(NS):]
    if tag == "moveTo":
        return "Cmd::Move({}, {})".format(*points(cmd))
    if tag == "lnTo":
        return "Cmd::Line({}, {})".format(*points(cmd))
    if tag == "quadBezTo":
        return "Cmd::Quad([{}])".format(", ".join(points(cmd)))
    if tag == "cubicBezTo":
        return "Cmd::Cubic([{}])".format(", ".join(points(cmd)))
    if tag == "arcTo":
        return "Cmd::Arc {{ wr: {}, hr: {}, st: {}, sw: {} }}".format(
            *(rust_str(cmd.get(a)) for a in ("wR", "hR", "stAng", "swAng"))
        )
    if tag == "close":
        return "Cmd::Close"
    raise ValueError(f"unknown path command {tag}")


def path(p) -> str:
    def dim(attr):
        v = p.get(attr)
        return f"Some({int(v)})" if v is not None else "None"

    stroke = p.get("stroke", "1") not in ("0", "false")
    cmds = ", ".join(command(c) for c in p)
    return (
        f"PathDef {{ w: {dim('w')}, h: {dim('h')}, fill: PathFill::{FILLS[p.get('fill')]}, "
        f"stroke: {'true' if stroke else 'false'}, cmds: &[{cmds}] }}"
    )


def main() -> None:
    root = ET.parse(sys.argv[1]).getroot()
    # The 2016 annex defines upDownArrow twice, identically. A repeat is dropped; a repeat
    # that differs is refused, since which one is meant cannot be told.
    by_name = {}
    for s in root:
        seen = by_name.setdefault(s.tag, s)
        if seen is not s and ET.tostring(seen) != ET.tostring(s):
            raise ValueError(f"{s.tag} is defined twice, differently")
    shapes = sorted(by_name.values(), key=lambda s: s.tag)
    out = [
        "//! The DrawingML preset shapes: adjust values, guide formulas, text rectangle and paths.",
        "//!",
        "//! @generated by `scripts/gen_preset_geometry.py` from `presetShapeDefinitions.xml`, the",
        "//! electronic annex of ECMA-376 Part 1 (5th edition, 2016). Do not edit by hand; regenerate.",
        "//!",
        "//! The definitions are © Ecma International, used under its copyright license:",
        "//!",
    ]
    out += [("//! " + line).rstrip() for line in ECMA_NOTICE.splitlines()]
    out += [
        "",
        "use super::{Cmd, Guide, PathDef, PathFill, Preset};",
        "",
        f"/// Every preset shape, sorted by name ({len(shapes)} shapes).",
        "pub(crate) static PRESETS: &[Preset] = &[",
    ]
    for s in shapes:
        rect = s.find(f"{NS}rect")
        text_rect = (
            "Some([{}])".format(", ".join(rust_str(rect.get(k)) for k in ("l", "t", "r", "b")))
            if rect is not None
            else "None"
        )
        paths = ", ".join(path(p) for p in s.find(f"{NS}pathLst"))
        out.append(
            f"    Preset {{ name: {rust_str(s.tag)}, av: {guides(s.find(f'{NS}avLst'))}, "
            f"gd: {guides(s.find(f'{NS}gdLst'))}, text_rect: {text_rect}, paths: &[{paths}] }},"
        )
    out.append("];")
    OUT.write_text("\n".join(out) + "\n", encoding="utf-8", newline="\n")
    print(f"wrote {OUT} ({len(shapes)} presets)")


if __name__ == "__main__":
    main()
