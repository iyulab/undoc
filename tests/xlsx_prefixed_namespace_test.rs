//! A spreadsheet part is identified by the namespace its elements are in, not by the
//! prefix the writer happened to choose.
//!
//! The Open XML SDK serialises parts built from typed elements with an explicit prefix
//! (`<x:worksheet xmlns:x="…/spreadsheetml/2006/main"><x:sheetData><x:row>…`). Such a
//! workbook is valid and other readers open it, but a reader that compares the written
//! qualified name (`x:row`) with `row` finds no sheet, no row and no cell: the result is an
//! empty document and no error.
//!
//! Every test builds the same workbook twice — once with the default namespace (the
//! control) and once with a prefix — and requires the two renderings to be identical.

use std::io::{Cursor, Write};

use undoc::parse_bytes;
use undoc::render::{to_markdown, RenderOptions};
use zip::write::SimpleFileOptions;

const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// How a part declares the spreadsheet namespace.
#[derive(Clone, Copy)]
struct Style {
    /// `""` for the default namespace, otherwise a prefix such as `"x"`.
    prefix: &'static str,
    /// Prefix bound to the relationships namespace for `…:id` attributes.
    rel_prefix: &'static str,
}

impl Style {
    /// `<name` or `<x:name`.
    fn open(self, name: &str) -> String {
        if self.prefix.is_empty() {
            name.to_string()
        } else {
            format!("{}:{name}", self.prefix)
        }
    }

    /// The root's namespace declaration.
    fn decl(self) -> String {
        if self.prefix.is_empty() {
            format!(r#"xmlns="{MAIN_NS}" xmlns:{}="{REL_NS}""#, self.rel_prefix)
        } else {
            format!(
                r#"xmlns:{}="{MAIN_NS}" xmlns:{}="{REL_NS}""#,
                self.prefix, self.rel_prefix
            )
        }
    }
}

const DEFAULT: Style = Style {
    prefix: "",
    rel_prefix: "r",
};
const PREFIXED: Style = Style {
    prefix: "x",
    rel_prefix: "r",
};
/// A prefix nobody writes by convention, and a non-`r` relationships prefix.
const ODD: Style = Style {
    prefix: "ns0",
    rel_prefix: "rel",
};

fn element(s: Style, name: &str, attrs: &str, inner: &str) -> String {
    let tag = s.open(name);
    if inner.is_empty() {
        format!("<{tag} {attrs}/>")
    } else {
        format!("<{tag} {attrs}>{inner}</{tag}>")
    }
}

fn workbook_part(s: Style) -> String {
    let sheet = element(
        s,
        "sheet",
        &format!(r#"name="Data" sheetId="1" {}:id="rId1""#, s.rel_prefix),
        "",
    );
    let sheets = element(s, "sheets", "", &sheet);
    let root = s.open("workbook");
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><{root} {}>{sheets}</{root}>"#,
        s.decl()
    )
}

fn shared_strings_part(s: Style) -> String {
    let si = |text: &str| {
        let t = element(s, "t", "", text);
        element(s, "si", "", &t)
    };
    let items = format!("{}{}{}", si("name"), si("amount"), si("total"));
    let root = s.open("sst");
    let decl = s.decl();
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><{root} {decl} count="3" uniqueCount="3">{items}</{root}>"#
    )
}

fn styles_part(s: Style) -> String {
    // One custom number format (id 164) used by cell style 1.
    let num_fmt = element(s, "numFmt", r#"numFmtId="164" formatCode="0.00""#, "");
    let num_fmts = element(s, "numFmts", r#"count="1""#, &num_fmt);
    let xf0 = element(
        s,
        "xf",
        r#"numFmtId="0" fontId="0" fillId="0" borderId="0""#,
        "",
    );
    let xf1 = element(
        s,
        "xf",
        r#"numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1""#,
        "",
    );
    let cell_xfs = element(s, "cellXfs", r#"count="2""#, &format!("{xf0}{xf1}"));
    let root = s.open("styleSheet");
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><{root} {}>{num_fmts}{cell_xfs}</{root}>"#,
        s.decl()
    )
}

fn worksheet_part(s: Style) -> String {
    // Header row of shared strings; a numeric row styled with the custom format; a row
    // of inline strings; a merged range.
    let shared = |r: &str, idx: u32| {
        let v = element(s, "v", "", &idx.to_string());
        element(s, "c", &format!(r#"r="{r}" t="s""#), &v)
    };
    let number = |r: &str, n: &str| {
        let v = element(s, "v", "", n);
        element(s, "c", &format!(r#"r="{r}" s="1""#), &v)
    };
    let inline = |r: &str, text: &str| {
        let t = element(s, "t", "", text);
        let is = element(s, "is", "", &t);
        element(s, "c", &format!(r#"r="{r}" t="inlineStr""#), &is)
    };

    let row1 = element(
        s,
        "row",
        r#"r="1""#,
        &format!("{}{}", shared("A1", 0), shared("B1", 1)),
    );
    let row2 = element(
        s,
        "row",
        r#"r="2""#,
        &format!("{}{}", inline("A2", "alpha"), number("B2", "12.5")),
    );
    let row3 = element(
        s,
        "row",
        r#"r="3""#,
        &format!("{}{}", shared("A3", 2), number("B3", "7")),
    );
    let sheet_data = element(s, "sheetData", "", &format!("{row1}{row2}{row3}"));
    let merge = element(s, "mergeCell", r#"ref="A4:B4""#, "");
    let merges = element(s, "mergeCells", r#"count="1""#, &merge);
    let root = s.open("worksheet");
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><{root} {}>{sheet_data}{merges}</{root}>"#,
        s.decl()
    )
}

/// A relationships part; `prefix` is `""` for the default namespace.
fn relationships(prefix: &str, items: &[(&str, &str, &str)]) -> String {
    let (root_tag, tag, decl) = if prefix.is_empty() {
        (
            "Relationships".to_string(),
            "Relationship".to_string(),
            r#"xmlns="http://schemas.openxmlformats.org/package/2006/relationships""#.to_string(),
        )
    } else {
        (
            format!("{prefix}:Relationships"),
            format!("{prefix}:Relationship"),
            format!(
                r#"xmlns:{prefix}="http://schemas.openxmlformats.org/package/2006/relationships""#
            ),
        )
    };
    let body: String = items
        .iter()
        .map(|(id, kind, target)| {
            format!(
                r#"<{tag} Id="{id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}" Target="{target}"/>"#
            )
        })
        .collect();
    format!(r#"<?xml version="1.0" encoding="UTF-8"?><{root_tag} {decl}>{body}</{root_tag}>"#)
}

fn build(s: Style) -> Vec<u8> {
    build_with_rels(s, "")
}

fn build_with_rels(s: Style, rels_prefix: &str) -> Vec<u8> {
    let files: Vec<(&str, String)> = vec![
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"#
                .to_string(),
        ),
        (
            "_rels/.rels",
            relationships(
                rels_prefix,
                &[(
                    "rId1",
                    "officeDocument",
                    "xl/workbook.xml",
                )],
            ),
        ),
        ("xl/workbook.xml", workbook_part(s)),
        (
            "xl/_rels/workbook.xml.rels",
            relationships(
                rels_prefix,
                &[
                    ("rId1", "worksheet", "worksheets/sheet1.xml"),
                    ("rId2", "sharedStrings", "sharedStrings.xml"),
                    ("rId3", "styles", "styles.xml"),
                ],
            ),
        ),
        ("xl/worksheets/sheet1.xml", worksheet_part(s)),
        ("xl/sharedStrings.xml", shared_strings_part(s)),
        ("xl/styles.xml", styles_part(s)),
    ];

    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default();
    for (name, body) in files {
        zip.start_file(name, options).expect("start zip entry");
        zip.write_all(body.as_bytes()).expect("write zip entry");
    }
    zip.finish().expect("finish zip").into_inner()
}

fn render(bytes: &[u8]) -> String {
    let doc = parse_bytes(bytes).expect("workbook parses");
    to_markdown(&doc, &RenderOptions::default()).expect("renders")
}

fn markdown(s: Style) -> String {
    render(&build(s))
}

#[test]
fn the_control_workbook_reads_its_content() {
    let md = markdown(DEFAULT);
    for expected in ["name", "amount", "alpha", "12.5", "total"] {
        assert!(
            md.contains(expected),
            "control is missing {expected:?}:\n{md}"
        );
    }
}

#[test]
fn a_prefixed_spreadsheet_namespace_reads_like_the_default_one() {
    let control = markdown(DEFAULT);
    let prefixed = markdown(PREFIXED);
    assert!(
        !prefixed.trim().is_empty(),
        "prefixed workbook read as empty"
    );
    assert_eq!(prefixed, control);
}

#[test]
fn an_unconventional_prefix_and_relationship_prefix_read_like_the_default_one() {
    // `ns0` for the spreadsheet namespace and `rel` for the relationships namespace: the
    // sheet's `rel:id` is what ties it to its part.
    assert_eq!(markdown(ODD), markdown(DEFAULT));
}

#[test]
fn prefixed_relationship_parts_resolve_the_sheet_and_its_helpers() {
    // The package and workbook relationship parts are namespace-qualified too: with the
    // elements read by their written name there is no relationship to follow, hence no
    // sheet part, shared strings or styles.
    let control = markdown(DEFAULT);
    assert_eq!(render(&build_with_rels(DEFAULT, "pr")), control);
    assert_eq!(render(&build_with_rels(PREFIXED, "pr")), control);
}
