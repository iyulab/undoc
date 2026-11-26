//! Paragraphs whose style marks them as code come out as one fenced block, their lines and
//! whitespace as written.
//!
//! The style names are the ones writers use for code: `Source Code` (pandoc, which its own
//! reader turns back into a code block), Word's `HTML Preformatted`, and a custom `Code`.
//! A name alone is not enough where the text is set in a proportional face: documents keep
//! `HTML Preformatted` on ordinary prose pasted from a web page.

use std::io::{Cursor, Write};

use undoc::parse_bytes;
use undoc::render::{to_markdown, RenderOptions};
use zip::write::SimpleFileOptions;

const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"#;

const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;

const DOC_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#;

const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/// `Source Code` with no face of its own, a style based on it, a localized id for Word's
/// `HTML Preformatted` (Courier New), a custom `Code` in Consolas and one without a face, and
/// the character style pandoc puts on the runs.
fn styles() -> String {
    let style = |kind: &str, id: &str, name: &str, based_on: &str, font: &str| {
        let based_on = match based_on {
            "" => String::new(),
            id => format!(r#"<w:basedOn w:val="{id}"/>"#),
        };
        let font = match font {
            "" => String::new(),
            face => format!(r#"<w:rPr><w:rFonts w:ascii="{face}" w:hAnsi="{face}"/></w:rPr>"#),
        };
        format!(
            r#"<w:style w:type="{kind}" w:styleId="{id}"><w:name w:val="{name}"/>{based_on}{font}</w:style>"#
        )
    };
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="{W_NS}">{}{}{}{}{}{}<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style></w:styles>"#,
        style("paragraph", "SourceCode", "Source Code", "", ""),
        style("paragraph", "Listing", "Listing", "SourceCode", ""),
        style("paragraph", "HTML", "HTML Preformatted", "", "Courier New"),
        style("paragraph", "Code", "Code", "", "Consolas"),
        style("paragraph", "PlainCode", "code", "", ""),
        style("character", "VerbatimChar", "Verbatim Char", "", "Consolas"),
    )
}

fn run(text: &str, rpr: &str) -> String {
    format!(r#"<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>"#)
}

fn para(style: &str, runs: &str) -> String {
    let ppr = match style {
        "" => String::new(),
        id => format!(r#"<w:pPr><w:pStyle w:val="{id}"/></w:pPr>"#),
    };
    format!("<w:p>{ppr}{runs}</w:p>")
}

fn line(style: &str, text: &str) -> String {
    para(style, &run(text, ""))
}

fn render(body: &str, options: &RenderOptions) -> String {
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="{W_NS}"><w:body>{body}</w:body></w:document>"#
    );
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for (name, data) in [
        ("[Content_Types].xml", CONTENT_TYPES.to_string()),
        ("_rels/.rels", ROOT_RELS.to_string()),
        ("word/_rels/document.xml.rels", DOC_RELS.to_string()),
        ("word/styles.xml", styles()),
        ("word/document.xml", document),
    ] {
        zip.start_file(name, SimpleFileOptions::default()).unwrap();
        zip.write_all(data.as_bytes()).unwrap();
    }
    let bytes = zip.finish().unwrap().into_inner();
    let doc = parse_bytes(&bytes).expect("document parses");
    to_markdown(&doc, options).expect("renders")
}

fn md(body: &str) -> String {
    render(body, &RenderOptions::default())
}

#[test]
fn consecutive_code_paragraphs_are_one_fenced_block() {
    let body = [
        line("", "Run:"),
        line("SourceCode", "var p = Create(path);"),
        line("SourceCode", "    await p.Run();   // *indented*"),
        para("SourceCode", ""),
        line("Listing", "# after a blank line"),
        line("", "Done."),
    ]
    .concat();
    assert_eq!(
        md(&body),
        "Run:\n\n```\nvar p = Create(path);\n    await p.Run();   // *indented*\n\n# after a blank line\n```\n\nDone."
    );
}

#[test]
fn breaks_and_tabs_inside_a_code_paragraph_are_kept() {
    let runs = format!(
        r#"{}<w:r><w:br/></w:r><w:r><w:tab/></w:r>{}"#,
        run("if a {", r#"<w:rStyle w:val="VerbatimChar"/>"#),
        run("b()", r#"<w:b/>"#),
    );
    assert_eq!(md(&para("SourceCode", &runs)), "```\nif a {\n\tb()\n```");
}

#[test]
fn blank_paragraphs_at_the_edges_of_a_block_are_not_part_of_it() {
    let body = [
        para("SourceCode", ""),
        line("SourceCode", "x = 1"),
        para("SourceCode", ""),
    ]
    .concat();
    assert_eq!(md(&body), "```\nx = 1\n```");
    assert_eq!(md(&para("SourceCode", "")), "");
}

#[test]
fn the_fence_is_longer_than_any_backtick_run_inside() {
    let body = line("SourceCode", "```rust");
    assert_eq!(md(&body), "````\n```rust\n````");
}

#[test]
fn words_html_preformatted_and_a_monospaced_code_style_are_code() {
    let body = [line("HTML", "a  =  1"), line("Code", "b  =  2")].concat();
    assert_eq!(md(&body), "```\na  =  1\nb  =  2\n```");
}

#[test]
fn a_code_style_on_proportional_text_is_prose() {
    let times = r#"<w:rFonts w:ascii="Times New Roman"/>"#;
    for style in ["HTML", "SourceCode", "Code"] {
        let body = para(style, &run("An ordinary sentence.", times));
        assert_eq!(md(&body), "An ordinary sentence.", "{style}");
    }
}

#[test]
fn a_custom_code_style_without_a_monospaced_face_is_prose() {
    assert_eq!(md(&line("PlainCode", "Penal code.")), "Penal code.");
}

#[test]
fn a_monospaced_face_alone_does_not_make_a_block() {
    let consolas = r#"<w:rFonts w:ascii="Consolas"/>"#;
    let body = para("", &run("dotnet restore", consolas));
    assert_eq!(md(&body), "dotnet restore");
}

#[test]
fn cleanup_leaves_the_text_of_a_block_alone() {
    // Cleanup rewrites typographic quotes, collapses runs of spaces and drops a line that is
    // only a number -- none of which may touch code.
    let body = [
        line("", "Before  the   block."),
        line("SourceCode", "s  =  \u{201c}x\u{201d}"),
        line("SourceCode", "42"),
        line("", "After."),
    ]
    .concat();
    let options = RenderOptions::default().with_cleanup();
    assert_eq!(
        render(&body, &options),
        "Before the block.\n\n```\ns  =  \u{201c}x\u{201d}\n42\n```\n\nAfter."
    );
}

#[test]
fn a_code_paragraph_in_a_table_cell_is_a_code_span() {
    let cell = |content: &str| format!("<w:tc>{content}</w:tc>");
    let table = format!(
        r#"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>{}{}</w:tr><w:tr>{}{}</w:tr></w:tbl>"#,
        cell(&line("", "Command")),
        cell(&line("", "Use")),
        cell(&line("SourceCode", "a |  b")),
        cell(&line("", "pipe")),
    );
    let out = md(&table);
    assert!(out.contains(r"| `a \|  b` | pipe |"), "{out}");
}

#[test]
fn heading_analysis_reads_the_paragraphs_around_a_block_as_before() {
    let body = [
        line("Heading1", "Setup"),
        line("SourceCode", "make"),
        line("SourceCode", "make install"),
        line("Heading1", "Usage"),
        line("", "Run it."),
    ]
    .concat();
    assert_eq!(
        render(&body, &RenderOptions::default().with_heading_analysis()),
        "# Setup

```
make
make install
```

# Usage

Run it."
    );
}

#[cfg(feature = "refine")]
#[test]
fn refinement_leaves_the_text_of_a_block_as_it_is() {
    let body = [
        line("SourceCode", "| a  |  b |"),
        line("SourceCode", "    1. not a list"),
    ]
    .concat();
    // Refinement writes the document again and may choose its own fence; the lines inside
    // are the code's.
    let out = render(&body, &RenderOptions::default().with_refine());
    assert!(
        out.contains(
            "```
| a  |  b |
    1. not a list
```"
        ),
        "{out:?}"
    );
}
