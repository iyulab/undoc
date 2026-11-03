//! Namespace-prefix normalisation for the WordprocessingML readers.
//!
//! An XML element is identified by its namespace URI and local name; the prefix a writer
//! binds to the namespace is arbitrary. The DOCX readers, however, match the *written*
//! qualified name (`w:p`, `w:t`, `r:id`, `a:blip`, …). A document whose writer chose another
//! prefix — Python's `ElementTree` writes `ns0:` unless a namespace is registered, and some
//! generators bind the main namespace as the default — is valid, yet every paragraph, run and
//! table is then invisible to the reader and the result is an empty document with no error.
//!
//! [`canonical_prefixes`] gives the readers the spelling they match: every element and
//! attribute in a known OOXML namespace is written with that namespace's conventional prefix.
//! Parts that already use the conventional prefixes — nearly all — pass through untouched,
//! without being parsed. The text between tags is never altered.
//!
//! Not rewritten: namespace *declarations* (the readers do not resolve prefixes), and
//! prefixes that occur inside attribute values (`mc:Ignorable="w14"`), which no reader here
//! consumes.

use quick_xml::events::Event;
use quick_xml::name::{QName, ResolveResult};
use quick_xml::reader::NsReader;

/// Conventional prefix and the namespace URIs (ECMA-376 transitional, then strict) it stands
/// for. Only namespaces whose elements or attributes a reader matches by prefix are listed.
const KNOWN: &[(&str, &[&str])] = &[
    (
        "w",
        &[
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "http://purl.oclc.org/ooxml/wordprocessingml/main",
        ],
    ),
    (
        "r",
        &[
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "http://purl.oclc.org/ooxml/officeDocument/relationships",
        ],
    ),
    (
        "a",
        &[
            "http://schemas.openxmlformats.org/drawingml/2006/main",
            "http://purl.oclc.org/ooxml/drawingml/main",
        ],
    ),
    (
        "wp",
        &[
            "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
            "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing",
        ],
    ),
    (
        "pic",
        &[
            "http://schemas.openxmlformats.org/drawingml/2006/picture",
            "http://purl.oclc.org/ooxml/drawingml/picture",
        ],
    ),
    (
        "mc",
        &["http://schemas.openxmlformats.org/markup-compatibility/2006"],
    ),
    ("v", &["urn:schemas-microsoft-com:vml"]),
    ("o", &["urn:schemas-microsoft-com:office:office"]),
];

/// The conventional prefix of a known namespace URI.
fn conventional_prefix(uri: &str) -> Option<&'static str> {
    KNOWN
        .iter()
        .find(|(_, uris)| uris.contains(&uri))
        .map(|(prefix, _)| *prefix)
}

/// Whether `prefix` is one of the conventional prefixes.
fn is_conventional(prefix: &str) -> bool {
    KNOWN.iter().any(|(p, _)| *p == prefix)
}

/// Rewrite `xml` so every element and attribute in a known namespace carries that
/// namespace's conventional prefix. Returns the input unchanged when it already does, and
/// also when it is not well-formed (the readers report that themselves).
pub(crate) fn canonical_prefixes(xml: String) -> String {
    if !needs_rewrite(&xml) {
        return xml;
    }
    match rewrite(&xml) {
        Some(rewritten) => rewritten,
        None => xml,
    }
}

/// Scan the namespace declarations: a rewrite is needed when a known namespace is bound to
/// anything but its conventional prefix, or a conventional prefix is bound to another
/// namespace.
fn needs_rewrite(xml: &str) -> bool {
    for (at, _) in xml.match_indices("xmlns") {
        let rest = &xml[at + "xmlns".len()..];
        let (declared, rest) = match rest.strip_prefix(':') {
            Some(after) => match after.find(['=', ' ', '\t', '\r', '\n']) {
                Some(end) => (&after[..end], &after[end..]),
                None => continue,
            },
            None => ("", rest),
        };
        let rest = rest.trim_start();
        let Some(rest) = rest.strip_prefix('=') else {
            continue;
        };
        let rest = rest.trim_start();
        let Some(quote) = rest.chars().next().filter(|c| matches!(c, '"' | '\'')) else {
            continue;
        };
        let value = &rest[1..];
        let Some(end) = value.find(quote) else {
            continue;
        };
        let uri = &value[..end];
        match conventional_prefix(uri) {
            Some(conventional) if conventional != declared => return true,
            None if !declared.is_empty() && is_conventional(declared) => return true,
            _ => {}
        }
    }
    false
}

/// The name an element or attribute is written under after normalisation, or `None` to keep
/// it as written.
fn canonical_name(name: QName<'_>, resolved: &Option<String>, is_element: bool) -> Option<String> {
    let written: &str = name.as_ref();
    let (prefix, local) = match written.split_once(':') {
        Some((prefix, local)) => (Some(prefix), local),
        None => (None, written),
    };
    // An unprefixed attribute has no namespace, whatever the default namespace is.
    if prefix.is_none() && !is_element {
        return None;
    }
    if prefix == Some("xmlns") || written == "xmlns" {
        return None;
    }
    match resolved.as_deref().and_then(conventional_prefix) {
        Some(conventional) => {
            if prefix == Some(conventional) {
                None
            } else {
                Some(format!("{conventional}:{local}"))
            }
        }
        // A conventional prefix bound to a namespace that is not its own: move it out of the
        // way, so it cannot be mistaken for the real thing.
        None => match prefix {
            Some(p) if is_conventional(p) => Some(format!("_{p}:{local}")),
            _ => None,
        },
    }
}

fn namespace_of(result: ResolveResult<'_>) -> Option<String> {
    match result {
        ResolveResult::Bound(ns) => Some(ns.as_ref().to_owned()),
        _ => None,
    }
}

fn rewrite(xml: &str) -> Option<String> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().allow_dangling_amp = true;

    let mut out = String::with_capacity(xml.len() + xml.len() / 16);
    let mut copied = 0usize;
    let mut buf = Vec::new();

    loop {
        let before = reader.buffer_position() as usize;
        let (resolved, event) = {
            let (result, event) = reader.read_resolved_event_into(&mut buf).ok()?;
            (namespace_of(result), event.into_owned())
        };
        let after = reader.buffer_position() as usize;

        let replacement = match &event {
            Event::Start(e) | Event::Empty(e) => {
                let name = canonical_name(e.name(), &resolved, true);
                let mut attrs = String::new();
                let mut changed = name.is_some();
                for attr in e.attributes().with_checks(false) {
                    let attr = attr.ok()?;
                    let value: &str = &attr.value;
                    if value.contains('"') {
                        // Re-quoting would need escaping; leave the part as written.
                        return None;
                    }
                    let key: &str = attr.key.as_ref();
                    let attr_ns = if key.contains(':') {
                        namespace_of(reader.resolver().resolve_attribute(attr.key).0)
                    } else {
                        None
                    };
                    let renamed = canonical_name(attr.key, &attr_ns, false);
                    changed |= renamed.is_some();
                    attrs.push(' ');
                    attrs.push_str(renamed.as_deref().unwrap_or(key));
                    attrs.push_str("=\"");
                    attrs.push_str(value);
                    attrs.push('"');
                }
                if changed {
                    let written = e.name().as_ref().to_owned();
                    let tag = name.unwrap_or(written);
                    let close = if matches!(event, Event::Empty(_)) {
                        "/>"
                    } else {
                        ">"
                    };
                    Some(format!("<{tag}{attrs}{close}"))
                } else {
                    None
                }
            }
            Event::End(e) => canonical_name(e.name(), &resolved, true).map(|n| format!("</{n}>")),
            Event::Eof => {
                out.push_str(&xml[copied..]);
                return Some(out);
            }
            _ => None,
        };

        if let Some(tag) = replacement {
            out.push_str(&xml[copied..before]);
            out.push_str(&tag);
            copied = after;
        }
        buf.clear();
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    #[test]
    fn conventional_parts_are_returned_untouched() {
        let xml = format!(
            r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:p><w:r><w:t>hi &amp; bye</w:t></w:r></w:p></w:document>"#
        );
        assert!(!needs_rewrite(&xml));
        assert_eq!(canonical_prefixes(xml.clone()), xml);
    }

    #[test]
    fn an_unconventional_prefix_is_rewritten_for_elements_and_attributes() {
        let xml = format!(
            r#"<ns0:document xmlns:ns0="{W}" xmlns:ns1="{R}"><ns0:p><ns0:hyperlink ns1:id="rId4"><ns0:r><ns0:t>x &lt; y</ns0:t></ns0:r></ns0:hyperlink></ns0:p></ns0:document>"#
        );
        let out = canonical_prefixes(xml);
        assert!(out.contains("<w:document "), "{out}");
        assert!(out.contains(r#"<w:hyperlink r:id="rId4">"#), "{out}");
        assert!(
            out.contains("<w:t>x &lt; y</w:t></w:r></w:hyperlink>"),
            "{out}"
        );
        assert!(out.ends_with("</w:p></w:document>"), "{out}");
        assert!(!out.contains("ns0:p"), "{out}");
    }

    #[test]
    fn a_default_namespace_gets_the_conventional_prefix() {
        let xml = format!(r#"<document xmlns="{W}"><body><p w="1"/></body></document>"#);
        let out = canonical_prefixes(xml);
        assert_eq!(
            out,
            format!(r#"<w:document xmlns="{W}"><w:body><w:p w="1"/></w:body></w:document>"#)
        );
    }

    #[test]
    fn a_conventional_prefix_bound_elsewhere_is_moved_aside() {
        // `w` is the VML-like foreign namespace here; the word namespace is `ns0`.
        let xml = format!(
            r#"<ns0:document xmlns:ns0="{W}" xmlns:w="urn:example:other"><w:t>not text</w:t><ns0:t>text</ns0:t></ns0:document>"#
        );
        let out = canonical_prefixes(xml);
        assert!(out.contains("<_w:t>not text</_w:t>"), "{out}");
        assert!(out.contains("<w:t>text</w:t>"), "{out}");
    }

    #[test]
    fn text_comments_and_cdata_are_preserved() {
        let xml = format!(
            r#"<?xml version="1.0"?><ns0:a xmlns:ns0="{W}"><!-- c --><ns0:t xml:space="preserve">  a  </ns0:t><![CDATA[<x>]]></ns0:a>"#
        );
        let out = canonical_prefixes(xml);
        assert!(out.starts_with(r#"<?xml version="1.0"?><w:a "#), "{out}");
        assert!(out.contains("<!-- c -->"), "{out}");
        assert!(
            out.contains(r#"<w:t xml:space="preserve">  a  </w:t>"#),
            "{out}"
        );
        assert!(out.contains("<![CDATA[<x>]]>"), "{out}");
    }

    #[test]
    fn malformed_input_is_returned_as_written() {
        let xml = format!(r#"<ns0:a xmlns:ns0="{W}"><ns0:b></ns0:a>"#);
        assert_eq!(canonical_prefixes(xml.clone()), xml);
    }
}
