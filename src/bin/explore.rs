//! Utility to explore DOCX structure for development
use undoc::container::OoxmlContainer;

fn main() {
    let path = std::env::args().nth(1).unwrap_or("test-files/file-sample_1MB.docx".to_string());
    let container = OoxmlContainer::open(&path).expect("Failed to open file");

    println!("=== Files in archive ===");
    for file in container.list_files() {
        println!("  {}", file);
    }

    println!("\n=== [Content_Types].xml ===");
    if let Ok(content) = container.read_xml("[Content_Types].xml") {
        println!("{}", &content[..content.len().min(2000)]);
    }

    println!("\n=== word/document.xml (first 3000 chars) ===");
    if let Ok(content) = container.read_xml("word/document.xml") {
        println!("{}", &content[..content.len().min(3000)]);
    }

    println!("\n=== word/styles.xml (first 2000 chars) ===");
    if let Ok(content) = container.read_xml("word/styles.xml") {
        println!("{}", &content[..content.len().min(2000)]);
    }

    println!("\n=== _rels/.rels ===");
    if let Ok(content) = container.read_xml("_rels/.rels") {
        println!("{}", content);
    }

    println!("\n=== word/_rels/document.xml.rels ===");
    if let Ok(content) = container.read_xml("word/_rels/document.xml.rels") {
        println!("{}", content);
    }
}
