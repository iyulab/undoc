//! Utility to explore PPTX structure for development
use undoc::container::OoxmlContainer;

fn main() {
    let path = std::env::args().nth(1).unwrap_or("test-files/file_example_PPT_1MB.pptx".to_string());
    let container = OoxmlContainer::open(&path).expect("Failed to open file");

    println!("=== Files in archive ===");
    for file in container.list_files() {
        println!("  {}", file);
    }

    println!("\n=== [Content_Types].xml ===");
    if let Ok(content) = container.read_xml("[Content_Types].xml") {
        println!("{}", &content[..content.len().min(2000)]);
    }

    println!("\n=== ppt/presentation.xml ===");
    if let Ok(content) = container.read_xml("ppt/presentation.xml") {
        println!("{}", &content[..content.len().min(2000)]);
    }

    println!("\n=== ppt/slides/slide1.xml (first 3000 chars) ===");
    if let Ok(content) = container.read_xml("ppt/slides/slide1.xml") {
        println!("{}", &content[..content.len().min(3000)]);
    }

    println!("\n=== ppt/notesSlides/notesSlide1.xml (first 2000 chars) ===");
    if let Ok(content) = container.read_xml("ppt/notesSlides/notesSlide1.xml") {
        println!("{}", &content[..content.len().min(2000)]);
    }

    println!("\n=== ppt/_rels/presentation.xml.rels ===");
    if let Ok(content) = container.read_xml("ppt/_rels/presentation.xml.rels") {
        println!("{}", content);
    }
}
