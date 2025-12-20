//! Utility to explore XLSX structure for development
use undoc::container::OoxmlContainer;

fn main() {
    let path = std::env::args()
        .nth(1)
        .unwrap_or("test-files/file_example_XLSX_5000.xlsx".to_string());
    let container = OoxmlContainer::open(&path).expect("Failed to open file");

    println!("=== Files in archive ===");
    for file in container.list_files() {
        println!("  {}", file);
    }

    println!("\n=== [Content_Types].xml ===");
    if let Ok(content) = container.read_xml("[Content_Types].xml") {
        println!("{}", &content[..content.len().min(2000)]);
    }

    println!("\n=== xl/workbook.xml ===");
    if let Ok(content) = container.read_xml("xl/workbook.xml") {
        println!("{}", &content[..content.len().min(2000)]);
    }

    println!("\n=== xl/sharedStrings.xml (first 2000 chars) ===");
    if let Ok(content) = container.read_xml("xl/sharedStrings.xml") {
        println!("{}", &content[..content.len().min(2000)]);
    }

    println!("\n=== xl/worksheets/sheet1.xml (first 3000 chars) ===");
    if let Ok(content) = container.read_xml("xl/worksheets/sheet1.xml") {
        println!("{}", &content[..content.len().min(3000)]);
    }

    println!("\n=== xl/styles.xml (first 2000 chars) ===");
    if let Ok(content) = container.read_xml("xl/styles.xml") {
        println!("{}", &content[..content.len().min(2000)]);
    }

    println!("\n=== xl/_rels/workbook.xml.rels ===");
    if let Ok(content) = container.read_xml("xl/_rels/workbook.xml.rels") {
        println!("{}", content);
    }
}
