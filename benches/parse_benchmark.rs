//! Benchmarks for undoc parsing performance.
//!
//! Run with: cargo bench
//!
//! These benchmarks test parsing performance at various document sizes.

use criterion::{black_box, criterion_group, criterion_main, BenchmarkId, Criterion, Throughput};
use std::io::Cursor;

/// Creates a synthetic DOCX document with the given number of paragraphs.
fn create_test_docx(paragraph_count: usize) -> Vec<u8> {
    use std::io::Write;
    use zip::write::SimpleFileOptions;
    use zip::ZipWriter;

    let mut buffer = Vec::new();
    let mut zip = ZipWriter::new(Cursor::new(&mut buffer));

    let options = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);

    // [Content_Types].xml
    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#,
    )
    .unwrap();

    // _rels/.rels
    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
    )
    .unwrap();

    // word/_rels/document.xml.rels
    zip.start_file("word/_rels/document.xml.rels", options)
        .unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#,
    )
    .unwrap();

    // Generate document content
    let mut content = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>"#,
    );

    for i in 0..paragraph_count {
        content.push_str(&format!(
            r#"
    <w:p>
      <w:r>
        <w:t>This is paragraph {} with some test content for benchmarking purposes.</w:t>
      </w:r>
    </w:p>"#,
            i
        ));
    }

    content.push_str(
        r#"
  </w:body>
</w:document>"#,
    );

    zip.start_file("word/document.xml", options).unwrap();
    zip.write_all(content.as_bytes()).unwrap();

    zip.finish().unwrap();
    buffer
}

/// Benchmark DOCX parsing at various sizes.
fn bench_docx_parsing(c: &mut Criterion) {
    let mut group = c.benchmark_group("docx_parsing");

    for para_count in [10, 100, 500, 1000].iter() {
        let data = create_test_docx(*para_count);
        let size = data.len() as u64;

        group.throughput(Throughput::Bytes(size));
        group.bench_with_input(
            BenchmarkId::new("paragraphs", para_count),
            &data,
            |b, data| {
                b.iter(|| {
                    let _ = undoc::parse_bytes(black_box(data));
                });
            },
        );
    }

    group.finish();
}

/// Benchmark document rendering to Markdown.
fn bench_markdown_rendering(c: &mut Criterion) {
    let mut group = c.benchmark_group("markdown_rendering");

    for para_count in [10, 100, 500].iter() {
        let data = create_test_docx(*para_count);
        let document = undoc::parse_bytes(&data).unwrap();

        group.bench_with_input(
            BenchmarkId::new("paragraphs", para_count),
            &document,
            |b, doc| {
                b.iter(|| {
                    let options = undoc::RenderOptions::default();
                    let _ = undoc::render::to_markdown(black_box(doc), &options);
                });
            },
        );
    }

    group.finish();
}

/// Benchmark text extraction.
fn bench_text_extraction(c: &mut Criterion) {
    let mut group = c.benchmark_group("text_extraction");

    for para_count in [10, 100, 500, 1000].iter() {
        let data = create_test_docx(*para_count);
        let document = undoc::parse_bytes(&data).unwrap();

        group.bench_with_input(
            BenchmarkId::new("paragraphs", para_count),
            &document,
            |b, doc| {
                b.iter(|| {
                    let _ = black_box(doc).plain_text();
                });
            },
        );
    }

    group.finish();
}

criterion_group!(
    benches,
    bench_docx_parsing,
    bench_markdown_rendering,
    bench_text_extraction,
);
criterion_main!(benches);
