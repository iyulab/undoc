# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**undoc** is a high-performance Rust library for extracting Microsoft Office documents (DOCX, XLSX, PPTX) into structured Markdown with assets. It provides:
- CLI tool (`undoc-cli` crate)
- Rust library (`undoc` crate)
- C-ABI FFI bindings for C#/.NET integration

## Quality & Performance Goals

This library aims for **industry-leading quality and performance**:

- **Correctness**: 100% fidelity in text extraction; zero data loss
- **Structure Preservation**: Maintain document hierarchy, headings, lists, tables
- **Performance**: Process 100MB+ documents efficiently; parallel processing
- **Robustness**: Graceful degradation on malformed files; never panic
- **Zero-copy**: Minimize allocations; use borrowed data where feasible
- **Async-ready**: Design for optional async I/O integration

## Build Commands

```bash
cargo build                    # Build library
cargo build --release          # Release build
cargo test                     # Run all tests
cargo test <test_name>         # Run specific test
cargo test -p undoc            # Run tests for core crate
cargo clippy                   # Lint
cargo fmt                      # Format code
cargo bench                    # Run benchmarks
cargo doc --open               # Generate and view documentation

# CLI
cargo run -p undoc-cli -- document.docx

# FFI build
cargo build --release --features ffi
# Output: target/release/undoc.dll (Windows) | libundoc.so (Linux) | libundoc.dylib (macOS)
```

## Architecture

### OOXML Format Structure

All Office Open XML formats are ZIP archives with this common structure:

```
document.docx/.xlsx/.pptx
├── [Content_Types].xml        # MIME type definitions for all parts
├── _rels/
│   └── .rels                  # Package-level relationships
├── docProps/
│   ├── core.xml               # Dublin Core metadata (title, author, dates)
│   └── app.xml                # Application metadata
└── word/ | xl/ | ppt/         # Format-specific content folder
```

### Format Detection Strategy

Files are identified by:
1. **ZIP magic bytes**: `50 4B 03 04`
2. **Content type inspection**: Parse `[Content_Types].xml` for main part type
   - DOCX: `application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml`
   - XLSX: `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml`
   - PPTX: `application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml`

### DOCX (WordprocessingML)

**Key Parts:**
- `word/document.xml` - Main document body
- `word/styles.xml` - Style definitions
- `word/numbering.xml` - List numbering
- `word/_rels/document.xml.rels` - Relationships

**Namespaces:**
```xml
xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
```

**Element Mapping:**
| XML Element | Model | Notes |
|-------------|-------|-------|
| `<w:p>` | Paragraph | Contains runs |
| `<w:r>` | TextRun | Formatting container |
| `<w:t>` | Text | Actual text content |
| `<w:tbl>` | Table | Grid structure |
| `<w:pStyle val="Heading1">` | h1 | Style-based heading |
| `<w:b/>` | Bold | In run properties |
| `<w:i/>` | Italic | In run properties |
| `<w:drawing>` | Image | References media via rId |

### XLSX (SpreadsheetML)

**Key Parts:**
- `xl/workbook.xml` - Workbook structure
- `xl/worksheets/sheet*.xml` - Sheet content
- `xl/sharedStrings.xml` - String table
- `xl/styles.xml` - Cell formatting

**Namespaces:**
```xml
xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
```

**Cell Types (t attribute):**
- `s` - Shared string (value is index into sharedStrings.xml)
- `n` - Number
- `b` - Boolean
- `d` - Date (ISO 8601)
- `e` - Error
- `str` - Inline string

### PPTX (PresentationML)

**Key Parts:**
- `ppt/presentation.xml` - Slide list
- `ppt/slides/slide*.xml` - Individual slides
- `ppt/notesSlides/notesSlide*.xml` - Speaker notes
- `ppt/slideMasters/` - Master layouts

**Namespaces:**
```xml
xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
```

**Element Mapping:**
| XML Element | Model | Notes |
|-------------|-------|-------|
| `<p:sld>` | Slide | Root element |
| `<p:sp>` | Shape | Text container |
| `<p:txBody>` | TextBody | Rich text |
| `<a:p>` | Paragraph | DrawingML paragraph |
| `<a:r>` | Run | Text run |
| `<a:t>` | Text | Text content |

### Crate Structure

```
undoc/                          # Workspace root
├── Cargo.toml                  # Workspace manifest
├── src/                        # Core library
│   ├── lib.rs                  # Public API exports
│   ├── error.rs                # Error types (thiserror)
│   ├── detect.rs               # Format detection
│   ├── container.rs            # ZIP abstraction
│   ├── model/                  # Intermediate representation
│   ├── docx/                   # Word parser
│   ├── xlsx/                   # Excel parser
│   ├── pptx/                   # PowerPoint parser
│   ├── render/                 # Output renderers
│   ├── cleanup.rs              # LLM data cleanup
│   ├── async_api.rs            # Async API (optional)
│   └── ffi.rs                  # C-ABI bindings (optional)
├── cli/                        # CLI crate
│   └── src/main.rs
├── benches/                    # Benchmarks
└── examples/                   # Usage examples
```

## Key Dependencies

| Crate | Purpose |
|-------|---------|
| `zip` 2.2 | OOXML container extraction |
| `quick-xml` 0.37 | High-performance XML parsing |
| `serde` + `serde_json` | JSON serialization |
| `thiserror` 2.0 | Ergonomic error types |
| `clap` 4.5 | CLI argument parsing |
| `self_update` | GitHub release updates |

## Feature Flags

| Flag | Purpose | Default |
|------|---------|---------|
| `docx` | Word support | ✅ |
| `xlsx` | Excel support | ✅ |
| `pptx` | PowerPoint support | ✅ |
| `async` | Tokio async I/O | ❌ |
| `ffi` | C-ABI bindings | ❌ |

## Version Bump Checklist

When bumping version, **ALL** of the following files must be updated simultaneously:

```
Cargo.toml                           # Root library version
cli/Cargo.toml                       # CLI version (must match)
bindings/python/pyproject.toml       # Python package version
bindings/csharp/Undoc/Undoc.csproj   # C# package version
```

**Important**: CLI version mismatch causes "update available" message to appear even after updating. All versions must be in sync before creating a GitHub release tag.

## Key Implementation Notes

### Container Access Pattern
Use `RefCell<ZipArchive<Cursor<Vec<u8>>>>` for interior mutability, allowing multiple reads without &mut self.

### Relationship Resolution
All OOXML formats use relationship IDs (rId) to reference parts. Parse `_rels/*.rels` files first to build lookup table before processing content.

### Shared Strings (XLSX)
Cell values with type `s` are indices into `xl/sharedStrings.xml`. Load this table before processing worksheets.

### Markdown Conversion Strategy

| Source | Markdown |
|--------|----------|
| Heading styles | `#`, `##`, `###` |
| Bold (`<w:b/>`) | `**text**` |
| Italic (`<w:i/>`) | `*text*` |
| Strikethrough (`<w:strike/>`) | `~~text~~` |
| Tables | Pipe syntax, HTML fallback for merged cells |
| Images | Extract to assets/, reference as `![alt](path)` |
| Hyperlinks | `[text](url)` |

### Cleanup Pipeline (LLM Training Data)

Four-stage normalization:
1. **Normalize strings**: Unicode NFC, bullet standardization
2. **Clean lines**: Remove headers/footers, page numbers, TOC
3. **Filter structure**: Empty paragraphs, orphaned elements
4. **Final normalize**: Consecutive whitespace, trailing spaces

## Related Project

This project follows the architecture established by [unhwp](https://github.com/iyulab/unhwp), a similar library for Korean HWP documents. See `claudedocs/IMPLEMENTATION_PLAN.md` for the detailed phased implementation strategy.
