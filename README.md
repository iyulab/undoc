# undoc

[![Crates.io](https://img.shields.io/crates/v/undoc.svg)](https://crates.io/crates/undoc)
[![Documentation](https://docs.rs/undoc/badge.svg)](https://docs.rs/undoc)
[![PyPI](https://img.shields.io/pypi/v/undoc.svg)](https://pypi.org/project/undoc/)
[![NuGet](https://img.shields.io/nuget/v/Undoc.svg)](https://www.nuget.org/packages/Undoc/)
[![CI](https://github.com/iyulab/undoc/actions/workflows/ci.yml/badge.svg)](https://github.com/iyulab/undoc/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A high-performance Rust library for extracting content from Microsoft Office documents (DOCX, XLSX, PPTX) to Markdown, plain text, and JSON.

## Features

- **Multi-format support**: DOCX (Word), XLSX (Excel), PPTX (PowerPoint)
- **Multiple output formats**: Markdown, Plain Text, JSON (with full metadata)
- **Structure preservation**: Headings, lists, tables, inline formatting
- **Smart heading detection**: Style-based heading recognition (English/Korean)
- **Table cell alignment**: Proper left/center/right alignment in Markdown
- **Merged cells keep their columns**: a merged cell renders at the column it starts in,
  with the columns it covers left empty — so grouped/multi-row headers stay above their
  own data instead of shifting
- **PPTX table extraction**: Full table parsing from PowerPoint slides
- **CJK text support**: Smart spacing for Korean, Chinese, Japanese content
- **Asset extraction**: Images, charts, and embedded media with resolved paths (XLSX drawings included)
- **Rich content**: Footnotes/endnotes, headers/footers, text boxes, cell comments, hyperlinks
- **Section markers**: `<!-- slide N: Name -->` / `<!-- sheet N: Name -->` boundary markers for PPTX/XLSX
- **Text cleanup**: Multiple presets for LLM training data preparation
- **Self-update**: Built-in update mechanism via GitHub releases
- **C-ABI FFI**: Native library for C#, Python, and other languages
- **Parallel processing**: Uses Rayon for multi-section documents
- **Streaming pipeline** (0.3.0+): `parse_file_streaming` yields sections as they parse; peak memory bounded regardless of document size
- **WebAssembly** (0.4.0+): `@iyulab/undoc` npm package — parse Office documents in the browser, no server required

---

## Table of Contents

- [Installation](#installation)
- [WASM / Browser](#wasm--browser)
- [CLI Usage](#cli-usage)
- [Rust Library Usage](#rust-library-usage)
- [C# / .NET Integration](#c--net-integration)
- [Output Formats](#output-formats)
- [Feature Flags](#feature-flags)
- [License](#license)

---

## WASM / Browser

Parse Office documents in the browser — no server, no upload:

```js
import init, { parse } from '@iyulab/undoc';
await init();
const doc = parse(new Uint8Array(await file.arrayBuffer()));
console.log(doc.format());      // "docx" | "xlsx" | "pptx"
console.log(doc.toMarkdown());
```

Install: `npm install @iyulab/undoc`

**[Live Playground →](https://iyulab.github.io/undoc/)**

---

## Installation

### Pre-built Binaries (Recommended)

Download the latest release from [GitHub Releases](https://github.com/iyulab/undoc/releases/latest).

#### Windows (x64)

```powershell
# Download and extract (replace VERSION with the actual version, e.g. v0.3.1)
$VERSION = (Invoke-RestMethod "https://api.github.com/repos/iyulab/undoc/releases/latest").tag_name
Invoke-WebRequest -Uri "https://github.com/iyulab/undoc/releases/latest/download/undoc-windows-x86_64-${VERSION}.zip" -OutFile "undoc.zip"
Expand-Archive -Path "undoc.zip" -DestinationPath "."

# Move to a directory in PATH (optional)
Move-Item -Path "undoc.exe" -Destination "$env:LOCALAPPDATA\Microsoft\WindowsApps\"

# Verify installation
undoc version
```

#### Linux (x64)

```bash
# Download and extract (replace VERSION with the actual version, e.g. v0.3.1)
VERSION=$(curl -s "https://api.github.com/repos/iyulab/undoc/releases/latest" | grep '"tag_name"' | cut -d'"' -f4)
curl -LO "https://github.com/iyulab/undoc/releases/latest/download/undoc-linux-x86_64-${VERSION}.tar.gz"
tar -xzf "undoc-linux-x86_64-${VERSION}.tar.gz"

# Install to /usr/local/bin (requires sudo)
sudo mv undoc /usr/local/bin/

# Or install to user directory
mkdir -p ~/.local/bin
mv undoc ~/.local/bin/
echo 'export PATH="$HOME/.local/bin:$PATH"' >> ~/.bashrc
source ~/.bashrc

# Verify installation
undoc version
```

#### macOS

```bash
# Intel Mac (replace VERSION with the actual version, e.g. v0.3.1)
VERSION=$(curl -s "https://api.github.com/repos/iyulab/undoc/releases/latest" | grep '"tag_name"' | cut -d'"' -f4)
curl -LO "https://github.com/iyulab/undoc/releases/latest/download/undoc-macos-x86_64-${VERSION}.tar.gz"
tar -xzf "undoc-macos-x86_64-${VERSION}.tar.gz"

# Apple Silicon (M1/M2/M3/M4)
curl -LO "https://github.com/iyulab/undoc/releases/latest/download/undoc-macos-aarch64-${VERSION}.tar.gz"
tar -xzf "undoc-macos-aarch64-${VERSION}.tar.gz"

# Install
sudo mv undoc /usr/local/bin/

# Verify
undoc version
```

#### Available Binaries

| Platform | Architecture | File |
|----------|--------------|------|
| Windows | x64 | `undoc-windows-x86_64-{version}.zip` |
| Linux | x64 | `undoc-linux-x86_64-{version}.tar.gz` |
| macOS | Intel | `undoc-macos-x86_64-{version}.tar.gz` |
| macOS | Apple Silicon | `undoc-macos-aarch64-{version}.tar.gz` |

### Updating

undoc includes a built-in self-update mechanism:

```bash
# Check for updates
undoc update --check

# Update to latest version
undoc update

# Force reinstall (even if on latest)
undoc update --force
```

### Install via Cargo

If you have Rust installed:

```bash
# Install CLI
cargo install undoc-cli

# Add library to your project
cargo add undoc
```

---

## CLI Usage

### Quick Start

```bash
# Convert to Markdown + extract media (default)
undoc document.docx

# Specify output directory
undoc document.docx ./output

# With text cleanup for LLM training
undoc document.docx --cleanup aggressive
```

### Output Structure

```
document_output/
├── extract.md      # Markdown output
└── media/          # Extracted images and media
    └── image1.jpeg
```

Use `undoc convert <file> --all` to produce all three formats at once:

```
document_output/
├── extract.md      # Markdown output
├── extract.txt     # Plain text output
├── content.json    # Full structured JSON
└── media/
```

### Commands

```bash
undoc <file> [output]              # Convert to Markdown + extract media (default)
undoc convert <file> [OPTIONS]     # Convert with full format/streaming control
undoc markdown <file> [OPTIONS]    # Convert to Markdown only (alias: md)
undoc text <file> [OPTIONS]        # Convert to plain text only
undoc json <file> [OPTIONS]        # Convert to JSON only
undoc info <file>                  # Show document information
undoc extract <file> [OPTIONS]     # Extract resources only
undoc update [OPTIONS]             # Self-update to latest version
undoc version                      # Show version information
```

### Convert (multi-format streaming pipeline)

```bash
# Markdown only (default)
undoc convert document.docx

# All formats + media
undoc convert document.docx --all -o ./output

# Specific formats
undoc convert document.docx --formats md,txt,json

# Skip media extraction
undoc convert document.docx --no-images

# Lossless mode (page breaks + headers/footers)
undoc convert document.docx --lossless

# Insert section boundary markers for PPTX/XLSX
undoc convert presentation.pptx --section-markers
```

#### Convert Options

| Option | Description | Default |
|--------|-------------|---------|
| `-o, --output` | Output directory | `<stem>_output/` |
| `--formats` | Comma-separated formats: `md,txt,json` | `md` |
| `--all` | Output all formats (MD + TXT + JSON) | false |
| `--no-images` | Skip media extraction | false |
| `--cleanup` | Text cleanup: `minimal`, `standard`, `aggressive` | none |
| `--section-markers` | Insert `<!-- slide/sheet N: Name -->` markers (PPTX/XLSX) | false |
| `--emit-page-breaks` | Emit `---` for hard page breaks | false |
| `--include-headers-footers` | Include section headers/footers as blockquotes | false |
| `--lossless` | Enable both `--emit-page-breaks` and `--include-headers-footers` | false |
| `-q, --quiet` | Suppress progress output | false |

### Convert to Markdown

```bash
# Basic conversion (output to stdout)
undoc markdown document.docx

# Save to file
undoc markdown document.docx -o output.md

# With YAML frontmatter
undoc markdown document.docx --frontmatter -o output.md

# With text cleanup for LLM training
undoc markdown document.docx --cleanup standard -o cleaned.md

# Table rendering options
undoc markdown spreadsheet.xlsx --table-mode html -o output.md
# `markdown` keeps a merged cell's column position but flattens the merge itself —
# Markdown has no colspan. Use `html` when the merge has to survive.

# Limit heading depth
undoc markdown document.docx --max-heading 3 -o output.md

# Insert section boundary markers for PPTX/XLSX
undoc markdown presentation.pptx --section-markers -o slides.md

# Lossless mode (preserve page breaks and headers/footers)
undoc markdown document.docx --lossless -o output.md
```

#### Markdown Options

| Option | Description | Default |
|--------|-------------|---------|
| `-o, --output` | Output file path | stdout |
| `-f, --frontmatter` | Include YAML frontmatter | false |
| `--table-mode` | Table rendering: `markdown`, `html`, `ascii` | markdown |
| `--cleanup` | Text cleanup: `minimal`, `standard`, `aggressive` | none |
| `--max-heading` | Maximum heading level (1-6) | 4 |
| `--section-markers` | Insert `<!-- slide/sheet N: Name -->` markers (PPTX/XLSX) | false |
| `--emit-page-breaks` | Emit `---` for hard page breaks | false |
| `--include-headers-footers` | Include section headers/footers as blockquotes | false |
| `--lossless` | Enable both `--emit-page-breaks` and `--include-headers-footers` | false |

### Convert to Plain Text

```bash
# Basic extraction
undoc text document.docx

# With cleanup
undoc text document.docx --cleanup standard -o output.txt
```

### Convert to JSON

```bash
# Pretty-printed JSON
undoc json document.docx -o output.json

# Compact JSON
undoc json document.docx --compact -o output.json
```

### Show Document Information

```bash
undoc info document.docx
```

Output:
```
Document Information
────────────────────────────────────────
File: document.docx
Format: Docx
Sections: 5
Resources: 3
Title: My Document
Author: John Doe
Pages/Slides/Sheets: 10
Created: 2025-01-15T10:30:00Z
Modified: 2025-01-20T14:45:00Z

Content Statistics
────────────────────────────────────────
Words: 2500
Characters: 15000
```

### Extract Resources

```bash
# Extract to current directory
undoc extract presentation.pptx

# Extract to specific directory
undoc extract presentation.pptx -o ./media
```

### Self-Update

```bash
# Check for updates
undoc update --check

# Update to latest version
undoc update

# Force reinstall
undoc update --force
```

### Examples

```bash
# Convert Word document to Markdown with frontmatter
undoc md report.docx --frontmatter -o report.md

# Convert Excel to Markdown tables
undoc md data.xlsx -o tables.md

# Convert PowerPoint to Markdown with slide markers
undoc md presentation.pptx --section-markers -o slides.md

# Extract all images from a document
undoc extract report.docx -o ./images

# Get document metadata
undoc info document.docx

# Convert with aggressive cleanup for AI training
undoc md document.docx --cleanup aggressive -o cleaned.md

# Batch conversion (shell)
for f in *.docx; do undoc md "$f" -o "${f%.docx}.md"; done

# Batch conversion (PowerShell)
Get-ChildItem *.docx | ForEach-Object { undoc md $_.FullName -o "$($_.BaseName).md" }
```

---

## Rust Library Usage

### Quick Start

```rust
use undoc::{parse_file, render};

fn main() -> undoc::Result<()> {
    // Parse document
    let doc = parse_file("document.docx")?;

    // Convert to Markdown
    let options = render::RenderOptions::default();
    let markdown = render::to_markdown(&doc, &options)?;
    println!("{}", markdown);

    // Get plain text
    let text = render::to_text(&doc, &options)?;

    // Get JSON
    let json = render::to_json(&doc, render::JsonFormat::Pretty)?;

    Ok(())
}
```

### Convenience Functions

```rust
// One-shot conversions without building a Document first
let text     = undoc::extract_text("document.docx")?;
let markdown = undoc::to_markdown("document.docx")?;
let json     = undoc::to_json("document.docx", undoc::render::JsonFormat::Pretty)?;

// Parse from bytes
let doc = undoc::parse_bytes(&file_bytes)?;
```

### Render Options

```rust
use undoc::render::{RenderOptions, CleanupPreset, TableFallback};
use undoc::SectionMarkerStyle;

let options = RenderOptions::new()
    .with_frontmatter(true)
    .with_table_fallback(TableFallback::Html)
    .with_cleanup_preset(CleanupPreset::Aggressive)
    .with_max_heading(3)
    .with_section_markers(SectionMarkerStyle::Comment); // <!-- slide N: Name -->

let markdown = render::to_markdown(&doc, &options)?;

// Lossless mode: preserve page breaks and headers/footers
let lossless = RenderOptions::lossless();
let markdown = render::to_markdown(&doc, &lossless)?;
```

### Working with Document Structure

```rust
use undoc::parse_file;

let doc = parse_file("document.docx")?;

// Access metadata
println!("Title: {:?}", doc.metadata.title);
println!("Author: {:?}", doc.metadata.author);
println!("Created: {:?}", doc.metadata.created);

// Iterate sections
for section in &doc.sections {
    println!("Section: {:?}", section.name);
    for element in &section.elements {
        // Process paragraphs, tables, etc.
    }
}

// Extract resources
for (id, resource) in &doc.resources {
    let filename = resource.suggested_filename(id);
    std::fs::write(&filename, &resource.data)?;
}
```

### Streaming Pipeline

Supported for PPTX (per slide) and XLSX (per sheet). DOCX is not yet supported.

```rust
use std::ops::ControlFlow;
use undoc::{parse_file_streaming, ParseEvent, SectionStreamOptions};

parse_file_streaming("slides.pptx", SectionStreamOptions::default(), |event| {
    match event {
        ParseEvent::DocumentStart { metadata, section_count, .. } => {
            println!("{} sections", section_count);
        }
        ParseEvent::SectionParsed(section) => {
            // section memory is freed after this callback returns
            println!("Slide: {:?}", section.name);
        }
        ParseEvent::ResourceExtracted { name, data } => {
            std::fs::write(&name, &data).ok();
        }
        _ => {}
    }
    ControlFlow::Continue(())
})?;
```

### Format Detection

```rust
use undoc::{detect_format_from_path, detect_format_from_bytes, FormatType};

// From file path
let format = detect_format_from_path("document.docx")?;
assert_eq!(format, FormatType::Docx);

// From bytes
let data = std::fs::read("document.docx")?;
let format = detect_format_from_bytes(&data)?;
```

---

## C# / .NET Integration

undoc ships a .NET package that bundles the native library for every supported platform.

```bash
dotnet add package Undoc
```

### Usage

```csharp
using Undoc;

// Parse and convert to Markdown
using var doc = UndocDocument.ParseFile("document.docx");
string markdown = doc.ToMarkdown(new MarkdownOptions { IncludeFrontmatter = true });
Console.WriteLine(markdown);

// Document metadata — null means "not set", not "failed"
Console.WriteLine($"Title: {doc.Title}");
Console.WriteLine($"Author: {doc.Author}");
Console.WriteLine($"Sections: {doc.SectionCount}");
Console.WriteLine($"Resources: {doc.ResourceCount}");

// Other output formats
string text = doc.ToText();
string json = doc.ToJson(compact: false);

// Resources
string[] resourceIds = doc.GetResourceIds();
JsonDocument? info = doc.GetResourceInfo("rId1");
byte[]? imageData = doc.GetResourceData("rId1");
```

`UndocDocument.ParseBytes(byte[])` parses content already in memory.

### Classifying failures

`UndocException.Kind` says *why* a call failed, so callers can react to the reason
instead of matching on message text:

```csharp
try
{
    using var doc = UndocDocument.ParseFile(path);
    Console.WriteLine(doc.ToMarkdown());
}
catch (UndocException ex)
{
    switch (ex.Kind)
    {
        case UndocErrorKind.ZipArchive:
            Console.Error.WriteLine("The file is damaged.");
            break;
        case UndocErrorKind.UnknownFormat:
        case UndocErrorKind.UnsupportedFormat:
            Console.Error.WriteLine("Not a supported Office document.");
            break;
        case UndocErrorKind.Encrypted:
            Console.Error.WriteLine("The document is encrypted.");
            break;
        default:
            // Also the right branch for a reason this build has no name for.
            Console.Error.WriteLine($"Extraction failed ({ex.Kind}): {ex.Message}");
            break;
    }
}
```

The kind numbers are a stable ABI contract: a new failure reason takes the next free
number and existing ones are never renumbered. Always keep a `default` branch — treat an
unrecognised value as a generic failure so a newer library stays usable. `Kind` is
`Other` for failures raised by the wrapper itself, and never `None` (which means
success).

### Using the native library directly

Download from [GitHub Releases](https://github.com/iyulab/undoc/releases):

| Platform | Library File |
|----------|-------------|
| Windows x64 | `undoc.dll` |
| Linux x64 | `libundoc.so` |
| macOS | `libundoc.dylib` |

Or build from source:

```bash
cargo build --release --features ffi
```

The C API is declared in [include/undoc.h](include/undoc.h). Failures return `NULL` (or
`-1`); call `undoc_last_error()` for the message and `undoc_last_error_kind()` for the
`UndocErrorKind` classification.

---

## Output Formats

### Markdown

Structured Markdown with preserved formatting:

- **Headings**: Document headings → `#`, `##`, `###`
- **Lists**: Ordered and unordered with nesting
- **Tables**: Markdown tables (with HTML/ASCII fallback for complex layouts)
- **Inline styles**: Bold (`**`), italic (`*`), underline (`<u>`), strikethrough, superscript/subscript
- **Hyperlinks**: Preserved as Markdown links (DOCX, XLSX, PPTX)
- **Images**: Linked image references from document drawings
- **Footnotes**: Markdown reference-style (`[^N]` / `[^N]: text`)
- **Headers/Footers**: Rendered as blockquotes (opt-in via `--include-headers-footers`)

### Plain Text

Pure text content without formatting markers.

### JSON

Complete document structure with metadata:

```json
{
  "metadata": {
    "title": "Document Title",
    "author": "Author Name",
    "created": "2025-01-15T10:30:00Z",
    "modified": "2025-01-20T14:45:00Z"
  },
  "sections": [...],
  "resources": [...]
}
```

---

## Supported Formats

| Format | Extension | Status |
|--------|-----------|--------|
| Word | .docx | Supported |
| Excel | .xlsx | Supported |
| PowerPoint | .pptx | Supported |

---

## Feature Flags

| Feature | Description | Default |
|---------|-------------|---------|
| `ffi` | C-ABI foreign function interface | No |

```toml
# Cargo.toml - enable FFI
[dependencies]
undoc = { version = "0.3", features = ["ffi"] }
```

---

## Performance

- Parallel section/sheet/slide processing with Rayon
- Efficient XML parsing with quick-xml
- Memory-efficient handling of large documents
- Streaming pipeline for bounded peak memory on large files

---

## License

MIT License - see [LICENSE](LICENSE) for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Related Projects

- [unhwp](https://github.com/iyulab/unhwp) - Korean HWP document extraction
- [unpdf](https://github.com/iyulab/unpdf) - PDF document extraction
