# undoc

[![Crates.io](https://img.shields.io/crates/v/undoc.svg)](https://crates.io/crates/undoc)
[![Documentation](https://docs.rs/undoc/badge.svg)](https://docs.rs/undoc)
[![CI](https://github.com/iyulab/undoc/actions/workflows/ci.yml/badge.svg)](https://github.com/iyulab/undoc/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A high-performance Rust library for extracting content from Microsoft Office documents (DOCX, XLSX, PPTX) to Markdown, plain text, and JSON.

## Features

- **Multi-format support**: DOCX (Word), XLSX (Excel), PPTX (PowerPoint)
- **Multiple output formats**: Markdown, Plain Text, JSON (with full metadata)
- **Structure preservation**: Headings, lists, tables, inline formatting
- **Asset extraction**: Images, charts, and embedded media
- **Text cleanup**: Multiple presets for LLM training data preparation
- **Self-update**: Built-in update mechanism via GitHub releases
- **C-ABI FFI**: Native library for C#, Python, and other languages
- **Parallel processing**: Uses Rayon for multi-section documents

---

## Table of Contents

- [Installation](#installation)
- [CLI Usage](#cli-usage)
- [Rust Library Usage](#rust-library-usage)
- [C# / .NET Integration](#c--net-integration)
- [Output Formats](#output-formats)
- [Feature Flags](#feature-flags)
- [License](#license)

---

## Installation

### Pre-built Binaries (Recommended)

Download the latest release from [GitHub Releases](https://github.com/iyulab/undoc/releases/latest).

#### Windows (x64)

```powershell
# Download and extract
Invoke-WebRequest -Uri "https://github.com/iyulab/undoc/releases/latest/download/undoc-cli-x86_64-pc-windows-msvc.zip" -OutFile "undoc.zip"
Expand-Archive -Path "undoc.zip" -DestinationPath "."

# Move to a directory in PATH (optional)
Move-Item -Path "undoc.exe" -Destination "$env:LOCALAPPDATA\Microsoft\WindowsApps\"

# Verify installation
undoc version
```

#### Linux (x64)

```bash
# Download and extract
curl -LO https://github.com/iyulab/undoc/releases/latest/download/undoc-cli-x86_64-unknown-linux-gnu.tar.gz
tar -xzf undoc-cli-x86_64-unknown-linux-gnu.tar.gz

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
# Intel Mac
curl -LO https://github.com/iyulab/undoc/releases/latest/download/undoc-cli-x86_64-apple-darwin.tar.gz
tar -xzf undoc-cli-x86_64-apple-darwin.tar.gz

# Apple Silicon (M1/M2/M3)
curl -LO https://github.com/iyulab/undoc/releases/latest/download/undoc-cli-aarch64-apple-darwin.tar.gz
tar -xzf undoc-cli-aarch64-apple-darwin.tar.gz

# Install
sudo mv undoc /usr/local/bin/

# Verify
undoc version
```

#### Available Binaries

| Platform | Architecture | File |
|----------|--------------|------|
| Windows | x64 | `undoc-cli-x86_64-pc-windows-msvc.zip` |
| Linux | x64 | `undoc-cli-x86_64-unknown-linux-gnu.tar.gz` |
| macOS | Intel | `undoc-cli-x86_64-apple-darwin.tar.gz` |
| macOS | Apple Silicon | `undoc-cli-aarch64-apple-darwin.tar.gz` |

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
# Extract all formats (Markdown, text, JSON) + media to output directory
undoc document.docx

# Specify output directory
undoc document.docx ./output

# With text cleanup for LLM training
undoc document.docx --cleanup aggressive
```

### Output Structure

```
document_output/
├── extract.md      # Markdown output with frontmatter
├── extract.txt     # Plain text output
├── content.json    # Full structured JSON
└── media/          # Extracted images and media
    └── image1.jpeg
```

### Commands

```bash
undoc <file> [output]              # Extract all formats (default)
undoc convert <file> [OPTIONS]     # Same as above, explicit command
undoc markdown <file> [OPTIONS]    # Convert to Markdown only (alias: md)
undoc text <file> [OPTIONS]        # Convert to plain text only
undoc json <file> [OPTIONS]        # Convert to JSON only
undoc info <file>                  # Show document information
undoc extract <file> [OPTIONS]     # Extract resources only
undoc update [OPTIONS]             # Self-update to latest version
undoc version                      # Show version information
```

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

# Limit heading depth
undoc markdown document.docx --max-heading 3 -o output.md
```

#### Markdown Options

| Option | Description | Default |
|--------|-------------|---------|
| `-o, --output` | Output file path | stdout |
| `-f, --frontmatter` | Include YAML frontmatter | false |
| `--table-mode` | Table rendering: `markdown`, `html`, `ascii` | markdown |
| `--cleanup` | Text cleanup: `minimal`, `standard`, `aggressive` | none |
| `--max-heading` | Maximum heading level (1-6) | 6 |

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

# Convert PowerPoint to Markdown
undoc md presentation.pptx -o slides.md

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

### Render Options

```rust
use undoc::render::{RenderOptions, CleanupPreset, TableFallback};

let options = RenderOptions::new()
    .with_frontmatter(true)
    .with_table_fallback(TableFallback::Html)
    .with_cleanup_preset(CleanupPreset::Aggressive)
    .with_max_heading(3);

let markdown = render::to_markdown(&doc, &options)?;
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

undoc provides C-ABI compatible bindings for integration with C# and .NET applications.

### Getting the Native Library

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

### C# Wrapper Usage

```csharp
using Undoc;

// Parse and convert to Markdown
string markdown = UndocNative.ToMarkdown("document.docx");

// Parse and convert to plain text
string text = UndocNative.ToText("document.docx");

// Parse and convert to JSON
string json = UndocNative.ToJson("document.docx");

// From byte array
byte[] data = File.ReadAllBytes("document.docx");
string markdown = UndocNative.ToMarkdownFromBytes(data);
```

See [bindings/csharp/Undoc.cs](bindings/csharp/Undoc.cs) for the complete wrapper implementation.

---

## Output Formats

### Markdown

Structured Markdown with preserved formatting:

- **Headings**: Document headings → `#`, `##`, `###`
- **Lists**: Ordered and unordered with nesting
- **Tables**: Markdown tables (with HTML/ASCII fallback for complex layouts)
- **Inline styles**: Bold (`**`), italic (`*`), underline, strikethrough
- **Hyperlinks**: Preserved as Markdown links
- **Images**: Reference-style image links

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
undoc = { version = "0.1", features = ["ffi"] }
```

---

## Performance

- Parallel section/sheet/slide processing with Rayon
- Efficient XML parsing with quick-xml
- Memory-efficient handling of large documents

---

## License

MIT License - see [LICENSE](LICENSE) for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Related Projects

- [unhwp](https://github.com/iyulab/unhwp) - Korean HWP document extraction
