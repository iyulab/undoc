# undoc

[![Crates.io](https://img.shields.io/crates/v/undoc.svg)](https://crates.io/crates/undoc)
[![Documentation](https://docs.rs/undoc/badge.svg)](https://docs.rs/undoc)
[![CI](https://github.com/iyulab/undoc/actions/workflows/ci.yml/badge.svg)](https://github.com/iyulab/undoc/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A high-performance Rust library for extracting Microsoft Office documents (DOCX, XLSX, PPTX) into structured Markdown with assets.

## Features

- **Multi-format support**: DOCX (Word), XLSX (Excel), PPTX (PowerPoint)
- **Multiple output formats**: Markdown, Plain Text, JSON (with full metadata)
- **Structure preservation**: Headings, lists, tables, inline formatting, slides
- **Asset extraction**: Images, charts, and embedded media
- **Self-update**: Built-in update mechanism via GitHub releases
- **C-ABI FFI**: Native library for C#, Python, and other languages
- **Parallel processing**: Uses Rayon for multi-section documents
- **Async support**: Optional Tokio integration

---

## Table of Contents

- [Installation](#installation)
  - [Pre-built Binaries (Recommended)](#pre-built-binaries-recommended)
  - [Updating](#updating)
  - [Install via Cargo](#install-via-cargo)
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
undoc --version
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
undoc --version
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
undoc --version
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

### Basic Conversion

```bash
# Convert DOCX/XLSX/PPTX to Markdown (creates <filename>_output/ directory)
undoc document.docx
undoc spreadsheet.xlsx
undoc presentation.pptx

# Specify output directory
undoc document.docx ./output

# Using subcommand
undoc convert document.docx -o ./output
```

### Output Structure

**DOCX Output:**
```
document_output/
├── extract.md      # Markdown output
├── extract.txt     # Plain text output
├── content.json    # Full structured JSON
└── media/          # Extracted images
    ├── image1.png
    └── image2.jpg
```

**XLSX Output:**
```
spreadsheet_output/
├── extract.md      # All sheets as Markdown tables
├── extract.txt     # Tab-separated values
├── content.json    # Full structured JSON with formulas
└── charts/         # Extracted chart images
    └── chart1.png
```

**PPTX Output:**
```
presentation_output/
├── extract.md      # Slides as Markdown sections
├── extract.txt     # Plain text (slide by slide)
├── content.json    # Full structured JSON with notes
└── media/          # Extracted images and media
    ├── slide1_image1.png
    └── slide2_video.mp4
```

### Cleanup Options (for LLM Training Data)

```bash
# Default cleanup - balanced normalization
undoc document.docx --cleanup

# Minimal cleanup - essential normalization only
undoc document.docx --cleanup-minimal

# Aggressive cleanup - maximum purification
undoc document.docx --cleanup-aggressive
```

### Commands

```bash
undoc --help                    # Show help
undoc --version                 # Show version
undoc version                   # Show detailed version info
undoc update --check            # Check for updates
undoc update                    # Self-update to latest version
undoc convert FILE [OPTIONS]    # Convert with explicit subcommand
```

### Examples

```bash
# Basic conversion
undoc report.docx

# Convert with cleanup for AI training
undoc report.docx ./cleaned --cleanup-aggressive

# Convert Excel with specific sheets
undoc data.xlsx --sheets "Sheet1,Summary"

# Convert PowerPoint with speaker notes
undoc slides.pptx --include-notes

# Batch conversion (shell)
for f in *.docx; do undoc "$f" --cleanup; done

# Batch conversion (PowerShell)
Get-ChildItem *.docx,*.xlsx,*.pptx | ForEach-Object { undoc $_.FullName --cleanup }
```

---

## Rust Library Usage

### Quick Start

```rust
use undoc::{parse_file, to_markdown};

fn main() -> undoc::Result<()> {
    // Simple text extraction
    let text = undoc::extract_text("document.docx")?;
    println!("{}", text);

    // Convert to Markdown
    let markdown = to_markdown("document.docx")?;
    std::fs::write("output.md", markdown)?;

    Ok(())
}
```

### Format-Specific APIs

```rust
use undoc::{Docx, Xlsx, Pptx};

// Word documents
let doc = Docx::open("report.docx")?;
let markdown = doc.to_markdown()?;
let images = doc.extract_images()?;

// Excel spreadsheets
let workbook = Xlsx::open("data.xlsx")?;
for sheet in workbook.sheets() {
    println!("Sheet: {}", sheet.name());
    let table = sheet.to_markdown_table()?;
    println!("{}", table);
}

// PowerPoint presentations
let ppt = Pptx::open("slides.pptx")?;
for (i, slide) in ppt.slides().enumerate() {
    println!("## Slide {}\n", i + 1);
    println!("{}", slide.to_markdown()?);
    if let Some(notes) = slide.speaker_notes() {
        println!("\n> Notes: {}", notes);
    }
}
```

## Output Formats

undoc provides four complementary output formats:

| Format | Method | Description |
|--------|--------|-------------|
| **RawContent** | `doc.raw_content()` | JSON with full metadata, styles, structure |
| **RawText** | `doc.plain_text()` | Pure text without formatting |
| **Markdown** | `to_markdown()` | Structured Markdown |
| **Media** | `doc.resources` | Extracted binary assets |

### RawContent (JSON)

Get the complete document structure with all metadata:

```rust
let doc = undoc::parse_file("document.docx")?;
let json = doc.raw_content();

// JSON includes:
// - metadata: title, author, created, modified
// - sections: paragraphs, tables (DOCX)
// - sheets: cells, formulas, merged ranges (XLSX)
// - slides: shapes, text, notes, transitions (PPTX)
// - styles: bold, italic, underline, font, color
// - images, charts, embedded objects
```

## Builder API

```rust
use undoc::{Undoc, TableFallback};

let markdown = Undoc::new()
    .with_images(true)
    .with_image_dir("./assets")
    .with_table_fallback(TableFallback::Html)
    .with_frontmatter()
    .lenient()  // Skip invalid sections
    .parse("document.docx")?
    .to_markdown()?;
```

### Excel-Specific Options

```rust
use undoc::Xlsx;

let workbook = Xlsx::open("data.xlsx")?
    .with_formulas(true)      // Include formula expressions
    .with_hidden_sheets(false) // Skip hidden sheets
    .with_merged_cells(true);  // Handle merged cell ranges

let markdown = workbook.to_markdown()?;
```

### PowerPoint-Specific Options

```rust
use undoc::Pptx;

let presentation = Pptx::open("slides.pptx")?
    .with_speaker_notes(true)   // Include speaker notes
    .with_slide_numbers(true)   // Add slide numbers
    .with_animations(false);    // Skip animation metadata

let markdown = presentation.to_markdown()?;
```

## C# / .NET Integration

undoc provides C-ABI compatible bindings for seamless integration with C# and .NET applications.

### Getting the Native Library

Build from source or download from [GitHub Releases](https://github.com/iyulab/undoc/releases):

| Platform | Library File |
|----------|-------------|
| Windows x64 | `undoc.dll` |
| Linux x64 | `libundoc.so` |
| macOS | `libundoc.dylib` |

```bash
# Build native library from source
cargo build --release
# Output: target/release/undoc.dll (Windows)
#         target/release/libundoc.so (Linux)
#         target/release/libundoc.dylib (macOS)
```

### Quick Start

```csharp
using Undoc;

// Parse document once, access multiple outputs
using var doc = OfficeDocument.Parse("document.docx");

// Get Markdown
string markdown = doc.Markdown;
File.WriteAllText("output.md", markdown);

// Get plain text
string text = doc.RawText;

// Get full structured JSON (metadata, styles, formatting)
string json = doc.RawContent;

// Extract all images
foreach (var image in doc.Images)
{
    image.SaveTo($"./images/{image.Name}");
    Console.WriteLine($"Saved: {image.Name} ({image.Size} bytes)");
}

// Document statistics
Console.WriteLine($"Format: {doc.Format}");  // Docx, Xlsx, or Pptx
Console.WriteLine($"Pages/Sheets/Slides: {doc.Count}");
```

### Excel-Specific Usage

```csharp
using var workbook = ExcelDocument.Parse("data.xlsx");

foreach (var sheet in workbook.Sheets)
{
    Console.WriteLine($"Sheet: {sheet.Name}");
    Console.WriteLine($"Rows: {sheet.RowCount}, Columns: {sheet.ColumnCount}");
    
    // Access cells
    var value = sheet.GetCell("A1");
    var formula = sheet.GetFormula("B2");
    
    // Export sheet as Markdown table
    File.WriteAllText($"{sheet.Name}.md", sheet.ToMarkdown());
}
```

### PowerPoint-Specific Usage

```csharp
using var presentation = PptxDocument.Parse("slides.pptx");

for (int i = 0; i < presentation.SlideCount; i++)
{
    var slide = presentation.GetSlide(i);
    Console.WriteLine($"Slide {i + 1}: {slide.Title}");
    
    if (slide.HasSpeakerNotes)
    {
        Console.WriteLine($"  Notes: {slide.SpeakerNotes}");
    }
}
```

### With Cleanup Options

```csharp
var options = new ConversionOptions
{
    EnableCleanup = true,
    CleanupPreset = CleanupPreset.Aggressive,  // For LLM training
    IncludeFrontmatter = true,
    TableFallback = TableFallback.Html
};

using var doc = OfficeDocument.Parse("document.docx", options);
File.WriteAllText("cleaned.md", doc.Markdown);
```

### Static Methods (Simple API)

```csharp
// One-liner conversion
string markdown = OfficeConverter.ToMarkdown("document.docx");

// With cleanup
string cleanedMarkdown = OfficeConverter.ToMarkdown("document.docx", enableCleanup: true);

// Plain text extraction
string text = OfficeConverter.ExtractText("spreadsheet.xlsx");

// From byte array or stream
byte[] data = File.ReadAllBytes("presentation.pptx");
string md = OfficeConverter.BytesToMarkdown(data);
```

### ASP.NET Core Example

```csharp
[ApiController]
[Route("api/[controller]")]
public class DocumentController : ControllerBase
{
    [HttpPost("convert")]
    public async Task<IActionResult> ConvertDocument(IFormFile file)
    {
        if (file == null) return BadRequest("No file");

        var ext = Path.GetExtension(file.FileName).ToLower();
        if (!new[] { ".docx", ".xlsx", ".pptx" }.Contains(ext))
            return BadRequest("Unsupported format");

        using var ms = new MemoryStream();
        await file.CopyToAsync(ms);

        try
        {
            var markdown = OfficeConverter.BytesToMarkdown(ms.ToArray(), enableCleanup: true);
            return Ok(new { markdown });
        }
        catch (UndocException ex)
        {
            return BadRequest(new { error = ex.Message });
        }
    }
}
```

See [C# Integration Guide](docs/csharp-integration.md) for complete documentation.

## Supported Formats

| Format | Extension | Container | Status |
|--------|-----------|-----------|--------|
| Word | .docx | ZIP/XML (OOXML) | ✅ Supported |
| Excel | .xlsx | ZIP/XML (OOXML) | ✅ Supported |
| PowerPoint | .pptx | ZIP/XML (OOXML) | ✅ Supported |
| Legacy Word | .doc | OLE/CFB | 🔜 Planned |
| Legacy Excel | .xls | OLE/CFB | 🔜 Planned |
| Legacy PowerPoint | .ppt | OLE/CFB | 🔜 Planned |

## Structure Preservation

undoc maintains document structure during conversion:

### DOCX (Word)
- **Headings**: Heading styles → `#`, `##`, `###`
- **Lists**: Ordered and unordered with nesting
- **Tables**: Cell spans, alignment, HTML fallback for complex tables
- **Images**: Extracted with Markdown references
- **Inline styles**: Bold (`**`), italic (`*`), underline (`<u>`), strikethrough (`~~`)
- **Hyperlinks**: Preserved as Markdown links
- **Footnotes/Endnotes**: Converted to reference-style notes

### XLSX (Excel)
- **Sheets**: Each sheet as a separate section
- **Tables**: Markdown tables with alignment
- **Formulas**: Optional formula preservation
- **Merged cells**: Proper span handling
- **Charts**: Extracted as images
- **Data types**: Numbers, dates, currencies formatted appropriately

### PPTX (PowerPoint)
- **Slides**: Each slide as a section with `---` separators
- **Titles**: Slide titles as headings
- **Bullet points**: Converted to Markdown lists
- **Tables**: Same as DOCX handling
- **Images**: Extracted with references
- **Speaker notes**: Optional inclusion as blockquotes
- **Shapes with text**: Text content extracted

## Feature Flags

| Feature | Description | Default |
|---------|-------------|---------|
| `docx` | Word document support | ✅ |
| `xlsx` | Excel spreadsheet support | ✅ |
| `pptx` | PowerPoint presentation support | ✅ |
| `legacy` | Legacy .doc/.xls/.ppt support | ❌ |
| `async` | Async I/O with Tokio | ❌ |
| `ffi` | C-ABI foreign function interface | ❌ |

```toml
# Cargo.toml - customize features
[dependencies]
undoc = { version = "0.1", default-features = false, features = ["docx", "xlsx"] }
```

## Performance

- Parallel section/sheet/slide processing with Rayon
- Zero-copy XML parsing where possible
- Memory-efficient streaming for large documents
- Lazy image extraction

Run benchmarks:
```bash
cargo bench
```

## Comparison with Similar Tools

| Feature | undoc | python-docx | Apache POI | pandoc |
|---------|-------|-------------|------------|--------|
| Language | Rust | Python | Java | Haskell |
| DOCX | ✅ | ✅ | ✅ | ✅ |
| XLSX | ✅ | ❌ | ✅ | ❌ |
| PPTX | ✅ | ❌ | ✅ | ❌ |
| Markdown output | ✅ | ❌ | ❌ | ✅ |
| C# bindings | ✅ | ❌ | ❌ | ❌ |
| Parallel processing | ✅ | ❌ | ❌ | ❌ |
| Self-update | ✅ | ❌ | ❌ | ❌ |

## License

MIT License - see [LICENSE](LICENSE) for details.