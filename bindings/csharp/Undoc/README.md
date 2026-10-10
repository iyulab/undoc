# Undoc

High-performance Microsoft Office document extraction to Markdown for .NET.

## Installation

```bash
dotnet add package Undoc
```

## Usage

### Basic Usage

```csharp
using Undoc;

// Parse a document
using var doc = UndocDocument.ParseFile("document.docx");

// Convert to Markdown
var markdown = doc.ToMarkdown();
Console.WriteLine(markdown);

// Convert to plain text
var text = doc.ToText();

// Convert to JSON
var json = doc.ToJson();
```

### With Markdown Options

```csharp
using Undoc;

using var doc = UndocDocument.ParseFile("document.xlsx");

var options = new MarkdownOptions
{
    IncludeFrontmatter = true,
    ParagraphSpacing = true
};

var markdown = doc.ToMarkdown(options);
```

### Parse from Bytes

```csharp
using Undoc;

byte[] data = File.ReadAllBytes("document.pptx");

using var doc = UndocDocument.ParseBytes(data);
var markdown = doc.ToMarkdown();
```

### Extract Resources (Images)

```csharp
using Undoc;

using var doc = UndocDocument.ParseFile("document.docx");

// Get all resource IDs
var resourceIds = doc.GetResourceIds();

foreach (var id in resourceIds)
{
    // Get resource metadata
    using var info = doc.GetResourceInfo(id);
    // "primary": an image the document shows. "alternate" (a picture's SVG original) and
    // "layer" (an HD Photo effects layer) belong to the primary named by "companion_of".
    if (info is null || info.RootElement.GetProperty("role").GetString() != "primary")
    {
        continue;
    }
    var filename = info.RootElement.GetProperty("filename").GetString();
    Console.WriteLine($"Resource: {filename}");

    // Get resource binary data
    var data = doc.GetResourceData(id);
    if (data != null && filename != null)
    {
        File.WriteAllBytes(filename, data);
    }
}
```

### Document Metadata

```csharp
using Undoc;

using var doc = UndocDocument.ParseFile("document.docx");

Console.WriteLine($"Title: {doc.Title}");
Console.WriteLine($"Author: {doc.Author}");
Console.WriteLine($"Sections: {doc.SectionCount}");
Console.WriteLine($"Resources: {doc.ResourceCount}");
Console.WriteLine($"Library Version: {UndocDocument.Version}");
```

### Render a Slide

A slide of a `.pptx` renders to PNG from the parsed document, without reading the file again:

```csharp
using Undoc;

using var doc = UndocDocument.ParseFile("deck.pptx");
var slide = doc.RenderSection(0, new RenderSectionOptions { Dpi = 150 });
File.WriteAllBytes("slide1.png", slide.Png);

if (!slide.Gaps.IsEmpty)
    Console.WriteLine($"{slide.Gaps.Charts} charts and {slide.Gaps.GraphicFrames} other graphic frames not painted");
```

What the renderer cannot paint yet — charts, embedded objects and other graphic frames, custom geometry,
pictures other than PNG and JPEG, text in a script that needs shaping — is left out and counted
in `Gaps`; the rest of the slide is painted. No font is bundled. Text is drawn in the directories you name
(`FontDirectories`), then the system's. A Linux container without fonts draws no text — install a font
package (Noto Sans CJK covers Latin and East Asian text) or pass a font directory — and
reports the runs as gaps. A section the document does not have throws `UndocErrorKind.SectionOutOfRange`; a
document that is not a presentation throws `UnsupportedFormat`.

### Handling Failures

`UndocException.Kind` says *why* a call failed, so you can react to the reason instead of
matching on message text:

```csharp
using Undoc;

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

The numbers behind `UndocErrorKind` are a stable ABI contract: a new reason takes the next
free number and existing ones are never renumbered. Always keep a `default` branch so an
unrecognised value degrades to a generic failure rather than going unhandled. `Kind` is
`Other` for failures raised by the wrapper itself, and never `None` (which means success).

## Supported Formats

- **DOCX** - Microsoft Word documents
- **XLSX** - Microsoft Excel spreadsheets
- **PPTX** - Microsoft PowerPoint presentations

## Features

- **RAG-Ready Output**: Structured Markdown optimized for RAG/LLM applications
- **High Performance**: Native Rust implementation via P/Invoke
- **Asset Extraction**: Images and embedded resources
- **Metadata Preservation**: Document properties, styles, formatting
- **Cross-Platform**: Windows, Linux, macOS (Intel & ARM)

## API Reference

### UndocDocument Class

#### Static Methods

- `ParseFile(string path)` - Parse document from file path
- `ParseBytes(byte[] data)` - Parse document from bytes

#### Instance Methods

- `ToMarkdown(MarkdownOptions? options)` - Convert to Markdown
- `ToText()` - Convert to plain text
- `ToJson(bool compact)` - Convert to JSON
- `PlainText()` - Get plain text (fast extraction)
- `GetTables(bool tsv = false)` - Every table as CSV (RFC 4180), or tab-separated, in reading order (`IReadOnlyList<TableText>`): `Section`, `Index` (its place in the section, from 1) and `Text`. A table nested in a cell is a table of its own, right after the one that holds it; a merged cell's text is in its top-left position and the positions it covers are empty.
- `GetResourceIds()` - List of resource IDs
- `GetResourceInfo(string id)` - Resource metadata as JsonDocument
- `GetResourceData(string id)` - Resource binary data
- `RenderSection(int index, RenderSectionOptions? options)` - A slide as PNG, its size, and what was not painted (`RenderedSection`)

#### Properties

- `Title` - Document title
- `Author` - Document author
- `SectionCount` - Number of sections
- `ResourceCount` - Number of resources
- `Version` (static) - Library version

### MarkdownOptions Class

- `IncludeFrontmatter` - Include YAML frontmatter
- `EscapeSpecialChars` - Escape special characters
- `ParagraphSpacing` - Add extra paragraph spacing

## License

MIT License - see [LICENSE](../../LICENSE) for details.
