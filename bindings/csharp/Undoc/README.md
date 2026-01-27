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
    var filename = info?.RootElement.GetProperty("filename").GetString();
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
- `GetResourceIds()` - List of resource IDs
- `GetResourceInfo(string id)` - Resource metadata as JsonDocument
- `GetResourceData(string id)` - Resource binary data

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
