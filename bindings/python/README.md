# undoc

High-performance Microsoft Office document extraction to Markdown.

## Installation

```bash
pip install undoc
```

## Usage

### Basic Usage

```python
from undoc import parse_file

# Parse a document
doc = parse_file("document.docx")

# Convert to Markdown
markdown = doc.to_markdown()
print(markdown)

# Convert to plain text
text = doc.to_text()

# Convert to JSON
json_data = doc.to_json()
```

### With Context Manager

```python
from undoc import parse_file

with parse_file("document.xlsx") as doc:
    print(doc.to_markdown(frontmatter=True))
    print(f"Sections: {doc.section_count}")
    print(f"Resources: {doc.resource_count}")
```

### Parse from Bytes

```python
from undoc import parse_bytes

with open("document.pptx", "rb") as f:
    data = f.read()

doc = parse_bytes(data)
markdown = doc.to_markdown()
```

### Extract Resources (Images)

```python
from undoc import parse_file

doc = parse_file("document.docx")

# Get all resource IDs
resource_ids = doc.get_resource_ids()

for rid in resource_ids:
    # Get resource metadata
    info = doc.get_resource_info(rid)
    print(f"Resource: {info['filename']} ({info['mime_type']})")

    # Get resource binary data
    data = doc.get_resource_data(rid)

    # Save to file
    with open(info['filename'], 'wb') as f:
        f.write(data)
```

### Document Metadata

```python
from undoc import parse_file

doc = parse_file("document.docx")

print(f"Title: {doc.title}")
print(f"Author: {doc.author}")
print(f"Sections: {doc.section_count}")
print(f"Resources: {doc.resource_count}")
```

## Supported Formats

- **DOCX** - Microsoft Word documents
- **XLSX** - Microsoft Excel spreadsheets
- **PPTX** - Microsoft PowerPoint presentations

## Features

- **RAG-Ready Output**: Structured Markdown optimized for RAG/LLM applications
- **High Performance**: Native Rust implementation via FFI
- **Asset Extraction**: Images and embedded resources
- **Metadata Preservation**: Document properties, styles, formatting
- **Cross-Platform**: Windows, Linux, macOS (Intel & ARM)

## API Reference

### Functions

- `parse_file(path)` - Parse document from file path
- `parse_bytes(data)` - Parse document from bytes
- `version()` - Get library version

### Undoc Class

#### Conversion Methods

- `to_markdown(frontmatter=False, escape_special=False, paragraph_spacing=False)` - Convert to Markdown
- `to_text()` - Convert to plain text
- `to_json(compact=False)` - Convert to JSON
- `plain_text()` - Get plain text (fast extraction)

#### Properties

- `title` - Document title
- `author` - Document author
- `section_count` - Number of sections
- `resource_count` - Number of resources

#### Resource Methods

- `get_resource_ids()` - List of resource IDs
- `get_resource_info(id)` - Resource metadata
- `get_resource_data(id)` - Resource binary data

## License

MIT License - see [LICENSE](../../LICENSE) for details.
