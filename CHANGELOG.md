# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.0] - 2025-01-20

### Added

- **Core Library**
  - DOCX (Word) document parsing with full structure extraction
  - XLSX (Excel) spreadsheet parsing with shared strings and cell formatting
  - PPTX (PowerPoint) presentation parsing with slide content and notes
  - Common OOXML container handling for all Office formats
  - Automatic format detection from file extension and magic bytes

- **Document Model**
  - Unified document model for all Office formats
  - Metadata extraction (title, author, created, modified dates)
  - Section-based content organization
  - Paragraph model with text runs and styling
  - Table model with cell spans and alignment
  - Resource/media extraction support

- **Rendering**
  - Markdown output with configurable options
  - Plain text extraction
  - JSON serialization (pretty and compact)
  - YAML frontmatter generation
  - Table rendering modes: Markdown, HTML, ASCII
  - Text cleanup presets: Minimal, Standard, Aggressive
  - Configurable maximum heading depth

- **CLI Tool**
  - `markdown` / `md` command for Markdown conversion
  - `text` command for plain text extraction
  - `json` command for JSON output
  - `info` command for document metadata display
  - `extract` command for resource extraction
  - `update` command for self-updating from GitHub releases
  - `version` command for version information
  - Cross-platform support (Windows, Linux, macOS)

- **FFI (Foreign Function Interface)**
  - C-ABI compatible library for native bindings
  - Thread-safe error handling
  - Functions for file and byte array parsing
  - Markdown, text, and JSON rendering
  - C header file for integration
  - C# wrapper class for .NET applications

- **CI/CD**
  - GitHub Actions CI workflow with multi-platform testing
  - Automated release workflow triggered by version changes
  - Multi-platform binary builds (Windows, Linux, macOS Intel/ARM)
  - Automatic GitHub releases with library and CLI artifacts
  - crates.io publishing support

### Technical Details

- Built with Rust for performance and safety
- Parallel processing with Rayon for multi-section documents
- Efficient XML parsing with quick-xml
- ZIP container handling with zip crate
- Self-update mechanism using self_update crate

## [Unreleased]

### Planned

- Legacy format support (.doc, .xls, .ppt)
- Async I/O with Tokio
- Additional output formats (HTML, RST)
- Image optimization options
- Batch processing mode
- Plugin system for custom processors
