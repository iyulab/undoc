# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Fixed

- **Tables with merged cells kept their columns aligned.** A merged cell is now anchored
  at the column it starts in, and the columns it covers render as empty cells — in both
  the Markdown and the plain-text renderer.

  Previously a row of merged group labels was narrower than the table (Markdown has no
  colspan), and the renderer made up the difference by padding the *left* of the header
  row while padding the *right* of data rows. Labels ended up under the wrong columns.
  Vertical merges were not tracked at all, so rows below one started a column too far
  left. Neither failure raised anything: the output was a well-formed table that said
  something the document did not.

- **No `#` is invented in a table's first column.** A short header row used to have a
  literal `#` inserted as a stand-in row-number heading. It came from no cell in the
  input, and it appeared in both renderers.

**Output change.** Documents whose tables contain merged cells render differently — that
is the fix. Tables without merges are unchanged. `TableFallback::Html` is unaffected and
still available for callers that want merges expressed rather than flattened.

## [0.7.0] - 2026-07-31

### Upgrade notes

`RenderOptions` with no cleanup configured no longer collapses blank lines — see **Changed**
below. Output produced with a cleanup preset is unaffected. Rust callers that set
`CleanupOptions::detect_mojibake` must drop the field; it never did anything.

### Removed
- **`CleanupOptions::detect_mojibake`.** No stage ever read it, so a preset that set it
  promised a behaviour that never ran. The `detect_mojibake()` function stays: it reports
  what it finds and changes nothing, which is a diagnostic a caller invokes deliberately
  rather than a pipeline option. Callers that set the field should drop it.

### Changed
- **`cleanup: None` now means no post-processing at all.** Blank-line collapsing used to run
  regardless of the cleanup options, justified as lossless under CommonMark — true of the
  rendered result, but not of the Markdown itself, which a consumer may diff, cite by line
  number, or chunk for retrieval. Collapsing is part of whitespace normalization now and is
  governed by `CleanupOptions::final_normalize`, which already covers whitespace. Every
  shipped preset enables that option, so **preset-configured output is unchanged**. Two other
  cases do change: `cleanup: None` no longer collapses anything, and `CleanupOptions` built by
  hand with `final_normalize: false` no longer collapses either — previously it did, since the
  pass ran outside the options entirely. This also makes Markdown and text output follow the
  same policy, which they previously did not.

### Fixed
- **Whitespace at a run boundary was rendered inside the markup wrapping the run.** A
  document stores the space around a word in the runs themselves, so a trailing space
  belongs *between* runs: inside the delimiters it becomes `[label ](url)`, where the space
  sits within the link, or `**bold **`, which is not emphasis at all because CommonMark
  refuses a closing delimiter preceded by whitespace. The markup now wraps the trimmed text
  and the whitespace is emitted outside it. A run consisting only of whitespace is treated
  as the separator it is, rather than being wrapped in delimiters around nothing.
- **A thematic break could emit three consecutive newlines.** It prepended its own blank
  line to a paragraph that already ended with one, and the unconditional collapsing pass
  hid the result. The separator is now emitted to fit what precedes it.
- **Nested lists were flattened by whitespace normalization.** The final cleanup stage
  collapsed every run of whitespace, including the indentation at the start of a line —
  and in Markdown that indentation is the only expression of list nesting, so sub-items
  came back out at the top level. The renderer emits two spaces per level for exactly this
  purpose, so cleanup was undoing it. Leading whitespace is now preserved; only trailing
  whitespace and runs inside a line are normalized.
- **Cleanup deleted list items whose text was a number.** A leading dash was treated as
  page-number decoration, so `- 5` and `- 2026` were removed as page furniture — while
  `- 5 -`, which really is page decoration, survived. The two are now told apart by
  symmetry: a dash on both sides is decoration, a dash on the left alone is a Markdown
  list marker. Labelled (`Page 5`), trailing-decorated (`5 -`) and bare short numbers are
  still removed.
- **`preserve_frontmatter` now holds for the whole cleanup pipeline, not just one stage
  of it.** It was passed only to the line filter, so the later stages still saw the
  frontmatter: whitespace normalisation collapsed runs of spaces, which re-nests a YAML
  block, and structure filtering could drop a single-character line. The block is now
  detached before any stage runs and reattached afterwards, so no stage has to know
  about it. An opening fence with no closing fence is not a block and is left in the body.
- **The CLI update notification went to stdout, corrupting piped output.** `md`, `json`
  and `text` emit document data on stdout, so the notification line landed in the middle
  of it — `undoc json … | jq .` failed to parse. It now goes to stderr, where it still
  appears in an interactive terminal.
- **WebAssembly: every thrown error now carries its `kind`, not just the one from
  `parse`.** 0.6.0 documented the property as part of the error contract, but the five
  `OfficeDocument` methods (`fromBytes`, `toMarkdown`, `toText`, `toJson`, `metadata`)
  still threw bare strings — so a caller who followed the documentation and branched on
  `err.kind` got `undefined` from all but one entry point. All of them now throw a real
  `Error` with the same numeric reason the C ABI reports. A failure to serialise metadata
  is reported as `Render`, since producing output is rendering.

## [0.6.0] - 2026-07-30

### Added
- **Structured error classification.** Failures now carry a machine-readable reason
  alongside their message, so consumers can branch on *why* a call failed instead of
  matching on message text.
  - Rust: `ErrorKind` (`#[repr(i32)]`) and `Error::kind()`.
  - C ABI: `undoc_last_error_kind()` and the `UndocErrorKind` enum in `include/undoc.h`.
    Written and cleared in lockstep with `undoc_last_error()`, so a message is never
    paired with a stale reason.
  - C#: `UndocException.Kind` and `UndocErrorKind`.
  - Python: `UndocError.kind` and `ErrorKind`.
  - WebAssembly: the thrown `Error` carries a numeric `kind` property.
  - The numbers are a stable ABI contract: a new reason takes the next free number and
    existing ones are never reused or renumbered, so an unrecognised value can safely be
    treated as a generic failure. Unknown values pass through unchanged in every binding
    rather than being collapsed or rejected. `ErrorKind` is `#[non_exhaustive]`, so Rust
    callers should match it with a `_ =>` arm for the same reason.

### Changed
- Container and XML failures are no longer flattened into a single reason. A damaged
  archive, an unsupported one, a password-protected one, an absent part, an I/O failure
  and an encoding failure are now distinguishable — reported through the existing error
  variants, with no new kind numbers.
- OLE/CFB containers no longer surface as a damaged ZIP archive or an unrecognised file,
  and the two kinds that share that header are now told apart. An OOXML document
  protected with ECMA-376 encryption is a CFB wrapping the encrypted package, so its
  directory is walked for the `EncryptedPackage` stream:
  - Found → reported as **encrypted**, which tells the caller to supply a password.
  - Not found → reported as an **unsupported format**, naming the legacy binary format
    when its well-known stream identifies it (`Word 97-2003 (.doc)`,
    `Excel 97-2003 (.xls)`, `PowerPoint 97-2003 (.ppt)`).
  - A CFB whose directory cannot be read stays unsupported and says so, rather than
    guessing which kind it was.

  Previously both were reported as unsupported, which sent a caller holding a
  password-protected document looking for a file-format converter. Detection happens for
  both the path- and byte-based entry points, and no new kind numbers were needed.
- A password-protected entry inside a container is now reported as encrypted.
- Dependency: added `cfb` 0.10 (with `byteorder`, `fnv`, `uuid`) for the CFB directory
  walk above. Used by format detection only. Pure Rust, so the wasm32 build is unaffected.
- Internal: the FFI last-error/panic-guard plumbing now runs on the shared `uncore`
  crate (thread-local slot, panic guard, boundary-reason helpers) instead of a
  hand-rolled implementation duplicated across the `un*` extraction family. No
  observable change — every `ErrorKind` discriminant, every exported C symbol's name
  and signature, and every failure message stay exactly as they were.

### Fixed
- JSON serialization failures were reported as XML parse errors, pointing callers at the
  input document for a problem in output rendering. They are now render failures.
- `read_xml_optional` silently degraded a part that exists but cannot be read into
  "part absent", contrary to its documented behaviour: every failure to open an entry
  was treated as a missing component. Only a genuinely absent entry is now treated as
  absent.
- `undoc_get_title` / `undoc_get_author` returned NULL with no recorded reason when the
  value existed but held an interior NUL byte, indistinguishable from "not set". The two
  cases are now distinguishable.
- `undoc_section_count` / `undoc_resource_count` did not clear a previous failure, so a
  successful call could leave a stale error recorded.
- The README's C# examples documented an API that does not exist in the published
  package (wrong namespace, wrong method names, wrong return types) and pointed at a
  legacy wrapper file that was not part of any build. The examples now match the shipped
  package, and the unused file has been removed.
- The Python package reported a stale version from `undoc.__version__`. CI now verifies
  it alongside the other version-bearing files.

## [0.5.5] - 2026-07-15

### Fixed
- DOCX header/footer and text-box content could be lost during extraction, including
  table text inside those parts.
- Dependency: crossbeam-epoch 0.9.18 → 0.9.20 (RUSTSEC-2026-0204).

## [0.5.4] - 2026-07-08

### Added
- DOCX headers and footers are now extracted, including first-page and even-page
  variants and content inside text boxes.

## [0.5.3] - 2026-07-05

### Fixed
- Dependency: quick-xml → 0.41 (RUSTSEC-2026-0194, RUSTSEC-2026-0195). quick-xml 0.40+
  emits entity references as separate events, which silently dropped every entity
  (`&amp;`, `&#13;`, …) across text accumulation sites. Entity references are now
  resolved and preserved, with graceful degradation on a stray `&`.

## [0.5.2] - 2026-06-12

### Added
- **FEAT (#8)** — DOCX legacy VML images (`w:pict`/`w:object` > `v:imagedata`), typical of .doc → .docx conversions, are now extracted as inline images. VML inside `mc:Fallback` is still skipped to avoid duplicating the paired DrawingML `mc:Choice` image.
- **FEAT (#8)** — XLSX "Place in Cell" rich-value images (cells with `t="e" vm="N"` and a `#VALUE!` placeholder) are now resolved through the `xl/metadata.xml` → `xl/richData/*` chain and rendered as inline images in their table cell, replacing the placeholder. Files with a missing or foreign rich-data chain degrade gracefully.

### Fixed
- Image-only paragraphs are no longer considered empty, so they survive empty-paragraph filtering (previously dropped VML images inside DOCX table cells).
- CI: npm publish of `@iyulab/undoc` failed with `ENEEDAUTH` — the workflow now configures registry auth via `actions/setup-node`.

## [0.5.1] - 2026-06-12

### Fixed
- **BUG (#7)** — Bare carriage returns (`\r`, stored by Excel/openpyxl as `&#13;` character references) no longer leak into Markdown output, where they split pipe-table rows across physical lines. CR and CRLF are normalized to LF at the shared XML decode layer (covers DOCX, XLSX, and PPTX, including literal CR/CRLF per XML §2.11), and the Markdown/text table renderers additionally treat any remaining CR as a line break when flattening cell text.

## [0.5.0] - 2026-06-01

### Added
- **FEAT-1** — PPTX slide layout / slide master text inheritance. Placeholder shapes that are empty in a slide now inherit text from the slideLayout XML, falling back to the slideMaster XML. This recovers titles and subtitles defined only at the layout or master level.
- **FEAT-2** — DOCX streaming support via `parse_file_streaming`. The entire document is parsed and its sections are delivered as `SectionParsed` events, consistent with PPTX and XLSX streaming APIs. The previous `UnsupportedFormat` error is removed.

## [0.4.1] - 2026-06-01

### Fixed
- **BUG-1** — DOCX paragraphs with both a heading style (`<w:pStyle>`) and list numbering (`<w:numPr>`) no longer render as `# - item`. List items always take precedence over heading formatting.
- **BUG-2** — DOCX tables with vertically merged cells (`<w:vMerge>`) now correctly compute `rowspan` on the originating cell in the HTML table fallback. The `vMerge` origin is tracked at parse-time using a column-cursor + origin-map; continuation cells remain excluded from the cell list as before.

### Added
- **FEAT-3** — Markdown frontmatter now includes a `format: docx|xlsx|pptx` field, making the source document type explicit for LLM pipelines.

## [0.4.0] - 2026-05-31

### Added
- **WASM support**: `undoc-wasm` crate — parse DOCX/XLSX/PPTX in browser/Node.js via wasm-bindgen
- **GitHub Pages playground**: drag-drop live demo at https://iyulab.github.io/undoc/
- `@iyulab/undoc` npm package published on release via `wasm-pack`
- `parse()` module-level function and `OfficeDocument` class with `fromBytes`, `toMarkdown`, `toText`, `toJson`, `format`, `metadata` methods

### Changed
- Filesystem-dependent APIs (`parse_file`, `extract_text`, `to_markdown`, `to_text`, `to_json`, `parse_file_streaming`, parser `open()` methods) now cfg-gated on `not(target_arch = "wasm32")` — no behavior change on native targets
- `zip` dependency: disabled `lzma`/`bzip2`/`xz`/`zstd` compression features (OOXML uses Deflate only; removes C-compiler dependency for WASM builds)

## [0.3.0] - 2026-05-12

### Added

#### Streaming API
- `parse_file_streaming()` — processes PPTX slides and XLSX sheets with bounded memory
- `ParseEvent` enum: `DocumentStart` (with `image_map`), `SectionParsed`, `SectionFailed`, `DocumentEnd`, `ResourceExtracted`
- `SectionStreamOptions` — configure lenient mode and resource extraction for streaming
- `render_section_to_string()` — render a single section to Markdown (streaming renderer)

#### CLI Improvements (parity with unhwp v0.3.0)
- `convert` default output is now **Markdown only** (`extract.md` + `images/`); use `--all` or `--formats` for additional formats
- `--formats <md,txt,json>` — select output formats (comma-separated)
- `--all` — shorthand for `--formats md,txt,json`
- `--no-images` — skip binary resource extraction
- `--quiet` / `-q` — suppress progress output
- `--cleanup none` — explicit no-cleanup option added to `CleanupMode`

#### Architecture
- `cli/src/writer.rs` — `MultiFormatWriter` and `StreamingWriter` separated from CLI logic
- `cmd_convert` rewired to streaming pipeline for PPTX/XLSX; DOCX uses full-parse batch path

### Fixed
- CLI path sanitization in resource extraction (prevent path traversal attacks)
- `cli/Cargo.toml` dep version now exact (`"0.3.0"`) so CI version-check catches drift
- CI `version-check` job now also verifies `cli -> undoc dep` version alignment

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

## [0.1.1] - 2025-12-20

### Added

- **PPTX Table Parsing**
  - Full table extraction from PowerPoint slides (`a:tbl` elements)
  - Header row auto-detection for proper Markdown table rendering
  - Table content ordering (text before tables on each slide)

- **Smart Text Spacing**
  - CJK (Korean, Chinese, Japanese) character detection
  - Automatic spacing between CJK and ASCII characters
  - Intelligent run merging with `merge_adjacent_runs()`

### Fixed

- **Markdown Over-escaping**
  - Context-aware escaping for `*` and `_` characters
  - Fixed `(\* note)` patterns being incorrectly escaped
  - Fixed `*SYNC:` at line start being over-escaped
  - Properly handle emphasis markers near punctuation

### Changed

- **Code Refactoring**
  - Extracted `parse_core_metadata()` to shared container module
  - Removed ~90 lines of duplicate code across DOCX/PPTX/XLSX parsers
  - Improved code maintainability and single source of truth

## [0.1.2] - 2025-12-21

### Fixed

- **FFI Release Build**
  - Fixed GitHub Actions workflow where CLI build would overwrite the FFI-enabled library
  - FFI library artifacts are now preserved before CLI build to prevent filename collision
  - Added FFI export verification step to ensure `undoc_version` and other functions are properly exported
  - Release DLL now correctly contains all C-ABI functions (~1.5MB instead of 0.5MB)

### Changed

- **CI/CD Improvements**
  - Separated FFI library preservation step in release workflow
  - Added automated verification of FFI exports for all platforms
  - Improved error messages for missing exports

## [0.1.3] - 2025-12-21

### Fixed

- **Korean Text Quality**
  - Fixed word-level spacing in Korean DOCX conversion
  - Improved table cell text formatting

## [0.1.4] - 2025-12-21

### Added

- **Korean Word Spacing**
  - Smart word boundary detection for Korean text
  - Automatic spacing between CJK characters and ASCII

### Fixed

- **Table Rendering**
  - Fixed table cell content alignment issues
  - Improved nested table detection

## [0.1.5] - 2025-12-21

### Added

- **Image Extraction (Document Body)**
  - Extract images from `w:drawing` elements in document body
  - Support for alt text extraction from `wp:docPr`

### Fixed

- **Korean Word Spacing**
  - Source fidelity maintained (not a bug - follows original document)

## [0.1.6] - 2025-12-21

### Fixed

- **Image Parsing in Table Cells**
  - Added `w:drawing` element handling to `parse_table()` function
  - Images in table cells now correctly parsed to `para.images` vector
  - Support for `wp:docPr` alt text and `a:blip` resource references

## [0.1.7] - 2025-12-21

### Fixed

- **Image Rendering in Table Cells**
  - Fixed `render_cell_content()` to iterate over `para.images` vector
  - Images now correctly rendered as `![alt](path)` in markdown output
  - Root cause: Two-stage pipeline (parse → render) was incomplete

## [0.1.8] - 2025-12-21

### Added

- **FFI Resource Access API**
  - `undoc_get_resource_ids()`: Get all resource IDs as JSON array
  - `undoc_get_resource_info()`: Get resource metadata as JSON
  - `undoc_get_resource_data()`: Get binary data with length
  - `undoc_free_bytes()`: Free binary data allocated by `undoc_get_resource_data`
  - ID-based access pattern (vs index-based) for natural OOXML alignment
  - Enables C# object-oriented wrapper: `result.Images`, `result.Markdown`

## [0.2.0] - 2026-04-19

### Breaking

- **Strict root-part integrity (XLSX/PPTX)**
  - XLSX files missing `xl/workbook.xml` now return `Error::MissingComponent("xl/workbook.xml")`; previously returned an empty `Document`.
  - PPTX files missing `ppt/presentation.xml` now return `Error::MissingComponent("ppt/presentation.xml")`; previously returned an empty `Document`.
  - Consistent with 0.1.21 behavior for *malformed* root parts (already surfaces `Error::Encoding`). Missing root parts are the same integrity category and no longer fall through silently.
  - Migration: if prior code relied on empty-`Document` behavior for structurally-corrupt inputs, match on `Error::MissingComponent(path)` at the call site and construct an empty `Document` explicitly.

### Fixed

- **Mixed-entity round-trip across all OOXML parsers**
  - Text nodes containing both legitimate entities (e.g. `&amp;`) and malformed entities (e.g. `&bogus;`) in the same span now decode legitimate entities and preserve malformed tokens verbatim.
  - Previously the `quick_xml::escape::unescape` all-or-nothing failure caused the whole span to fall back to raw bytes, leaving legitimate entities over-escaped.
  - Affects DOCX body/textbox/nested tables, PPTX slide text, XLSX shared strings and inline `str` cells, chart labels, and OOXML metadata.

### Added

- **`src/decode.rs` module** — new crate-private module owning lenient XML entity decoding.
  - `lenient_unescape(&str) -> Cow<'_, str>` — fast path via `quick_xml::escape::unescape`; slow path scans `&...;` tokens within a 16-byte window and decodes each independently.
  - `decode_text_lossy(&BytesText) -> String` — content-text wrapper with `String::from_utf8_lossy` substitution.
  - `decode_text_strict(&BytesText, location) -> Result<String>` — metadata wrapper requiring valid UTF-8, surfacing `Error::xml_parse_with_context` on failure.

### Changed

- **Eliminated decoder duplication** — five duplicate `decode_*_lossless` helpers across `src/docx/parser.rs`, `src/pptx/parser.rs`, `src/xlsx/parser.rs`, `src/xlsx/shared_strings.rs`, `src/charts.rs` removed. 15 call sites now route through `crate::decode::decode_text_lossy`.
- **`container::metadata_text_or_raw`** delegates to `crate::decode::decode_text_strict`, gaining mixed-entity decoding while preserving strict-UTF-8 semantics.

## [0.2.1] - 2026-04-27

### Changed (behavior — review before upgrading)

- **`Block::PageBreak` no longer emits `\n\n---\n\n` by default.** Markdown has
  no page concept, and the rule fragments reading flow. Consumers that need
  the old behavior can opt back in via
  `RenderOptions::default().with_emit_page_breaks(true)` or the new
  `RenderOptions::lossless()` preset (which restores both page breaks and
  headers/footers in one shot).
- **DOCX section headers/footers are no longer included by default.** Their
  page-chrome content was contaminating LLM training data. Set
  `with_include_headers_footers(true)` (or use `lossless()`) to restore.
- **Heading paragraphs no longer emit redundant `**…**` wrappers** when the
  whole heading text is uniformly bold (a Word styling artifact). Partial
  bold runs inside headings are still preserved as authored emphasis. The
  same rule applies to header cells in tables. Toggle via
  `with_strip_redundant_emphasis_in_headings(false)` to revert.
- **`undoc md` (Markdown subcommand) now uses the heading analyzer by
  default**, matching the default `Convert` command. Documents that style
  headings via font size + bold (no explicit `Heading` paragraph style) are
  now detected as headings on this code path too.
- **`undoc text` (Text subcommand) now also enables the heading analyzer
  config** for parity with `md` and `convert`. The plain-text renderer
  currently ignores heading levels, so this is a no-op for output today; it
  unifies the three commands' option-building so future heading-aware text
  rendering can light up uniformly.

### Added

- `RenderOptions::lossless()` constructor — opt back into the previous
  rendering behavior in one call (page breaks + headers/footers).
- `RenderOptions::with_emit_page_breaks`,
  `with_include_headers_footers`, `with_callout_blockquote`,
  `with_strip_redundant_emphasis_in_headings` builder methods.
- `RenderOptions::callout_blockquote` (default `false`) — when enabled,
  single-row, single-column tables whose entire content is bold render as
  `> **…**` blockquotes instead of 1×1 markdown tables. Off by default; turn
  on only when the corpus is known to use 1×1 tables exclusively for
  callouts.
- **CLI: `--emit-page-breaks`, `--include-headers-footers`, `--lossless`
  flags on `undoc md` and `undoc convert`** — surface the new
  `RenderOptions` toggles at the CLI so users can opt back into the pre-0.2.1
  page-break / header-footer rendering without writing Rust. `--lossless` is
  a shortcut equivalent to setting both individual toggles. Without these
  flags the CLI follows the new 0.2.1 defaults (page breaks and
  headers/footers off).

### Fixed

- **Markdown escape policy** is now context-aware:
  - `|` is no longer escaped in regular paragraphs — only inside table
    cells where it is the column delimiter. (`v1.0 | 2026-04-27` was
    previously emitted as `v1.0 \| 2026-04-27`.)
  - Intra-word `_` is no longer escaped, per the CommonMark flanking rule
    that intra-word `_` cannot open or close emphasis. Identifiers like
    `YESUNG_OMS_backup`, `snake_case`, `in_house` now pass through
    verbatim instead of `YESUNG\_OMS\_backup`.
- **Tight lists** — consecutive list paragraphs are now joined by a single
  newline so they render as tight markdown lists. Previously every list
  item was separated by a blank line, forcing renderers into "loose list"
  mode with oversized vertical spacing.
- **Cell alignment fallback** — when a table cell has no explicit
  `<w:tcPr>/<w:jc>` (the common case), the renderer now falls back to the
  alignment of the cell's first paragraph (`<w:pPr>/<w:jc>`), recovering
  the visual intent that authors typically express via paragraph
  properties.

## [Unreleased]

### Planned

- Legacy format support (.doc, .xls, .ppt)
- Async I/O with Tokio
- Additional output formats (HTML, RST)
- Image optimization options
- Batch processing mode
- Plugin system for custom processors
- `undoc_get_paragraph_count()` FFI function
