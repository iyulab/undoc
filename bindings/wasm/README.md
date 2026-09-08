# @iyulab/undoc (WASM)

Browser/Node.js WebAssembly package for [undoc](https://github.com/iyulab/undoc).

Extracts DOCX, XLSX, and PPTX files to Markdown, plain text, or JSON — no server required.

## Install

```bash
npm install @iyulab/undoc
```

## Usage (ESM)

> This package is built with `wasm-pack --target bundler`. It is an ES module and the
> WebAssembly binary is initialised for you, so there is no `init()` to await — import
> the functions and call them. Use it through a bundler (webpack, Vite, Rollup, esbuild).

```js
import { parse } from '@iyulab/undoc';

const bytes = new Uint8Array(await file.arrayBuffer());
const doc = parse(bytes);

console.log(doc.format());       // "docx" | "xlsx" | "pptx"
console.log(doc.toMarkdown());   // Markdown string
console.log(doc.toText());       // Plain text string
console.log(doc.toJson());       // JSON string
```

## API

### `supportedFormats(): string`

JSON array of the formats this package can parse:

```js
import { supportedFormats } from '@iyulab/undoc';

JSON.parse(supportedFormats());
// [{ extension: "docx", name: "Word Document" },
//  { extension: "xlsx", name: "Excel Workbook" },
//  { extension: "pptx", name: "PowerPoint Presentation" }]
```

Ask rather than hardcoding the extensions on your side — a local copy of the list goes stale
the moment this package learns a new format, and nothing tells you it has.

### `parse(data: Uint8Array): OfficeDocument`

Parse a DOCX, XLSX, or PPTX byte array. Throws if the format is unrecognized.

### `OfficeDocument`

| Method | Returns | Description |
|--------|---------|-------------|
| `fromBytes(data)` | `OfficeDocument` | Alias for module-level `parse()` |
| `format()` | `string` | `"docx"` \| `"xlsx"` \| `"pptx"` |
| `toMarkdown()` | `string` | Full document as Markdown |
| `toText()` | `string` | Plain text extraction |
| `toJson()` | `string` | Structured JSON |
| `metadata()` | `string` | JSON of title/author/subject etc. |

## Playground

Live demo: https://iyulab.github.io/undoc/
