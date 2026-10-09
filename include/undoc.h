/**
 * undoc - Microsoft Office Document Extraction Library
 *
 * High-performance library for extracting content from DOCX, XLSX, and PPTX files.
 * Converts documents to Markdown, plain text, or JSON, and renders slides to PNG.
 *
 * Copyright (c) 2024 iyulab
 * MIT License
 */

#ifndef UNDOC_H
#define UNDOC_H

#include <stddef.h>
#include <stdint.h>

#ifdef __cplusplus
extern "C" {
#endif

/* Opaque document handle */
typedef struct UndocDocument UndocDocument;

/* Flags for markdown rendering */
#define UNDOC_FLAG_FRONTMATTER      1  /* Include YAML frontmatter */
#define UNDOC_FLAG_ESCAPE_SPECIAL   2  /* Escape special Markdown characters */
#define UNDOC_FLAG_PARAGRAPH_SPACING 4 /* Add blank lines between paragraphs */
#define UNDOC_FLAG_REFINE            8 /* Apply the shape-refinement pass to the Markdown */

/* JSON format options */
#define UNDOC_JSON_PRETTY   0  /* Pretty-printed JSON with indentation */
#define UNDOC_JSON_COMPACT  1  /* Compact JSON without whitespace */

/**
 * Why the last call failed, as returned by undoc_last_error_kind().
 *
 * Values 1..=13 and 300..=399 mirror the library's own failure reasons; values
 * 100..=199 are raised at the FFI boundary and have no library-side counterpart. These numbers are a stable
 * ABI contract: a new reason takes the next free number and existing ones are never
 * reused or renumbered. Treat an unrecognised value as a generic failure rather than
 * as an error, so that a newer library stays usable by older callers.
 */
typedef enum UndocErrorKind {
    UNDOC_ERROR_NONE               = 0,   /* The last call succeeded */
    UNDOC_ERROR_OTHER              = 1,   /* Failure with no more specific reason */
    UNDOC_ERROR_IO                 = 2,   /* Missing or unreadable file */
    UNDOC_ERROR_UNKNOWN_FORMAT     = 3,   /* Not a recognised Office document */
    UNDOC_ERROR_UNSUPPORTED_FORMAT = 4,   /* Recognised but not supported */
    UNDOC_ERROR_ZIP_ARCHIVE        = 5,   /* The OOXML container could not be read */
    UNDOC_ERROR_XML_PARSE          = 6,   /* XML content could not be parsed */
    UNDOC_ERROR_INVALID_DATA       = 7,   /* Malformed data inside the document */
    UNDOC_ERROR_MISSING_COMPONENT  = 8,   /* A required document part is absent */
    UNDOC_ERROR_ENCODING           = 9,   /* Text encoding conversion failed */
    UNDOC_ERROR_STYLE_NOT_FOUND    = 10,  /* A referenced style is absent */
    UNDOC_ERROR_RESOURCE_NOT_FOUND = 11,  /* A referenced resource is absent */
    UNDOC_ERROR_ENCRYPTED          = 12,  /* The document is encrypted */
    UNDOC_ERROR_RENDER             = 13,  /* Rendering the output failed */
    UNDOC_ERROR_INVALID_ARGUMENT   = 100, /* An argument was NULL or not valid UTF-8 */
    UNDOC_ERROR_PANIC              = 101, /* A panic was caught at the boundary */
    UNDOC_ERROR_INVALID_OUTPUT     = 102, /* Output holds a NUL byte, cannot cross ABI */
    UNDOC_ERROR_SECTION_OUT_OF_RANGE = 300 /* A section index the document does not have */
} UndocErrorKind;

/**
 * Get the library version.
 *
 * @return Static version string (do not free)
 */
const char* undoc_version(void);

/**
 * Get the last error message.
 *
 * Call this after a function returns NULL to get the error description.
 *
 * @return Error message or NULL if no error. Do not free.
 */
const char* undoc_last_error(void);

/**
 * Classify the last error without parsing its message.
 *
 * Returns UNDOC_ERROR_NONE (0) when the last call on this thread succeeded. Written
 * and cleared in lockstep with undoc_last_error(), so a message is never paired with
 * a stale kind.
 *
 * @return A UndocErrorKind value; treat an unrecognised value as a generic failure.
 */
int undoc_last_error_kind(void);

/**
 * Parse a document from a file path.
 *
 * Automatically detects format from file extension and content.
 * Supports .docx, .xlsx, and .pptx files.
 *
 * @param path Path to the document file (UTF-8 encoded)
 * @return Document handle or NULL on error. Must be freed with undoc_free_document().
 */
UndocDocument* undoc_parse_file(const char* path);

/**
 * Parse a document from a byte buffer.
 *
 * @param data Pointer to document data
 * @param len Length of data in bytes
 * @return Document handle or NULL on error. Must be freed with undoc_free_document().
 */
UndocDocument* undoc_parse_bytes(const uint8_t* data, size_t len);

/**
 * Free a document handle.
 *
 * @param doc Document handle (may be NULL)
 */
void undoc_free_document(UndocDocument* doc);

/**
 * Convert a document to Markdown.
 *
 * @param doc Document handle
 * @param flags Bitwise OR of UNDOC_FLAG_* constants
 * @return Markdown string or NULL on error. Must be freed with undoc_free_string().
 */
char* undoc_to_markdown(const UndocDocument* doc, int flags);

/**
 * Convert a document to plain text.
 *
 * @param doc Document handle
 * @return Plain text string or NULL on error. Must be freed with undoc_free_string().
 */
char* undoc_to_text(const UndocDocument* doc);

/**
 * Convert a document to JSON.
 *
 * @param doc Document handle
 * @param format UNDOC_JSON_PRETTY or UNDOC_JSON_COMPACT
 * @return JSON string or NULL on error. Must be freed with undoc_free_string().
 */
char* undoc_to_json(const UndocDocument* doc, int format);

/**
 * Get plain text content directly.
 *
 * @param doc Document handle
 * @return Plain text or NULL on error. Must be freed with undoc_free_string().
 */
char* undoc_plain_text(const UndocDocument* doc);

/**
 * Get the number of sections in a document.
 *
 * For Word documents, sections are page sections.
 * For Excel, sections are worksheets.
 * For PowerPoint, sections are slides.
 *
 * @param doc Document handle
 * @return Section count or -1 on error
 */
int undoc_section_count(const UndocDocument* doc);

/**
 * Get the number of embedded resources.
 *
 * Resources include images, media files, and other embedded objects.
 *
 * @param doc Document handle
 * @return Resource count or -1 on error
 */
int undoc_resource_count(const UndocDocument* doc);

/**
 * Get the document title.
 *
 * @param doc Document handle
 * @return Title or NULL if not set. Must be freed with undoc_free_string().
 */
char* undoc_get_title(const UndocDocument* doc);

/**
 * Get the document author.
 *
 * @param doc Document handle
 * @return Author or NULL if not set. Must be freed with undoc_free_string().
 */
char* undoc_get_author(const UndocDocument* doc);

/**
 * Get the ids of all embedded resources as a JSON array.
 *
 * A document with no resources returns "[]", never NULL.
 *
 * @param doc Document handle
 * @return JSON array such as ["rId1", "rId2"], or NULL on error.
 *         Must be freed with undoc_free_string().
 */
char* undoc_get_resource_ids(const UndocDocument* doc);

/**
 * Get one resource's metadata as a JSON object, without its bytes.
 *
 * The object carries id, type, filename, mime_type, size, width, height, alt_text, role
 * and companion_of. role is "primary" for an image the document shows (or media it plays),
 * "alternate" for another encoding of one (a picture's SVG original) and "layer" for a
 * layer composited onto one; companion_of is the id of the primary it belongs to.
 *
 * @param doc Document handle
 * @param resource_id Resource id (UTF-8, NUL-terminated)
 * @return JSON object, or NULL if the resource does not exist or on error.
 *         Must be freed with undoc_free_string().
 */
char* undoc_get_resource_info(const UndocDocument* doc, const char* resource_id);

/**
 * Get one resource's bytes.
 *
 * @param doc Document handle
 * @param resource_id Resource id (UTF-8, NUL-terminated)
 * @param out_len Receives the length of the returned buffer. Left untouched when an
 *                argument is rejected.
 * @return Buffer, or NULL if the resource does not exist or on error.
 *         Must be freed with undoc_free_bytes() together with *out_len.
 */
uint8_t* undoc_get_resource_data(const UndocDocument* doc, const char* resource_id, size_t* out_len);

/**
 * Render a section to a PNG. A section of a .pptx presentation is a slide; it is painted
 * by the parser the handle keeps, with no second read of the file.
 *
 * Anything the renderer cannot paint yet (charts, tables and other graphic frames, custom
 * geometry, pictures other than PNG and JPEG, text no face covers) is left out and counted
 * in out_info; the rest of the slide is painted. No font is bundled: on a host without
 * fonts, pass font_dirs, or text is counted as a gap.
 *
 * @param doc Document handle
 * @param index 0-based section index, in presentation order.
 * @param options_json NULL, or {"dpi": 150, "font_dirs": ["..."], "system_fonts": true}:
 *        resolution, directories searched (with subdirectories) for font files, and
 *        whether the system's font directories are searched too.
 * @param out_len Receives the PNG length in bytes (0 on error). Left untouched when doc or
 *                out_len is NULL.
 * @param out_info NULL, or receives {"width":N,"height":N,"gaps":{"shapes":N,"images":N,
 *        "text_runs":N,"charts":N,"graphic_frames":N,"approximated_fills":N},
 *        "substituted_text_runs":N} (must be freed with undoc_free_string).
 * @return PNG bytes (must be freed with undoc_free_bytes() together with *out_len), or
 *         NULL on error (UNDOC_ERROR_SECTION_OUT_OF_RANGE, UNDOC_ERROR_UNSUPPORTED_FORMAT
 *         for a document that is not a .pptx presentation, UNDOC_ERROR_INVALID_ARGUMENT).
 */
uint8_t* undoc_render_section(const UndocDocument* doc,
                              int index,
                              const char* options_json,
                              size_t* out_len,
                              char** out_info);

/**
 * Free a string allocated by this library.
 *
 * @param str String pointer (may be NULL)
 */
void undoc_free_string(char* str);

/**
 * Free a buffer returned by undoc_get_resource_data() or undoc_render_section().
 *
 * @param data Buffer pointer (may be NULL)
 * @param len The length the call wrote to out_len
 */
void undoc_free_bytes(uint8_t* data, size_t len);

#ifdef __cplusplus
}
#endif

#endif /* UNDOC_H */
