/**
 * undoc - Microsoft Office Document Extraction Library
 *
 * High-performance library for extracting content from DOCX, XLSX, and PPTX files.
 * Converts documents to Markdown, plain text, or JSON.
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

/* JSON format options */
#define UNDOC_JSON_PRETTY   0  /* Pretty-printed JSON with indentation */
#define UNDOC_JSON_COMPACT  1  /* Compact JSON without whitespace */

/**
 * Why the last call failed, as returned by undoc_last_error_kind().
 *
 * Values 1..=13 mirror the library's own failure reasons; values 100+ are raised at
 * the FFI boundary and have no library-side counterpart. These numbers are a stable
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
    UNDOC_ERROR_INVALID_OUTPUT     = 102  /* Output holds a NUL byte, cannot cross ABI */
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
 * Free a string allocated by this library.
 *
 * @param str String pointer (may be NULL)
 */
void undoc_free_string(char* str);

#ifdef __cplusplus
}
#endif

#endif /* UNDOC_H */
