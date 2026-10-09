"""Main undoc API for Python."""

import json
import os
from dataclasses import dataclass
from enum import IntEnum
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Union

from ._native import (
    get_library,
    UNDOC_FLAG_FRONTMATTER,
    UNDOC_FLAG_NO_ESCAPE,
    UNDOC_FLAG_PARAGRAPH_SPACING,
    UNDOC_FLAG_REFINE,
    UNDOC_JSON_PRETTY,
    UNDOC_JSON_COMPACT,
)
import ctypes


class ErrorKind(IntEnum):
    """Why an undoc call failed, so callers can branch on the reason.

    Values 1-13 mirror the library's own failure reasons; values 100+ are raised at
    the interop boundary and have no library-side counterpart. The numbers are part
    of the native ABI (``UndocErrorKind`` in ``undoc.h``): a new reason takes the next
    free number and existing ones are never renumbered, so an unrecognised value is
    kept as a plain :class:`int` rather than rejected.
    """

    NONE = 0
    OTHER = 1
    IO = 2
    UNKNOWN_FORMAT = 3
    UNSUPPORTED_FORMAT = 4
    ZIP_ARCHIVE = 5
    XML_PARSE = 6
    INVALID_DATA = 7
    MISSING_COMPONENT = 8
    ENCODING = 9
    STYLE_NOT_FOUND = 10
    RESOURCE_NOT_FOUND = 11
    ENCRYPTED = 12
    RENDER = 13
    INVALID_ARGUMENT = 100
    PANIC = 101
    INVALID_OUTPUT = 102
    SECTION_OUT_OF_RANGE = 300


class UndocError(Exception):
    """Exception raised when undoc operations fail.

    Attributes:
        kind: An :class:`ErrorKind`, or the raw integer if the native library reported
            a reason this build does not know about. Never :attr:`ErrorKind.NONE`,
            which means success. Defaults to :attr:`ErrorKind.OTHER` for failures that
            did not come from the native library.
    """

    def __init__(self, message: str, kind: int = ErrorKind.OTHER) -> None:
        super().__init__(message)
        try:
            self.kind: int = ErrorKind(kind)
        except ValueError:
            # Forward compatibility: a newer native library may report a number this
            # build has no name for. Keep it rather than losing the classification.
            self.kind = kind


def _decode_utf8_ptr(ptr: int) -> str:
    """Copy a null-terminated UTF-8 string from a non-null native pointer."""
    if not ptr:
        raise ValueError("native pointer is null")
    return ctypes.string_at(ptr).decode("utf-8")


def _copy_and_free_utf8_ptr(lib, ptr: int) -> str:
    """Copy a Rust-owned UTF-8 string, then free the original allocation."""
    try:
        return _decode_utf8_ptr(ptr)
    finally:
        lib.undoc_free_string(ptr)


def _get_last_error(lib=None) -> str:
    """Get the last error message from the native library."""
    lib = lib or get_library()
    error = lib.undoc_last_error()
    if error:
        return _decode_utf8_ptr(error)
    return "Unknown error"


def _get_last_error_kind(lib=None) -> int:
    """Get the classification of the last error from the native library.

    An unrecognised number passes through unchanged so a newer native library stays
    usable. Zero is the one value that cannot stand: it means success, and we only ask
    while building a failure.
    """
    lib = lib or get_library()
    kind = lib.undoc_last_error_kind()
    return ErrorKind.OTHER if kind == ErrorKind.NONE else kind


def _native_failure(action: str, lib=None) -> UndocError:
    """Build the error for a failed native call, with both message and classification.

    Every native failure goes through here so no raise site can quietly drop the
    classification and leave the caller with ``OTHER``.
    """
    lib = lib or get_library()
    return UndocError(f"{action}: {_get_last_error(lib)}", _get_last_error_kind(lib))


def _require_result_ptr(lib, ptr: Optional[int], action: str) -> int:
    """Require a non-null native result pointer for operations that signal failure via NULL."""
    if ptr:
        return ptr
    raise _native_failure(action, lib)


def version() -> str:
    """Get the undoc library version."""
    lib = get_library()
    ver = lib.undoc_version()
    return _decode_utf8_ptr(ver) if ver else "unknown"


def parse_file(path: Union[str, Path]) -> "Undoc":
    """Parse a document from a file path.

    Args:
        path: Path to the document file (.docx, .xlsx, .pptx, .doc, .xls, or .ppt)

    Returns:
        Undoc: Parsed document object

    Raises:
        UndocError: If parsing fails
        FileNotFoundError: If file doesn't exist
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")

    lib = get_library()
    handle = lib.undoc_parse_file(str(path).encode("utf-8"))
    if not handle:
        raise _native_failure(f"Failed to parse {path}")

    return Undoc(handle)


def parse_bytes(data: bytes) -> "Undoc":
    """Parse a document from bytes.

    Args:
        data: Document content as bytes

    Returns:
        Undoc: Parsed document object

    Raises:
        UndocError: If parsing fails
    """
    lib = get_library()
    data_ptr = (ctypes.c_uint8 * len(data)).from_buffer_copy(data)
    handle = lib.undoc_parse_bytes(data_ptr, len(data))
    if not handle:
        raise _native_failure("Failed to parse bytes")

    return Undoc(handle)


@dataclass(frozen=True)
class RenderedSection:
    """A rendered section: a PNG, its size in pixels, and what it could not show.

    ``gaps`` counts, by kind, what the renderer left out — ``shapes`` (custom geometry),
    ``images`` (pictures other than PNG and JPEG), ``text_runs`` (text no face covers, in a
    script that needs shaping, or vertical), ``charts``, ``graphic_frames`` (embedded objects, SmartArt with no drawing)
    and ``approximated_fills`` (fills drawn as a stand-in: a pattern in its foreground color,
    a rectangular or shape-following gradient as a radial one, a tiled picture stretched). All
    zero means everything was painted; otherwise the rest of the slide still was.

    ``substituted_text_runs`` counts text drawn in a face standing in for the one it asks
    for — readable, but not the slide's own typeface. It is not a gap.
    """

    png: bytes
    width: int
    height: int
    gaps: "dict[str, int]"
    substituted_text_runs: int = 0


class Undoc:
    """Represents a parsed Office document.

    This class provides methods to extract content from DOCX, XLSX, and PPTX
    documents in various formats (Markdown, plain text, JSON).
    """

    def __init__(self, handle: ctypes.c_void_p):
        """Initialize with a native document handle.

        Args:
            handle: Native document handle from undoc_parse_file/undoc_parse_bytes
        """
        self._handle = handle
        self._lib = get_library()

    def __del__(self):
        """Free the native document handle."""
        if hasattr(self, "_handle") and self._handle:
            self._lib.undoc_free_document(self._handle)
            self._handle = None

    def __enter__(self) -> "Undoc":
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._handle:
            self._lib.undoc_free_document(self._handle)
            self._handle = None

    def to_markdown(
        self,
        frontmatter: bool = False,
        escape_special: bool = True,
        paragraph_spacing: bool = False,
        refine: bool = False,
    ) -> str:
        """Convert document to Markdown.

        Args:
            frontmatter: Include YAML frontmatter with metadata
            escape_special: Escape special Markdown characters, so text that
                reads as Markdown syntax stays text (default, as in Rust)
            paragraph_spacing: Add extra spacing between paragraphs
            refine: Apply the lossless, idempotent markdown shape-refinement
                pass (table shape, ordered-list numbering, link/image paths,
                frontmatter, section anchors) after rendering

        Returns:
            Markdown string

        Raises:
            UndocError: If conversion fails
        """
        flags = 0
        if frontmatter:
            flags |= UNDOC_FLAG_FRONTMATTER
        if not escape_special:
            flags |= UNDOC_FLAG_NO_ESCAPE
        if paragraph_spacing:
            flags |= UNDOC_FLAG_PARAGRAPH_SPACING
        if refine:
            flags |= UNDOC_FLAG_REFINE

        result = self._lib.undoc_to_markdown(self._handle, flags)
        if not result:
            raise _native_failure("Failed to convert to markdown")

        return _copy_and_free_utf8_ptr(self._lib, result)

    def to_text(self) -> str:
        """Convert document to plain text.

        Returns:
            Plain text string

        Raises:
            UndocError: If conversion fails
        """
        result = self._lib.undoc_to_text(self._handle)
        if not result:
            raise _native_failure("Failed to convert to text")

        return _copy_and_free_utf8_ptr(self._lib, result)

    def to_json(self, compact: bool = False) -> str:
        """Convert document to JSON.

        Args:
            compact: Use compact JSON format (no indentation)

        Returns:
            JSON string

        Raises:
            UndocError: If conversion fails
        """
        fmt = UNDOC_JSON_COMPACT if compact else UNDOC_JSON_PRETTY
        result = self._lib.undoc_to_json(self._handle, fmt)
        if not result:
            raise _native_failure("Failed to convert to JSON")

        return _copy_and_free_utf8_ptr(self._lib, result)

    def plain_text(self) -> str:
        """Get plain text content (faster than to_text for simple extraction).

        Returns:
            Plain text string

        Raises:
            UndocError: If extraction fails
        """
        result = self._lib.undoc_plain_text(self._handle)
        if not result:
            raise _native_failure("Failed to get plain text")

        return _copy_and_free_utf8_ptr(self._lib, result)

    @property
    def section_count(self) -> int:
        """Get the number of sections in the document."""
        count = self._lib.undoc_section_count(self._handle)
        if count < 0:
            raise _native_failure("Failed to get section count")
        return count

    @property
    def resource_count(self) -> int:
        """Get the number of resources (images, etc.) in the document."""
        count = self._lib.undoc_resource_count(self._handle)
        if count < 0:
            raise _native_failure("Failed to get resource count")
        return count

    @property
    def title(self) -> Optional[str]:
        """Get the document title, if set."""
        result = self._lib.undoc_get_title(self._handle)
        if result:
            return _copy_and_free_utf8_ptr(self._lib, result)
        return None

    @property
    def author(self) -> Optional[str]:
        """Get the document author, if set."""
        result = self._lib.undoc_get_author(self._handle)
        if result:
            return _copy_and_free_utf8_ptr(self._lib, result)
        return None

    def get_resource_ids(self) -> List[str]:
        """Get list of resource IDs in the document.

        Returns:
            List of resource ID strings
        """
        result = self._lib.undoc_get_resource_ids(self._handle)
        result = _require_result_ptr(self._lib, result, "Failed to get resource IDs")
        ids = json.loads(_copy_and_free_utf8_ptr(self._lib, result))
        return ids

    def get_resource_info(self, resource_id: str) -> Optional[Dict]:
        """Get metadata for a resource.

        Args:
            resource_id: The resource ID

        Returns:
            Dictionary with resource metadata, or None if not found
        """
        result = self._lib.undoc_get_resource_info(
            self._handle, resource_id.encode("utf-8")
        )
        if not result:
            return None

        return json.loads(_copy_and_free_utf8_ptr(self._lib, result))

    def get_resource_data(self, resource_id: str) -> Optional[bytes]:
        """Get binary data for a resource.

        Args:
            resource_id: The resource ID

        Returns:
            Resource data as bytes, or None if not found
        """
        length = ctypes.c_size_t()
        data_ptr = self._lib.undoc_get_resource_data(
            self._handle, resource_id.encode("utf-8"), ctypes.byref(length)
        )
        if not data_ptr:
            return None

        # Copy data before freeing
        data = bytes(data_ptr[: length.value])
        self._lib.undoc_free_bytes(data_ptr, length.value)

        return data

    def render_section(
        self,
        index: int,
        dpi: float = 150.0,
        font_dirs: Sequence[Union[str, "os.PathLike[str]"]] = (),
        system_fonts: bool = True,
    ) -> RenderedSection:
        """Render a section to a PNG. A section of a presentation is a slide; it is painted
        from the package this document was parsed from, without reading the file again.

        Anything the renderer cannot paint yet (charts, embedded objects and other graphic frames,
        custom geometry, pictures other than PNG and JPEG, text no face covers) is left out
        and counted in :attr:`RenderedSection.gaps`; the rest of the slide is painted.

        No font is bundled. Text is drawn in faces found in ``font_dirs`` (searched with
        their subdirectories), then — with ``system_fonts`` — in the system's font
        directories; on a host without fonts, pass a directory of font files.

        Args:
            index: Section index (0-based), in presentation order.
            dpi: Resolution; a slide point is ``dpi / 72`` pixels.
            font_dirs: Directories searched for font files.
            system_fonts: Whether the system's font directories are searched too.

        Raises:
            UndocError: ``kind == ErrorKind.SECTION_OUT_OF_RANGE`` for a section the
                document does not have; ``UNSUPPORTED_FORMAT`` for a document that is not
                a .pptx presentation; ``RENDER`` for a resolution it cannot draw.
        """
        options = json.dumps(
            {
                "dpi": dpi,
                "font_dirs": [os.fspath(d) for d in font_dirs],
                "system_fonts": system_fonts,
            }
        ).encode("utf-8")
        out_len = ctypes.c_size_t(0)
        info = ctypes.c_void_p(None)
        result = self._lib.undoc_render_section(
            self._handle, index, options, ctypes.byref(out_len), ctypes.byref(info)
        )
        if not result:
            raise _native_failure(f"Failed to render section {index}", self._lib)
        try:
            png = ctypes.string_at(result, out_len.value)
        finally:
            self._lib.undoc_free_bytes(result, out_len.value)
        report = json.loads(_copy_and_free_utf8_ptr(self._lib, info.value)) if info.value else {}
        return RenderedSection(
            png=png,
            width=int(report.get("width", 0)),
            height=int(report.get("height", 0)),
            gaps=dict(report.get("gaps", {})),
            substituted_text_runs=int(report.get("substituted_text_runs", 0)),
        )
