"""Main undoc API for Python."""

import json
from pathlib import Path
from typing import Dict, List, Optional, Union

from ._native import (
    get_library,
    UNDOC_FLAG_FRONTMATTER,
    UNDOC_FLAG_ESCAPE_SPECIAL,
    UNDOC_FLAG_PARAGRAPH_SPACING,
    UNDOC_JSON_PRETTY,
    UNDOC_JSON_COMPACT,
)
import ctypes


class UndocError(Exception):
    """Exception raised when undoc operations fail."""

    pass


def _get_last_error() -> str:
    """Get the last error message from the native library."""
    lib = get_library()
    error = lib.undoc_last_error()
    if error:
        return error.decode("utf-8")
    return "Unknown error"


def version() -> str:
    """Get the undoc library version."""
    lib = get_library()
    ver = lib.undoc_version()
    return ver.decode("utf-8") if ver else "unknown"


def parse_file(path: Union[str, Path]) -> "Undoc":
    """Parse a document from a file path.

    Args:
        path: Path to the document file (.docx, .xlsx, or .pptx)

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
        raise UndocError(f"Failed to parse {path}: {_get_last_error()}")

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
        raise UndocError(f"Failed to parse bytes: {_get_last_error()}")

    return Undoc(handle)


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
        escape_special: bool = False,
        paragraph_spacing: bool = False,
    ) -> str:
        """Convert document to Markdown.

        Args:
            frontmatter: Include YAML frontmatter with metadata
            escape_special: Escape special Markdown characters
            paragraph_spacing: Add extra spacing between paragraphs

        Returns:
            Markdown string

        Raises:
            UndocError: If conversion fails
        """
        flags = 0
        if frontmatter:
            flags |= UNDOC_FLAG_FRONTMATTER
        if escape_special:
            flags |= UNDOC_FLAG_ESCAPE_SPECIAL
        if paragraph_spacing:
            flags |= UNDOC_FLAG_PARAGRAPH_SPACING

        result = self._lib.undoc_to_markdown(self._handle, flags)
        if not result:
            raise UndocError(f"Failed to convert to markdown: {_get_last_error()}")

        return result.decode("utf-8")

    def to_text(self) -> str:
        """Convert document to plain text.

        Returns:
            Plain text string

        Raises:
            UndocError: If conversion fails
        """
        result = self._lib.undoc_to_text(self._handle)
        if not result:
            raise UndocError(f"Failed to convert to text: {_get_last_error()}")

        return result.decode("utf-8")

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
            raise UndocError(f"Failed to convert to JSON: {_get_last_error()}")

        return result.decode("utf-8")

    def plain_text(self) -> str:
        """Get plain text content (faster than to_text for simple extraction).

        Returns:
            Plain text string

        Raises:
            UndocError: If extraction fails
        """
        result = self._lib.undoc_plain_text(self._handle)
        if not result:
            raise UndocError(f"Failed to get plain text: {_get_last_error()}")

        return result.decode("utf-8")

    @property
    def section_count(self) -> int:
        """Get the number of sections in the document."""
        count = self._lib.undoc_section_count(self._handle)
        if count < 0:
            raise UndocError(f"Failed to get section count: {_get_last_error()}")
        return count

    @property
    def resource_count(self) -> int:
        """Get the number of resources (images, etc.) in the document."""
        count = self._lib.undoc_resource_count(self._handle)
        if count < 0:
            raise UndocError(f"Failed to get resource count: {_get_last_error()}")
        return count

    @property
    def title(self) -> Optional[str]:
        """Get the document title, if set."""
        result = self._lib.undoc_get_title(self._handle)
        if result:
            return result.decode("utf-8")
        return None

    @property
    def author(self) -> Optional[str]:
        """Get the document author, if set."""
        result = self._lib.undoc_get_author(self._handle)
        if result:
            return result.decode("utf-8")
        return None

    def get_resource_ids(self) -> List[str]:
        """Get list of resource IDs in the document.

        Returns:
            List of resource ID strings
        """
        result = self._lib.undoc_get_resource_ids(self._handle)
        if not result:
            return []

        ids = json.loads(result.decode("utf-8"))
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

        return json.loads(result.decode("utf-8"))

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
