"""undoc - High-performance Microsoft Office document extraction.

This package provides Python bindings for the undoc library, which extracts
DOCX, XLSX, and PPTX documents into structured Markdown with assets.
"""

from .undoc import Undoc, UndocError, parse_file, parse_bytes, version

__all__ = ["Undoc", "UndocError", "parse_file", "parse_bytes", "version"]
__version__ = "0.1.11"
