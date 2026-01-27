"""Native library loading for undoc."""

import ctypes
import platform
import os
from pathlib import Path

# Library filename by platform
_LIB_NAMES = {
    "Windows": "undoc.dll",
    "Linux": "libundoc.so",
    "Darwin": "libundoc.dylib",
}

# Runtime identifier
_RUNTIME_IDS = {
    ("Windows", "AMD64"): "win-x64",
    ("Windows", "x86_64"): "win-x64",
    ("Linux", "x86_64"): "linux-x64",
    ("Darwin", "x86_64"): "osx-x64",
    ("Darwin", "arm64"): "osx-arm64",
}


def _get_lib_path() -> Path:
    """Get the path to the native library."""
    system = platform.system()
    machine = platform.machine()

    lib_name = _LIB_NAMES.get(system)
    if not lib_name:
        raise OSError(f"Unsupported platform: {system}")

    runtime_id = _RUNTIME_IDS.get((system, machine))
    if not runtime_id:
        raise OSError(f"Unsupported architecture: {system}/{machine}")

    # Look for the library in the package
    package_dir = Path(__file__).parent
    lib_path = package_dir / "lib" / runtime_id / lib_name

    if lib_path.exists():
        return lib_path

    # Fallback: look in package root
    lib_path = package_dir / "lib" / lib_name
    if lib_path.exists():
        return lib_path

    # Fallback: look in current directory
    lib_path = Path(lib_name)
    if lib_path.exists():
        return lib_path

    # Fallback: system library path
    return Path(lib_name)


def _load_library() -> ctypes.CDLL:
    """Load the native undoc library."""
    lib_path = _get_lib_path()

    try:
        if platform.system() == "Windows":
            # On Windows, use LoadLibraryEx with LOAD_WITH_ALTERED_SEARCH_PATH
            return ctypes.CDLL(str(lib_path), winmode=0)
        else:
            return ctypes.CDLL(str(lib_path))
    except OSError as e:
        raise OSError(
            f"Failed to load undoc native library from {lib_path}: {e}\n"
            f"Make sure the library is installed for your platform ({platform.system()}/{platform.machine()})."
        ) from e


# Load the library
_lib = _load_library()

# Define function signatures
_lib.undoc_version.argtypes = []
_lib.undoc_version.restype = ctypes.c_char_p

_lib.undoc_last_error.argtypes = []
_lib.undoc_last_error.restype = ctypes.c_char_p

_lib.undoc_parse_file.argtypes = [ctypes.c_char_p]
_lib.undoc_parse_file.restype = ctypes.c_void_p

_lib.undoc_parse_bytes.argtypes = [ctypes.POINTER(ctypes.c_uint8), ctypes.c_size_t]
_lib.undoc_parse_bytes.restype = ctypes.c_void_p

_lib.undoc_free_document.argtypes = [ctypes.c_void_p]
_lib.undoc_free_document.restype = None

_lib.undoc_to_markdown.argtypes = [ctypes.c_void_p, ctypes.c_int]
_lib.undoc_to_markdown.restype = ctypes.c_char_p

_lib.undoc_to_text.argtypes = [ctypes.c_void_p]
_lib.undoc_to_text.restype = ctypes.c_char_p

_lib.undoc_to_json.argtypes = [ctypes.c_void_p, ctypes.c_int]
_lib.undoc_to_json.restype = ctypes.c_char_p

_lib.undoc_plain_text.argtypes = [ctypes.c_void_p]
_lib.undoc_plain_text.restype = ctypes.c_char_p

_lib.undoc_section_count.argtypes = [ctypes.c_void_p]
_lib.undoc_section_count.restype = ctypes.c_int

_lib.undoc_resource_count.argtypes = [ctypes.c_void_p]
_lib.undoc_resource_count.restype = ctypes.c_int

_lib.undoc_get_title.argtypes = [ctypes.c_void_p]
_lib.undoc_get_title.restype = ctypes.c_char_p

_lib.undoc_get_author.argtypes = [ctypes.c_void_p]
_lib.undoc_get_author.restype = ctypes.c_char_p

_lib.undoc_free_string.argtypes = [ctypes.c_char_p]
_lib.undoc_free_string.restype = None

_lib.undoc_get_resource_ids.argtypes = [ctypes.c_void_p]
_lib.undoc_get_resource_ids.restype = ctypes.c_char_p

_lib.undoc_get_resource_info.argtypes = [ctypes.c_void_p, ctypes.c_char_p]
_lib.undoc_get_resource_info.restype = ctypes.c_char_p

_lib.undoc_get_resource_data.argtypes = [
    ctypes.c_void_p,
    ctypes.c_char_p,
    ctypes.POINTER(ctypes.c_size_t),
]
_lib.undoc_get_resource_data.restype = ctypes.POINTER(ctypes.c_uint8)

_lib.undoc_free_bytes.argtypes = [ctypes.POINTER(ctypes.c_uint8), ctypes.c_size_t]
_lib.undoc_free_bytes.restype = None

# Export constants
UNDOC_FLAG_FRONTMATTER = 1
UNDOC_FLAG_ESCAPE_SPECIAL = 2
UNDOC_FLAG_PARAGRAPH_SPACING = 4

UNDOC_JSON_PRETTY = 0
UNDOC_JSON_COMPACT = 1


def get_library():
    """Get the loaded native library."""
    return _lib
