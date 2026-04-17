"""Tests for undoc Python bindings."""

import ctypes
import io
import os
import platform
import pytest
import subprocess
import sys
import zipfile
from pathlib import Path

# Import eagerly so native-backed verification can fail hard when requested.
NATIVE_IMPORT_ERROR = None

try:
    import undoc.undoc as undoc_module
    from undoc import Undoc, UndocError, parse_file, parse_bytes, version

    LIBRARY_AVAILABLE = True
except OSError as exc:
    undoc_module = None
    Undoc = None
    UndocError = Exception
    parse_file = None
    parse_bytes = None
    version = None
    LIBRARY_AVAILABLE = False
    NATIVE_IMPORT_ERROR = exc


# Get test files directory
TEST_FILES_DIR = Path(__file__).parent.parent.parent.parent / "test-files"


def _native_library_filename() -> str:
    system = platform.system()
    if system == "Windows":
        return "undoc.dll"
    if system == "Darwin":
        return "libundoc.dylib"
    return "libundoc.so"


if os.environ.get("UNDOC_REQUIRE_NATIVE") == "1" and not LIBRARY_AVAILABLE:
    configured_path = os.environ.get("UNDOC_LIB_PATH")
    configured_suffix = (
        f" (UNDOC_LIB_PATH={configured_path})" if configured_path else ""
    )
    raise RuntimeError(
        "UNDOC_REQUIRE_NATIVE=1 but undoc native bindings failed to load"
        f"{configured_suffix}: {NATIVE_IMPORT_ERROR}"
    ) from NATIVE_IMPORT_ERROR


def create_minimal_docx_bytes(text: str = "Привет из Python") -> bytes:
    """Create a tiny DOCX fixture without relying on external test files."""
    document_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>{text}</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"""

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_STORED) as zf:
        zf.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>""",
        )
        zf.writestr(
            "_rels/.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>""",
        )
        zf.writestr(
            "word/_rels/document.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>""",
        )
        zf.writestr("word/document.xml", document_xml)
    return buf.getvalue()


class TestNativeVerificationHarness:
    def test_strict_native_mode_fails_closed_on_invalid_library_path(self, tmp_path):
        bad_library = tmp_path / _native_library_filename()
        bad_library.write_text("not a real shared library", encoding="utf-8")

        env = os.environ.copy()
        env["UNDOC_REQUIRE_NATIVE"] = "1"
        env["UNDOC_LIB_PATH"] = str(bad_library)
        env["PYTHONPATH"] = str(Path(__file__).resolve().parents[1] / "src")

        result = subprocess.run(
            [
                sys.executable,
                "-m",
                "pytest",
                str(Path(__file__)),
                "-k",
                "test_version_returns_string",
                "-q",
            ],
            cwd=Path(__file__).resolve().parents[3],
            capture_output=True,
            text=True,
            env=env,
            check=False,
        )

        assert result.returncode != 0
        combined_output = result.stdout + result.stderr
        assert "UNDOC_REQUIRE_NATIVE=1" in combined_output
        assert str(bad_library) in combined_output


class FakeStringLibrary:
    """Minimal fake library for ownership/UTF-8 regression tests."""

    def __init__(self):
        self._buffers = []
        self.freed = []

    def _alloc(self, text: str) -> int:
        buf = ctypes.create_string_buffer(text.encode("utf-8"))
        self._buffers.append(buf)
        return ctypes.addressof(buf)

    def undoc_last_error(self):
        return self._alloc("Ошибка native")

    def undoc_version(self):
        return self._alloc("1.2.3")

    def undoc_free_string(self, ptr):
        self.freed.append(int(ptr))

    def undoc_free_document(self, _handle):
        return None

    def undoc_to_markdown(self, _handle, _flags):
        return self._alloc("Привет из Markdown")

    def undoc_plain_text(self, _handle):
        return self._alloc("")

    def undoc_get_title(self, _handle):
        return self._alloc("Заголовок")

    def undoc_get_author(self, _handle):
        return self._alloc("Автор")

    def undoc_get_resource_ids(self, _handle):
        return self._alloc('["rId1"]')

    def undoc_get_resource_info(self, _handle, _resource_id):
        return self._alloc('{"filename":"Пример.png"}')


class TestVersion:
    def test_version_returns_string(self):
        v = version()
        assert isinstance(v, str)
        assert len(v) > 0

    def test_version_format(self):
        v = version()
        # Should be semver-like
        parts = v.split(".")
        assert len(parts) >= 2


class TestParseFile:
    def test_parse_nonexistent_file(self):
        with pytest.raises(FileNotFoundError):
            parse_file("nonexistent.docx")

    @pytest.mark.skipif(
        not (TEST_FILES_DIR / "file-sample_1MB.docx").exists(),
        reason="Test file not available",
    )
    def test_parse_docx(self):
        doc = parse_file(TEST_FILES_DIR / "file-sample_1MB.docx")
        assert doc is not None
        assert doc.section_count >= 0

    @pytest.mark.skipif(
        not (TEST_FILES_DIR / "sample-xlsx-file.xlsx").exists(),
        reason="Test file not available",
    )
    def test_parse_xlsx(self):
        doc = parse_file(TEST_FILES_DIR / "sample-xlsx-file.xlsx")
        assert doc is not None
        assert doc.section_count >= 0

    @pytest.mark.skipif(
        not (TEST_FILES_DIR / "file_example_PPT_1MB.pptx").exists(),
        reason="Test file not available",
    )
    def test_parse_pptx(self):
        doc = parse_file(TEST_FILES_DIR / "file_example_PPT_1MB.pptx")
        assert doc is not None
        assert doc.section_count >= 0


class TestConversion:
    @pytest.fixture
    def sample_docx(self):
        path = TEST_FILES_DIR / "file-sample_1MB.docx"
        if not path.exists():
            pytest.skip("Test file not available")
        return parse_file(path)

    def test_to_markdown(self, sample_docx):
        md = sample_docx.to_markdown()
        assert isinstance(md, str)
        assert len(md) > 0

    def test_to_markdown_with_frontmatter(self, sample_docx):
        md = sample_docx.to_markdown(frontmatter=True)
        assert "---" in md

    def test_to_text(self, sample_docx):
        text = sample_docx.to_text()
        assert isinstance(text, str)
        assert len(text) > 0

    def test_to_json(self, sample_docx):
        json_str = sample_docx.to_json()
        assert isinstance(json_str, str)
        assert json_str.startswith("{")

    def test_to_json_compact(self, sample_docx):
        json_str = sample_docx.to_json(compact=True)
        assert isinstance(json_str, str)
        # Compact JSON has no indentation
        assert "\n  " not in json_str

    def test_plain_text(self, sample_docx):
        text = sample_docx.plain_text()
        assert isinstance(text, str)


class TestMetadata:
    @pytest.fixture
    def sample_docx(self):
        path = TEST_FILES_DIR / "file-sample_1MB.docx"
        if not path.exists():
            pytest.skip("Test file not available")
        return parse_file(path)

    def test_section_count(self, sample_docx):
        assert isinstance(sample_docx.section_count, int)
        assert sample_docx.section_count >= 0

    def test_resource_count(self, sample_docx):
        assert isinstance(sample_docx.resource_count, int)
        assert sample_docx.resource_count >= 0

    def test_title(self, sample_docx):
        title = sample_docx.title
        # Title may be None or string
        assert title is None or isinstance(title, str)

    def test_author(self, sample_docx):
        author = sample_docx.author
        # Author may be None or string
        assert author is None or isinstance(author, str)


class TestContextManager:
    @pytest.mark.skipif(
        not (TEST_FILES_DIR / "file-sample_1MB.docx").exists(),
        reason="Test file not available",
    )
    def test_context_manager(self):
        with parse_file(TEST_FILES_DIR / "file-sample_1MB.docx") as doc:
            md = doc.to_markdown()
            assert len(md) > 0
        # After exiting, the document should be freed
        # (we can't easily test this, but at least it shouldn't crash)


class TestParseBytes:
    @pytest.mark.skipif(
        not (TEST_FILES_DIR / "file-sample_1MB.docx").exists(),
        reason="Test file not available",
    )
    def test_parse_bytes(self):
        path = TEST_FILES_DIR / "file-sample_1MB.docx"
        with open(path, "rb") as f:
            data = f.read()

        doc = parse_bytes(data)
        assert doc is not None

        md = doc.to_markdown()
        assert len(md) > 0


class TestResources:
    @pytest.fixture
    def docx_with_images(self):
        # Try to find a document with images
        for name in ["file-sample_1MB.docx", "sample-docx-file.docx"]:
            path = TEST_FILES_DIR / name
            if path.exists():
                doc = parse_file(path)
                if doc.resource_count > 0:
                    return doc
        pytest.skip("No test file with resources available")

    def test_get_resource_ids(self, docx_with_images):
        ids = docx_with_images.get_resource_ids()
        assert isinstance(ids, list)
        assert len(ids) > 0

    def test_get_resource_info(self, docx_with_images):
        ids = docx_with_images.get_resource_ids()
        if ids:
            info = docx_with_images.get_resource_info(ids[0])
            assert info is not None
            assert "filename" in info

    def test_get_resource_data(self, docx_with_images):
        ids = docx_with_images.get_resource_ids()
        if ids:
            data = docx_with_images.get_resource_data(ids[0])
            assert data is not None
            assert len(data) > 0

    def test_get_nonexistent_resource(self, docx_with_images):
        info = docx_with_images.get_resource_info("nonexistent_id")
        assert info is None

        data = docx_with_images.get_resource_data("nonexistent_id")
        assert data is None


class TestFfiOwnershipAndUtf8:
    def test_rust_owned_strings_are_copied_and_freed(self, monkeypatch):
        fake_lib = FakeStringLibrary()
        monkeypatch.setattr(undoc_module, "get_library", lambda: fake_lib)

        doc = undoc_module.Undoc(123)
        markdown = doc.to_markdown()
        expected_ptr = ctypes.addressof(fake_lib._buffers[-1])

        assert markdown == "Привет из Markdown"
        assert fake_lib.freed == [expected_ptr]

    def test_last_error_uses_utf8_without_free(self, monkeypatch):
        fake_lib = FakeStringLibrary()
        monkeypatch.setattr(undoc_module, "get_library", lambda: fake_lib)

        assert undoc_module._get_last_error() == "Ошибка native"
        assert fake_lib.freed == []

    def test_version_uses_utf8_without_free(self, monkeypatch):
        fake_lib = FakeStringLibrary()
        monkeypatch.setattr(undoc_module, "get_library", lambda: fake_lib)

        assert undoc_module.version() == "1.2.3"
        assert fake_lib.freed == []

    def test_metadata_and_resource_json_are_copied_before_free(self, monkeypatch):
        fake_lib = FakeStringLibrary()
        monkeypatch.setattr(undoc_module, "get_library", lambda: fake_lib)

        doc = undoc_module.Undoc(123)
        title = doc.title
        author = doc.author
        resource_ids = doc.get_resource_ids()
        info = doc.get_resource_info("rId1")

        expected_freed = [ctypes.addressof(buf) for buf in fake_lib._buffers]

        assert title == "Заголовок"
        assert author == "Автор"
        assert resource_ids == ["rId1"]
        assert info == {"filename": "Пример.png"}
        assert fake_lib.freed == expected_freed

    def test_plain_text_preserves_valid_empty_string(self, monkeypatch):
        fake_lib = FakeStringLibrary()
        monkeypatch.setattr(undoc_module, "get_library", lambda: fake_lib)

        doc = undoc_module.Undoc(123)
        text = doc.plain_text()

        assert text == ""
        assert fake_lib.freed == [ctypes.addressof(fake_lib._buffers[-1])]

    def test_get_resource_ids_preserves_valid_empty_list(self, monkeypatch):
        fake_lib = FakeStringLibrary()
        fake_lib.undoc_get_resource_ids = lambda _handle: fake_lib._alloc("[]")
        monkeypatch.setattr(undoc_module, "get_library", lambda: fake_lib)

        doc = undoc_module.Undoc(123)
        resource_ids = doc.get_resource_ids()

        assert resource_ids == []
        assert fake_lib.freed == [ctypes.addressof(fake_lib._buffers[-1])]

    def test_get_resource_ids_raises_on_native_null(self, monkeypatch):
        fake_lib = FakeStringLibrary()
        fake_lib.undoc_get_resource_ids = lambda _handle: 0
        monkeypatch.setattr(undoc_module, "get_library", lambda: fake_lib)

        doc = undoc_module.Undoc(123)

        with pytest.raises(
            undoc_module.UndocError, match="Failed to get resource IDs: Ошибка native"
        ):
            doc.get_resource_ids()

    def test_parse_bytes_generated_docx_preserves_unicode(self):
        doc = parse_bytes(create_minimal_docx_bytes())
        text = doc.to_text()

        assert "Привет из Python" in text

    def test_parse_file_generated_docx_preserves_unicode(self, tmp_path):
        path = tmp_path / "unicode.docx"
        path.write_bytes(create_minimal_docx_bytes("Привет из файла"))

        with parse_file(path) as doc:
            markdown = doc.to_markdown()

        assert "Привет из файла" in markdown
