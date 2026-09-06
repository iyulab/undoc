"""Tests for undoc Python bindings."""

import ctypes
import io
import json
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


# Documents are assembled in-process rather than committed as binaries, the same way the
# Rust suite builds its OOXML packages. A generated fixture is always present, so a test
# that needs one can fail loudly instead of quietly skipping itself.
SAMPLE_TEXT = "Paragraph the binding must carry"


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


# A 1x1 opaque PNG. Small enough to inline, real enough to be read back as resource
# bytes — which is what the resource tests need a document to carry.
_ONE_PIXEL_PNG = (
    b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01"
    b"\x08\x02\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDAT\x08\xd7c\xf8\xcf"
    b"\xc0\x00\x00\x03\x01\x01\x00\x18\xdd\x8d\xb0\x00\x00\x00\x00IEND\xaeB`\x82"
)


def create_docx_with_image_bytes(text: str = "Document with a picture") -> bytes:
    """A DOCX carrying one embedded image, for the resource-inventory tests."""
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
  <Default Extension="png" ContentType="image/png"/>
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
  <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>""",
        )
        zf.writestr("word/document.xml", document_xml)
        zf.writestr("word/media/image1.png", _ONE_PIXEL_PNG)
    return buf.getvalue()


def create_minimal_xlsx_bytes(text: str = "Spreadsheet cell") -> bytes:
    """Create a tiny XLSX fixture without relying on external test files."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_STORED) as zf:
        zf.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>""",
        )
        zf.writestr(
            "_rels/.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""",
        )
        zf.writestr(
            "xl/workbook.xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>""",
        )
        zf.writestr(
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>""",
        )
        zf.writestr(
            "xl/worksheets/sheet1.xml",
            f"""<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>{text}</t></is></c></row>
  </sheetData>
</worksheet>""",
        )
    return buf.getvalue()


def create_minimal_pptx_bytes(text: str = "Slide text") -> bytes:
    """Create a tiny PPTX fixture without relying on external test files."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_STORED) as zf:
        zf.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>""",
        )
        zf.writestr(
            "_rels/.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>""",
        )
        zf.writestr(
            "ppt/presentation.xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
</p:presentation>""",
        )
        zf.writestr(
            "ppt/_rels/presentation.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>""",
        )
        zf.writestr(
            "ppt/slides/slide1.xml",
            f"""<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:sp><p:txBody>
      <a:bodyPr/>
      <a:p><a:r><a:t>{text}</a:t></a:r></a:p>
    </p:txBody></p:sp>
  </p:spTree></p:cSld>
</p:sld>""",
        )
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
        # What the native layer would report as the failure classification. Tests set
        # this to stand in for a specific reason, including numbers this build has no
        # name for.
        self.last_error_kind = 5  # ZIP_ARCHIVE

    def _alloc(self, text: str) -> int:
        buf = ctypes.create_string_buffer(text.encode("utf-8"))
        self._buffers.append(buf)
        return ctypes.addressof(buf)

    def undoc_last_error(self):
        return self._alloc("Ошибка native")

    def undoc_last_error_kind(self):
        return self.last_error_kind

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

    def test_parse_docx(self, tmp_path):
        path = tmp_path / "sample.docx"
        path.write_bytes(create_minimal_docx_bytes("Word content"))

        doc = parse_file(path)
        assert doc.section_count == 1
        assert "Word content" in doc.to_markdown()

    def test_parse_xlsx(self, tmp_path):
        path = tmp_path / "sample.xlsx"
        path.write_bytes(create_minimal_xlsx_bytes("Spreadsheet content"))

        doc = parse_file(path)
        assert doc.section_count == 1
        assert "Spreadsheet content" in doc.to_markdown()

    def test_parse_pptx(self, tmp_path):
        path = tmp_path / "sample.pptx"
        path.write_bytes(create_minimal_pptx_bytes("Presentation content"))

        doc = parse_file(path)
        assert doc.section_count == 1
        assert "Presentation content" in doc.to_markdown()


class TestConversion:
    @pytest.fixture
    def sample_docx(self):
        return parse_bytes(create_minimal_docx_bytes(SAMPLE_TEXT))

    def test_to_markdown(self, sample_docx):
        assert SAMPLE_TEXT in sample_docx.to_markdown()

    def test_to_markdown_with_frontmatter(self, sample_docx):
        md = sample_docx.to_markdown(frontmatter=True)
        assert "---" in md
        assert SAMPLE_TEXT in md

    def test_to_markdown_with_refine(self, sample_docx):
        # Refining reshapes markdown; it must not drop the document's text.
        assert SAMPLE_TEXT in sample_docx.to_markdown(refine=True)

    def test_to_text(self, sample_docx):
        assert SAMPLE_TEXT in sample_docx.to_text()

    def test_to_json(self, sample_docx):
        # `startswith("{")` alone would pass on a truncated payload, so decode it.
        payload = json.loads(sample_docx.to_json())

        assert payload["format"] == "docx"
        assert SAMPLE_TEXT in json.dumps(payload, ensure_ascii=False)

    def test_to_json_compact(self, sample_docx):
        json_str = sample_docx.to_json(compact=True)
        assert isinstance(json_str, str)
        # Compact JSON has no indentation
        assert "\n  " not in json_str

    def test_plain_text(self, sample_docx):
        assert SAMPLE_TEXT in sample_docx.plain_text()


class TestMetadata:
    @pytest.fixture
    def sample_docx(self):
        return parse_bytes(create_minimal_docx_bytes(SAMPLE_TEXT))

    def test_section_count(self, sample_docx):
        assert sample_docx.section_count == 1

    def test_resource_count(self, sample_docx):
        # This package carries no media part; the resource tests cover the other case.
        assert sample_docx.resource_count == 0

    def test_title(self, sample_docx):
        # The package has no core properties, so an absent title must surface as None
        # rather than an empty string or a raised error.
        assert sample_docx.title is None

    def test_author(self, sample_docx):
        assert sample_docx.author is None


class TestContextManager:
    def test_context_manager(self, tmp_path):
        path = tmp_path / "sample.docx"
        path.write_bytes(create_minimal_docx_bytes(SAMPLE_TEXT))

        with parse_file(path) as doc:
            assert SAMPLE_TEXT in doc.to_markdown()
        # After exiting, the document should be freed
        # (we can't easily test this, but at least it shouldn't crash)


class TestParseBytes:
    def test_parse_bytes(self):
        doc = parse_bytes(create_minimal_docx_bytes(SAMPLE_TEXT))

        assert SAMPLE_TEXT in doc.to_markdown()


class TestResources:
    @pytest.fixture
    def docx_with_images(self):
        doc = parse_bytes(create_docx_with_image_bytes())
        assert doc.resource_count > 0, "the generated document must carry its image"
        return doc

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


class TestErrorKind:
    """The classification channel: a caller must be able to branch on the reason."""

    def test_kind_travels_with_the_error(self, monkeypatch):
        fake_lib = FakeStringLibrary()
        fake_lib.last_error_kind = int(undoc_module.ErrorKind.ZIP_ARCHIVE)
        fake_lib.undoc_get_resource_ids = lambda _handle: 0
        monkeypatch.setattr(undoc_module, "get_library", lambda: fake_lib)

        doc = undoc_module.Undoc(123)

        with pytest.raises(undoc_module.UndocError) as excinfo:
            doc.get_resource_ids()

        assert excinfo.value.kind is undoc_module.ErrorKind.ZIP_ARCHIVE

    def test_unknown_kind_value_is_preserved_not_rejected(self, monkeypatch):
        """Forward compatibility: a newer native library may report a reason this
        build has no name for. It must arrive as a plain int, not raise ValueError."""
        fake_lib = FakeStringLibrary()
        fake_lib.last_error_kind = 9999
        fake_lib.undoc_get_resource_ids = lambda _handle: 0
        monkeypatch.setattr(undoc_module, "get_library", lambda: fake_lib)

        doc = undoc_module.Undoc(123)

        with pytest.raises(undoc_module.UndocError) as excinfo:
            doc.get_resource_ids()

        assert excinfo.value.kind == 9999
        assert not isinstance(excinfo.value.kind, undoc_module.ErrorKind)

    def test_unclassified_failure_is_other_never_none(self, monkeypatch):
        """A failure the native layer left unclassified must not read as success."""
        fake_lib = FakeStringLibrary()
        fake_lib.last_error_kind = 0
        fake_lib.undoc_get_resource_ids = lambda _handle: 0
        monkeypatch.setattr(undoc_module, "get_library", lambda: fake_lib)

        doc = undoc_module.Undoc(123)

        with pytest.raises(undoc_module.UndocError) as excinfo:
            doc.get_resource_ids()

        assert excinfo.value.kind is undoc_module.ErrorKind.OTHER
        assert excinfo.value.kind != undoc_module.ErrorKind.NONE

    def test_error_raised_without_a_kind_defaults_to_other(self):
        assert undoc_module.UndocError("wrapper-side").kind is undoc_module.ErrorKind.OTHER

    def test_corrupted_archive_reports_zip_archive(self):
        """End to end against the real library: a damaged container is recognisable
        from the error alone, without reading its message."""
        corrupted = b"PK\x03\x04" + b"truncated garbage with no central directory"

        with pytest.raises(UndocError) as excinfo:
            parse_bytes(corrupted)

        assert excinfo.value.kind is undoc_module.ErrorKind.ZIP_ARCHIVE

    def test_non_office_input_reports_unknown_format(self):
        with pytest.raises(UndocError) as excinfo:
            parse_bytes(b"not an office document at all")

        assert excinfo.value.kind is undoc_module.ErrorKind.UNKNOWN_FORMAT

    def test_distinct_failures_report_distinct_kinds(self):
        with pytest.raises(UndocError) as damaged:
            parse_bytes(b"PK\x03\x04garbage")
        with pytest.raises(UndocError) as foreign:
            parse_bytes(b"plain text file")

        assert damaged.value.kind != foreign.value.kind

    def test_discriminants_match_the_native_abi(self):
        kinds = undoc_module.ErrorKind
        assert (kinds.NONE, kinds.OTHER, kinds.IO) == (0, 1, 2)
        assert (kinds.UNKNOWN_FORMAT, kinds.UNSUPPORTED_FORMAT) == (3, 4)
        assert (kinds.ZIP_ARCHIVE, kinds.XML_PARSE, kinds.INVALID_DATA) == (5, 6, 7)
        assert (kinds.MISSING_COMPONENT, kinds.ENCODING) == (8, 9)
        assert (kinds.STYLE_NOT_FOUND, kinds.RESOURCE_NOT_FOUND) == (10, 11)
        assert (kinds.ENCRYPTED, kinds.RENDER) == (12, 13)
        assert (kinds.INVALID_ARGUMENT, kinds.PANIC, kinds.INVALID_OUTPUT) == (100, 101, 102)
