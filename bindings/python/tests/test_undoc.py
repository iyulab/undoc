"""Tests for undoc Python bindings."""

import os
import pytest
from pathlib import Path

# Check if library is available
try:
    from undoc import Undoc, UndocError, parse_file, parse_bytes, version
    LIBRARY_AVAILABLE = True
except OSError:
    LIBRARY_AVAILABLE = False


# Get test files directory
TEST_FILES_DIR = Path(__file__).parent.parent.parent.parent / "test-files"


@pytest.mark.skipif(not LIBRARY_AVAILABLE, reason="Native library not available")
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


@pytest.mark.skipif(not LIBRARY_AVAILABLE, reason="Native library not available")
class TestParseFile:
    def test_parse_nonexistent_file(self):
        with pytest.raises(FileNotFoundError):
            parse_file("nonexistent.docx")

    @pytest.mark.skipif(
        not (TEST_FILES_DIR / "file-sample_1MB.docx").exists(),
        reason="Test file not available"
    )
    def test_parse_docx(self):
        doc = parse_file(TEST_FILES_DIR / "file-sample_1MB.docx")
        assert doc is not None
        assert doc.section_count >= 0

    @pytest.mark.skipif(
        not (TEST_FILES_DIR / "sample-xlsx-file.xlsx").exists(),
        reason="Test file not available"
    )
    def test_parse_xlsx(self):
        doc = parse_file(TEST_FILES_DIR / "sample-xlsx-file.xlsx")
        assert doc is not None
        assert doc.section_count >= 0

    @pytest.mark.skipif(
        not (TEST_FILES_DIR / "file_example_PPT_1MB.pptx").exists(),
        reason="Test file not available"
    )
    def test_parse_pptx(self):
        doc = parse_file(TEST_FILES_DIR / "file_example_PPT_1MB.pptx")
        assert doc is not None
        assert doc.section_count >= 0


@pytest.mark.skipif(not LIBRARY_AVAILABLE, reason="Native library not available")
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


@pytest.mark.skipif(not LIBRARY_AVAILABLE, reason="Native library not available")
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


@pytest.mark.skipif(not LIBRARY_AVAILABLE, reason="Native library not available")
class TestContextManager:
    @pytest.mark.skipif(
        not (TEST_FILES_DIR / "file-sample_1MB.docx").exists(),
        reason="Test file not available"
    )
    def test_context_manager(self):
        with parse_file(TEST_FILES_DIR / "file-sample_1MB.docx") as doc:
            md = doc.to_markdown()
            assert len(md) > 0
        # After exiting, the document should be freed
        # (we can't easily test this, but at least it shouldn't crash)


@pytest.mark.skipif(not LIBRARY_AVAILABLE, reason="Native library not available")
class TestParseBytes:
    @pytest.mark.skipif(
        not (TEST_FILES_DIR / "file-sample_1MB.docx").exists(),
        reason="Test file not available"
    )
    def test_parse_bytes(self):
        path = TEST_FILES_DIR / "file-sample_1MB.docx"
        with open(path, "rb") as f:
            data = f.read()

        doc = parse_bytes(data)
        assert doc is not None

        md = doc.to_markdown()
        assert len(md) > 0


@pytest.mark.skipif(not LIBRARY_AVAILABLE, reason="Native library not available")
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
