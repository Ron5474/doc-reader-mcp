import pytest
from mcp.types import TextContent, ImageContent
from readers.pdf import read_pdf


def test_read_pdf_returns_text(tmp_path):
    import fitz
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((50, 100), "Hello PDF World")
    path = tmp_path / "test.pdf"
    doc.save(str(path))
    doc.close()

    result = read_pdf(str(path))

    assert len(result) >= 1
    texts = [r for r in result if isinstance(r, TextContent)]
    assert any("Hello PDF World" in t.text for t in texts)


def test_read_pdf_file_not_found():
    with pytest.raises(ValueError, match="not found"):
        read_pdf("/nonexistent/path/file.pdf")
