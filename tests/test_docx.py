import pytest
from mcp.types import TextContent, ImageContent
from readers.docx import read_docx


def test_read_docx_returns_text(simple_docx):
    result = read_docx(simple_docx)
    texts = [r for r in result if isinstance(r, TextContent)]
    assert len(texts) >= 1
    combined = "\n".join(t.text for t in texts)
    assert "Hello from DOCX" in combined


def test_read_docx_renders_table_as_markdown(docx_with_table):
    result = read_docx(docx_with_table)
    texts = [r for r in result if isinstance(r, TextContent)]
    combined = "\n".join(t.text for t in texts)
    assert "|" in combined  # markdown table uses pipe chars
    assert "A" in combined
    assert "B" in combined


def test_read_docx_file_not_found():
    with pytest.raises(ValueError, match="not found"):
        read_docx("/nonexistent/path/file.docx")
