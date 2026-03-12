import pytest
from mcp.types import TextContent, ImageContent
from readers.pptx import read_pptx


def test_read_pptx_returns_text(pptx_with_notes):
    result = read_pptx(pptx_with_notes)
    texts = [r for r in result if isinstance(r, TextContent)]
    assert len(texts) >= 1
    combined = "\n".join(t.text for t in texts)
    assert "Slide Title" in combined


def test_read_pptx_includes_speaker_notes(pptx_with_notes):
    result = read_pptx(pptx_with_notes)
    texts = [r for r in result if isinstance(r, TextContent)]
    combined = "\n".join(t.text for t in texts)
    assert "These are speaker notes" in combined
    assert "--- Notes ---" in combined


def test_read_pptx_file_not_found():
    with pytest.raises(ValueError, match="not found"):
        read_pptx("/nonexistent/path/file.pptx")
