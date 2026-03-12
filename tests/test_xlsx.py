import pytest
from mcp.types import TextContent
from readers.xlsx import read_xlsx


def test_read_xlsx_returns_sheet_content(simple_xlsx):
    result = read_xlsx(simple_xlsx)
    assert len(result) == 1
    assert isinstance(result[0], TextContent)
    assert "Sheet1" in result[0].text
    assert "Alice" in result[0].text
    assert "90" in result[0].text


def test_read_xlsx_multi_sheet(xlsx_multi_sheet):
    result = read_xlsx(xlsx_multi_sheet)
    assert len(result) == 2
    sheet_labels = [r.text.split("\n")[0] for r in result]
    assert any("Alpha" in label for label in sheet_labels)
    assert any("Beta" in label for label in sheet_labels)


def test_read_xlsx_file_not_found():
    with pytest.raises(ValueError, match="not found"):
        read_xlsx("/nonexistent/path/file.xlsx")
