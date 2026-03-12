from mcp.server.fastmcp import FastMCP
from readers.pdf import read_pdf as _read_pdf
from readers.pptx import read_pptx as _read_pptx
from readers.docx import read_docx as _read_docx
from readers.xlsx import read_xlsx as _read_xlsx

mcp = FastMCP("slides-reader")


@mcp.tool()
def read_pdf(path: str) -> list:
    """Extract text, tables, and images from a PDF file, page by page."""
    return _read_pdf(path)


@mcp.tool()
def read_pptx(path: str) -> list:
    """Extract text, images, and speaker notes from a PowerPoint file, slide by slide."""
    return _read_pptx(path)


@mcp.tool()
def read_docx(path: str) -> list:
    """Extract text, tables, and images from a Word document."""
    return _read_docx(path)


@mcp.tool()
def read_xlsx(path: str) -> list:
    """Extract data from an Excel file as markdown tables, one per sheet."""
    return _read_xlsx(path)


if __name__ == "__main__":
    mcp.run()
