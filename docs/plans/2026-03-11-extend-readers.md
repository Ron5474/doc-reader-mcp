# Extend Readers Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Refactor slides-reader into a `readers/` module structure, add DOCX and XLSX support, and enrich PPTX extraction with speaker notes.

**Architecture:** Move each format's logic into `readers/<format>.py`, keeping `main.py` as a thin MCP wiring layer. Add `read_docx` and `read_xlsx` MCP tools. Enrich `read_pptx` with speaker notes appended to each slide's `TextContent`.

**Tech Stack:** Python 3.12, FastMCP, pymupdf, python-pptx, python-docx, openpyxl, pandas, pytest

---

### Task 1: Add dependencies and set up test infrastructure

**Files:**
- Modify: `pyproject.toml`
- Create: `tests/__init__.py`
- Create: `tests/conftest.py`

**Step 1: Add new dependencies**

Edit `pyproject.toml` dependencies list to add:
```toml
dependencies = [
    "mcp[cli]>=1.26.0",
    "pandas>=3.0.1",
    "pymupdf>=1.27.1",
    "python-pptx>=1.0.2",
    "tabulate>=0.10.0",
    "python-docx>=1.1.0",
    "openpyxl>=3.1.0",
    "pytest>=8.0.0",
]
```

**Step 2: Sync dependencies**

Run: `uv sync`
Expected: resolves and installs python-docx, openpyxl, pytest

**Step 3: Create test package**

Create `tests/__init__.py` — empty file.

**Step 4: Create conftest with fixture helpers**

Create `tests/conftest.py`:
```python
import pytest
import io
from docx import Document
from pptx import Presentation
from pptx.util import Inches
import openpyxl
import fitz


@pytest.fixture
def simple_docx(tmp_path):
    doc = Document()
    doc.add_paragraph("Hello from DOCX")
    doc.add_paragraph("Second paragraph")
    path = tmp_path / "test.docx"
    doc.save(str(path))
    return str(path)


@pytest.fixture
def docx_with_table(tmp_path):
    doc = Document()
    doc.add_paragraph("Before table")
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "A"
    table.cell(0, 1).text = "B"
    table.cell(1, 0).text = "1"
    table.cell(1, 1).text = "2"
    path = tmp_path / "table.docx"
    doc.save(str(path))
    return str(path)


@pytest.fixture
def simple_xlsx(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Name", "Score"])
    ws.append(["Alice", 90])
    ws.append(["Bob", 85])
    path = tmp_path / "test.xlsx"
    wb.save(str(path))
    return str(path)


@pytest.fixture
def xlsx_multi_sheet(tmp_path):
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Alpha"
    ws1.append(["X", "Y"])
    ws1.append([1, 2])
    ws2 = wb.create_sheet("Beta")
    ws2.append(["P", "Q"])
    ws2.append([3, 4])
    path = tmp_path / "multi.xlsx"
    wb.save(str(path))
    return str(path)


@pytest.fixture
def pptx_with_notes(tmp_path):
    prs = Presentation()
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = "Slide Title"
    slide.placeholders[1].text = "Slide body text"
    notes_slide = slide.notes_slide
    notes_slide.notes_text_frame.text = "These are speaker notes"
    path = tmp_path / "notes.pptx"
    prs.save(str(path))
    return str(path)
```

**Step 5: Verify pytest runs**

Run: `uv run pytest tests/ -v`
Expected: "no tests ran" with exit 0 (or exit 5 for "no tests collected" — both are fine)

**Step 6: Commit**

```bash
git add pyproject.toml uv.lock tests/
git commit -m "chore: add docx/xlsx deps and test infrastructure"
```

---

### Task 2: Create readers/ package and migrate pdf.py

**Files:**
- Create: `readers/__init__.py`
- Create: `readers/pdf.py`
- Create: `tests/test_pdf.py`

**Step 1: Write the failing test**

Create `tests/test_pdf.py`:
```python
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
    import pytest
    with pytest.raises(ValueError, match="not found"):
        read_pdf("/nonexistent/path/file.pdf")
```

**Step 2: Run to verify it fails**

Run: `uv run pytest tests/test_pdf.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'readers'`

**Step 3: Create readers package**

Create `readers/__init__.py` — empty file.

**Step 4: Create readers/pdf.py**

Move logic from `main.py` into `readers/pdf.py`:
```python
from mcp.types import TextContent, ImageContent
import fitz
import base64
import os


def read_pdf(path: str) -> list:
    """Extract text, tables, and images from a PDF file, page by page."""
    if not os.path.exists(path):
        raise ValueError(f"File not found: {path}")

    try:
        doc = fitz.open(path)
    except Exception as e:
        raise ValueError(f"Could not open PDF: {e}") from e

    output = []
    for i, page in enumerate(doc, start=1):
        text = page.get_text().strip()

        tables = page.find_tables()
        if tables.tables:
            for table in tables.tables:
                df = table.to_pandas()
                markdown_table = df.to_markdown(index=False)
                for col in df.columns:
                    for val in df[col].astype(str):
                        text = text.replace(val, "", 1)
                text = text.strip() + f"\n\n{markdown_table}"

        if text:
            output.append(TextContent(type="text", text=f"--- Page {i} ---\n{text}"))

        for img in page.get_images():
            xref = img[0]
            base_img = doc.extract_image(xref)
            img_bytes = base_img["image"]
            ext = base_img["ext"]
            b64 = base64.standard_b64encode(img_bytes).decode("utf-8")
            output.append(ImageContent(type="image", data=b64, mimeType=f"image/{ext}"))

    doc.close()
    return output
```

**Step 5: Run tests to verify they pass**

Run: `uv run pytest tests/test_pdf.py -v`
Expected: 2 PASSED

**Step 6: Commit**

```bash
git add readers/ tests/test_pdf.py
git commit -m "feat: extract pdf reader into readers/pdf.py"
```

---

### Task 3: Migrate and enhance pptx.py with speaker notes

**Files:**
- Create: `readers/pptx.py`
- Create: `tests/test_pptx.py`

**Step 1: Write the failing tests**

Create `tests/test_pptx.py`:
```python
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
```

**Step 2: Run to verify they fail**

Run: `uv run pytest tests/test_pptx.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'readers.pptx'`

**Step 3: Create readers/pptx.py**

```python
from mcp.types import TextContent, ImageContent
from pptx import Presentation
import base64
import os


def read_pptx(path: str) -> list:
    """Extract text, images, and speaker notes from a PowerPoint file, slide by slide."""
    if not os.path.exists(path):
        raise ValueError(f"File not found: {path}")

    try:
        prs = Presentation(path)
    except Exception as e:
        raise ValueError(f"Could not open PPTX: {e}") from e

    output = []
    for i, slide in enumerate(prs.slides, start=1):
        texts = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    line = para.text.strip()
                    if line:
                        texts.append(line)

        notes = ""
        if slide.has_notes_slide:
            notes_text = slide.notes_slide.notes_text_frame.text.strip()
            if notes_text:
                notes = f"\n\n--- Notes ---\n{notes_text}"

        if texts or notes:
            content = f"--- Slide {i} ---\n" + "\n".join(texts) + notes
            output.append(TextContent(type="text", text=content))

        for shape in slide.shapes:
            if shape.shape_type == 13:
                img = shape.image
                b64 = base64.standard_b64encode(img.blob).decode("utf-8")
                mime = img.content_type
                output.append(ImageContent(type="image", data=b64, mimeType=mime))

    return output
```

**Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/test_pptx.py -v`
Expected: 3 PASSED

**Step 5: Commit**

```bash
git add readers/pptx.py tests/test_pptx.py
git commit -m "feat: extract pptx reader with speaker notes support"
```

---

### Task 4: Create docx.py reader

**Files:**
- Create: `readers/docx.py`
- Create: `tests/test_docx.py`

**Step 1: Write the failing tests**

Create `tests/test_docx.py`:
```python
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
```

**Step 2: Run to verify they fail**

Run: `uv run pytest tests/test_docx.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'readers.docx'`

**Step 3: Create readers/docx.py**

```python
from mcp.types import TextContent, ImageContent
from docx import Document
from docx.oxml.ns import qn
import pandas as pd
import base64
import os


def _table_to_markdown(table) -> str:
    data = [[cell.text for cell in row.cells] for row in table.rows]
    if not data:
        return ""
    df = pd.DataFrame(data[1:], columns=data[0])
    return df.to_markdown(index=False)


def read_docx(path: str) -> list:
    """Extract text, tables, and images from a Word document."""
    if not os.path.exists(path):
        raise ValueError(f"File not found: {path}")

    try:
        doc = Document(path)
    except Exception as e:
        raise ValueError(f"Could not open DOCX: {e}") from e

    output = []
    texts = []

    for block in doc.element.body:
        tag = block.tag.split("}")[-1]

        if tag == "p":
            # Find matching paragraph object
            for para in doc.paragraphs:
                if para._element is block:
                    line = para.text.strip()
                    if line:
                        texts.append(line)
                    break

        elif tag == "tbl":
            # Find matching table object
            for table in doc.tables:
                if table._element is block:
                    md = _table_to_markdown(table)
                    if md:
                        texts.append(md)
                    break

    if texts:
        output.append(TextContent(type="text", text="\n".join(texts)))

    # Extract inline images
    for rel in doc.part.rels.values():
        if "image" in rel.reltype:
            img_part = rel.target_part
            b64 = base64.standard_b64encode(img_part.blob).decode("utf-8")
            mime = img_part.content_type
            output.append(ImageContent(type="image", data=b64, mimeType=mime))

    return output
```

**Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/test_docx.py -v`
Expected: 3 PASSED

**Step 5: Commit**

```bash
git add readers/docx.py tests/test_docx.py
git commit -m "feat: add docx reader with table support"
```

---

### Task 5: Create xlsx.py reader

**Files:**
- Create: `readers/xlsx.py`
- Create: `tests/test_xlsx.py`

**Step 1: Write the failing tests**

Create `tests/test_xlsx.py`:
```python
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
```

**Step 2: Run to verify they fail**

Run: `uv run pytest tests/test_xlsx.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'readers.xlsx'`

**Step 3: Create readers/xlsx.py**

```python
from mcp.types import TextContent
import pandas as pd
import os


def read_xlsx(path: str) -> list:
    """Extract data from an Excel file, one TextContent per sheet rendered as markdown."""
    if not os.path.exists(path):
        raise ValueError(f"File not found: {path}")

    try:
        xl = pd.ExcelFile(path, engine="openpyxl")
    except Exception as e:
        raise ValueError(f"Could not open XLSX: {e}") from e

    output = []
    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name)
        md = df.to_markdown(index=False)
        output.append(TextContent(type="text", text=f"--- Sheet: {sheet_name} ---\n{md}"))

    return output
```

**Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/test_xlsx.py -v`
Expected: 3 PASSED

**Step 5: Commit**

```bash
git add readers/xlsx.py tests/test_xlsx.py
git commit -m "feat: add xlsx reader with per-sheet markdown output"
```

---

### Task 6: Rewrite main.py as thin MCP wiring layer

**Files:**
- Modify: `main.py`

**Step 1: Run full test suite first to confirm green baseline**

Run: `uv run pytest tests/ -v`
Expected: all tests PASS

**Step 2: Rewrite main.py**

```python
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
```

**Step 3: Run full test suite again**

Run: `uv run pytest tests/ -v`
Expected: all tests still PASS

**Step 4: Commit**

```bash
git add main.py
git commit -m "refactor: thin main.py delegates to readers/ modules"
```

---

### Task 7: Update README

**Files:**
- Modify: `README.md`

**Step 1: Update the Tools section**

Replace the Tools section in README.md with:

```markdown
## Tools

- `read_pdf` — Extracts text, tables (as Markdown), and images from a PDF file, page by page.
- `read_pptx` — Extracts text, images, and speaker notes from a PowerPoint file, slide by slide.
- `read_docx` — Extracts text, tables (as Markdown), and images from a Word document.
- `read_xlsx` — Extracts data from an Excel file as Markdown tables, one section per sheet.
```

**Step 2: Commit**

```bash
git add README.md
git commit -m "docs: update README with new tools"
```
