# Design: Extend slides-reader with DOCX, XLSX, and Richer Extraction

Date: 2026-03-11

## Overview

Extend the slides-reader MCP server with two new file format readers (DOCX, XLSX) and richer extraction from the existing PPTX reader (speaker notes). Refactor the codebase into a `readers/` module structure for maintainability.

## Architecture

```
slides-reader/
├── main.py              # MCP tool registrations only
└── readers/
    ├── __init__.py
    ├── pdf.py           # existing read_pdf logic
    ├── pptx.py          # existing read_pptx logic
    ├── docx.py          # new
    └── xlsx.py          # new
```

`main.py` becomes thin — imports readers and registers `@mcp.tool()` decorators only.

## Extraction Scope

| Format | Text | Tables | Images | Extras |
|--------|------|--------|--------|--------|
| PDF | per page | markdown tables | embedded images | — |
| PPTX | per slide | — | embedded images | speaker notes |
| DOCX | per paragraph/section | markdown tables | embedded images | — |
| XLSX | — | per sheet as markdown | — | sheet names |

### PPTX speaker notes
Speaker notes appear after slide text as a `--- Notes ---` block within the same `TextContent` item.

### XLSX output
One `TextContent` per sheet, labeled `--- Sheet: <name> ---`, table rendered as markdown.

### DOCX tables
Same approach as PDF: pandas DataFrame → `to_markdown()`.

## Error Handling

Each reader raises a `ValueError` with a clear message for:
- File not found
- Wrong/unexpected file extension
- Corrupt or unreadable file (library exceptions wrapped)

FastMCP surfaces these as error responses to the client. No silent failures, no partial results on corrupt files.

## Dependencies to Add

- `python-docx` — for DOCX reading
- `openpyxl` — for XLSX reading (pandas dependency for `.read_excel()`)
