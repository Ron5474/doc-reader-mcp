# slides-reader

An MCP (Model Context Protocol) server that extracts text and images from PDF and PowerPoint files, making them accessible to Claude and other MCP-compatible AI clients.

## Tools

- `read_pdf` — Extracts text, tables (as Markdown), and images from a PDF file, page by page.
- `read_pptx` — Extracts text, images, and speaker notes from a PowerPoint file, slide by slide.
- `read_docx` — Extracts text, tables (as Markdown), and images from a Word document.
- `read_xlsx` — Extracts data from an Excel file as Markdown tables, one section per sheet.

## Requirements

- Python 3.12+
- [`uv`](https://docs.astral.sh/uv/) (recommended)

## Installation

```bash
git clone https://github.com/Ron5474/slides-reader
cd slides-reader
uv sync
```

## Running the server

```bash
uv run python main.py
```

## MCP Configuration

### Claude Code (via CLI)

The quickest way to add this MCP server to Claude Code is with the `claude mcp add` command:

```bash
claude mcp add slides-reader -- uv --directory /absolute/path/to/slides-reader run python main.py
```

Replace `/absolute/path/to/slides-reader` with the actual absolute path to this repository.

To verify it was added:

```bash
claude mcp list
```

### Claude Code (via JSON config)

Alternatively, edit `~/.claude/claude_code_config.json`:

```json
{
  "mcpServers": {
    "slides-reader": {
      "command": "uv",
      "args": [
        "--directory",
        "/absolute/path/to/slides-reader",
        "run",
        "python",
        "main.py"
      ]
    }
  }
}
```

### Claude Desktop (via JSON config)

- **Mac:** `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows:** `%APPDATA%\Claude\claude_desktop_config.json`

```json
{
  "mcpServers": {
    "slides-reader": {
      "command": "uv",
      "args": [
        "--directory",
        "/absolute/path/to/slides-reader",
        "run",
        "python",
        "main.py"
      ]
    }
  }
}
```

After editing the config, restart Claude for the changes to take effect.

## Usage

Once configured, you can ask Claude to read any PDF or PPTX file by providing its absolute path:

> "Read the PDF at /home/user/documents/report.pdf"

> "Summarize the slides at /home/user/presentations/lecture.pptx"
