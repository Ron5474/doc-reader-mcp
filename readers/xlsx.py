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
    with xl:
        for sheet_name in xl.sheet_names:
            df = xl.parse(sheet_name)
            md = df.to_markdown(index=False)
            if md:
                output.append(TextContent(type="text", text=f"--- Sheet: {sheet_name} ---\n{md}"))

    return output
