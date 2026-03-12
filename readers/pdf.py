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
    try:
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
    finally:
        doc.close()
    return output
