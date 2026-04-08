"""Extract text and table data from order PDFs using pdfplumber."""

from __future__ import annotations

import base64
from pathlib import Path

import pdfplumber


def extract_text(pdf_path: str | Path) -> str:
    """Extract all text from a PDF, page by page."""
    pdf_path = Path(pdf_path)
    pages_text: list[str] = []

    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, 1):
            text = page.extract_text() or ""
            if text.strip():
                pages_text.append(f"--- Page {i} ---\n{text}")

    return "\n\n".join(pages_text)


def extract_tables(pdf_path: str | Path) -> list[list[list[str | None]]]:
    """Extract tables from a PDF. Returns list of tables, each table is list of rows."""
    pdf_path = Path(pdf_path)
    all_tables: list[list[list[str | None]]] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if tables:
                all_tables.extend(tables)

    return all_tables


def pdf_to_base64_images(pdf_path: str | Path, max_pages: int = 5) -> list[str]:
    """Convert PDF pages to base64-encoded PNG images for vision API.

    Used for PDFs that are image-heavy (e.g. Konfigurationsblatt).
    """
    pdf_path = Path(pdf_path)
    images: list[str] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages[:max_pages]:
            img = page.to_image(resolution=200)
            import io

            buf = io.BytesIO()
            img.original.save(buf, format="PNG")
            b64 = base64.b64encode(buf.getvalue()).decode("utf-8")
            images.append(b64)

    return images


def get_pdf_info(pdf_path: str | Path) -> dict:
    """Get basic info about a PDF."""
    pdf_path = Path(pdf_path)
    with pdfplumber.open(pdf_path) as pdf:
        return {
            "path": str(pdf_path),
            "filename": pdf_path.name,
            "num_pages": len(pdf.pages),
            "file_size_kb": round(pdf_path.stat().st_size / 1024, 1),
        }
