"""Render first PDF page as PNG for fidelity report."""

from __future__ import annotations

from pathlib import Path


def render_pdf_preview(pdf_path: Path, output_path: Path) -> Path | None:
    try:
        import fitz  # PyMuPDF
    except ImportError:
        return None

    if not pdf_path.exists():
        return None

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with fitz.open(str(pdf_path)) as doc:
        if len(doc) == 0:
            return None
        pix = doc[0].get_pixmap(dpi=150)
        pix.save(str(output_path))
        if len(doc) > 1:
            pix2 = doc[1].get_pixmap(dpi=150)
            pix2.save(str(output_path.with_name("preview_page2.png")))
    return output_path
