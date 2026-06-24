"""Add interactive PDF signature field, with anchor-based placement."""

from __future__ import annotations

import shutil
from pathlib import Path

ANCHOR_PATTERNS = (
    "חתימה (שדה לחתימה דיגיטלית)",
    ")חתימה (שדה לחתימה דיגיטלית",
    "שדה לחתימה דיגיטלית",
    "לחתימה דיגיטלית",
)


def add_signature_field(
    pdf_path: Path,
    output_path: Path,
    field_name: str,
    width: float,
    height: float,
    margin_bottom: float,
    margin_left: float,
    page: str = "last",
    anchor_text: str | None = None,
) -> dict:
    box = _resolve_box(
        pdf_path=pdf_path,
        page=page,
        width=width,
        height=height,
        margin_bottom=margin_bottom,
        margin_left=margin_left,
        anchor_text=anchor_text,
    )

    if pdf_path.resolve() != output_path.resolve():
        shutil.copy2(pdf_path, output_path)
        work_path = output_path
    else:
        work_path = pdf_path

    _add_signature_widget_fitz(work_path, field_name, box, page)
    return {"field_name": field_name, "box": box}


def _add_signature_widget_fitz(
    pdf_path: Path,
    field_name: str,
    box: tuple[float, float, float, float],
    page: str,
) -> None:
    import fitz

    x1, y1, x2, y2 = box
    with fitz.open(str(pdf_path)) as doc:
        page_index = _resolve_page_index_fitz(doc, page)
        pg = doc[page_index]
        page_height = pg.rect.height

        for widget in list(pg.widgets() or []):
            if widget.field_name == field_name:
                pg.delete_widget(widget)

        fitz_rect = fitz.Rect(x1, page_height - y2, x2, page_height - y1)
        widget = fitz.Widget()
        widget.field_type = fitz.PDF_WIDGET_TYPE_SIGNATURE
        widget.field_name = field_name
        widget.rect = fitz_rect
        widget.field_flags = 0
        pg.add_widget(widget)
        doc.saveIncr()


def _resolve_box(
    pdf_path: Path,
    page: str,
    width: float,
    height: float,
    margin_bottom: float,
    margin_left: float,
    anchor_text: str | None,
) -> tuple[float, float, float, float]:
    if anchor_text:
        anchored = _box_from_anchor(pdf_path, page, anchor_text, width, height)
        if anchored:
            return anchored

    x1 = margin_left
    y1 = margin_bottom
    return (x1, y1, x1 + width, y1 + height)


def _page_height(pdf_path: Path, page: str) -> float:
    import fitz

    with fitz.open(str(pdf_path)) as doc:
        page_index = _resolve_page_index_fitz(doc, page)
        return float(doc[page_index].rect.height)


def _box_from_anchor(
    pdf_path: Path,
    page: str,
    anchor_text: str,
    width: float,
    height: float,
) -> tuple[float, float, float, float] | None:
    try:
        import fitz
    except ImportError:
        return None

    patterns = [anchor_text, *ANCHOR_PATTERNS]
    seen: set[str] = set()
    patterns = [p for p in patterns if p and not (p in seen or seen.add(p))]

    with fitz.open(str(pdf_path)) as doc:
        page_index = _resolve_page_index_fitz(doc, page)
        pg = doc[page_index]
        page_height = pg.rect.height
        best = None

        for pattern in patterns:
            hits = pg.search_for(pattern)
            if not hits:
                continue
            rect = max(hits, key=lambda r: r.y0)
            best = rect
            break

        if best is None:
            return None

        fitz_top = best.y1 + 6
        fitz_bottom = fitz_top + height
        x1 = max(20, best.x0 - 8)
        y1 = page_height - fitz_bottom
        y2 = page_height - fitz_top
        return (x1, y1, x1 + width, y2)


def _resolve_page_index_fitz(doc, page: str) -> int:
    if page == "last":
        return len(doc) - 1
    index = int(page)
    if index < 0 or index >= len(doc):
        raise ValueError(f"Page index {index} out of range (0-{len(doc) - 1})")
    return index


def _widget_rect(pdf_path: Path, field_name: str) -> list[float] | None:
    try:
        import fitz
    except ImportError:
        return None

    with fitz.open(str(pdf_path)) as doc:
        for pg in doc:
            for widget in pg.widgets() or []:
                if widget.field_name == field_name and widget.rect:
                    r = widget.rect
                    return [float(r.x0), float(r.y0), float(r.x1), float(r.y1)]
    return None


def verify_signature_field(pdf_path: Path, field_name: str) -> dict:
    from pypdf import PdfReader

    reader = PdfReader(str(pdf_path))
    fields = reader.get_fields() or {}
    if field_name not in fields:
        return {
            "found": False,
            "field_name": field_name,
            "message": f"Field '{field_name}' not found. Available: {list(fields.keys())}",
        }

    field = fields[field_name]
    field_type = field.get("/FT", "")
    is_sig = str(field_type) == "/Sig" or field_type == "/Sig"

    rect = _widget_rect(pdf_path, field_name)
    annot_flags = _annotation_flags(pdf_path, field_name)

    pdf_rect = None
    if rect:
        page_height = _page_height(pdf_path, "last")
        pdf_rect = [
            rect[0],
            page_height - rect[3],
            rect[2],
            page_height - rect[1],
        ]

    locked = bool(annot_flags and (int(annot_flags) & 128))
    interactive = is_sig and not locked

    return {
        "found": True,
        "field_name": field_name,
        "field_type": str(field_type),
        "is_signature_field": is_sig,
        "rect": rect,
        "rect_pdf": pdf_rect,
        "annot_flags": annot_flags,
        "locked": locked,
        "interactive": interactive,
        "message": (
            "Valid interactive /Sig field"
            if interactive
            else "Signature field is locked or not interactive"
        ),
    }


def _annotation_flags(pdf_path: Path, field_name: str) -> int | None:
    from pypdf import PdfReader

    reader = PdfReader(str(pdf_path))
    for page in reader.pages:
        for annot_ref in page.get("/Annots") or []:
            annot = annot_ref.get_object()
            if annot.get("/T") == field_name or str(annot.get("/T", "")) == field_name:
                flags = annot.get("/F")
                return int(flags) if flags is not None else None
    return None
