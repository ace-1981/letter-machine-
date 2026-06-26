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

SIG_BOX_MARKER = "⟦SIG_BOX⟧"
DATE_BOX_MARKER = "⟦DATE_BOX⟧"
DATE_ANCHOR = "תאריך:"
SIG_BORDER_RGB = (0.45, 0.45, 0.45)
SIG_FILL_RGB = (1.0, 1.0, 1.0)


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


def add_date_field(
    pdf_path: Path,
    output_path: Path,
    field_name: str,
    width: float,
    height: float,
    page: str = "last",
) -> dict:
    box = _resolve_date_box(pdf_path, page, width, height)
    if box is None:
        raise ValueError("Could not place date field — anchor not found in PDF")

    if pdf_path.resolve() != output_path.resolve():
        shutil.copy2(pdf_path, output_path)
        work_path = output_path
    else:
        work_path = pdf_path

    _add_text_widget_fitz(work_path, field_name, box, page)
    return {"field_name": field_name, "box": box}


def _draw_field_box(pg, fitz_rect) -> None:
    pg.draw_rect(
        fitz_rect,
        color=SIG_BORDER_RGB,
        width=0.75,
        overlay=True,
    )


def _add_text_widget_fitz(
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
        _draw_field_box(pg, fitz_rect)

        widget = fitz.Widget()
        widget.field_type = fitz.PDF_WIDGET_TYPE_TEXT
        widget.field_name = field_name
        widget.rect = fitz_rect
        widget.field_flags = 0
        widget.border_width = 0
        widget.border_color = (1, 1, 1)
        widget.fill_color = SIG_FILL_RGB
        widget.text_color = (0, 0, 0)
        widget.text_fontsize = 10
        widget.field_value = ""
        widget.text_maxlen = 40
        pg.add_widget(widget)
        for w in pg.widgets() or []:
            if w.field_name == field_name:
                w.field_value = ""
                w.text_color = (0, 0, 0)
                w.update()
                break
        doc.saveIncr()


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
        _draw_field_box(pg, fitz_rect)

        widget = fitz.Widget()
        widget.field_type = fitz.PDF_WIDGET_TYPE_SIGNATURE
        widget.field_name = field_name
        widget.rect = fitz_rect
        widget.field_flags = 0
        widget.border_width = 0
        widget.border_color = (1, 1, 1)
        widget.fill_color = SIG_FILL_RGB
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

    with fitz.open(str(pdf_path)) as doc:
        page_index = _resolve_page_index_fitz(doc, page)
        pg = doc[page_index]
        page_height = pg.rect.height
        page_width = pg.rect.width

        label = _find_label_rect(pg, anchor_text)
        if label is None:
            return None

        date_rect = _find_date_rect(pg)
        gap = 32.0
        max_x2 = page_width - 36.0
        if date_rect is not None:
            max_x2 = min(max_x2, float(date_rect.x0) - gap)

        box_top, box_bottom = _resolve_box_vertical(pg, label, height)
        x2 = min(max_x2, float(label.x1) + 4)
        x1 = max(36.0, x2 - width)
        if x2 - x1 < width:
            x2 = x1 + width

        y1 = page_height - box_bottom
        y2 = page_height - box_top
        return (x1, y1, x2, y2)


def _resolve_date_box(
    pdf_path: Path,
    page: str,
    width: float,
    height: float,
) -> tuple[float, float, float, float] | None:
    try:
        import fitz
    except ImportError:
        return None

    with fitz.open(str(pdf_path)) as doc:
        page_index = _resolve_page_index_fitz(doc, page)
        pg = doc[page_index]
        page_height = pg.rect.height

        date_label = _find_date_rect(pg)
        if date_label is None:
            return None

        marker = _find_date_marker_rect(pg) or _find_marker_rect(pg)
        if marker is not None:
            top = float(marker.y0) - 2
        elif date_label is not None:
            top = float(date_label.y1) + 8
        else:
            return None

        bottom = top + height
        if date_label is not None:
            x2 = float(date_label.x0) - 10
        else:
            x2 = pg.rect.width - 40
        x1 = max(280.0, x2 - width)

        y1 = page_height - bottom
        y2 = page_height - top
        return (x1, y1, x2, y2)


def _find_date_marker_rect(pg):
    import fitz

    for pattern in (DATE_BOX_MARKER, "DATE_BOX"):
        hits = pg.search_for(pattern)
        if hits:
            return max(hits, key=lambda r: r.y0)
    return None


def _find_label_rect(pg, anchor_text: str):
    import fitz

    patterns = [anchor_text, *ANCHOR_PATTERNS]
    seen: set[str] = set()
    patterns = [p for p in patterns if p and not (p in seen or seen.add(p))]

    best = None
    for pattern in patterns:
        hits = pg.search_for(pattern)
        if not hits:
            continue
        rect = max(hits, key=lambda r: r.y0)
        best = rect
        break
    return best


def _find_date_rect(pg):
    import fitz

    hits = pg.search_for(DATE_ANCHOR)
    if not hits:
        hits = pg.search_for("תאריך")
    if not hits:
        return None
    footer_min_y = pg.rect.height * 0.15
    footer_hits = [h for h in hits if h.y0 > footer_min_y]
    if not footer_hits:
        footer_hits = hits
    return max(footer_hits, key=lambda r: r.y0)


def _find_marker_rect(pg):
    import fitz

    for pattern in (SIG_BOX_MARKER, "SIG_BOX"):
        hits = pg.search_for(pattern)
        if hits:
            return hits[0]
    return None


def _find_signature_visual_rect(pg, label) -> object | None:
    """Locate the single signature rectangle below the label."""
    import fitz

    label_bottom = float(label.y1)
    area_candidates: list[fitz.Rect] = []
    hlines: list[fitz.Rect] = []
    vlines: list[fitz.Rect] = []

    for drawing in pg.get_drawings():
        rect = drawing.get("rect")
        if rect is None:
            continue
        r = fitz.Rect(rect)
        if r.y0 < label_bottom + 8 or r.y0 > label_bottom + 70:
            continue
        if r.width >= 90 and r.height >= 18:
            area_candidates.append(r)
        elif r.height <= 2.5 and r.width >= 90:
            hlines.append(r)
        elif r.width <= 2.5 and r.height >= 18:
            vlines.append(r)

    if area_candidates:
        return max(area_candidates, key=lambda r: r.width * r.height)

    if len(hlines) >= 2 and len(vlines) >= 2:
        x0 = min(r.x0 for r in vlines)
        x1 = max(r.x1 for r in vlines)
        y0 = min(r.y0 for r in hlines)
        y1 = max(r.y1 for r in hlines)
        if x1 - x0 >= 90 and y1 - y0 >= 18:
            return fitz.Rect(x0, y0, x1, y1)
    return None


def _box_from_visual_rect(
    visual,
    page_height: float,
    *,
    pad: float = 1.0,
) -> tuple[float, float, float, float]:
    x1 = float(visual.x0) + pad
    x2 = float(visual.x1) - pad
    fitz_top = float(visual.y0) + pad
    fitz_bottom = float(visual.y1) - pad
    y1 = page_height - fitz_bottom
    y2 = page_height - fitz_top
    return (x1, y1, x2, y2)


def _resolve_box_vertical(pg, label, height: float) -> tuple[float, float]:
    marker = _find_marker_rect(pg)
    if marker is not None:
        top = float(marker.y0) - 2
        return top, top + height
    top = float(label.y1) + 12
    return top, top + height


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
