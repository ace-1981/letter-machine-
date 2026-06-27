"""Diagnose signature vs blue overlay in PDF."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import fitz
from pypdf import PdfReader

BLUE = (0.122, 0.306, 0.475)  # DARK_BLUE approx


def blue_rects(pg, y_min=0, y_max=9999):
    out = []
    for d in pg.get_drawings():
        fill = d.get("fill")
        rect = d.get("rect")
        if rect is None or not fill:
            continue
        r = fitz.Rect(rect)
        if r.y1 < y_min or r.y0 > y_max:
            continue
        if len(fill) >= 3 and fill[0] < 0.2 and fill[1] < 0.4 and fill[2] > 0.4:
            out.append(r)
    return out


def overlap(a: fitz.Rect, b: fitz.Rect) -> bool:
    return bool(a & b)


def analyze(pdf_path: Path) -> None:
    print("===", pdf_path.name, "===")
    doc = fitz.open(str(pdf_path))
    pg = doc[0]
    sig = next((w for w in (pg.widgets() or []) if w.field_name == "MemberSignature"), None)
    sig_rect = fitz.Rect(sig.rect) if sig and sig.rect else None
    print("sig_rect", sig_rect)

    blues = blue_rects(pg, 500, 700)
    print("blue rects near signature zone (y 500-700):", len(blues))
    for i, r in enumerate(blues[:15]):
        ov = overlap(r, sig_rect) if sig_rect else False
        print(f"  blue[{i}] {tuple(round(x,1) for x in (r.x0,r.y0,r.x1,r.y1))} overlap_sig={ov}")

    # text positions
    for label in ["חתימה", "הנחיות"]:
        hits = [h for h in pg.search_for(label) if h.y0 > 400]
        if hits:
            r = hits[0]
            print(label, "text", tuple(round(x, 1) for x in (r.x0, r.y0, r.x1, r.y1)))

    # annotation order
    reader = PdfReader(str(pdf_path))
    page = reader.pages[0]
    annots = page.get("/Annots") or []
    print("annot count", len(annots))
    for ref in annots:
        a = ref.get_object()
        print(" ", a.get("/T"), a.get("/Subtype"), a.get("/Rect"))

    doc.close()


def simulate_sign(src: Path, dst: Path) -> None:
    """Add fake signature ink on page to test layering."""
    import shutil

    shutil.copy2(src, dst)
    doc = fitz.open(str(dst))
    pg = doc[0]
    sig = next(w for w in (pg.widgets() or []) if w.field_name == "MemberSignature")
    r = fitz.Rect(sig.rect)
    # draw visible ink on page content stream (simulates some viewers)
    shape = pg.new_shape()
    shape.draw_line(fitz.Point(r.x0 + 10, r.y0 + r.height / 2), fitz.Point(r.x1 - 10, r.y0 + r.height / 2))
    shape.finish(color=(0, 0, 0), width=2)
    shape.commit(overlay=True)
    doc.saveIncr()
    doc.close()


if __name__ == "__main__":
    pdf = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "samples" / "test.pdf"
    analyze(pdf)
    if len(sys.argv) > 2 and sys.argv[2] == "simulate":
        out = pdf.with_name(pdf.stem + "_signed_test.pdf")
        simulate_sign(pdf, out)
        analyze(out)
