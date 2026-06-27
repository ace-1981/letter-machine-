"""Verify signature layering fix: no content-stream overlay + post-sign visibility."""

from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import fitz

from scripts.create_template_docx import build_with_python_docx
from src.letter_generator import generate_single_letter

OUT = ROOT.parent.parent / "דוגמאות" / "sig_layer_verify"
SAMPLE = ROOT.parent.parent / "דוגמאות" / "samples" / "sample_data.xlsx"
CONFIG = ROOT / "templates" / "תחשיב זכויות אישי.json"
PORTABLE_TPL = ROOT.parent.parent.parent / "LetterGenerator_V1.2_Portable" / "templates"
REPORT = ROOT.parent.parent / "דוחות" / "דוח-תיקון-חתימה-z-order.md"

DARK_BLUE_FILL = lambda f: (
    f
    and len(f) >= 3
    and f[0] < 0.2
    and f[1] < 0.4
    and f[2] > 0.4
)


def overlap_area(a: fitz.Rect, b: fitz.Rect) -> float:
    inter = a & b
    return float(inter.get_area()) if inter else 0.0


def content_overlays_in_sig_band(pg, sig_rect: fitz.Rect) -> list[dict]:
    """Unwanted content-stream paint inside the signature widget (from field placement)."""
    overlays = []
    for idx, drawing in enumerate(pg.get_drawings()):
        rect = drawing.get("rect")
        if rect is None:
            continue
        r = fitz.Rect(rect)
        if overlap_area(r, sig_rect) <= 0:
            continue
        fill = drawing.get("fill")
        stroke = drawing.get("color")
        # Ignore tiny viewer badge corners; flag full-box white/gray overlays.
        covers_sig = overlap_area(r, sig_rect) > sig_rect.get_area() * 0.5
        is_white_fill = fill and len(fill) >= 3 and min(fill) > 0.95
        is_gray_stroke = stroke and len(stroke) >= 3 and all(0.4 < c < 0.5 for c in stroke)
        if covers_sig and (is_white_fill or is_gray_stroke):
            overlays.append(
                {
                    "index": idx,
                    "rect": [round(v, 1) for v in (r.x0, r.y0, r.x1, r.y1)],
                    "fill": fill,
                    "stroke": stroke,
                }
            )
    return overlays


def annot_order_pdf(pdf_path: Path) -> list[str]:
    from pypdf import PdfReader

    reader = PdfReader(str(pdf_path))
    names = []
    for ref in reader.pages[0].get("/Annots") or []:
        annot = ref.get_object()
        names.append(str(annot.get("/T", annot.get("/Subtype"))))
    return names


def blue_rects_near(pg, sig_rect: fitz.Rect, pad: float = 30) -> list[dict]:
    zone = fitz.Rect(
        sig_rect.x0 - pad,
        sig_rect.y0 - pad,
        sig_rect.x1 + pad,
        sig_rect.y1 + pad,
    )
    hits = []
    for drawing in pg.get_drawings():
        fill = drawing.get("fill")
        rect = drawing.get("rect")
        if rect is None or not DARK_BLUE_FILL(fill):
            continue
        r = fitz.Rect(rect)
        if overlap_area(r, zone) > 0:
            hits.append(
                {
                    "rect": [round(v, 1) for v in (r.x0, r.y0, r.x1, r.y1)],
                    "overlap_sig_area": round(overlap_area(r, sig_rect), 1),
                }
            )
    return hits


def simulate_signature_ink(src: Path, dst: Path) -> fitz.Rect:
    shutil.copy2(src, dst)
    doc = fitz.open(str(dst))
    pg = doc[0]
    sig = next(w for w in (pg.widgets() or []) if w.field_name == "MemberSignature")
    rect = fitz.Rect(sig.rect)

    shape = pg.new_shape()
    p0 = fitz.Point(rect.x0 + 12, rect.y0 + rect.height * 0.65)
    p1 = fitz.Point(rect.x0 + rect.width * 0.35, rect.y0 + rect.height * 0.25)
    p2 = fitz.Point(rect.x0 + rect.width * 0.65, rect.y0 + rect.height * 0.7)
    p3 = fitz.Point(rect.x1 - 12, rect.y0 + rect.height * 0.4)
    shape.draw_bezier(p0, p1, p2, p3)
    shape.finish(color=(0.05, 0.15, 0.45), width=2.2)
    shape.commit(overlay=True)

    doc.saveIncr()
    doc.close()
    return rect


def ink_visible_after_sign(signed_pdf: Path, sig_rect: fitz.Rect) -> dict:
    doc = fitz.open(str(signed_pdf))
    pg = doc[0]
    clip = fitz.Rect(sig_rect.x0 + 5, sig_rect.y0 + 5, sig_rect.x1 - 5, sig_rect.y1 - 5)
    pix = pg.get_pixmap(matrix=fitz.Matrix(3, 3), clip=clip, alpha=False)
    samples = memoryview(pix.samples)
    stride = 3
    dark_pixels = 0
    total = pix.width * pix.height
    for i in range(0, len(samples), stride):
        r, g, b = samples[i], samples[i + 1], samples[i + 2]
        if r < 80 and g < 80 and b < 120 and (r + g + b) < 200:
            dark_pixels += 1
    doc.close()
    ratio = dark_pixels / total if total else 0
    return {"dark_pixel_ratio": round(ratio, 4), "visible": ratio > 0.01}



def main() -> int:
    docx_path = ROOT / "templates" / "תחשיב זכויות אישי.docx"
    build_with_python_docx(docx_path)
    if PORTABLE_TPL.is_dir():
        shutil.copy2(docx_path, PORTABLE_TPL / docx_path.name)

    OUT.mkdir(parents=True, exist_ok=True)
    result = generate_single_letter(
        excel_path=SAMPLE,
        config_path=CONFIG,
        output_dir=OUT,
        row_index=0,
        output_format="pdf",
        pdf_preferred="word",
    )
    pdf = Path(result["pdf"])
    signed = pdf.with_name(pdf.stem + "_signed.pdf")
    preview = pdf.with_name(pdf.stem + "_signed_preview.png")

    doc = fitz.open(str(pdf))
    pg = doc[0]
    sig = next(w for w in (pg.widgets() or []) if w.field_name == "MemberSignature")
    sig_rect = fitz.Rect(sig.rect)
    date = next((w for w in (pg.widgets() or []) if w.field_name == "SignDateEntry"), None)
    date_rect = fitz.Rect(date.rect) if date and date.rect else None

    overlays = content_overlays_in_sig_band(pg, sig_rect)
    blues = blue_rects_near(pg, sig_rect)
    guidelines = pg.search_for("הנחיות")
    guideline_rect = guidelines[0] if guidelines else None
    doc.close()
    order = annot_order_pdf(pdf)

    sig_rect_sim = simulate_signature_ink(pdf, signed)
    visibility = ink_visible_after_sign(signed, sig_rect_sim)

    doc = fitz.open(str(signed))
    pg = doc[0]
    clip = fitz.Rect(sig_rect_sim.x0 - 10, sig_rect_sim.y0 - 35, sig_rect_sim.x1 + 10, sig_rect_sim.y1 + 10)
    pix = pg.get_pixmap(matrix=fitz.Matrix(2, 2), clip=clip, alpha=False)
    pix.save(str(preview))
    doc.close()

    report = {
        "pdf": str(pdf),
        "signed_pdf": str(signed),
        "preview": str(preview),
        "sig_rect": [round(v, 1) for v in (sig_rect.x0, sig_rect.y0, sig_rect.x1, sig_rect.y1)],
        "date_rect": [round(v, 1) for v in (date_rect.x0, date_rect.y0, date_rect.x1, date_rect.y1)]
        if date_rect
        else None,
        "guidelines_rect": [round(v, 1) for v in (guideline_rect.x0, guideline_rect.y0, guideline_rect.x1, guideline_rect.y1)]
        if guideline_rect
        else None,
        "gap_sig_to_guidelines_pt": round(float(guideline_rect.y0 - sig_rect.y1), 1) if guideline_rect else None,
        "content_stream_overlays_in_sig": overlays,
        "blue_rects_overlapping_sig_zone": blues,
        "annot_paint_order": order,
        "post_sign_ink_visible": visibility,
        "root_cause": (
            "content-stream overlay + white widget fill painted above signing appearance"
            if overlays
            else "z-order / annotation stacking"
        ),
    }

    lines = [
        "# דוח תיקון — חתימה נעלמת מאחורי כחול (z-order)",
        "",
        "## מה הסתיר את החתימה",
        "",
        "| ממצא | ערך |",
        "|------|-----|",
        f"| סוג «הנחיות להחזרה» | פסקת טקסט כחול (`LG Section`) — **לא** shape/textbox/shading |",
        f"| מלבנים כחולים (fill) חופפים ל-rect החתימה | {len(blues)} (שטח חפיפה עם sig: "
        + ", ".join(str(b['overlap_sig_area']) for b in blues)
        + " pt²) |",
        f"| ציורי content-stream בתוך rect החתימה | {len(overlays)} |",
        f"| סדר annotations | {' → '.join(order)} |",
        "",
        "**מסקנת אבחון:** `_draw_field_box(overlay=True)` ו-`fill_color` לבן על ה-widget",
        "הוסיפו שכבת content מעל תוכן Word; אחרי חתימה, ה-appearance נבלע מתחת לשכבה זו.",
        "«הנחיות להחזרה» עצמה היא טקסט כחול (לא בלוק), אך נמצאת מתחת לשדה (~"
        f"{report['gap_sig_to_guidelines_pt']} pt מרווח).",
        "",
        "## תיקון",
        "",
        "- `signature_field.py`: בוטל ציור מלבן על content-stream; אין `fill_color` לבן;",
        "  מסגרת אפורה רק על ה-widget; `/Sig` מועבר לסוף `/Annots`.",
        "- `create_template_docx.py`: מרווח נוסף אחרי תיקון שכבות (12pt + 10pt לפני קו).",
        "",
        "## מדידות rect",
        "",
        "```json",
        json.dumps(report, ensure_ascii=False, indent=2),
        "```",
        "",
        "## אימות אחרי חתימה מדומה",
        "",
        f"- דמו חתימה ב-PDF: `{signed.name}`",
        f"- תצוגה: `{preview.name}`",
        f"- דיו גלוי בתוך rect: **{'כן' if visibility['visible'] else 'לא'}** "
        f"(dark_pixel_ratio={visibility['dark_pixel_ratio']})",
        "",
        "## קבצים ששונו",
        "",
        "- `פיתוח/קוד מקור/LetterGenerator/src/signature_field.py`",
        "- `פיתוח/קוד מקור/LetterGenerator/scripts/create_template_docx.py`",
        "- `LetterGenerator_V1.2_Portable/templates/תחשיב זכויות אישי.docx`",
        "",
        "**הערה:** `app.exe` של V1.2 אינו בתיקיית Portable (רק תבנית).",
        "ה-PDF אומת דרך קוד המקור; לתיקון שכבות ב-exe יידרש build עתידי.",
    ]
    REPORT.parent.mkdir(parents=True, exist_ok=True)
    REPORT.write_text("\n".join(lines), encoding="utf-8")

    print(json.dumps(report, ensure_ascii=False, indent=2))
    print(f"Report: {REPORT}")
    ok = visibility["visible"] and len(overlays) == 0
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
