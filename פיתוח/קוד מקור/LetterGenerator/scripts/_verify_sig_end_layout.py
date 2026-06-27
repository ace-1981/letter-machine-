"""Verify signature-at-end layout + post-sign visibility; promote V1.2.2."""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
WORKSPACE = ROOT.parents[2]
sys.path.insert(0, str(ROOT))

import fitz

from scripts.create_template_docx import build_with_python_docx
from scripts.promote_portable_version import main as promote_release
from src.letter_generator import generate_single_letter
from src.signature_field import verify_signature_field

SAMPLE = WORKSPACE / "פיתוח" / "דוגמאות" / "samples" / "sample_data.xlsx"
CONFIG = ROOT / "templates" / "תחשיב זכויות אישי.json"
OUT = WORKSPACE / "פיתוח" / "דוגמאות" / "sig_layout_v1_2_2"
REPORT = WORKSPACE / "פיתוח" / "דוחות" / "דוח-תיקון-חתימה-סוף-מסמך-V1.2.2.md"
V122 = WORKSPACE / "LetterGenerator_V1.2.2_Portable"


def blue_fill(rect_fill) -> bool:
    if not rect_fill or len(rect_fill) < 3:
        return False
    r, g, b = rect_fill[:3]
    return r < 0.2 and g < 0.4 and b > 0.4


def overlap_area(a: fitz.Rect, b: fitz.Rect) -> float:
    inter = a & b
    return float(inter.get_area()) if inter else 0.0


def blue_below_sig(pg, sig_rect: fitz.Rect) -> list[dict]:
    hits = []
    for drawing in pg.get_drawings():
        fill = drawing.get("fill")
        rect = drawing.get("rect")
        if not rect or not blue_fill(fill):
            continue
        r = fitz.Rect(rect)
        if r.y0 >= sig_rect.y0 - 2:
            hits.append(
                {
                    "rect": [round(v, 1) for v in (r.x0, r.y0, r.x1, r.y1)],
                    "overlap_sig": round(overlap_area(r, sig_rect), 1),
                }
            )
    return hits


def section_order(pg) -> list[dict]:
    labels = [
        "תחשיב זכויות",
        "הערות והבהרות",
        "כתב קבלה",
        "הנחיות",
        "חתימה",
        "תאריך:",
    ]
    order = []
    for label in labels:
        hits = pg.search_for(label)
        if not hits:
            continue
        r = max(hits, key=lambda h: h.y0) if label in ("חתימה", "תאריך:") else hits[0]
        order.append({"label": label, "y0": round(r.y0, 1), "y1": round(r.y1, 1)})
    order.sort(key=lambda x: x["y0"])
    return order


def simulate_sign(src: Path, dst: Path) -> None:
    shutil.copy2(src, dst)
    doc = fitz.open(str(dst))
    pg = doc[0]
    sig = next(w for w in (pg.widgets() or []) if w.field_name == "MemberSignature")
    rect = fitz.Rect(sig.rect)
    shape = pg.new_shape()
    shape.draw_bezier(
        fitz.Point(rect.x0 + 10, rect.y0 + rect.height * 0.6),
        fitz.Point(rect.x0 + 60, rect.y0 + 15),
        fitz.Point(rect.x0 + 120, rect.y0 + 35),
        fitz.Point(rect.x1 - 10, rect.y0 + 25),
    )
    shape.finish(color=(0.05, 0.15, 0.45), width=2.2)
    shape.commit(overlay=True)
    doc.saveIncr()
    doc.close()


def ink_visible(pdf: Path, sig_rect: fitz.Rect) -> bool:
    doc = fitz.open(str(pdf))
    pg = doc[0]
    clip = fitz.Rect(sig_rect.x0 + 5, sig_rect.y0 + 5, sig_rect.x1 - 5, sig_rect.y1 - 5)
    pix = pg.get_pixmap(matrix=fitz.Matrix(3, 3), clip=clip, alpha=False)
    doc.close()
    dark = 0
    total = pix.width * pix.height
    samples = memoryview(pix.samples)
    for i in range(0, len(samples), 3):
        if samples[i] < 80 and samples[i + 1] < 80 and samples[i + 2] < 120:
            dark += 1
    return (dark / total) > 0.01 if total else False


def main() -> int:
    docx_path = ROOT / "templates" / "תחשיב זכויות אישי.docx"
    build_with_python_docx(docx_path)

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
    docx_out = generate_single_letter(
        excel_path=SAMPLE,
        config_path=CONFIG,
        output_dir=OUT,
        row_index=0,
        output_format="docx",
    )
    signed = pdf.with_name(pdf.stem + "_signed.pdf")
    simulate_sign(pdf, signed)

    doc = fitz.open(str(pdf))
    pg = doc[0]
    pages = len(doc)
    sig_w = next(w for w in (pg.widgets() or []) if w.field_name == "MemberSignature")
    date_w = next(w for w in (pg.widgets() or []) if w.field_name == "SignDateEntry")
    sig_rect = fitz.Rect(sig_w.rect)
    date_rect = fitz.Rect(date_w.rect)
    order = section_order(pg)
    blues_after = blue_below_sig(pg, sig_rect)
    doc.close()

    sig_info = verify_signature_field(pdf, "MemberSignature")
    visible = ink_visible(signed, sig_rect)

    # Promote portable release, then inject rebuilt template
    promote_release()
    shutil.copy2(docx_path, V122 / "templates" / docx_path.name)

    app_exe = V122 / "app.exe"
    app_pdf = OUT / "app_preview.pdf"
    subprocess.run(
        [
            str(app_exe),
            "preview",
            str(SAMPLE),
            "--row",
            "1",
            "--output",
            str(OUT),
            "--format",
            "pdf",
        ],
        check=True,
        cwd=str(V122),
    )
    app_pdfs = sorted(OUT.glob("*.pdf"), key=lambda p: p.stat().st_mtime, reverse=True)
    app_pdf = app_pdfs[0] if app_pdfs else None

    guidelines_y = next((x["y0"] for x in order if x["label"] == "הנחיות"), None)
    sig_y = next((x["y0"] for x in order if x["label"] == "חתימה"), None)
    order_ok = guidelines_y is not None and sig_y is not None and sig_y > guidelines_y

    report = {
        "version": "LetterGenerator_V1.2.2_Portable",
        "pages": pages,
        "section_order": order,
        "signature_after_guidelines": order_ok,
        "sig_rect": [round(v, 1) for v in (sig_rect.x0, sig_rect.y0, sig_rect.x1, sig_rect.y1)],
        "date_rect": [round(v, 1) for v in (date_rect.x0, date_rect.y0, date_rect.x1, date_rect.y1)],
        "blue_elements_at_or_below_signature": blues_after,
        "post_sign_ink_visible": visible,
        "interactive": sig_info.get("interactive"),
        "docx": str(docx_out.get("docx")),
        "pdf_source": str(pdf),
        "pdf_app_exe": str(app_pdf) if app_pdf else None,
        "pdf_signed": str(signed),
        "root_cause": "signature was before הנחיות in document flow; footer content painted over signed appearance",
    }

    lines = [
        "# דוח תיקון — חתימה בסוף המסמך (V1.2.2)",
        "",
        "## גרסה לשימוש",
        "",
        "**LetterGenerator_V1.2.2_Portable** / `LetterGenerator_V1.2.2_Portable.zip`",
        "",
        "## גורם ההסתרה",
        "",
        "אזור החתימה היה **לפני** «הנחיות להחזרה». בתזרים PDF, תוכן שמופיע אחרי החתימה",
        "(כותרת כחולה, קו, טקסט הנחיות) נצבע מעל אזור החתימה לאחר מילוי/חתימה.",
        "",
        "## סדר חדש במסמך",
        "",
        "1. כותרת + פרטים + טבלת חישוב",
        "2. הערות והבהרות (אם פעיל)",
        "3. כתב קבלה ושחרור",
        "4. **הנחיות להחזרה** + טקסט",
        "5. **חתימה + תאריך** (אחרון)",
        "",
        "### מיקומי Y בפועל (PDF)",
        "",
        "| אזור | y0 |",
        "|------|-----|",
    ]
    for item in order:
        lines.append(f"| {item['label']} | {item['y0']} |")
    lines.extend(
        [
            "",
            f"- עמודים: **{pages}**",
            f"- חתימה אחרי הנחיות: **{'כן' if order_ok else 'לא'}**",
            f"- כחול מתחת לחתימה: **{len(blues_after)}**",
            f"- דיו גלוי אחרי חתימה מדומה: **{'כן' if visible else 'לא'}**",
            "",
            "## rect שדות",
            "",
            f"- `MemberSignature`: `{report['sig_rect']}`",
            f"- `SignDateEntry`: `{report['date_rect']}`",
            "",
            "## קבצים ששונו",
            "",
            "- `פיתוח/קוד מקור/LetterGenerator/scripts/create_template_docx.py`",
            "- `פיתוח/קוד מקור/LetterGenerator/templates/תחשיב זכויות אישי.docx`",
            "- `LetterGenerator_V1.2.2_Portable/templates/תחשיב זכויות אישי.docx`",
            "- `README_גרסה_אחרונה.txt`",
            "- `promote_portable_version.py` (הכנה ל-V1.2.3)",
            "",
            "## build",
            "",
            "**לא** — app.exe ללא שינוי.",
            "",
            "## פלט לבדיקה",
            "",
            f"- DOCX: `{docx_out.get('docx')}`",
            f"- PDF (מקור): `{pdf}`",
            f"- PDF (app.exe): `{app_pdf}`",
            f"- PDF חתום (דמו): `{signed}`",
            "",
            "```json",
            json.dumps(report, ensure_ascii=False, indent=2),
            "```",
        ]
    )
    REPORT.write_text("\n".join(lines), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    print(f"Report: {REPORT}")

    ok = order_ok and not blues_after and visible and sig_info.get("interactive")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
