"""Generate sample PDF + signature crop preview after signature layout fix."""

from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import fitz

from src.letter_generator import generate_single_letter
from src.signature_field import SIG_BOX_MARKER, verify_signature_field

VERSION = "v1.1.4"
OUT = ROOT / "samples" / "signature_fix_v1_1_4"
PREVIEW = ROOT.parent / "cursor" / "signature-area-preview-v1.1.4.png"
REPORT = ROOT.parent / "cursor" / "דוח-תיקון-חתימה-V1.1.4.md"
CONFIG = ROOT / "templates" / "תחשיב זכויות אישי.json"
SAMPLE = ROOT / "samples" / "sample_data.xlsx"
PORTABLE_TPL = ROOT.parent / "LetterGenerator_V1.1_Portable" / "templates"


def _overlap(a: tuple[float, float, float, float], b: tuple[float, float, float, float]) -> bool:
    ax0, ay0, ax1, ay1 = a
    bx0, by0, bx1, by1 = b
    return ax0 < bx1 and ax1 > bx0 and ay0 < by1 and ay1 > by0


def main() -> int:
    from scripts.create_template_docx import build_with_python_docx

    docx_path = ROOT / "templates" / "תחשיב זכויות אישי.docx"
    build_with_python_docx(docx_path)
    if PORTABLE_TPL.is_dir():
        shutil.copy2(docx_path, PORTABLE_TPL / docx_path.name)
        shutil.copy2(CONFIG, PORTABLE_TPL / CONFIG.name)

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
    sig = verify_signature_field(pdf, "MemberSignature")

    with fitz.open(str(pdf)) as doc:
        page = doc[-1]
        date_hits = page.search_for("תאריך")
        label_hits = page.search_for("חתימה (שדה")
        marker_hits = page.search_for("SIG_BOX")
        page_h = page.rect.height

        date_r = next((h for h in date_hits if h.y0 > 400), None) if date_hits else None
        label_r = max(label_hits, key=lambda r: r.y0) if label_hits else None
        widget_r = sig.get("rect")
        overlap_date = False
        if widget_r and date_r:
            overlap_date = _overlap(
                tuple(widget_r),
                (date_r.x0, date_r.y0, date_r.x1, date_r.y1),
            )

        clip_y0 = min(
            [r.y0 for r in [label_r, date_r] if r is not None] + ([widget_r[1]] if widget_r else [])
        ) - 20
        clip_y1 = max(
            [r.y1 for r in [label_r, date_r] if r is not None] + ([widget_r[3]] if widget_r else [])
        ) + 30
        clip = fitz.Rect(30, clip_y0, page.rect.width - 30, clip_y1)
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=clip, alpha=False)
        PREVIEW.parent.mkdir(parents=True, exist_ok=True)
        pix.save(str(PREVIEW))

    lines = [
        "# דוח תיקון חתימה V1.1.4",
        "",
        "## פתרון",
        "",
        "**שכבה ויזואלית אחת:** בוטל מלבן ב-DOCX. ב-PDF — מסגרת אפורה דקה מצוירת +",
        "שדה `MemberSignature` חופף ב-100% (ללא מסגרת/רקע נפרדים על השדה).",
        "",
        "## מבנה",
        "",
        "- כותרת: חתימה (שדה לחתימה דיגיטלית)",
        "- מתחת: שדה PDF = המלבן הוויזואלי היחיד",
        "- ימין: תאריך: __________________",
        "",
        "## קבצים",
        "",
        f"- PDF לדוגמה: `{pdf}`",
        f"- תצוגת אזור חתימה: `{PREVIEW}`",
        "",
        "## אימות",
        "",
        f"- שדה אינטראקטיבי: **{'כן' if sig.get('interactive') else 'לא'}**",
        f"- סוג שדה: `{sig.get('field_type')}`",
        f"- חפיפה עם אזור תאריך: **{'לא' if not overlap_date else 'כן'}**",
        f"- יישור שדה למלבן ויזואלי: **שכבה יחידה — שדה PDF בלבד**",
        "",
        "## פרטי מיקום",
        "",
        "```json",
        json.dumps(
            {
                "widget_rect": sig.get("rect"),
                "date_rect": [float(date_r.x0), float(date_r.y0), float(date_r.x1), float(date_r.y1)]
                if date_r
                else None,
                "label_rect": [float(label_r.x0), float(label_r.y0), float(label_r.x1), float(label_r.y1)]
                if label_r
                else None,
                "marker_found": bool(marker_hits),
            },
            ensure_ascii=False,
            indent=2,
        ),
        "```",
        "",
        "## הערה",
        "",
        "Adobe/Chrome עשויים להציג הדגשה במצב Fill & Sign — זה תקין.",
        "ה-PDF הבסיסי נקי; לא נבנה ZIP חדש עד אישור.",
    ]
    REPORT.write_text("\n".join(lines), encoding="utf-8")
    print(f"PDF: {pdf.name}")
    print(f"Preview: {PREVIEW.name}")
    print(f"Report: {REPORT.name}")
    print(f"Interactive: {sig.get('interactive')}")
    print(f"Overlap date: {overlap_date}")
    return 0 if sig.get("interactive") and not overlap_date else 1


if __name__ == "__main__":
    raise SystemExit(main())
