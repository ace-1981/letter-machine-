"""Verify V1.2.3 rollback layout and PDF signature layering."""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

WORKSPACE = Path(__file__).resolve().parents[4]
ROOT = Path(__file__).resolve().parent.parent
NEW_DIR = WORKSPACE / "LetterGenerator_V1.2.3_Portable"
ARCHIVE = WORKSPACE / "OLD"
SAMPLE = WORKSPACE / "פיתוח" / "דוגמאות" / "samples" / "sample_data.xlsx"
OUT = WORKSPACE / "פיתוח" / "דוגמאות" / "rollback_v1_2_3"
REPORT = WORKSPACE / "פיתוח" / "דוחות" / "דוח-rollback-V1.2.3.md"

sys.path.insert(0, str(ROOT))
import fitz
from pypdf import PdfReader

from src.letter_generator import generate_single_letter


def tail(docx: Path) -> list[str]:
    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    with zipfile.ZipFile(docx) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    items = []
    for p in root.iter(f"{w}p"):
        text = "".join(t.text or "" for t in p.iter(f"{w}t")).strip()
        if text:
            items.append(text)
    return items[-8:]


def main() -> int:
    v121 = ARCHIVE / "LetterGenerator_V1.2.1_Portable" / "templates" / "תחשיב זכויות אישי.docx"
    v123 = NEW_DIR / "templates" / "תחשיב זכויות אישי.docx"
    restored_tail = tail(v123)
    layout_ok = restored_tail == tail(v121)

    OUT.mkdir(parents=True, exist_ok=True)
    cfg = NEW_DIR / "templates" / "תחשיב זכויות אישי.json"
    pdf = Path(
        generate_single_letter(
            SAMPLE, cfg, OUT, output_format="pdf", pdf_preferred="word"
        )["pdf"]
    )
    generate_single_letter(SAMPLE, cfg, OUT, output_format="docx")

    signed = pdf.with_name(pdf.stem + "_signed.pdf")
    shutil.copy2(pdf, signed)
    doc = fitz.open(str(signed))
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
    pix = pg.get_pixmap(
        matrix=fitz.Matrix(3, 3),
        clip=fitz.Rect(rect.x0 + 5, rect.y0 + 5, rect.x1 - 5, rect.y1 - 5),
        alpha=False,
    )
    doc.saveIncr()
    doc.close()

    doc = fitz.open(str(pdf))
    pg = doc[0]
    guidelines_y = max(h.y0 for h in pg.search_for("הנחיות"))
    sig_y = max(h.y0 for h in pg.search_for("חתימה"))
    doc.close()

    order = [
        str(a.get_object().get("/T"))
        for a in (PdfReader(str(pdf)).pages[0].get("/Annots") or [])
    ]
    sig_rect = [round(v, 1) for v in rect]
    dark = sum(
        1
        for i in range(0, len(pix.samples), 3)
        if pix.samples[i] < 80 and pix.samples[i + 1] < 80 and pix.samples[i + 2] < 120
    )
    ink_visible = dark / (pix.width * pix.height) > 0.01

    try:
        subprocess.run(
            [str(NEW_DIR / "app.exe"), "preview", str(SAMPLE), "--row", "1", "--output", "output", "--format", "pdf"],
            check=True,
            cwd=str(NEW_DIR),
        )
        app_pdf = NEW_DIR / "output" / pdf.name
    except subprocess.CalledProcessError:
        app_pdf = None

    result = {
        "layout_ok": layout_ok,
        "sig_before_guidelines": sig_y < guidelines_y,
        "sig_rect": sig_rect,
        "annot_order": order,
        "sig_last": order[-1] == "MemberSignature" if order else False,
        "ink_visible_source": ink_visible,
        "app_pdf": str(app_pdf) if app_pdf else None,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))

    lines = [
        "# דוח Rollback — V1.2.3",
        "",
        "## גרסה לשימוש",
        "",
        "**LetterGenerator_V1.2.3_Portable** / `LetterGenerator_V1.2.3_Portable.zip`",
        "",
        "## Rollback",
        "",
        "| פריט | ערך |",
        "|------|-----|",
        "| בסיס | V1.2.1 (עימוד תקין) |",
        "| לא מבוסס על | V1.2.2 (שבורה, בארכיון) |",
        "",
        "## מה היה שבור ב-V1.2.2",
        "",
        "1. **החלפת סדר פסקאות:** חתימה+תאריך הועברו **אחרי** «הנחיות להחזרה»",
        "2. אינדקסים: הנחיות 51→47, חתימה 49→56",
        "3. בעיית החתימה לא נפתרה",
        "",
        "## מה שוחזר",
        "",
        f"- tail DOCX זהה ל-V1.2.1: **{'כן' if layout_ok else 'לא'}**",
        f"- חתימה לפני הנחיות (PDF): **{'כן' if sig_y < guidelines_y else 'לא'}**",
        "",
        "```",
        *restored_tail,
        "```",
        "",
        "## תיקון חתימה (קוד מקור — לא ב-app.exe)",
        "",
        "- `letter_generator.py`: תאריך לפני חתימה",
        "- `signature_field.py`: ללא overlay על content; `/Sig` בסוף `/Annots`",
        "- **build חדש: לא בוצע — נדרש אישור**",
        "",
        "## PDF",
        "",
        f"- `MemberSignature` rect: `{sig_rect}`",
        f"- `/Annots`: `{' → '.join(order)}`",
        f"- חתימה אחרונה: **{'כן' if order and order[-1] == 'MemberSignature' else 'לא'}**",
        f"- דיו גלוי (מקור, מדומה): **{'כן' if ink_visible else 'לא'}**",
        "",
        "## app.exe",
        "",
        "**לא השתנה** — זהה ל-V1.2.1",
        "",
        "## בדיקה",
        "",
        f"- PDF מקור: `{pdf}`",
        f"- PDF חתום (דמו): `{signed}`",
        f"- PDF app.exe: `{app_pdf}`",
    ]
    REPORT.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Report: {REPORT}")
    ok = layout_ok and sig_y < guidelines_y and ink_visible
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
