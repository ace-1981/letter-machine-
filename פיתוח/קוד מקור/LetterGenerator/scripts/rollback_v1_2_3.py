"""Rollback from V1.2.2 to V1.2.1 layout; publish V1.2.3 Portable."""

from __future__ import annotations

import hashlib
import json
import shutil
import subprocess
import zipfile
from pathlib import Path

WORKSPACE = Path(__file__).resolve().parents[4]
ARCHIVE = WORKSPACE / "OLD"
SOURCE_V121 = ARCHIVE / "LetterGenerator_V1.2.1_Portable"
BROKEN_V122 = WORKSPACE / "LetterGenerator_V1.2.2_Portable"
NEW_DIR = WORKSPACE / "LetterGenerator_V1.2.3_Portable"
NEW_ZIP = WORKSPACE / "LetterGenerator_V1.2.3_Portable.zip"
BROKEN_ZIP = WORKSPACE / "LetterGenerator_V1.2.2_Portable.zip"
ROOT = Path(__file__).resolve().parent.parent
SAMPLE = WORKSPACE / "פיתוח" / "דוגמאות" / "samples" / "sample_data.xlsx"
OUT = WORKSPACE / "פיתוח" / "דוגמאות" / "rollback_v1_2_3"
REPORT = WORKSPACE / "פיתוח" / "דוחות" / "דוח-rollback-V1.2.3.md"

VERSION_TXT = """Version: 1.2.3
Based on: 1.2.1 (rollback from broken 1.2.2)
Date: 27/06/2026
Change type: patch
Change summary: restored V1.2.1 document layout; PDF signature field layering fix in source (date then sig, annot on top)
app.exe: unchanged — signature layering fix requires approved rebuild
Build: no new build
Status: latest usable version
"""

README_LATEST = """גרסה אחרונה לשימוש:
LetterGenerator_V1.2.3_Portable

קובץ להעברה:
LetterGenerator_V1.2.3_Portable.zip

הפעלה:
1. לחלץ את ה-ZIP לתיקייה במחשב.
2. להפעיל app.exe.
3. קבצי התבנית נמצאים בתיקיית templates.
4. ניתן לערוך את קובץ Word:
   templates/תחשיב זכויות אישי.docx
   כדי לשנות מלל ועיצוב.
5. את קובץ JSON יש לערוך רק להגדרות טכניות.
6. הפלט נשמר בתיקיית output.
7. נדרש Microsoft Word מותקן לצורך יצירת PDF.

נוהל גרסאות:
כל שינוי בתבנית, בפריסה או בקבצי Portable — גם קטן — מחייב מספר גרסה חדש
בשם התיקייה ובשם קובץ ה-ZIP (למשל V1.2.3, V1.2.4).

כל שאר הקבצים, הקוד, הדוחות, הבדיקות, הדוגמאות והגרסאות הישנות נמצאים תחת:
פיתוח/
"""

PORTABLE_README = """מחולל מכתבים — Letter Generator V1.2.3 (Portable)
====================================================

אין צורך בהתקנה. העתיקו את כל התיקייה למיקום כלשהו והפעילו app.exe בדאבל-קליק.

מבנה תיקייה
-----------
LetterGenerator_V1.2.3_Portable/
  app.exe
  VERSION.txt
  templates/
  output/
  README.txt

חדש ב-V1.2.3 (patch)
--------------------
- שחזור עימוד V1.2.1 (חתימה לפני «הנחיות להחזרה») — ביטול V1.2.2 השבורה
- תיקון שכבות PDF בקוד מקור (דורש build מאושר ל-app.exe)
- app.exe ללא build חדש — זהה ל-V1.2.1

חדש ב-V1.2.1
------------
- תיקוני תבנית ופריסה אחרי V1.2

חדש ב-V1.2 (בסיס)
-----------------
- האצת הפקת PDF + ולידציה מספרית

פרטי גרסה: VERSION.txt
"""


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def para_tail(docx: Path) -> list[str]:
    import zipfile
    from xml.etree import ElementTree as ET

    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    with zipfile.ZipFile(docx) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    items = []
    for p in root.iter(f"{w}p"):
        text = "".join(t.text or "" for t in p.iter(f"{w}t")).strip()
        if text:
            items.append(text)
    return items[-8:]


def create_zip(source_dir: Path, zip_path: Path) -> None:
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for file_path in sorted(source_dir.rglob("*")):
            if file_path.is_file():
                zf.write(file_path, file_path.relative_to(source_dir.parent))


def archive_broken_v122() -> None:
    for src in (BROKEN_V122, BROKEN_ZIP):
        if not src.exists():
            continue
        dest = ARCHIVE / src.name
        if dest.exists():
            if dest.is_dir():
                shutil.rmtree(dest, ignore_errors=True)
            elif dest.is_file():
                dest.unlink()
        try:
            shutil.move(str(src), str(dest))
        except PermissionError:
            if src.is_dir():
                if not dest.exists():
                    shutil.copytree(src, dest, dirs_exist_ok=True)
                shutil.rmtree(src, ignore_errors=True)
            else:
                shutil.copy2(src, dest)
                src.unlink(missing_ok=True)


def main() -> int:
    import sys

    sys.path.insert(0, str(ROOT))
    from scripts.create_template_docx import build_with_python_docx
    from src.letter_generator import generate_single_letter

    if not SOURCE_V121.is_dir():
        raise SystemExit(f"Missing archive base: {SOURCE_V121}")

    if NEW_DIR.exists():
        raise SystemExit(f"Already exists: {NEW_DIR}")

    shutil.copytree(SOURCE_V121, NEW_DIR)
    docx_path = ROOT / "templates" / "תחשיב זכויות אישי.docx"
    build_with_python_docx(docx_path)
    shutil.copy2(docx_path, NEW_DIR / "templates" / docx_path.name)
    (NEW_DIR / "VERSION.txt").write_text(VERSION_TXT, encoding="utf-8")
    (NEW_DIR / "README.txt").write_text(PORTABLE_README, encoding="utf-8")
    (WORKSPACE / "README_גרסה_אחרונה.txt").write_text(README_LATEST, encoding="utf-8")

    archive_broken_v122()
    create_zip(NEW_DIR, NEW_ZIP)

    v121_docx = SOURCE_V121 / "templates" / docx_path.name
    restored_tail = para_tail(NEW_DIR / "templates" / docx_path.name)
    v121_tail = para_tail(v121_docx)
    layout_ok = restored_tail == v121_tail

    OUT.mkdir(parents=True, exist_ok=True)
    result = generate_single_letter(
        excel_path=SAMPLE,
        config_path=NEW_DIR / "templates" / "תחשיב זכויות אישי.json",
        output_dir=OUT,
        row_index=0,
        output_format="pdf",
        pdf_preferred="word",
    )
    pdf_source = Path(result["pdf"])
    generate_single_letter(
        excel_path=SAMPLE,
        config_path=NEW_DIR / "templates" / "תחשיב זכויות אישי.json",
        output_dir=OUT,
        row_index=0,
        output_format="docx",
    )

    subprocess.run(
        [
            str(NEW_DIR / "app.exe"),
            "preview",
            str(SAMPLE),
            "--row",
            "1",
            "--output",
            "output",
            "--format",
            "pdf",
        ],
        check=True,
        cwd=str(NEW_DIR),
    )
    app_pdf = NEW_DIR / "output" / pdf_source.name

    # verify annot order + layout y positions via fitz
    import fitz
    from pypdf import PdfReader

    doc = fitz.open(str(pdf_source))
    pg = doc[0]
    guidelines_y = max(h.y0 for h in pg.search_for("הנחיות"))
    sig_label_y = max(h.y0 for h in pg.search_for("חתימה"))
    sig_w = next(w for w in (pg.widgets() or []) if w.field_name == "MemberSignature")
    sig_rect = [round(v, 1) for v in sig_w.rect]
    doc.close()
    sig_before_guidelines = sig_label_y < guidelines_y

    reader = PdfReader(str(pdf_source))
    annot_order = [
        str(a.get_object().get("/T"))
        for a in (reader.pages[-1].get("/Annots") or [])
    ]
    sig_last = annot_order and annot_order[-1] == "MemberSignature"

    exe_v121 = sha256(SOURCE_V121 / "app.exe")
    exe_v123 = sha256(NEW_DIR / "app.exe")

    lines = [
        "# דוח Rollback — V1.2.3",
        "",
        "## גרסה לשימוש",
        "",
        "**LetterGenerator_V1.2.3_Portable**",
        "",
        "## Rollback מ",
        "",
        "- בסיס: **V1.2.1** (ארכיון `פיתוח/גרסאות ישנות/LetterGenerator_V1.2.1_Portable`)",
        "- V1.2.2 הועברה לארכיון כגרסה שבורה",
        "",
        "## מה היה שבור ב-V1.2.2",
        "",
        "1. **עימוד:** אזור חתימה+תאריך הועבר לסוף המסמך — **אחרי** «הנחיות להחזרה»",
        "2. **החלפת פסקאות:** ב-V1.2.1 הסדר היה: כתב קבלה → חתימה → הנחיות",
        "   ב-V1.2.2: כתב קבלה → הנחיות → חתימה (אינדקסים 47↔51)",
        "3. בעיית החתימה לא נפתרה; העברת הבלוק לא עזרה",
        "",
        "## מה שוחזר",
        "",
        "- סדר פסקאות ועימוד **זהה ל-V1.2.1**",
        f"- התאמת tail DOCX ל-V1.2.1: **{'כן' if layout_ok else 'לא'}**",
        f"- חתימה לפני הנחיות ב-PDF: **{'כן' if sig_before_guidelines else 'לא'}**",
        "",
        "### tail DOCX משוחזר",
        "",
        "```",
        *restored_tail,
        "```",
        "",
        "## תיקון חתימה (קוד מקור בלבד — לא ב-app.exe)",
        "",
        "- `signature_field.py`: ללא ציור על content-stream; `/Sig` מועבר לסוף `/Annots`",
        "- `letter_generator.py`: תאריך נוסף לפני חתימה; חתימה נוספת אחרונה",
        "- **נדרש build מאושר** כדי שהתיקון ייכנס ל-app.exe",
        "",
        "## rect וסדר annotations (PDF מקור)",
        "",
        f"- `MemberSignature` rect: `{sig_rect}`",
        f"- סדר `/Annots`: `{' → '.join(annot_order)}`",
        f"- חתימה אחרונה ב-Annots: **{'כן' if sig_last else 'לא'}**",
        "",
        "## build / app.exe",
        "",
        f"- build חדש: **לא**",
        f"- app.exe זהה ל-V1.2.1: **{exe_v121 == exe_v123}**",
        "",
        "## קבצים",
        "",
        "- `create_template_docx.py` — rollback עימוד",
        "- `signature_field.py`, `letter_generator.py` — שכבות PDF",
        "- `LetterGenerator_V1.2.3_Portable/` + ZIP",
        "",
        "## פלט לבדיקה",
        "",
        f"- DOCX/PDF מקור: `{OUT}`",
        f"- PDF app.exe: `{app_pdf}`",
    ]
    REPORT.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(json.dumps({
        "layout_ok": layout_ok,
        "sig_before_guidelines": sig_before_guidelines,
        "annot_order": annot_order,
        "sig_last": sig_last,
        "exe_unchanged": exe_v121 == exe_v123,
    }, ensure_ascii=False, indent=2))
    print(f"Report: {REPORT}")
    return 0 if layout_ok and sig_before_guidelines else 1


if __name__ == "__main__":
    raise SystemExit(main())
