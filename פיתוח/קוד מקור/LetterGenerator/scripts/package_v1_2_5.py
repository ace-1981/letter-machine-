"""Package LetterGenerator_V1.2.5_Portable — template-only patch (no app.exe rebuild).

V1.2.5 ships the RTL-robust template: paragraphs/styles/docDefaults now use
physical jc="right" + bidi instead of the fragile legacy "left"+bidi mirror, so
Word edits (adding a letter, bold, underline, ...) no longer flip the document
to the left. app.exe is unchanged (templates are external), so it is reused
as-is from V1.2.4.
"""

from __future__ import annotations

import hashlib
import shutil
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent          # .../LetterGenerator
WORKSPACE = ROOT.parents[2]                              # .../מכונת מכתבים
OLD = WORKSPACE / "OLD"
NEW_NAME = "LetterGenerator_V1.2.5_Portable"
PREV_NAME = "LetterGenerator_V1.2.4_Portable"
NEW_DIR = WORKSPACE / NEW_NAME
NEW_ZIP = WORKSPACE / f"{NEW_NAME}.zip"
PREV_DIR = WORKSPACE / PREV_NAME
PREV_ZIP = WORKSPACE / f"{PREV_NAME}.zip"
SRC_TEMPLATES = ROOT / "templates"

VERSION_TXT = """Version: 1.2.5
Based on: 1.2.4
Date: 27/06/2026
Change type: patch (template only)
Change summary: RTL robustness fix - paragraphs/styles/document-defaults use physical jc=right + bidi instead of the legacy left+bidi mirror. Editing the Word template (adding a letter, bold, underline, formatting) no longer flips the document to the left; alignment stays right even if Word strips a paragraph's RTL properties.
app.exe: unchanged (reused from 1.2.4 - templates are external)
Build: no new PyInstaller build
Status: latest usable version
"""

README_LATEST = """גרסה אחרונה לשימוש:
LetterGenerator_V1.2.5_Portable

קובץ להעברה:
LetterGenerator_V1.2.5_Portable.zip

הפעלה:
1. לחלץ את ה-ZIP לתיקייה במחשב.
2. להפעיל app.exe.
3. קבצי התבנית נמצאים בתיקיית templates.
4. ניתן לערוך את קובץ Word:
   templates/תחשיב זכויות אישי.docx
   כדי לשנות מלל ועיצוב — כולל הוספת אותיות, בולד, קו תחתון —
   היישור לימין יישמר.
5. את קובץ JSON יש לערוך רק להגדרות טכניות.
6. הפלט נשמר בתיקיית output.
7. נדרש Microsoft Word מותקן לצורך יצירת PDF.

נוהל גרסאות:
כל שינוי בתבנית, בפריסה או בקבצי Portable — גם קטן — מחייב מספר גרסה חדש
בשם התיקייה ובשם קובץ ה-ZIP (למשל V1.2.5, V1.2.6).

גרסאות ישנות (Portable + ZIP) נמצאות תחת:
OLD/

קוד, דוחות, בדיקות ודוגמאות נמצאים תחת:
פיתוח/
"""

PORTABLE_README = """מחולל מכתבים — Letter Generator V1.2.5 (Portable)
====================================================

אין צורך בהתקנה. העתיקו את כל התיקייה למיקום כלשהו והפעילו app.exe בדאבל-קליק.

מבנה תיקייה
-----------
LetterGenerator_V1.2.5_Portable/
  app.exe
  VERSION.txt
  templates/
    תחשיב זכויות אישי.docx
    תחשיב זכויות אישי.json
  output/
  README.txt

חדש ב-V1.2.5 (תבנית בלבד)
-------------------------
- חוסן RTL: עריכת קובץ ה-Word (הוספת אות, בולד, קו תחתון, עיצוב) לא משנה
  יותר את יישור המסמך לשמאל. היישור נשאר ימני גם אם Word "מנקה" מאפייני
  כיווניות מפסקה, כי היישור הפיזי (jc=right) מוגדר גם בברירת המחדל של המסמך.
- app.exe ללא שינוי (זהה ל-1.2.4) — התבנית חיצונית.

דרישות מערכת
------------
- Microsoft Word מותקן — נדרש להמרת DOCX ל-PDF.
- Windows 10 ומעלה.

הפעלה
-----
1. הפעילו app.exe.
2. בחרו Excel, סוג פלט, והפיקו או Preview.
3. פלט נשמר בתיקיית output/.

הערות
-----
- חתימה דיגיטלית אינטראקטיבית: Adobe Acrobat / Foxit.
- פרטי גרסה: VERSION.txt
"""


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def create_zip(source_dir: Path, zip_path: Path) -> None:
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for fp in sorted(source_dir.rglob("*")):
            if fp.is_file():
                zf.write(fp, fp.relative_to(source_dir.parent))


def archive_prev() -> None:
    OLD.mkdir(exist_ok=True)
    for src in (PREV_DIR, PREV_ZIP):
        if not src.exists():
            continue
        dest = OLD / src.name
        if dest.exists():
            if dest.is_dir():
                shutil.rmtree(dest, ignore_errors=True)
            else:
                dest.unlink()
        try:
            shutil.move(str(src), str(dest))
        except PermissionError:
            if src.is_dir():
                shutil.copytree(src, dest, dirs_exist_ok=True)
                shutil.rmtree(src, ignore_errors=True)
            else:
                shutil.copy2(src, dest)
                src.unlink(missing_ok=True)


def main() -> int:
    # 1. Regenerate the robust template into source templates/
    sys.path.insert(0, str(ROOT))
    from scripts.create_template_docx import build_with_python_docx

    docx_path = SRC_TEMPLATES / "תחשיב זכויות אישי.docx"
    build_with_python_docx(docx_path)
    print(f"Template regenerated (robust RTL): {docx_path}")

    if not PREV_DIR.is_dir():
        print(f"Previous portable not found: {PREV_DIR}")
        return 1

    # 2. Clone V1.2.4 portable -> V1.2.5 (reuse app.exe)
    if NEW_DIR.exists():
        shutil.rmtree(NEW_DIR, ignore_errors=True)
    shutil.copytree(PREV_DIR, NEW_DIR)

    # 3. Replace template docx with the robust one; keep JSON
    shutil.copy2(docx_path, NEW_DIR / "templates" / "תחשיב זכויות אישי.docx")

    # 4. Clean any stray output, keep README only
    out_dir = NEW_DIR / "output"
    out_dir.mkdir(exist_ok=True)
    for f in out_dir.iterdir():
        if f.name.lower() != "readme.txt":
            if f.is_dir():
                shutil.rmtree(f, ignore_errors=True)
            else:
                f.unlink()
    out_readme = out_dir / "README.txt"
    if not out_readme.exists():
        src_readme = ROOT / "release" / "OUTPUT_README.txt"
        if src_readme.is_file():
            shutil.copy2(src_readme, out_readme)
        else:
            out_readme.write_text("קבצי הפלט (PDF/DOCX) נשמרים בתיקייה זו.\n", encoding="utf-8")

    # 5. Version + readme files
    (NEW_DIR / "VERSION.txt").write_text(VERSION_TXT, encoding="utf-8")
    (NEW_DIR / "README.txt").write_text(PORTABLE_README, encoding="utf-8")

    # 6. Verify required files
    required = [
        NEW_DIR / "app.exe",
        NEW_DIR / "README.txt",
        NEW_DIR / "VERSION.txt",
        NEW_DIR / "templates" / "תחשיב זכויות אישי.docx",
        NEW_DIR / "templates" / "תחשיב זכויות אישי.json",
        NEW_DIR / "output" / "README.txt",
    ]
    missing = [str(p.relative_to(NEW_DIR)) for p in required if not p.exists()]
    if missing:
        print(f"Incomplete release: {missing}")
        return 1

    # 7. ZIP + archive previous + README pointer
    create_zip(NEW_DIR, NEW_ZIP)
    archive_prev()
    (WORKSPACE / "README_גרסה_אחרונה.txt").write_text(README_LATEST, encoding="utf-8")

    same_exe = sha256(NEW_DIR / "app.exe")
    print("=== V1.2.5 PACKAGE COMPLETE ===")
    print(f"Dir: {NEW_DIR}")
    print(f"ZIP: {NEW_ZIP} ({NEW_ZIP.stat().st_size} bytes)")
    print(f"app.exe sha256: {same_exe[:16]} (reused from 1.2.4)")
    print(f"Archived previous: {PREV_NAME} -> OLD/")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
