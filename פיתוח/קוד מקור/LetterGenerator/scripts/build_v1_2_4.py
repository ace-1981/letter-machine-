"""Build app.exe (signature layering fix) and package LetterGenerator_V1.2.4_Portable.

- Rebuilds the template (table layout fix) into source templates/.
- Runs PyInstaller via build_release with the fixed source.
- Assembles V1.2.4 portable folder + ZIP in the workspace root.
- Archives V1.2.3 to OLD/.
"""

from __future__ import annotations

import hashlib
import os
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
WORKSPACE = ROOT.parents[2]
OLD = WORKSPACE / "OLD"
NEW_NAME = "LetterGenerator_V1.2.4_Portable"
PREV_NAME = "LetterGenerator_V1.2.3_Portable"
NEW_DIR = WORKSPACE / NEW_NAME
NEW_ZIP = WORKSPACE / f"{NEW_NAME}.zip"
PREV_DIR = WORKSPACE / PREV_NAME
PREV_ZIP = WORKSPACE / f"{PREV_NAME}.zip"
STAGED = ROOT.parent / NEW_NAME  # build_release stages here (ROOT.parent)
OUTPUT_README = ROOT / "release" / "OUTPUT_README.txt"

VERSION_TXT = """Version: 1.2.4
Based on: 1.2.3
Date: 27/06/2026
Change type: patch + rebuild
Change summary: table layout fix (סעיף column widened/centered, right gap from frame) + signature PDF layering fix (no white overlay, signature widget on top, date created before signature)
app.exe: REBUILT (includes signature layering fix)
Build: new PyInstaller build
Status: latest usable version
"""

README_LATEST = """גרסה אחרונה לשימוש:
LetterGenerator_V1.2.4_Portable

קובץ להעברה:
LetterGenerator_V1.2.4_Portable.zip

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
בשם התיקייה ובשם קובץ ה-ZIP (למשל V1.2.4, V1.2.5).

גרסאות ישנות (Portable + ZIP) נמצאות תחת:
OLD/

קוד, דוחות, בדיקות ודוגמאות נמצאים תחת:
פיתוח/
"""

PORTABLE_README = """מחולל מכתבים — Letter Generator V1.2.4 (Portable)
====================================================

אין צורך בהתקנה. העתיקו את כל התיקייה למיקום כלשהו והפעילו app.exe בדאבל-קליק.

מבנה תיקייה
-----------
LetterGenerator_V1.2.4_Portable/
  app.exe
  VERSION.txt
  templates/
    תחשיב זכויות אישי.docx
    תחשיב זכויות אישי.json
  output/
  README.txt

חדש ב-V1.2.4 (build חדש)
------------------------
- תיקון פריסת טבלה: עמודת «סעיף» הורחבה וממורכזת, מרווח מהמסגרת הימנית
- תיקון שכבות חתימה ב-PDF: בלי רקע לבן שמכסה, שדה החתימה בשכבה עליונה,
  שדה התאריך נוצר לפני שדה החתימה — החתימה נשארת גלויה אחרי חתימה
- app.exe נבנה מחדש עם התיקונים

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
    # 1. Rebuild template (table fix) into source templates/
    sys.path.insert(0, str(ROOT))
    from scripts.create_template_docx import build_with_python_docx

    docx_path = ROOT / "templates" / "תחשיב זכויות אישי.docx"
    build_with_python_docx(docx_path)
    print(f"Template rebuilt: {docx_path}")

    # 2. PyInstaller build via build_release (stages to ROOT.parent/NEW_NAME)
    env = {**os.environ, "LG_RELEASE_DIR": NEW_NAME, "PYTHONUTF8": "1"}
    rc = subprocess.run(
        [sys.executable, str(ROOT / "scripts" / "build_release.py")],
        cwd=ROOT,
        env=env,
    ).returncode
    if rc != 0:
        print("BUILD FAILED")
        return rc
    if not STAGED.is_dir():
        print(f"Staged release not found: {STAGED}")
        return 1

    # 3. Assemble final portable in workspace root
    if NEW_DIR.exists():
        shutil.rmtree(NEW_DIR, ignore_errors=True)
    shutil.move(str(STAGED), str(NEW_DIR))
    (NEW_DIR / "VERSION.txt").write_text(VERSION_TXT, encoding="utf-8")
    (NEW_DIR / "README.txt").write_text(PORTABLE_README, encoding="utf-8")
    (NEW_DIR / "output").mkdir(exist_ok=True)
    if OUTPUT_README.is_file():
        shutil.copy2(OUTPUT_README, NEW_DIR / "output" / "README.txt")

    # 4. Verify required files
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

    # 5. ZIP + archive previous + README pointer
    create_zip(NEW_DIR, NEW_ZIP)
    archive_prev()
    (WORKSPACE / "README_גרסה_אחרונה.txt").write_text(README_LATEST, encoding="utf-8")

    print("=== V1.2.4 BUILD COMPLETE ===")
    print(f"Dir: {NEW_DIR}")
    print(f"ZIP: {NEW_ZIP} ({NEW_ZIP.stat().st_size} bytes)")
    print(f"app.exe sha256: {sha256(NEW_DIR / 'app.exe')[:16]}")
    print(f"Archived previous: {PREV_NAME} -> OLD/")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
