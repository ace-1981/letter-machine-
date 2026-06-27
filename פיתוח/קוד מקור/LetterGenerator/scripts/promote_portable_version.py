"""Promote portable release folder + ZIP to next patch version (no rebuild)."""

from __future__ import annotations

import hashlib
import shutil
import zipfile
from datetime import date
from pathlib import Path

WORKSPACE = Path(__file__).resolve().parents[4]
OLD_NAME = "LetterGenerator_V1.2.3_Portable"
NEW_NAME = "LetterGenerator_V1.2.4_Portable"
ARCHIVE_DIR = WORKSPACE / "OLD"
OUTPUT_README_SRC = (
    Path(__file__).resolve().parents[1] / "release" / "OUTPUT_README.txt"
)

VERSION_TXT = """Version: 1.2.2
Based on: 1.2.1
Date: 27/06/2026
Change type: patch
Change summary: signature/date moved to end of document after all blue headings and הנחיות
app.exe: unchanged unless explicitly rebuilt
Build: no new build unless stated otherwise
Status: latest usable version
"""

README_LATEST = """גרסה אחרונה לשימוש:
LetterGenerator_V1.2.2_Portable

קובץ להעברה:
LetterGenerator_V1.2.2_Portable.zip

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
בשם התיקייה ובשם קובץ ה-ZIP (למשל V1.2.1, V1.2.2).

כל שאר הקבצים, הקוד, הדוחות, הבדיקות, הדוגמאות והגרסאות הישנות נמצאים תחת:
פיתוח/
"""

PORTABLE_README = """מחולל מכתבים — Letter Generator V1.2.2 (Portable)
====================================================

אין צורך בהתקנה. העתיקו את כל התיקייה למיקום כלשהו והפעילו app.exe בדאבל-קליק.

מבנה תיקייה
-----------
LetterGenerator_V1.2.2_Portable/
  app.exe
  VERSION.txt
  templates/
    תחשיב זכויות אישי.docx
    תחשיב זכויות אישי.json
  output/
  README.txt

חדש ב-V1.2.2 (patch)
--------------------
- אזור חתימה + תאריך הועבר לסוף המסמך — אחרי «הנחיות להחזרה»
- app.exe ללא build חדש — זהה ל-V1.2.1

חדש ב-V1.2.1 (patch)
--------------------
- תיקוני תבנית ופריסה אחרי V1.2

חדש ב-V1.2 (בסיס)
-----------------
- האצת הפקת PDF: Word נפתח פעם אחת לכל batch (לא לכל מכתב)
- תיקון ולידציה: סכום לא מספרי (עמודה P ושדות מספריים) — שגיאה + errors_report.csv

סוג פלט
-------
  PDF      — יוצר PDF בלבד (ברירת מחדל)
  Word/DOCX — יוצר DOCX בלבד

עריכת תבנית — בלי build מחדש
-----------------------------
טקסטים ועיצוב:  templates/תחשיב זכויות אישי.docx
מיפוי ותנאים:   templates/תחשיב זכויות אישי.json

דרישות מערכת
------------
- Microsoft Word מותקן — נדרש להמרת DOCX ל-PDF (מצב PDF בלבד).
- Windows 10 ומעלה.

הפעלה
-----
1. הפעילו app.exe.
2. בחרו Excel, סוג פלט, והפיקו או Preview.
3. פלט נשמר בתיקיית output/ (ברירת מחדל).

הערות
-----
- חתימה דיגיטלית אינטראקטיבית: Adobe Acrobat / Foxit.
- אם חסר DOCX או JSON — תוצג הודעה בהפעלה.
- פרטי גרסה: VERSION.txt
"""


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def create_zip(source_dir: Path, zip_path: Path) -> None:
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for file_path in sorted(source_dir.rglob("*")):
            if file_path.is_file():
                zf.write(file_path, file_path.relative_to(source_dir.parent))


def verify_release(release: Path) -> list[str]:
    required = [
        release / "app.exe",
        release / "README.txt",
        release / "VERSION.txt",
        release / "templates" / "תחשיב זכויות אישי.docx",
        release / "templates" / "תחשיב זכויות אישי.json",
        release / "output" / "README.txt",
    ]
    missing = [str(p.relative_to(release)) for p in required if not p.exists()]
    return missing


def main() -> int:
    old_dir = WORKSPACE / OLD_NAME
    new_dir = WORKSPACE / NEW_NAME
    old_zip = WORKSPACE / f"{OLD_NAME}.zip"
    new_zip = WORKSPACE / f"{NEW_NAME}.zip"
    report_path = WORKSPACE / "פיתוח" / "דוחות" / f"דוח-גרסה-{NEW_NAME.replace('LetterGenerator_', '').replace('_Portable', '')}.md"

    if not old_dir.is_dir():
        raise SystemExit(f"Missing source release: {old_dir}")
    if new_dir.exists():
        raise SystemExit(f"Target already exists: {new_dir}")

    old_exe_hash = sha256(old_dir / "app.exe")

    shutil.copytree(old_dir, new_dir)
    (new_dir / "VERSION.txt").write_text(VERSION_TXT, encoding="utf-8")
    (new_dir / "README.txt").write_text(PORTABLE_README, encoding="utf-8")
    (new_dir / "output").mkdir(exist_ok=True)
    shutil.copy2(OUTPUT_README_SRC, new_dir / "output" / "README.txt")

    missing = verify_release(new_dir)
    if missing:
        raise SystemExit(f"Release incomplete: {missing}")

    new_exe_hash = sha256(new_dir / "app.exe")
    exe_changed = old_exe_hash != new_exe_hash

    create_zip(new_dir, new_zip)

    ARCHIVE_DIR.mkdir(parents=True, exist_ok=True)
    for src in (old_dir, old_zip):
        if src.exists():
            dest = ARCHIVE_DIR / src.name
            if dest.exists():
                raise SystemExit(f"Archive target already exists: {dest}")
            shutil.move(str(src), str(dest))

    extra_root = WORKSPACE / "LetterGenerator"
    if extra_root.exists():
        dest_lg = ARCHIVE_DIR / "LetterGenerator"
        if dest_lg.exists():
            raise SystemExit(f"Archive target already exists: {dest_lg}")
        shutil.move(str(extra_root), str(dest_lg))

    (WORKSPACE / "README_גרסה_אחרונה.txt").write_text(README_LATEST, encoding="utf-8")

    allowed_root = {
        ".git",
        ".gitignore",
        NEW_NAME,
        f"{NEW_NAME}.zip",
        "README_גרסה_אחרונה.txt",
        "פיתוח",
    }
    root_items = {p.name for p in WORKSPACE.iterdir()}

    lines = [
        "# דוח גרסה — V1.2.1 Portable",
        "",
        f"**תאריך:** {date.today().strftime('%d/%m/%Y')}",
        "",
        "## סיכום",
        "",
        f"| פריט | ערך |",
        f"|------|-----|",
        f"| גרסה חדשה | `{NEW_NAME}/` |",
        f"| ZIP | `{NEW_NAME}.zip` |",
        f"| גרסה אחרונה לשימוש | **LetterGenerator_V1.2.1_Portable** |",
        f"| הועבר לארכיון | `{OLD_NAME}/`, `{OLD_NAME}.zip` → `OLD/` |",
        f"| app.exe השתנה | **{'כן' if exe_changed else 'לא'}** (SHA256 זהה ל-V1.2) |",
        f"| build חדש | **לא** |",
        "",
        "## תוכן VERSION.txt",
        "",
        "```",
        VERSION_TXT.strip(),
        "```",
        "",
        "## שורש הפרויקט (אחרי סידור)",
        "",
        "קבצים ותיקיות בשורש:",
        "",
    ]
    for name in sorted(root_items):
        lines.append(f"- `{name}`")
    if root_items - allowed_root:
        lines.extend(
            [
                "",
                "**הערה:** קיימים פריטים נוספים בשורש מעבר לרשימה המומלצת:",
                "",
            ]
        )
        for name in sorted(root_items - allowed_root):
            lines.append(f"- `{name}`")
    else:
        lines.append("")
        lines.append("שורש הפרויקט תואם את הרשימה המומלצת.")

    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text("\n".join(lines) + "\n", encoding="utf-8")

    print(f"Created: {new_dir}")
    print(f"ZIP: {new_zip}")
    print(f"Archived: {OLD_NAME} -> {ARCHIVE_DIR}")
    print(f"app.exe changed: {exe_changed}")
    print(f"Report: {report_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
