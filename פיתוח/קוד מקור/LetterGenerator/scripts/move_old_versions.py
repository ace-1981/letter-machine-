"""Move old portable releases to OLD/ folder."""

from __future__ import annotations

import shutil
from pathlib import Path

WORKSPACE = Path(__file__).resolve().parents[4]
OLD = WORKSPACE / "OLD"
CURRENT = "LetterGenerator_V1.2.3_Portable"
CURRENT_ZIP = f"{CURRENT}.zip"

KEEP_IN_ROOT = {
    ".git",
    ".gitignore",
    CURRENT,
    CURRENT_ZIP,
    "README_גרסה_אחרונה.txt",
    "פיתוח",
    "OLD",
}


def move_item(src: Path, dest: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
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
    OLD.mkdir(exist_ok=True)
    moved: list[str] = []

    # Root-level old portable folders
    for item in WORKSPACE.iterdir():
        if item.name in KEEP_IN_ROOT:
            continue
        if item.is_dir() and item.name.startswith("LetterGenerator"):
            move_item(item, OLD / item.name)
            moved.append(item.name)
        elif item.is_file() and item.name.startswith("LetterGenerator") and item.suffix == ".zip":
            if item.name != CURRENT_ZIP:
                move_item(item, OLD / item.name)
                moved.append(item.name)

    # Legacy archive folder
    legacy = WORKSPACE / "פיתוח" / "גרסאות ישנות"
    if legacy.is_dir():
        for item in legacy.iterdir():
            move_item(item, OLD / item.name)
            moved.append(f"גרסאות ישנות/{item.name}")
        try:
            legacy.rmdir()
        except OSError:
            pass

    # README pointer
    readme = WORKSPACE / "README_גרסה_אחרונה.txt"
    text = readme.read_text(encoding="utf-8")
    old_line = "כל שאר הקבצים, הקוד, הדוחות, הבדיקות, הדוגמאות והגרסאות הישנות נמצאים תחת:\nפיתוח/"
    new_line = (
        "גרסאות ישנות (Portable + ZIP) נמצאות תחת:\nOLD/\n\n"
        "קוד, דוחות, בדיקות ודוגמאות נמצאים תחת:\nפיתוח/"
    )
    if old_line in text:
        readme.write_text(text.replace(old_line, new_line), encoding="utf-8")

    (OLD / "README.txt").write_text(
        "תיקיית OLD — גרסאות Portable ישנות\n"
        "==================================\n\n"
        f"גרסה לשימוש: {CURRENT}\n\n"
        "תוכן: גרסאות קודמות שהועברו מהשורש ומארכיון פיתוח/גרסאות ישנות.\n"
        "אין למחוק — לשמירה והשוואה בלבד.\n",
        encoding="utf-8",
    )

    print(f"OLD folder: {OLD}")
    print(f"Moved {len(moved)} items:")
    for name in moved:
        print(f"  - {name}")
    print("Root now:")
    for item in sorted(WORKSPACE.iterdir(), key=lambda p: p.name):
        print(f"  {item.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
