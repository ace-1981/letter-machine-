"""Quick rebuild of app.exe / app_debug.exe after source changes (no full smoke)."""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
RELEASE_DIR = ROOT.parent / "LetterGenerator_release"


def rebuild(name: str, *, console: bool) -> Path:
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--name",
        name,
        "--distpath",
        str(ROOT / "dist"),
        "--workpath",
        str(ROOT / "build"),
        "--specpath",
        str(ROOT),
        "--collect-all",
        "PySide6",
        "--hidden-import",
        "fitz",
        "--hidden-import",
        "docxtpl",
        "--hidden-import",
        "jinja2",
        "--hidden-import",
        "lxml",
        "--hidden-import",
        "pandas",
        "--hidden-import",
        "openpyxl",
        "--hidden-import",
        "xlrd",
        "--hidden-import",
        "pypdf",
        "--hidden-import",
        "win32com",
        "--hidden-import",
        "win32com.client",
        "--console" if console else "--windowed",
        str(ROOT / "app.py"),
    ]
    subprocess.run(cmd, check=True, cwd=ROOT)
    exe = ROOT / "dist" / f"{name}.exe"
    if not exe.is_file():
        raise FileNotFoundError(exe)
    return exe


def main() -> int:
    rebuild("app_debug", console=True)
    app = rebuild("app", console=False)
    if RELEASE_DIR.is_dir():
        shutil.copy2(app, RELEASE_DIR / "app.exe")
        print(f"Updated {RELEASE_DIR / 'app.exe'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
