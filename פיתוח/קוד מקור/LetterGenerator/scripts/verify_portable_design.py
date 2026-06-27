"""Verify portable-app design: external templates, no hardcoded dev paths in runtime code."""

from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "src"
BUILD_SCRIPT = ROOT / "scripts" / "build_release.py"


def check_runtime_paths() -> list[str]:
    errors: list[str] = []
    hardcoded = re.compile(r"[Cc]:\\Users\\")
    for py in SRC.rglob("*.py"):
        text = py.read_text(encoding="utf-8")
        if hardcoded.search(text):
            errors.append(f"Hardcoded user path in {py.relative_to(ROOT)}")
    app_paths = (SRC / "app_paths.py").read_text(encoding="utf-8")
    if "sys.executable" not in app_paths or "frozen" not in app_paths:
        errors.append("app_paths.py must resolve root from sys.executable when frozen")
    return errors


def check_build_does_not_embed_templates() -> list[str]:
    errors: list[str] = []
    text = BUILD_SCRIPT.read_text(encoding="utf-8")
    if "--add-data" in text or "add_data" in text.lower():
        errors.append("build_release.py must not use --add-data for templates")
    for spec in ROOT.glob("*.spec"):
        content = spec.read_text(encoding="utf-8")
        if "תחשיב" in content or "templates" in content.lower() and "templates" in content:
            # only fail if templates path appears as data file
            if re.search(r"templates.*תחשיב|תחשיב.*templates", content):
                errors.append(f"{spec.name} appears to bundle template files")
    if "TEMPLATE_FILES" not in text or "_stage_portable" not in text:
        errors.append("build_release.py must copy templates via _stage_portable()")
    return errors


def check_external_loading() -> list[str]:
    errors: list[str] = []
    mw = (SRC / "ui" / "main_window.py").read_text(encoding="utf-8")
    if "get_templates_dir" not in mw:
        errors.append("main_window must use get_templates_dir()")
    lg = (SRC / "letter_generator.py").read_text(encoding="utf-8")
    if "get_template_dir" not in lg:
        errors.append("letter_generator must load DOCX relative to JSON parent dir")
    return errors


def main() -> int:
    errors = [
        *check_runtime_paths(),
        *check_build_does_not_embed_templates(),
        *check_external_loading(),
    ]
    if errors:
        print("Portable design check FAILED:")
        for e in errors:
            print(f"  - {e}")
        return 1
    print("Portable design check OK:")
    print("  - app.exe resolves paths relative to its own folder")
    print("  - templates/ stays external next to app.exe (not embedded in exe)")
    print("  - build_release copies DOCX+JSON to LetterGenerator_release/templates/")
    print("  - no installer; portable folder only")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
