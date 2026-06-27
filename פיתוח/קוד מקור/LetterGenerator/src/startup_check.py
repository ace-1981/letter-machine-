"""Validate external template files at application startup."""

from __future__ import annotations

from pathlib import Path

from src.app_paths import get_templates_dir
from src.config_loader import load_template_config


def check_startup_templates() -> list[str]:
    """Return user-facing error messages; empty list means templates are OK."""
    errors: list[str] = []
    templates_dir = get_templates_dir()

    if not templates_dir.is_dir():
        return ["חסר קובץ הגדרות JSON", "חסרה תבנית מכתב"]

    json_files = sorted(templates_dir.glob("*.json"))
    if not json_files:
        errors.append("חסר קובץ הגדרות JSON")
        return errors

    missing_docx = False
    for json_file in json_files:
        try:
            config = load_template_config(json_file)
        except (OSError, ValueError, KeyError):
            errors.append("חסר קובץ הגדרות JSON")
            continue

        docx_name = config.get("template_file", "")
        if not docx_name:
            missing_docx = True
            continue

        docx_path = templates_dir / docx_name
        if not docx_path.is_file():
            missing_docx = True

    if missing_docx:
        errors.append("חסרה תבנית מכתב")

    return list(dict.fromkeys(errors))
