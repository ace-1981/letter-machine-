"""Resolve application paths relative to app.exe or project root."""

from __future__ import annotations

import os
import sys
from pathlib import Path


def get_app_root() -> Path:
    """Directory containing app.exe (frozen) or LetterGenerator project root."""
    env_root = os.environ.get("LETTER_GEN_APP_ROOT")
    if env_root:
        return Path(env_root).resolve()
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def get_templates_dir() -> Path:
    return get_app_root() / "templates"


def get_default_output_dir() -> Path:
    return get_app_root() / "output"
