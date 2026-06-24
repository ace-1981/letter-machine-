"""Headless CLI for portable app testing and debugging."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from src.app_paths import get_app_root, get_default_output_dir, get_templates_dir
from src.config_loader import load_template_config
from src.letter_generator import generate_single_letter
from src.startup_check import check_startup_templates


def _configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


def _emit(text: str) -> None:
    """Print UTF-8 safely on Windows console builds."""
    data = (text + "\n").encode("utf-8", errors="replace")
    sys.stdout.buffer.write(data)
    sys.stdout.buffer.flush()


def _default_config() -> Path:
    json_files = sorted(get_templates_dir().glob("*.json"))
    if not json_files:
        raise FileNotFoundError("No JSON config found in templates/")
    return json_files[0]


def cmd_check() -> int:
    errors = check_startup_templates()
    if errors:
        for line in errors:
            print(line, file=sys.stderr)
        return 1
    _emit(f"OK  root={get_app_root()}")
    _emit(f"    templates={get_templates_dir()}")
    return 0


def cmd_preview(args: argparse.Namespace) -> int:
    excel = Path(args.excel).resolve()
    if not excel.is_file():
        print(f"Excel not found: {excel}", file=sys.stderr)
        return 1

    config = Path(args.config).resolve() if args.config else _default_config()
    output = Path(args.output).resolve() if args.output else get_default_output_dir()
    row_index = max(args.row - 1, 0)

    try:
        result = generate_single_letter(
            excel_path=excel,
            config_path=config,
            output_dir=output,
            row_index=row_index,
            pdf_preferred="word",
            keep_docx=False,
        )
    except Exception as exc:
        msg = str(exc)
        sys.stderr.buffer.write(msg.encode("utf-8", errors="replace") + b"\n")
        sys.stderr.buffer.flush()
        return 1
    _emit(str(result["pdf"]))
    return 0


def cmd_info() -> int:
    config_path = _default_config()
    config = load_template_config(config_path)
    info = {
        "app_root": str(get_app_root()),
        "templates_dir": str(get_templates_dir()),
        "config": str(config_path),
        "template_name": config.get("template_name"),
        "template_file": config.get("template_file"),
    }
    _emit(json.dumps(info, ensure_ascii=False, indent=2))
    return 0


def run_cli(argv: list[str] | None = None) -> int:
    _configure_stdio()
    parser = argparse.ArgumentParser(description="Letter Generator CLI")
    sub = parser.add_subparsers(dest="command", required=True)

    sub.add_parser("check", help="Validate external templates next to exe")

    preview = sub.add_parser("preview", help="Generate one preview PDF (headless)")
    preview.add_argument("excel", help="Path to Excel file")
    preview.add_argument("--row", type=int, default=1, help="Excel data row (1-based)")
    preview.add_argument("--output", help="Output directory (default: output/)")
    preview.add_argument("--config", help="JSON config path (default: first in templates/)")

    sub.add_parser("info", help="Print app root and loaded template info")

    args = parser.parse_args(argv)
    if args.command == "check":
        return cmd_check()
    if args.command == "preview":
        return cmd_preview(args)
    if args.command == "info":
        return cmd_info()
    return 2
