"""DOCX template rendering with docxtpl."""

from __future__ import annotations

from datetime import date
from pathlib import Path

from docxtpl import DocxTemplate

from src.excel_reader import apply_amount_formatting


def render_template(
    template_path: Path,
    output_path: Path,
    context: dict,
) -> Path:
    render_context = apply_amount_formatting(dict(context))
    render_context.setdefault("TODAY", date.today().strftime("%d/%m/%Y"))

    doc = DocxTemplate(str(template_path))
    doc.render(render_context)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path


def build_output_filename(context: dict, config: dict) -> str:
    pattern = config["output_filename"]["pattern"]
    values = dict(context)
    values["template_name"] = config["template_name"]
    return pattern.format(**values)
