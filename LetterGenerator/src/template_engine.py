"""DOCX template rendering with docxtpl."""

from __future__ import annotations

from datetime import date
from pathlib import Path

from docxtpl import DocxTemplate

from src.docx_table_builder import build_calc_table_subdoc
from src.excel_reader import apply_amount_formatting, format_amount


TABLE_ROW_DEFS = (
    {"num": "1", "desc": "מספר שנים לצורך התחשיב", "amount_key": "H", "is_total": False},
    {"num": "2", "desc": "סכום בגין שנת ותק אחת", "amount_static": "10 ₪", "is_total": False},
    {
        "num": "3",
        "desc": "סה\"כ בגין ותק (מספר שנים כפול סכום לשנת ותק)",
        "amount_key": "J",
        "is_total": True,
    },
    {"num": "4", "desc": "תוספת שווה לכל זכאי", "amount_key": "I", "is_total": False},
    {
        "num": "5",
        "desc": "תוספת בגין פטירה",
        "amount_key": "L",
        "condition": "show_DEATH_SECTION",
        "is_total": False,
    },
    {"num": "6", "desc": 'סה"כ זכאות', "amount_key": "M", "is_total": True},
    {
        "num": "7",
        "desc": "קיזוז בגין מענק עידוד עבודה",
        "amount_key": "O",
        "condition": "show_WORK_GRANT_SECTION",
        "is_total": False,
    },
    {
        "num": "ח",
        "desc": 'סה"כ אחרי קיזוז \'מענק עידוד עבודה\'',
        "amount_key": "P",
        "condition": "show_WORK_GRANT_SECTION",
        "is_total": True,
    },
)


def _build_table_rows(context: dict) -> list[dict]:
    rows: list[dict] = []
    for spec in TABLE_ROW_DEFS:
        condition = spec.get("condition")
        if condition and not context.get(condition):
            continue
        amount = spec.get("amount_static")
        if amount is None:
            key = spec["amount_key"]
            amount = context.get(key, "")
            if not isinstance(amount, str) or not amount.endswith("₪"):
                amount = format_amount(amount)
        rows.append(
            {
                "num": spec["num"],
                "desc": spec["desc"],
                "amount": amount,
                "is_total": spec.get("is_total", False),
            }
        )
    return rows


def render_template(
    template_path: Path,
    output_path: Path,
    context: dict,
) -> Path:
    render_context = apply_amount_formatting(dict(context))
    render_context.setdefault("TODAY", date.today().strftime("%d/%m/%Y"))
    render_context["table_rows"] = _build_table_rows(render_context)

    doc = DocxTemplate(str(template_path))
    render_context["calc_table"] = build_calc_table_subdoc(
        doc, render_context["table_rows"]
    )
    doc.render(render_context)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path


def build_output_filename(context: dict, config: dict, output_format: str = "pdf") -> str:
    from src.output_format import OUTPUT_DOCX, normalize_output_format

    pattern = config["output_filename"]["pattern"]
    if normalize_output_format(output_format) == OUTPUT_DOCX:
        pattern = pattern.replace(".pdf", ".docx")
    values = dict(context)
    values["template_name"] = config["template_name"]
    return pattern.format(**values)
