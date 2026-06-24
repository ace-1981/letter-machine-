"""Excel reading with support for xlsx (openpyxl) and legacy xls (xlrd)."""

from __future__ import annotations

import string
from pathlib import Path

import pandas as pd

AMOUNT_FIELDS = ("H", "I", "J", "L", "M", "O", "P")


def format_amount(value) -> str:
    """Format numeric value as right-aligned-friendly Hebrew amount: 12,345 ₪"""
    if value is None or value == "":
        return "0 ₪"
    if isinstance(value, str):
        stripped = value.strip()
        if stripped.endswith("₪"):
            return stripped
        try:
            num = float(stripped.replace(",", ""))
        except ValueError:
            return f"{stripped} ₪"
    else:
        try:
            num = float(value)
        except (TypeError, ValueError):
            return f"{value} ₪"

    if num == int(num):
        abs_part = f"{int(abs(num)):,}"
    else:
        abs_part = f"{abs(num):,.2f}".rstrip("0").rstrip(".")

    if num < 0:
        return f"\u200e-{abs_part} ₪"
    return f"\u200e{abs_part} ₪"


def apply_amount_formatting(context: dict) -> dict:
    formatted = dict(context)
    for field in AMOUNT_FIELDS:
        if field in formatted:
            formatted[field] = format_amount(formatted[field])
    return formatted

def column_letter_to_index(letter: str) -> int:
    letter = letter.strip().upper()
    index = 0
    for char in letter:
        if char not in string.ascii_uppercase:
            raise ValueError(f"Invalid column letter: {letter}")
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index - 1


def read_excel(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        return pd.read_excel(path, engine="openpyxl", header=0)
    if suffix == ".xls":
        try:
            return pd.read_excel(path, engine="xlrd", header=0)
        except ImportError as exc:
            raise ImportError(
                "Reading .xls files requires xlrd. Install with: pip install xlrd"
            ) from exc
        except Exception as exc:
            raise ValueError(
                f"Failed to read .xls file '{path.name}'. "
                "Ensure the file is a valid Excel 97-2003 workbook. "
                f"Details: {exc}"
            ) from exc
    raise ValueError(
        f"Unsupported Excel format '{suffix}'. Supported formats: .xlsx, .xls"
    )


def get_cell_by_column_letter(df: pd.DataFrame, letter: str):
    index = column_letter_to_index(letter)
    if index >= len(df.columns):
        return None
    return df.iloc[:, index]


def row_to_context(df: pd.DataFrame, row_index: int, config: dict) -> dict:
    """Build template context from one Excel row (0-based data row index)."""
    if row_index < 0 or row_index >= len(df):
        raise IndexError(f"Row index {row_index} out of range (0-{len(df) - 1})")

    context: dict = {}
    for var_name, column_letter in config["excel_columns"].items():
        value = get_cell_by_column_letter(df, column_letter).iloc[row_index]
        if pd.isna(value):
            context[var_name] = ""
        elif isinstance(value, float) and value == int(value):
            context[var_name] = int(value)
        else:
            context[var_name] = value

    for target, formula in config.get("computed_fields", {}).items():
        text = formula
        for key, val in context.items():
            text = text.replace("{" + key + "}", str(val))
        context[target] = text.strip()

    for flag_name, expression in config.get("conditions", {}).items():
        context[flag_name] = _evaluate_condition(expression, context)

    return context


def _evaluate_condition(expression: str, context: dict) -> bool:
    expr = expression.strip()
    for key, value in context.items():
        if isinstance(value, str):
            replacement = repr(value)
        else:
            replacement = str(value)
        expr = expr.replace(key, replacement)

    # Safe eval for simple comparisons only
    allowed = set("0123456789.<>!=+-eE'\"() ")
    if not all(c in allowed or c.isalpha() for c in expr.replace(" ", "")):
        raise ValueError(f"Unsupported condition expression: {expression}")

    try:
        return bool(eval(expr, {"__builtins__": {}}, {}))
    except Exception as exc:
        raise ValueError(f"Failed to evaluate condition '{expression}': {exc}") from exc


def describe_excel_support() -> str:
    return (
        "Supported formats: .xlsx (openpyxl), .xls (xlrd). "
        "For legacy .xls, xlrd 2.x reads Excel 97-2003 only (not .xlsb)."
    )
