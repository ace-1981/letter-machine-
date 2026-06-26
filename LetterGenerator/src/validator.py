"""Pre-production validation."""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from src.excel_reader import column_letter_to_index, read_excel
from src.pdf_converter import PdfConverterFactory


@dataclass
class ValidationResult:
    ok: bool
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def add_error(self, message: str) -> None:
        self.ok = False
        self.errors.append(message)

    def add_warning(self, message: str) -> None:
        self.warnings.append(message)


def validate_all(
    excel_path: Path,
    config: dict,
    template_docx: Path,
    output_dir: Path,
    pdf_preferred: str | None = None,
    output_format: str = "pdf",
) -> ValidationResult:
    from src.output_format import OUTPUT_PDF, normalize_output_format

    result = ValidationResult(ok=True)
    _validate_excel_columns(excel_path, config, result)
    _validate_template_variables(template_docx, config, result)
    _validate_conditions(config, result)
    _validate_output_dir(output_dir, result)
    if normalize_output_format(output_format) == OUTPUT_PDF:
        _validate_pdf_converter(pdf_preferred, result)
    return result


def _validate_excel_columns(excel_path: Path, config: dict, result: ValidationResult) -> None:
    try:
        df = read_excel(excel_path)
    except Exception as exc:
        result.add_error(f"Cannot read Excel file: {exc}")
        return

    if len(df) == 0:
        result.add_error("Excel file contains no data rows.")

    required = config.get("validation", {}).get("required_excel_columns", [])
    for letter in required:
        try:
            index = column_letter_to_index(letter)
        except ValueError as exc:
            result.add_error(str(exc))
            continue
        if index >= len(df.columns):
            result.add_error(f"Required Excel column {letter} is missing (file has {len(df.columns)} columns).")


def _validate_template_variables(docx_path: Path, config: dict, result: ValidationResult) -> None:
    if not docx_path.exists():
        result.add_error(f"Template DOCX not found: {docx_path}")
        return

    try:
        found_vars = _extract_docx_variables(docx_path)
    except Exception as exc:
        result.add_error(f"Cannot read template DOCX: {exc}")
        return

    required = set(config.get("validation", {}).get("required_template_variables", []))
    condition_flags = set(config.get("conditions", {}).keys())
    allowed_extras = {
        "show_DEATH_SECTION",
        "show_WORK_GRANT_SECTION",
        "show_NEW_MEMBER_SECTION",
        "show_BUILDING_DEBT_SECTION",
        "show_NOTES_SECTION",
    }
    allowed_extras.update(condition_flags)
    allowed_extras.add("calc_table")
    allowed_extras.add("table_rows")

    missing = required - found_vars
    for var in sorted(missing):
        result.add_error(f"Template missing required variable: {{{{{var}}}}}")

    unknown = found_vars - required - allowed_extras - {"TODAY"}
    for var in sorted(unknown):
        result.add_warning(f"Template contains variable not in validation list: {{{{{var}}}}}")


def _extract_docx_variables(docx_path: Path) -> set[str]:
    variables: set[str] = set()
    pattern = re.compile(r"\{\{?\s*([a-zA-Z_][a-zA-Z0-9_]*)\s*\}?\}")
    jinja_block = re.compile(r"\{%.*?%\}")

    with zipfile.ZipFile(docx_path, "r") as zf:
        xml_files = [n for n in zf.namelist() if n.startswith("word/") and n.endswith(".xml")]
        for name in xml_files:
            text = zf.read(name).decode("utf-8", errors="ignore")
            text = jinja_block.sub("", text)
            for match in pattern.finditer(text):
                variables.add(match.group(1))
    return variables


def _validate_conditions(config: dict, result: ValidationResult) -> None:
    from src.excel_reader import _evaluate_condition

    conditions = config.get("conditions", {})
    columns = set(config.get("excel_columns", {}).keys())
    sample_context = {key: 1 for key in columns}
    sample_context.update({k: "נקלט.ת" for k in columns if k == "T"})
    sample_context.update({k: "" for k in columns if k == "R"})

    for flag_name, expression in conditions.items():
        try:
            _evaluate_condition(expression, sample_context)
        except Exception as exc:
            result.add_error(f"Invalid condition '{flag_name}': {expression} — {exc}")


def _validate_output_dir(output_dir: Path, result: ValidationResult) -> None:
    if output_dir.exists():
        if not output_dir.is_dir():
            result.add_error(f"Output path exists but is not a directory: {output_dir}")
            return
    else:
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            result.add_error(f"Cannot create output directory: {exc}")
            return

    test_file = output_dir / ".write_test"
    try:
        test_file.write_text("ok", encoding="utf-8")
        test_file.unlink()
    except Exception as exc:
        result.add_error(f"Output directory is not writable: {exc}")


def _validate_pdf_converter(preferred: str | None, result: ValidationResult) -> None:
    try:
        converter = PdfConverterFactory.create(preferred)
        result.add_warning(f"PDF converter selected: {converter.name}")
    except Exception as exc:
        result.add_error(str(exc))


FIELD_LABELS = {
    "MEMBER_CODE": "קוד חבר",
    "LAST_NAME": "שם משפחה",
    "FIRST_NAME": "שם פרטי",
    "BANK_ACCOUNT": "חשבון בנק",
    "P": "סכום סופי",
    "H": "ותק",
    "I": "סכום קבוע",
    "J": "סכום לפי ותק",
    "L": "תוספת נפטרים",
    "M": "סהכ",
    "O": "חוב מענק",
}


def validate_row(context: dict, config: dict, excel_row_number: int) -> list[str]:
    """Return human-readable row errors (empty list = OK)."""
    validation = config.get("validation", {})
    errors: list[str] = []

    for field in validation.get("required_row_fields", []):
        value = context.get(field, "")
        if value is None or str(value).strip() == "":
            label = FIELD_LABELS.get(field, field)
            errors.append(f"שורה {excel_row_number} - {label} חסר")

    for field in validation.get("numeric_row_fields", []):
        value = context.get(field, "")
        if value is None or str(value).strip() == "":
            continue
        try:
            float(value)
        except (TypeError, ValueError):
            label = FIELD_LABELS.get(field, field)
            errors.append(f"שורה {excel_row_number} - {label} לא תקין")

    return errors
