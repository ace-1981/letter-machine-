"""Verify that template JSON actually controls application behavior."""

from __future__ import annotations

import copy
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

import pandas as pd

from src.config_loader import get_template_dir, load_template_config
from src.letter_generator import generate_single_letter
from src.signature_field import verify_signature_field
from src.validator import validate_all

CONFIG_PATH = ROOT / "templates" / "תחשיב זכויות אישי.json"
EXCEL_PATH = ROOT / "samples" / "sample_data.xlsx"
OUT_DIR = ROOT / "samples" / "json_influence_test"
DEATH_PHRASE = "תוספת בגין מקרה פטירה"


def main() -> int:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    original = load_template_config(CONFIG_PATH)
    backup = copy.deepcopy(original)
    results: list[dict] = []

    try:
        results.append(_test_filename_pattern(backup))
        results.append(_test_full_name_mapping(backup))
        results.append(_test_death_condition(backup))
        results.append(_test_required_fields(backup))
        results.append(_test_signature_field_name(backup))
    finally:
        _write_json(CONFIG_PATH, backup)
        restored = load_template_config(CONFIG_PATH)
        if restored != backup:
            raise RuntimeError("JSON was not restored to original values")

    report_path = ROOT.parent / "cursor" / "דוח-JSON-Influence-Test.md"
    _write_report(report_path, results, restored_ok=True)
    print(f"\nReport: {report_path}")

    failed = [r for r in results if not r["passed"]]
    return 0 if not failed else 1


def _write_json(path: Path, data: dict) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def _pdf_text(pdf_path: Path) -> str:
    import fitz

    with fitz.open(str(pdf_path)) as doc:
        return "\n".join(page.get_text() for page in doc)


def _preview(config: dict, excel: Path, row_index: int = 0) -> dict:
    temp_config = CONFIG_PATH.parent / "_json_influence_test.json"
    _write_json(temp_config, config)
    try:
        return generate_single_letter(
            excel,
            temp_config,
            OUT_DIR,
            row_index=row_index,
            pdf_preferred="word",
        )
    finally:
        temp_config.unlink(missing_ok=True)


def _test_filename_pattern(backup: dict) -> dict:
    name = "1. template_name / output filename"
    changed = copy.deepcopy(backup)
    changed["template_name"] = "תחשיב זכויות אישי TEST"
    _write_json(CONFIG_PATH, changed)

    result = _preview(changed, EXCEL_PATH)
    pdf_name = result["pdf"].name
    passed = "TEST" in pdf_name and pdf_name.endswith(".pdf")

    _write_json(CONFIG_PATH, backup)
    return {
        "test": name,
        "json_change": 'template_name → "תחשיב זכויות אישי TEST"',
        "output": f"PDF filename: {pdf_name}",
        "passed": passed,
        "restored": load_template_config(CONFIG_PATH)["template_name"] == backup["template_name"],
    }


def _test_full_name_mapping(backup: dict) -> dict:
    name = "2. FULL_NAME computed field"
    changed = copy.deepcopy(backup)
    changed["computed_fields"]["FULL_NAME"] = "{FIRST_NAME} {LAST_NAME}"
    _write_json(CONFIG_PATH, changed)

    result = _preview(changed, EXCEL_PATH)
    text = _pdf_text(result["pdf"])
    passed = "ישראל כהן" in text and "כהן ישראל" not in text

    _write_json(CONFIG_PATH, backup)
    return {
        "test": name,
        "json_change": 'computed_fields.FULL_NAME → "{FIRST_NAME} {LAST_NAME}"',
        "output": "PDF contains 'ישראל כהן' and not 'כהן ישראל'",
        "passed": passed,
        "restored": load_template_config(CONFIG_PATH)["computed_fields"] == backup["computed_fields"],
    }


def _test_death_condition(backup: dict) -> dict:
    name = "3. show_DEATH_SECTION condition"
    excel_l0 = OUT_DIR / "row_l_zero.xlsx"
    _create_l_zero_excel(excel_l0)

    changed = copy.deepcopy(backup)

    # Baseline: L > 0, row with L=0 → section hidden
    _write_json(CONFIG_PATH, changed)
    baseline = _preview(changed, excel_l0, row_index=0)
    hidden_with_gt = DEATH_PHRASE not in _pdf_text(baseline["pdf"])

    # Changed: L >= 0, same row → section visible
    changed["conditions"]["show_DEATH_SECTION"] = "L >= 0"
    _write_json(CONFIG_PATH, changed)
    modified = _preview(changed, excel_l0, row_index=0)
    visible_with_gte = DEATH_PHRASE in _pdf_text(modified["pdf"])

    _write_json(CONFIG_PATH, backup)
    passed = hidden_with_gt and visible_with_gte
    return {
        "test": name,
        "json_change": 'conditions.show_DEATH_SECTION: "L > 0" → "L >= 0" (row L=0)',
        "output": (
            f"L>0 + L=0 row: section hidden={hidden_with_gt}; "
            f"L>=0 + L=0 row: section visible={visible_with_gte}"
        ),
        "passed": passed,
        "restored": load_template_config(CONFIG_PATH)["conditions"]["show_DEATH_SECTION"] == "L > 0",
    }


def _test_required_fields(backup: dict) -> dict:
    name = "4. required_excel_columns validation"
    changed = copy.deepcopy(backup)
    changed["validation"]["required_excel_columns"] = list(
        changed["validation"]["required_excel_columns"]
    ) + ["Z"]
    _write_json(CONFIG_PATH, changed)

    config = load_template_config(CONFIG_PATH)
    template_docx = get_template_dir(CONFIG_PATH) / config["template_file"]
    validation = validate_all(EXCEL_PATH, config, template_docx, OUT_DIR, "word")
    errors = validation.errors
    passed = (not validation.ok) and any("column Z" in e for e in errors)

    _write_json(CONFIG_PATH, backup)
    return {
        "test": name,
        "json_change": 'validation.required_excel_columns + "Z"',
        "output": "; ".join(errors) if errors else "(no errors)",
        "passed": passed,
        "restored": "Z" not in load_template_config(CONFIG_PATH)["validation"]["required_excel_columns"],
    }


def _test_signature_field_name(backup: dict) -> dict:
    name = "5. signature_field.field_name"
    changed = copy.deepcopy(backup)
    changed["signature_field"]["field_name"] = "TestSignature"
    _write_json(CONFIG_PATH, changed)

    result = _preview(changed, EXCEL_PATH)
    verification = verify_signature_field(result["pdf"], "TestSignature")
    old_field = verify_signature_field(result["pdf"], "MemberSignature")
    passed = verification.get("found") and verification.get("is_signature_field") and not old_field.get("found")

    _write_json(CONFIG_PATH, backup)
    return {
        "test": name,
        "json_change": 'signature_field.field_name → "TestSignature"',
        "output": (
            f"TestSignature found={verification.get('found')}, "
            f"MemberSignature found={old_field.get('found')}"
        ),
        "passed": passed,
        "restored": load_template_config(CONFIG_PATH)["signature_field"]["field_name"] == "MemberSignature",
    }


def _create_l_zero_excel(path: Path) -> None:
    data = {
        "A": [""],
        "B": [""],
        "C": [99999],
        "D": [""],
        "E": ["לוי"],
        "F": ["דנה"],
        "G": [""],
        "H": [10],
        "I": [5000],
        "J": [50000],
        "K": [""],
        "L": [0],
        "M": [55000],
        "N": [""],
        "O": [0],
        "P": [55000],
        "Q": [""],
        "R": [""],
        "S": ["99-999-999999"],
        "T": [""],
    }
    pd.DataFrame(data).to_excel(path, index=False, engine="openpyxl")


def _write_report(path: Path, results: list[dict], restored_ok: bool) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    all_passed = all(r["passed"] for r in results)
    all_restored = all(r["restored"] for r in results) and restored_ok

    lines = [
        "# דוח JSON Influence Test",
        "",
        f"**קובץ נבדק:** `LetterGenerator/templates/תחשיב זכויות אישי.json`",
        "",
        f"**תוצאה כוללת:** {'עבר' if all_passed and all_restored else 'נכשל'}",
        "",
        "## בדיקות",
        "",
    ]

    for r in results:
        status = "✅" if r["passed"] else "❌"
        restore = "✅" if r["restored"] else "❌"
        lines.extend(
            [
                f"### {r['test']} {status}",
                "",
                f"- **שינוי ב-JSON:** {r['json_change']}",
                f"- **פלט:** {r['output']}",
                f"- **הוחזר למקור:** {restore}",
                "",
            ]
        )

    lines.extend(
        [
            "## מסקנה",
            "",
        ]
    )

    if all_passed and all_restored:
        lines.append(
            "קובץ ה-JSON **שולט בפועל** על: שם קובץ הפלט, מיפוי שם מלא, תנאי הצגת מקטעים, "
            "ולידציית שדות חובה, ושם שדה החתימה. כל הערכים הוחזרו למצב המקורי."
        )
    else:
        lines.append("נמצאו כשלים — יש לבדוק את הסעיפים שסומנו ב-❌.")

    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


if __name__ == "__main__":
    raise SystemExit(main())
