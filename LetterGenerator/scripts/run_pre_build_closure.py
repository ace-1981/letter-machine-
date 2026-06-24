"""Final pre-build closure checks before app.exe packaging."""

from __future__ import annotations

import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

import pandas as pd

from scripts.create_sample_excel import create_sample_excel
from scripts.create_template import create_template
from scripts.create_test_data import create_batch_excel
from src.config_loader import load_template_config
from src.letter_generator import generate_letters, generate_single_letter
from src.pdf_converter import PdfConverterFactory
from src.signature_field import verify_signature_field

CONFIG = ROOT / "templates" / "תחשיב זכויות אישי.json"
TEMPLATE = ROOT / "templates" / "תחשיב זכויות אישי.docx"
BACKUP_DIR = ROOT / "templates" / "backup_pre_build"
PREVIEW_OUT = ROOT / "samples" / "pre_build_output"
BATCH_OUT = ROOT / "samples" / "pre_build_batch"
BATCH_EXCEL = ROOT / "samples" / "pre_build_batch_20.xlsx"

CANONICAL_JSON = {
    "template_name": "תחשיב זכויות אישי",
    "template_file": "תחשיב זכויות אישי.docx",
    "excel_columns": {
        "MEMBER_CODE": "C",
        "LAST_NAME": "E",
        "FIRST_NAME": "F",
        "H": "H",
        "I": "I",
        "J": "J",
        "L": "L",
        "M": "M",
        "O": "O",
        "P": "P",
        "R": "R",
        "BANK_ACCOUNT": "S",
        "T": "T",
    },
    "computed_fields": {"FULL_NAME": "{LAST_NAME} {FIRST_NAME}"},
    "conditions": {
        "show_DEATH_SECTION": "L > 0",
        "show_WORK_GRANT_SECTION": "O > 0",
        "show_NEW_MEMBER_SECTION": "T == 'נקלט.ת'",
        "show_BUILDING_DEBT_SECTION": "R != ''",
    },
    "output_filename": {
        "pattern": "{MEMBER_CODE} {LAST_NAME} {FIRST_NAME} {template_name}.pdf"
    },
    "signature_field": {
        "field_name": "MemberSignature",
        "page": "last",
        "width": 180,
        "height": 42,
        "margin_bottom": 118,
        "margin_left": 108,
        "anchor_text": "חתימה (שדה לחתימה דיגיטלית)",
    },
    "validation": {
        "required_excel_columns": ["C", "E", "F", "H", "I", "J", "L", "M", "O", "P", "S"],
        "required_template_variables": [
            "FULL_NAME",
            "TODAY",
            "H",
            "I",
            "J",
            "L",
            "M",
            "O",
            "P",
            "BANK_ACCOUNT",
        ],
        "required_row_fields": ["MEMBER_CODE", "LAST_NAME", "FIRST_NAME", "BANK_ACCOUNT", "P"],
        "numeric_row_fields": ["H", "I", "J", "L", "M", "O", "P"],
    },
}


def main() -> int:
    report: dict = {"timestamp": datetime.now().isoformat(timespec="seconds"), "checks": {}}

    report["checks"]["json_restored"] = _check_json_restored()
    if not TEMPLATE.exists():
        create_template(TEMPLATE)
    report["checks"]["backup"] = _create_backups()
    report["checks"]["preview"] = _run_preview()
    report["checks"]["batch"] = _run_batch()
    report["checks"]["signature"] = _check_signature(report["checks"]["preview"]["pdf"])
    report["checks"]["json_influence"] = {"status": "PASS", "note": "אושר בנפרד (דוח JSON Influence Test)"}
    report["checks"]["pdf_visual"] = _check_pdf_basic(report["checks"]["preview"]["pdf"])
    report["checks"]["gui"] = _check_gui()
    report["checks"]["errors_report"] = report["checks"]["batch"].get("errors_report", {})
    report["checks"]["json_no_letter_text"] = _check_json_no_letter_text()
    report["checks"]["converter"] = _check_converters()

    report_path = ROOT.parent / "cursor" / "דוח-Final-Pre-Build.md"
    _write_report(report_path, report)
    print(f"Report: {report_path}")
    return 0 if _all_pass(report) else 1


def _check_json_restored() -> dict:
    current = load_template_config(CONFIG)
    ok = current == CANONICAL_JSON
    return {
        "status": "PASS" if ok else "FAIL",
        "matches_canonical": ok,
        "template_name": current.get("template_name"),
        "signature_field": current.get("signature_field", {}).get("field_name"),
        "death_condition": current.get("conditions", {}).get("show_DEATH_SECTION"),
    }


def _create_backups() -> dict:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    json_dst = BACKUP_DIR / f"תחשיב זכויות אישי_{stamp}.json"
    docx_dst = BACKUP_DIR / f"תחשיב זכויות אישי_{stamp}.docx"
    shutil.copy2(CONFIG, json_dst)
    shutil.copy2(TEMPLATE, docx_dst)
    latest_json = BACKUP_DIR / "תחשיב זכויות אישי.json"
    latest_docx = BACKUP_DIR / "תחשיב זכויות אישי.docx"
    shutil.copy2(CONFIG, latest_json)
    shutil.copy2(TEMPLATE, latest_docx)
    return {
        "status": "PASS",
        "json": str(json_dst),
        "docx": str(docx_dst),
        "latest_json": str(latest_json),
        "latest_docx": str(latest_docx),
    }


def _run_preview() -> dict:
    PREVIEW_OUT.mkdir(parents=True, exist_ok=True)
    excel = ROOT / "samples" / "sample_data.xlsx"
    create_sample_excel(excel)
    result = generate_single_letter(excel, CONFIG, PREVIEW_OUT, row_index=0, pdf_preferred="word")
    name = result["pdf"].name
    ok = "TEST" not in name and name.endswith(".pdf")
    return {
        "status": "PASS" if ok else "FAIL",
        "pdf": str(result["pdf"]),
        "filename": name,
        "no_test_suffix": ok,
    }


def _create_batch_20_excel(path: Path) -> Path:
    create_batch_excel(path)
    df = pd.read_excel(path, engine="openpyxl")
    if len(df) >= 20:
        return path
    extra = []
    families = ["כהן", "לוי", "מזרחי", "אברהם", "דוד"]
    first_names = ["ישראל", "דנה", "יוסף", "מירי", "אבי"]
    for i in range(15, 20):
        code = 10001 + i
        extra.append(
            {
                "A": "",
                "B": "",
                "C": code,
                "D": "",
                "E": families[i % len(families)],
                "F": first_names[i % len(first_names)],
                "G": "",
                "H": 12,
                "I": 5000,
                "J": 60000,
                "K": "",
                "L": 0,
                "M": 65000,
                "N": "",
                "O": 0,
                "P": 65000 + i,
                "Q": "",
                "R": "",
                "S": f"12-345-{679000 + i:06d}",
                "T": "פעיל",
            }
        )
    df = pd.concat([df, pd.DataFrame(extra)], ignore_index=True)
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(path, index=False, engine="openpyxl")
    return path


def _run_batch() -> dict:
    BATCH_OUT.mkdir(parents=True, exist_ok=True)
    _create_batch_20_excel(BATCH_EXCEL)
    result = generate_letters(BATCH_EXCEL, CONFIG, BATCH_OUT, pdf_preferred="word")
    csv_path = BATCH_OUT / "errors_report.csv"
    csv_ok = csv_path.exists()
    return {
        "status": "PASS" if result.total == 20 and csv_ok else "FAIL",
        "total": result.total,
        "success": result.success,
        "errors": len(result.errors),
        "errors_report": {
            "status": "PASS" if csv_ok else "FAIL",
            "path": str(csv_path),
            "exists": csv_ok,
            "error_rows": [e.get("excel_row") for e in result.errors],
        },
    }


def _check_signature(pdf_path: str) -> dict:
    v = verify_signature_field(Path(pdf_path), "MemberSignature")
    ok = v.get("interactive") and v.get("is_signature_field")
    return {
        "status": "PASS" if ok else "FAIL",
        "field_name": v.get("field_name"),
        "interactive": v.get("interactive"),
        "locked": v.get("locked"),
    }


def _check_pdf_basic(pdf_path: str) -> dict:
    import fitz

    with fitz.open(pdf_path) as doc:
        pages = len(doc)
        text = doc[0].get_text()
    ok = pages == 1 and "תחשיב זכויות אישי" in text and "כהן ישראל" in text
    return {
        "status": "PASS" if ok else "FAIL",
        "pages": pages,
        "has_title": "תחשיב זכויות אישי" in text,
        "has_name": "כהן ישראל" in text,
    }


def _check_gui() -> dict:
    try:
        from PySide6.QtWidgets import QApplication

        from src.ui.main_window import MainWindow

        app = QApplication.instance() or QApplication([])
        window = MainWindow()
        has_preview = hasattr(window, "preview_btn")
        has_batch = hasattr(window, "generate_btn")
        window.close()
        ok = has_preview and has_batch
        return {"status": "PASS" if ok else "FAIL", "preview_btn": has_preview, "generate_btn": has_batch}
    except Exception as exc:
        return {"status": "FAIL", "error": str(exc)}


def _check_json_no_letter_text() -> dict:
    config = load_template_config(CONFIG)
    allowed_paths = {
        "template_name",
        "template_file",
        "conditions.show_NEW_MEMBER_SECTION",
        "signature_field.anchor_text",
    }
    findings = []
    for key, value in _walk(config):
        if isinstance(value, str) and _contains_hebrew(value) and key not in allowed_paths:
            findings.append(f"{key}: {value[:60]}")
    ok = len(findings) == 0
    return {
        "status": "PASS" if ok else "FAIL",
        "note": "מותר: template_name, ערך תנאי מ-Excel, anchor_text טכני",
        "findings": findings,
    }


def _walk(obj, prefix=""):
    if isinstance(obj, dict):
        for k, v in obj.items():
            yield from _walk(v, f"{prefix}.{k}" if prefix else k)
    else:
        yield prefix, obj


def _contains_hebrew(text: str) -> bool:
    return any("\u0590" <= ch <= "\u05ff" for ch in text)


def _check_converters() -> dict:
    word = PdfConverterFactory.create("word")
    libre = None
    libre_ok = False
    try:
        from src.pdf_converter import LibreOfficePdfConverter

        libre = LibreOfficePdfConverter()
        libre_ok = libre.is_available()
    except Exception:
        pass
    return {
        "status": "PASS" if word.is_available() else "FAIL",
        "word": word.availability_message(),
        "libreoffice": libre.availability_message() if libre else "not checked",
        "active_converter": word.name,
        "libreoffice_available": libre_ok,
    }


def _all_pass(report: dict) -> bool:
    checks = report["checks"]
    keys = [
        "json_restored",
        "backup",
        "preview",
        "batch",
        "signature",
        "pdf_visual",
        "gui",
        "json_no_letter_text",
        "converter",
    ]
    for key in keys:
        if checks[key].get("status") != "PASS":
            return False
    if checks["errors_report"].get("status") != "PASS":
        return False
    return True


def _write_report(path: Path, report: dict) -> None:
    c = report["checks"]
    all_ok = _all_pass(report)
    lines = [
        "# דוח Final Pre-Build",
        "",
        f"**תאריך:** {report['timestamp']}",
        f"**מוכן ל-app.exe:** {'כן' if all_ok else 'לא — יש לתקן כשלים'}",
        "",
        "## סיכום בדיקות",
        "",
        "| בדיקה | סטטוס | פירוט |",
        "|--------|--------|--------|",
        f"| JSON במצב מקורי | {c['json_restored']['status']} | template_name={c['json_restored']['template_name']} |",
        f"| גיבוי docx/json | {c['backup']['status']} | `{c['backup']['latest_json']}` |",
        f"| Preview רגיל | {c['preview']['status']} | `{c['preview']['filename']}` |",
        f"| Batch 20 שורות | {c['batch']['status']} | {c['batch']['success']}/{c['batch']['total']} הצלחות, {c['batch']['errors']} שגיאות |",
        f"| errors_report.csv | {c['errors_report']['status']} | `{c['errors_report']['path']}` |",
        f"| חתימה דיגיטלית | {c['signature']['status']} | interactive={c['signature']['interactive']} |",
        f"| JSON משפיע | PASS | אושר בבדיקה נפרדת |",
        f"| PDF תקין | {c['pdf_visual']['status']} | {c['pdf_visual']['pages']} עמודים |",
        f"| GUI | {c['gui']['status']} | preview + batch buttons |",
        f"| אין טקסטי מכתב ב-JSON | {c['json_no_letter_text']['status']} | {c['json_no_letter_text']['note']} |",
        f"| PDF converter | {c['converter']['status']} | {c['converter']['active_converter']} |",
        "",
        "## גיבויים",
        "",
        f"- JSON: `{c['backup']['json']}`",
        f"- DOCX: `{c['backup']['docx']}`",
        "",
        "## Preview",
        "",
        f"- PDF: `{c['preview']['pdf']}`",
        "",
        "## Batch",
        "",
        f"- Excel: `{BATCH_EXCEL}`",
        f"- Output: `{BATCH_OUT}`",
        f"- שורות שגיאה: {c['batch']['errors_report'].get('error_rows', [])}",
        "",
    ]
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines), encoding="utf-8")


if __name__ == "__main__":
    raise SystemExit(main())
