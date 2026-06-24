"""Run all pre-GUI verification tests and print a report."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

from scripts.create_template import create_template
from scripts.create_test_data import (
    create_batch_excel,
    create_missing_column_excel,
    create_xls_from_xlsx,
)
from scripts.create_sample_excel import create_sample_excel
from src.config_loader import load_template_config
from src.excel_reader import describe_excel_support, read_excel
from src.letter_generator import generate_letters, generate_single_letter
from src.signature_field import verify_signature_field
from src.validator import validate_all

SAMPLES = ROOT / "samples"
TESTS = SAMPLES / "tests"
CONFIG = ROOT / "templates" / "תחשיב זכויות אישי.json"
TEMPLATE = ROOT / "templates" / "תחשיב זכויות אישי.docx"


def _header(title: str) -> None:
    print(f"\n{'=' * 60}")
    print(title)
    print("=" * 60)


def test_xls() -> dict:
    _header("TEST 2: xls support")
    xlsx = SAMPLES / "sample_data.xlsx"
    xls = SAMPLES / "sample_data.xls"
    create_sample_excel(xlsx)
    create_xls_from_xlsx(xlsx, xls)
    df = read_excel(xls)
    result = {
        "status": "PASS",
        "message": describe_excel_support(),
        "rows": len(df),
        "columns": len(df.columns),
        "file": str(xls),
    }
    print(result["message"])
    print(f"xls read OK: {result['rows']} rows, {result['columns']} columns")
    print(f"File: {xls}")
    return result


def test_negative_validate() -> dict:
    _header("TEST 3: negative validate")
    config = load_template_config(CONFIG)
    if not TEMPLATE.exists():
        create_template(TEMPLATE)

    cases = []

    missing_col = TESTS / "missing_column.xlsx"
    create_missing_column_excel(missing_col)
    r1 = validate_all(missing_col, config, TEMPLATE, TESTS / "out_neg", "word")
    cases.append(
        {
            "case": "עמודה חסרה",
            "ok": not r1.ok,
            "errors": r1.errors,
        }
    )
    print("Case: missing column")
    print("  ok:", r1.ok)
    for e in r1.errors:
        print("  -", e)

    batch = TESTS / "batch_15.xlsx"
    create_batch_excel(batch)
    r2 = validate_all(batch, config, TEMPLATE, TESTS / "out_neg", "word")
    cases.append({"case": "batch תקין structurally", "ok": r2.ok, "errors": r2.errors})

    from src.excel_reader import row_to_context
    from src.validator import validate_row

    df = read_excel(batch)
    row_ctx = row_to_context(df, 7, config)
    row_errs = validate_row(row_ctx, config, 9)
    cases.append(
        {
            "case": "שורה עם חשבון בנק חסר",
            "ok": len(row_errs) > 0,
            "errors": row_errs,
        }
    )
    print("Case: empty bank account row 9")
    for e in row_errs:
        print("  -", e)

    row_ctx2 = row_to_context(df, 11, config)
    row_errs2 = validate_row(row_ctx2, config, 13)
    cases.append(
        {
            "case": "שורה עם סכום לא תקין",
            "ok": len(row_errs2) > 0,
            "errors": row_errs2,
        }
    )
    print("Case: invalid amount row 13")
    for e in row_errs2:
        print("  -", e)

    passed = all(c["ok"] for c in cases)
    return {"status": "PASS" if passed else "FAIL", "cases": cases}


def test_batch() -> dict:
    _header("TEST 4: batch 15 rows")
    batch = TESTS / "batch_15.xlsx"
    create_batch_excel(batch)
    if not TEMPLATE.exists():
        create_template(TEMPLATE)

    out = TESTS / "batch_output"
    if out.exists():
        for f in out.glob("*"):
            f.unlink()

    result = generate_letters(batch, CONFIG, out, pdf_preferred="word")
    pdfs = list(out.glob("*.pdf"))
    csv_path = out / "errors_report.csv"

    info = {
        "status": "PASS" if result.success == 13 and len(result.errors) == 2 else "FAIL",
        "total": result.total,
        "success": result.success,
        "errors_count": len(result.errors),
        "pdfs_created": len(pdfs),
        "errors_report_exists": csv_path.exists(),
        "errors": result.errors,
    }
    print(f"Total: {info['total']}, Success: {info['success']}, Errors: {info['errors_count']}")
    print(f"PDFs created: {info['pdfs_created']}")
    print(f"errors_report.csv: {csv_path.exists()}")
    for err in result.errors:
        print(f"  - row {err['excel_row']}: {err['error']}")
    return info


def test_signature() -> dict:
    _header("TEST 1: signature field")
    out = SAMPLES / "poc_output"
    out.mkdir(parents=True, exist_ok=True)

    if not TEMPLATE.exists():
        create_template(TEMPLATE)
    xlsx = SAMPLES / "sample_data.xlsx"
    create_sample_excel(xlsx)

    gen = generate_single_letter(xlsx, CONFIG, out, row_index=0, pdf_preferred="word")
    config = load_template_config(CONFIG)
    field_name = config["signature_field"]["field_name"]
    verification = verify_signature_field(gen["pdf"], field_name)

    # Try to locate PDF viewers
    viewers = []
    candidates = [
        Path(r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"),
        Path(r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"),
        Path(r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"),
        Path(r"C:\Program Files (x86)\Foxit Software\Foxit PDF Reader\FoxitPDFReader.exe"),
    ]
    for c in candidates:
        if c.exists():
            viewers.append(str(c))

    opened_with = None
    if viewers:
        try:
            subprocess.Popen([viewers[0], str(gen["pdf"])])
            opened_with = viewers[0]
        except Exception:
            pass

    info = {
        "status": "PASS" if verification.get("is_signature_field") else "FAIL",
        "verification": verification,
        "pdf": str(gen["pdf"]),
        "viewers_found": viewers,
        "opened_with": opened_with,
        "manual_step": (
            "פתח את ה-PDF ב-Adobe Reader או Foxit → Fill & Sign → חתום ב-MemberSignature"
        ),
    }
    print(json.dumps(verification, ensure_ascii=False, indent=2))
    if viewers:
        print("PDF viewers found:", viewers)
        if opened_with:
            print("Opened PDF with:", opened_with)
    else:
        print("No Adobe/Foxit found — open PDF manually for signing test")
    print("Manual:", info["manual_step"])
    return info


def main() -> int:
    print("Letter Generator — Pre-GUI Test Suite")
    report = {
        "signature": test_signature(),
        "xls": test_xls(),
        "negative_validate": test_negative_validate(),
        "batch": test_batch(),
    }

    all_pass = all(
        report[k].get("status") == "PASS"
        for k in ("signature", "xls", "negative_validate", "batch")
    )

    _header("SUMMARY")
    for name, data in report.items():
        print(f"  {name}: {data.get('status')}")

    report_path = ROOT.parent / "cursor" / "דוח-בדיקות-לפני-GUI.md"
    _write_report(report_path, report, all_pass)
    print(f"\nReport written: {report_path}")

    return 0 if all_pass else 1


def _write_report(path: Path, report: dict, all_pass: bool) -> None:
    lines = [
        "# דוח בדיקות לפני GUI",
        "",
        f"**תוצאה כוללת:** {'עבר' if all_pass else 'נכשל'}",
        "",
        "## 1. שדה חתימה",
        "",
        f"- סטטוס אוטומטי: **{report['signature']['status']}**",
        f"- PDF: `{report['signature']['pdf']}`",
        "",
        "```json",
        json.dumps(report["signature"]["verification"], ensure_ascii=False, indent=2),
        "```",
        "",
        "**בדיקה ידנית:** " + report["signature"]["manual_step"],
        "",
    ]
    if report["signature"].get("opened_with"):
        lines.append(f"- PDF נפתח עם: `{report['signature']['opened_with']}`")
    lines.extend(
        [
            "",
            "## 2. תמיכת xls",
            "",
            f"- סטטוס: **{report['xls']['status']}**",
            f"- {report['xls']['message']}",
            f"- קובץ: `{report['xls']['file']}`",
            f"- שורות: {report['xls']['rows']}, עמודות: {report['xls']['columns']}",
            "",
            "## 3. validate שלילי",
            "",
        ]
    )
    for case in report["negative_validate"]["cases"]:
        lines.append(f"### {case['case']}")
        lines.append(f"- תוצאה: {'עבר' if case['ok'] else 'נכשל'}")
        for e in case.get("errors", []):
            lines.append(f"  - {e}")
        lines.append("")

    b = report["batch"]
    lines.extend(
        [
            "## 4. batch (15 שורות)",
            "",
            f"- סטטוס: **{b['status']}**",
            f"- סה\"כ: {b['total']}, הצליחו: {b['success']}, שגיאות: {b['errors_count']}",
            f"- PDFs שנוצרו: {b['pdfs_created']}",
            f"- errors_report.csv: {'כן' if b['errors_report_exists'] else 'לא'}",
            "",
        ]
    )
    for err in b.get("errors", []):
        lines.append(f"- שורה {err['excel_row']}: {err['error']}")

    path.write_text("\n".join(lines), encoding="utf-8")


if __name__ == "__main__":
    raise SystemExit(main())
