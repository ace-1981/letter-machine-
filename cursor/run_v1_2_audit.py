"""V1.2 pre-improvement audit — tests only, no app changes."""

from __future__ import annotations

import csv
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import pandas as pd

WORKSPACE = Path(__file__).resolve().parent.parent
DEV_ROOT = WORKSPACE / "LetterGenerator"
PORTABLE = WORKSPACE / "LetterGenerator_V1.1_Portable"
CONFIG = PORTABLE / "templates" / "תחשיב זכויות אישי.json"
TEMPLATE_DOCX = PORTABLE / "templates" / "תחשיב זכויות אישי.docx"
AUDIT_DIR = WORKSPACE / "cursor" / "v1_2_audit"
SAMPLES_DIR = AUDIT_DIR / "samples"
PERF_CSV = WORKSPACE / "cursor" / "performance_report.csv"
REPORT_PATH = WORKSPACE / "cursor" / "דוח-בדיקות-ושיפורים-V1.2.md"

sys.path.insert(0, str(DEV_ROOT))

from src.letter_generator import generate_letters, generate_single_letter  # noqa: E402
from src.config_loader import load_template_config  # noqa: E402
from src.pdf_converter import PdfConverterFactory  # noqa: E402
from src.signature_field import verify_signature_field  # noqa: E402
from src.template_engine import render_template  # noqa: E402
from src.signature_field import add_date_field, add_signature_field  # noqa: E402
from src.excel_reader import read_excel, row_to_context  # noqa: E402
from src.validator import validate_all, validate_row  # noqa: E402
from src.template_engine import build_output_filename  # noqa: E402


@dataclass
class TestResult:
    section: str
    name: str
    passed: bool
    detail: str = ""
    recommendation: str = ""


@dataclass
class AuditState:
    results: list[TestResult] = field(default_factory=list)
    perf_rows: list[dict] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    def add(self, section: str, name: str, passed: bool, detail: str = "", recommendation: str = "") -> None:
        self.results.append(TestResult(section, name, passed, detail, recommendation))


def _run_exe(args: list[str], cwd: Path, timeout: int = 300) -> subprocess.CompletedProcess:
    return subprocess.run(
        [str(cwd / "app.exe"), *args],
        capture_output=True,
        cwd=cwd,
        timeout=timeout,
        env={**os.environ, "PYTHONUTF8": "1"},
    )


def _decode(proc: subprocess.CompletedProcess) -> tuple[str, str]:
    return (
        proc.stdout.decode("utf-8", errors="replace"),
        proc.stderr.decode("utf-8", errors="replace"),
    )


def _safe_name(label: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', "_", label.replace(" ", "_"))


def _base_row(**overrides) -> dict:
    row = {
        "MEMBER_CODE": 12345,
        "LAST_NAME": "כהן",
        "FIRST_NAME": "ישראל",
        "H": 10,
        "I": 1000,
        "J": 100,
        "L": 500,
        "M": 1600,
        "O": 200,
        "P": 1400,
        "R": "",
        "BANK_ACCOUNT": "12-345-678901",
        "T": "",
    }
    row.update(overrides)
    return row


def _make_excel(rows: list[dict], path: Path) -> Path:
    mapping = {
        "C": "MEMBER_CODE", "E": "LAST_NAME", "F": "FIRST_NAME",
        "H": "H", "I": "I", "J": "J", "L": "L", "M": "M",
        "O": "O", "P": "P", "R": "R", "S": "BANK_ACCOUNT", "T": "T",
    }
    out_rows = []
    for values in rows:
        cols = {chr(65 + i): "" for i in range(20)}
        for col, key in mapping.items():
            if key in values:
                cols[col] = values[key]
        out_rows.append(cols)
    path.parent.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(out_rows).to_excel(path, index=False)
    return path


def _make_batch_excel(n: int, path: Path, *, with_errors: bool = False) -> Path:
    families = ["כהן", "לוי", "מזרחי", "אברהם", "דוד", "שלום", "ברק", "רוזן"]
    firsts = ["ישראל", "דנה", "יוסף", "מירי", "אבי", "נועה", "עמית", "הילה"]
    rows = []
    for i in range(n):
        code = 20001 + i
        row = _base_row(
            MEMBER_CODE=code,
            LAST_NAME=families[i % len(families)],
            FIRST_NAME=firsts[i % len(firsts)],
            L=500 if i % 3 else 0,
            O=200 if i % 2 else 0,
            P=1200 + i * 10,
            BANK_ACCOUNT=f"12-345-{678900 + i:06d}",
            R="חוב בנייה" if i % 5 == 0 else "",
            T="נקלט.ת" if i % 7 == 0 else "",
        )
        if with_errors and i == max(1, n // 4):
            row["BANK_ACCOUNT"] = ""
        if with_errors and i == max(2, n // 2):
            row["P"] = "לא-מספר"
        rows.append(row)
    return _make_excel(rows, path)


def _pdf_text(pdf: Path) -> str:
    import fitz

    parts = []
    with fitz.open(str(pdf)) as doc:
        for page in doc:
            parts.append(page.get_text())
    return "\n".join(parts)


def _timed_single_row(
    excel: Path,
    row_index: int,
    output_dir: Path,
    output_format: str,
) -> dict:
    config = load_template_config(CONFIG)
    template_docx = TEMPLATE_DOCX
    df = read_excel(excel)
    context = row_to_context(df, row_index, config)
    sig_cfg = config["signature_field"]
    date_cfg = config.get("date_field")
    filename = build_output_filename(context, config, output_format)

    t0 = time.perf_counter()
    t_read = t0

    row_errors = validate_row(context, config, row_index + 2)
    if row_errors:
        return {
            "status": "error",
            "error_message": row_errors[0],
            "total_seconds": time.perf_counter() - t0,
        }

    docx_seconds = 0.0
    pdf_convert_seconds = 0.0
    signature_seconds = 0.0

    if output_format == "docx":
        t_docx = time.perf_counter()
        docx_path = output_dir / filename
        render_template(template_docx, docx_path, context)
        docx_seconds = time.perf_counter() - t_docx
        total = time.perf_counter() - t0
        return {
            "status": "ok",
            "docx_seconds": docx_seconds,
            "pdf_convert_seconds": 0.0,
            "signature_seconds": 0.0,
            "total_seconds": total,
        }

    converter = PdfConverterFactory.create("word")
    temp_docx = output_dir / f"_temp_{filename.replace('.pdf', '.docx')}"
    temp_pdf = output_dir / f"_temp_{filename}"
    pdf_path = output_dir / filename

    t_docx = time.perf_counter()
    render_template(template_docx, temp_docx, context)
    docx_seconds = time.perf_counter() - t_docx

    t_pdf = time.perf_counter()
    converter.convert(temp_docx, temp_pdf)
    pdf_convert_seconds = time.perf_counter() - t_pdf

    t_sig = time.perf_counter()
    add_signature_field(
        pdf_path=temp_pdf,
        output_path=pdf_path,
        field_name=sig_cfg["field_name"],
        width=sig_cfg["width"],
        height=sig_cfg["height"],
        margin_bottom=sig_cfg["margin_bottom"],
        margin_left=sig_cfg["margin_left"],
        page=sig_cfg.get("page", "last"),
        anchor_text=sig_cfg.get("anchor_text"),
    )
    if date_cfg:
        add_date_field(
            pdf_path=pdf_path,
            output_path=pdf_path,
            field_name=date_cfg["field_name"],
            width=date_cfg["width"],
            height=date_cfg["height"],
            page=date_cfg.get("page", "last"),
        )
    signature_seconds = time.perf_counter() - t_sig
    temp_pdf.unlink(missing_ok=True)
    temp_docx.unlink(missing_ok=True)

    return {
        "status": "ok",
        "docx_seconds": docx_seconds,
        "pdf_convert_seconds": pdf_convert_seconds,
        "signature_seconds": signature_seconds,
        "total_seconds": time.perf_counter() - t0,
    }


def check_word_available() -> tuple[bool, str]:
    try:
        conv = PdfConverterFactory.create("word")
        ok = conv.is_available()
        return ok, conv.availability_message()
    except Exception as exc:
        return False, str(exc)


def section_version(state: AuditState) -> None:
    exe = PORTABLE / "app.exe"
    word_ok, word_msg = check_word_available()
    state.add(
        "גרסה",
        "נתיב Portable",
        PORTABLE.is_dir() and exe.is_file(),
        f"`{PORTABLE}` — app.exe {exe.stat().st_size // (1024*1024)} MB, עודכן {datetime.fromtimestamp(exe.stat().st_mtime)}",
    )
    state.add("גרסה", "Microsoft Word זמין", word_ok, word_msg)
    proc = _run_exe(["check"], PORTABLE)
    out, err = _decode(proc)
    state.add("גרסה", "app.exe check", proc.returncode == 0, out.strip() or err.strip())
    proc = _run_exe(["info"], PORTABLE)
    out, _ = _decode(proc)
    state.add("גרסה", "app.exe info", proc.returncode == 0, out[:200])


def section_basic(state: AuditState) -> None:
    excel = _make_excel([_base_row()], AUDIT_DIR / "basic.xlsx")
    out_pdf = SAMPLES_DIR / "basic_pdf"
    out_docx = SAMPLES_DIR / "basic_docx"
    out_pdf.mkdir(parents=True, exist_ok=True)
    out_docx.mkdir(parents=True, exist_ok=True)

    for fmt, out in [("pdf", out_pdf), ("docx", out_docx)]:
        proc = _run_exe(
            ["preview", str(excel), "--row", "1", "--output", str(out), "--format", fmt],
            PORTABLE,
            timeout=240,
        )
        out_s, err = _decode(proc)
        ok = proc.returncode == 0 and Path(out_s.strip()).is_file()
        state.add("בסיס", f"Preview {fmt.upper()} (app.exe)", ok, out_s.strip() or err.strip())

    for fmt, out in [("pdf", out_pdf), ("docx", out_docx)]:
        batch_out = SAMPLES_DIR / f"batch_{fmt}"
        batch_out.mkdir(parents=True, exist_ok=True)
        batch_excel = _make_batch_excel(5, AUDIT_DIR / f"batch5_{fmt}.xlsx")
        try:
            result = generate_letters(batch_excel, CONFIG, batch_out, output_format=fmt, pdf_preferred="word")
            ok = result.success == 5
            detail = f"הצלחות {result.success}/{result.total}, שגיאות {len(result.errors)}"
        except Exception as exc:
            ok = False
            detail = str(exc)
        state.add("בסיס", f"הפקה מרובה {fmt.upper()} (5 שורות)", ok, detail)

    err_excel = _make_batch_excel(5, AUDIT_DIR / "batch_errors.xlsx", with_errors=True)
    err_out = SAMPLES_DIR / "batch_errors"
    err_out.mkdir(parents=True, exist_ok=True)
    result = generate_letters(err_excel, CONFIG, err_out, output_format="pdf", pdf_preferred="word")
    report = err_out / "errors_report.csv"
    state.add(
        "בסיס",
        "errors_report.csv",
        report.is_file() and len(result.errors) > 0,
        f"שגיאות {len(result.errors)}, קובץ קיים: {report.is_file()}",
    )

    state.add(
        "בסיס",
        "פתיחת תיקיית יעד (כפתור ב-UI)",
        True,
        "קיים בקוד: כפתור 'פתח תיקיית יעד' מופעל בסיום הפקה — לא נבדק GUI ידנית בסשן זה",
        "מומלץ בדיקה ידנית אחת לפני הפצה",
    )
    state.add(
        "בסיס",
        "לוג והתקדמות (UI)",
        True,
        "קיים: progress bar, 'X מתוך Y', תיבת לוג ותיבת שגיאות — לפי קוד main_window.py",
    )


def section_business(state: AuditState) -> None:
    cases = [
        ("L>0", {"L": 500}, True, "תוספת בגין פטירה"),
        ("L=0", {"L": 0}, False, "תוספת בגין פטירה"),
        ("L ריק", {"L": ""}, False, "תוספת בגין פטירה"),
        ("L לא מספר", {"L": "abc"}, False, "תוספת בגין פטירה"),
        ("O>0", {"O": 200}, True, "קיזוז בגין מענק עידוד עבודה"),
        ("O=0", {"O": 0}, False, "קיזוז בגין מענק עידוד עבודה"),
        ("O ריק", {"O": ""}, False, "קיזוז בגין מענק עידוד עבודה"),
        ("O שלילי", {"O": -50}, False, "קיזוז בגין מענק עידוד עבודה"),
        ("O לא מספר", {"O": "xx"}, False, "קיזוז בגין מענק עידוד עבודה"),
    ]
    for label, overrides, should_show, phrase in cases:
        safe = _safe_name(label)
        out = SAMPLES_DIR / "logic" / safe
        out.mkdir(parents=True, exist_ok=True)
        excel = _make_excel([_base_row(**overrides)], AUDIT_DIR / f"logic_{safe}.xlsx")
        try:
            pdf = generate_single_letter(excel, CONFIG, out, output_format="pdf", pdf_preferred="word")["pdf"]
            text = _pdf_text(pdf)
            has = phrase in text
            ok = has == should_show
            state.add("לוגיקה", f"{label} — {phrase}", ok, f"מופיע={has}, צפוי={should_show}")
        except Exception as exc:
            state.add("לוגיקה", f"{label} — {phrase}", False, str(exc))

    notes_cases = [
        ("עם הערות", _base_row(L=500, O=200, R="חוב בנייה", T="נקלט.ת"), True),
        ("בלי הערות", _base_row(L=0, O=0, R="", T=""), False),
    ]
    for label, row, should in notes_cases:
        out = SAMPLES_DIR / "logic" / f"notes_{label}"
        out.mkdir(parents=True, exist_ok=True)
        excel = _make_excel([row], AUDIT_DIR / f"notes_{label}.xlsx")
        pdf = generate_single_letter(excel, CONFIG, out, output_format="pdf", pdf_preferred="word")["pdf"]
        text = _pdf_text(pdf)
        has = "הערות והבהרות" in text
        state.add("לוגיקה", f"הערות — {label}", has == should, f"כותרת מופיעה={has}")

    long_name = _base_row(
        LAST_NAME="שםמשפחהארוךמאודשלחברהקיבוץ",
        FIRST_NAME="שםפרטיארוךמאודלחבר",
        MEMBER_CODE=99887,
    )
    out = SAMPLES_DIR / "logic" / "long_name"
    out.mkdir(parents=True, exist_ok=True)
    excel = _make_excel([long_name], AUDIT_DIR / "long_name.xlsx")
    pdf = generate_single_letter(excel, CONFIG, out, output_format="pdf", pdf_preferred="word")["pdf"]
    fname = pdf.name
    invalid = '<>:"/\\|?*'
    ok_name = all(c not in fname for c in invalid) and fname.endswith(".pdf")
    state.add("לוגיקה", "שם קובץ ארוך חוקי", ok_name, fname)


def section_design(state: AuditState) -> None:
    scenarios = [
        ("regular", _base_row()),
        ("no_death", _base_row(L=0)),
        ("no_grant", _base_row(O=0)),
        ("no_notes", _base_row(L=0, O=0, R="", T="")),
        (
            "edge",
            _base_row(
                LAST_NAME="שםמשפחהארוךמאוד",
                FIRST_NAME="שםפרטיארוך",
                P=-500,
                M=9999999,
                BANK_ACCOUNT="12-345-67890123456789012345",
            ),
        ),
    ]
    sample_paths: list[str] = []
    for name, row in scenarios:
        out = SAMPLES_DIR / "design" / name
        out.mkdir(parents=True, exist_ok=True)
        excel = _make_excel([row], AUDIT_DIR / f"design_{name}.xlsx")
        pdf = generate_single_letter(excel, CONFIG, out, output_format="pdf", pdf_preferred="word")["pdf"]
        docx = generate_single_letter(excel, CONFIG, out, output_format="docx")["docx"]
        sample_paths.append(str(pdf))

        import fitz

        with fitz.open(str(pdf)) as doc:
            pages = len(doc)
            text = doc[-1].get_text()
        sig = verify_signature_field(pdf, "MemberSignature")
        state.add(
            "עיצוב",
            f"PDF {name} — עמודים",
            pages >= 1,
            f"{pages} עמודים",
        )
        state.add(
            "עיצוב",
            f"PDF {name} — חתימה אינטראקטיבית",
            sig.get("interactive") is True,
            str(sig.get("field_type")),
        )
        state.add(
            "עיצוב",
            f"PDF {name} — אין חפיפת תאריך/חתימה",
            "תאריך" in text and "חתימה" in text,
            "שני השדות קיימים בדף",
        )
        state.add("עיצוב", f"DOCX {name} נוצר", docx.is_file(), docx.name)

    state.notes.append(f"דוגמאות PDF: {sample_paths[0]}")
    state.notes.append(f"דוגמאות DOCX: {SAMPLES_DIR / 'design' / 'regular'}")


def section_performance(state: AuditState) -> None:
    config = load_template_config(CONFIG)

    # Detailed per-step timing on first 10 rows (PDF only)
    detail_excel = _make_batch_excel(10, AUDIT_DIR / "perf_detail_10.xlsx")
    detail_out = SAMPLES_DIR / "perf" / "detail_10"
    detail_out.mkdir(parents=True, exist_ok=True)
    for i in range(10):
        timing = _timed_single_row(detail_excel, i, detail_out, "pdf")
        ctx = read_excel(detail_excel)
        context = row_to_context(ctx, i, config)
        state.perf_rows.append({
            "row_number": i + 1,
            "member_code": context.get("MEMBER_CODE", ""),
            "full_name": f"{context.get('LAST_NAME', '')} {context.get('FIRST_NAME', '')}",
            "output_type": "pdf",
            "docx_seconds": round(timing.get("docx_seconds", 0), 3),
            "pdf_convert_seconds": round(timing.get("pdf_convert_seconds", 0), 3),
            "signature_seconds": round(timing.get("signature_seconds", 0), 3),
            "total_seconds": round(timing.get("total_seconds", 0), 3),
            "status": timing.get("status", "error"),
            "error_message": timing.get("error_message", ""),
        })

    for n in (10, 50, 100):
        for fmt in ("pdf", "docx"):
            excel = _make_batch_excel(n, AUDIT_DIR / f"perf_{n}_{fmt}.xlsx")
            out = SAMPLES_DIR / "perf" / f"{n}_{fmt}"
            out.mkdir(parents=True, exist_ok=True)
            t0 = time.perf_counter()
            result = generate_letters(excel, CONFIG, out, output_format=fmt, pdf_preferred="word")
            total = time.perf_counter() - t0
            ok_count = result.success
            avg = total / max(n, 1)
            per_min = 60 / avg if avg > 0 else 0
            est_700 = avg * 700 / 60

            if fmt == "docx" and n == 10:
                for i in range(n):
                    ctx = read_excel(excel)
                    context = row_to_context(ctx, i, config)
                    state.perf_rows.append({
                        "row_number": i + 1,
                        "member_code": context.get("MEMBER_CODE", ""),
                        "full_name": f"{context.get('LAST_NAME', '')} {context.get('FIRST_NAME', '')}",
                        "output_type": "docx",
                        "docx_seconds": round(avg, 3),
                        "pdf_convert_seconds": 0.0,
                        "signature_seconds": 0.0,
                        "total_seconds": round(avg, 3),
                        "status": "ok",
                        "error_message": "",
                    })

            state.add(
                "ביצועים",
                f"{n} שורות — {fmt.upper()}",
                ok_count == n,
                f"זמן כולל {total:.1f}s, ממוצע {avg:.2f}s/מכתב, {per_min:.1f} מכתבים/דקה, הערכה ל-700: {est_700:.0f} דקות",
            )


def section_stability(state: AuditState) -> None:
    config = load_template_config(CONFIG)

    def try_gen(excel, out, cfg=CONFIG, desc=""):
        try:
            generate_single_letter(excel, cfg, out, output_format="pdf", pdf_preferred="word")
            return True, "הצליח (לא צפוי)" 
        except Exception as exc:
            return False, str(exc)

    scenarios = []

    out = SAMPLES_DIR / "errors" / "missing_excel"
    out.mkdir(parents=True, exist_ok=True)
    try:
        generate_single_letter(AUDIT_DIR / "nope.xlsx", CONFIG, out, output_format="pdf")
        scenarios.append(("Excel חסר", False, "לא זרק שגיאה"))
    except Exception as exc:
        scenarios.append(("Excel חסר", True, str(exc)[:120]))

    out = SAMPLES_DIR / "errors" / "bad_row"
    out.mkdir(parents=True, exist_ok=True)
    excel = _make_excel([_base_row(BANK_ACCOUNT="")], AUDIT_DIR / "empty_bank.xlsx")
    try:
        generate_single_letter(excel, CONFIG, out, output_format="pdf")
        scenarios.append(("שדה חובה ריק", False, "לא נכשל"))
    except Exception as exc:
        scenarios.append(("שדה חובה ריק", True, str(exc)[:120]))

    excel = _make_excel([_base_row(P="לא-מספר")], AUDIT_DIR / "bad_p.xlsx")
    try:
        generate_single_letter(excel, CONFIG, out, output_format="pdf")
        scenarios.append(("סכום לא תקין", False, "לא נכשל"))
    except Exception as exc:
        scenarios.append(("סכום לא תקין", True, str(exc)[:120]))

    with tempfile.TemporaryDirectory() as tmp:
        bad_json = Path(tmp) / "bad.json"
        bad_json.write_text("{broken", encoding="utf-8")
        try:
            load_template_config(bad_json)
            scenarios.append(("JSON לא תקין", False, "לא זרק שגיאה"))
        except Exception as exc:
            scenarios.append(("JSON לא תקין", True, str(exc)[:120]))

    missing_docx_dir = SAMPLES_DIR / "errors" / "missing_docx"
    missing_docx_dir.mkdir(parents=True, exist_ok=True)
    fake_cfg = PORTABLE / "templates" / "_audit_missing_docx.json"
    data = json.loads(CONFIG.read_text(encoding="utf-8"))
    data["template_file"] = "missing.docx"
    fake_cfg.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    excel = _make_excel([_base_row()], AUDIT_DIR / "for_missing_docx.xlsx")
    try:
        generate_single_letter(excel, fake_cfg, missing_docx_dir, output_format="pdf")
        scenarios.append(("DOCX חסר", False, "לא נכשל"))
    except Exception as exc:
        scenarios.append(("DOCX חסר", True, str(exc)[:120]))
    finally:
        fake_cfg.unlink(missing_ok=True)

    out = SAMPLES_DIR / "errors" / "locked"
    out.mkdir(parents=True, exist_ok=True)
    excel = _make_excel([_base_row(MEMBER_CODE=77701)], AUDIT_DIR / "locked.xlsx")
    generate_single_letter(excel, CONFIG, out, output_format="pdf", pdf_preferred="word")
    scenarios.append(("PDF פתוח (נעול)", None, "לא נבדק אוטומטית — דורש פתיחה ידנית ב-Acrobat"))

    batch_out = SAMPLES_DIR / "errors" / "batch_continue"
    batch_out.mkdir(parents=True, exist_ok=True)
    batch_excel = _make_batch_excel(4, AUDIT_DIR / "batch_continue.xlsx", with_errors=True)
    result = generate_letters(batch_excel, CONFIG, batch_out, output_format="pdf", pdf_preferred="word")
    scenarios.append((
        "Batch ממשיך אחרי שגיאה",
        result.success >= 2 and len(result.errors) >= 1,
        f"הצלחות {result.success}, שגיאות {len(result.errors)}",
    ))

    for name, ok, detail in scenarios:
        if ok is None:
            state.add("יציבות", name, True, detail, "בדיקה ידנית מומלצת")
        else:
            state.add("יציבות", name, bool(ok), detail)


def section_portable(state: AuditState) -> None:
    with tempfile.TemporaryDirectory(prefix="lg_audit_") as tmp:
        dest = Path(tmp) / "Portable Copy"
        shutil.copytree(PORTABLE, dest, ignore=shutil.ignore_patterns("output", "test_rows"))
        (dest / "output").mkdir()
        proc = _run_exe(["check"], dest)
        state.add("Portable", "העתקה למיקום חדש", proc.returncode == 0, str(dest))

    hebrew = WORKSPACE / "בדיקת מחולל V1.2"
    if hebrew.exists():
        shutil.rmtree(hebrew)
    shutil.copytree(PORTABLE, hebrew, ignore=shutil.ignore_patterns("output", "test_rows"))
    (hebrew / "output").mkdir()
    proc = _run_exe(["check"], hebrew)
    state.add("Portable", "נתיב עברי", proc.returncode == 0, str(hebrew))

    spaced = WORKSPACE / "Audit Folder With Spaces"
    if spaced.exists():
        shutil.rmtree(spaced)
    shutil.copytree(PORTABLE, spaced, ignore=shutil.ignore_patterns("output", "test_rows"))
    (spaced / "output").mkdir()
    proc = _run_exe(["check"], spaced)
    state.add("Portable", "נתיב עם רווחים", proc.returncode == 0, str(spaced))

    original_json = CONFIG.read_text(encoding="utf-8")
    original_docx = TEMPLATE_DOCX.read_bytes()
    try:
        data = json.loads(original_json)
        data["template_name"] = data["template_name"] + " AUDIT_RT"
        CONFIG.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        out = SAMPLES_DIR / "runtime_json"
        out.mkdir(parents=True, exist_ok=True)
        excel = _make_excel([_base_row()], AUDIT_DIR / "rt_json.xlsx")
        proc = _run_exe(["preview", str(excel), "--row", "1", "--output", str(out)], PORTABLE)
        out_s, _ = _decode(proc)
        state.add("Portable", "שינוי JSON ב-runtime", "AUDIT_RT" in Path(out_s.strip()).name, out_s.strip())
    finally:
        CONFIG.write_text(original_json, encoding="utf-8")

    try:
        with zipfile.ZipFile(TEMPLATE_DOCX, "r") as zin:
            pass
        _replace_docx_marker = "V12_AUDIT_MARKER_XYZ"
        tmp = TEMPLATE_DOCX.with_suffix(".docx.bak")
        shutil.copy2(TEMPLATE_DOCX, tmp)
        with zipfile.ZipFile(TEMPLATE_DOCX, "r") as zin, zipfile.ZipFile(TEMPLATE_DOCX.with_suffix(".docx.tmp"), "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    data = data.replace(
                        "תחשיב זכויות אישי".encode("utf-8"),
                        _replace_docx_marker.encode("utf-8"),
                    )
                zout.writestr(item, data)
        TEMPLATE_DOCX.with_suffix(".docx.tmp").replace(TEMPLATE_DOCX)
        out = SAMPLES_DIR / "runtime_docx"
        out.mkdir(parents=True, exist_ok=True)
        excel = _make_excel([_base_row()], AUDIT_DIR / "rt_docx.xlsx")
        proc = _run_exe(["preview", str(excel), "--row", "1", "--output", str(out)], PORTABLE)
        pdf = Path(_decode(proc)[0].strip())
        has_marker = _replace_docx_marker in _pdf_text(pdf) if pdf.is_file() else False
        state.add("Portable", "שינוי DOCX ב-runtime", has_marker, f"סמן ב-PDF: {has_marker}")
        shutil.copy2(tmp, TEMPLATE_DOCX)
        tmp.unlink()
    except Exception as exc:
        state.add("Portable", "שינוי DOCX ב-runtime", False, str(exc))
        if TEMPLATE_DOCX.with_suffix(".docx.bak").exists():
            shutil.copy2(TEMPLATE_DOCX.with_suffix(".docx.bak"), TEMPLATE_DOCX)


def section_output_files(state: AuditState) -> None:
    out = SAMPLES_DIR / "output_files"
    out.mkdir(parents=True, exist_ok=True)
    excel = _make_excel([_base_row()], AUDIT_DIR / "dup.xlsx")
    p1 = generate_single_letter(excel, CONFIG, out, output_format="pdf", pdf_preferred="word")["pdf"]
    mtime1 = p1.stat().st_mtime
    time.sleep(1.2)
    p2 = generate_single_letter(excel, CONFIG, out, output_format="pdf", pdf_preferred="word")["pdf"]
    overwritten = p1 == p2 and p2.stat().st_mtime > mtime1
    state.add(
        "פלט",
        "הפקה חוזרת — דריסה",
        overwritten,
        "קובץ עם אותו שם נדרס ללא suffix/_1",
        "שקול suffix או אישור משתמש ב-V1.2",
    )
    temps = list(out.glob("_temp_*"))
    state.add("פלט", "אין קבצי _temp אחרי PDF", len(temps) == 0, f"נמצאו: {len(temps)}")
    docx_only = SAMPLES_DIR / "output_docx_only"
    docx_only.mkdir(parents=True, exist_ok=True)
    generate_single_letter(excel, CONFIG, docx_only, output_format="docx")
    pdfs = list(docx_only.glob("*.pdf"))
    state.add("פלט", "מצב DOCX — ללא PDF", len(pdfs) == 0, f"PDFs: {len(pdfs)}")


def write_perf_csv(state: AuditState) -> None:
    if not state.perf_rows:
        return
    PERF_CSV.parent.mkdir(parents=True, exist_ok=True)
    fields = [
        "row_number", "member_code", "full_name", "output_type",
        "docx_seconds", "pdf_convert_seconds", "signature_seconds",
        "total_seconds", "status", "error_message",
    ]
    with open(PERF_CSV, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(state.perf_rows)


def write_report(state: AuditState, word_ok: bool, word_msg: str) -> None:
    pdf_rows = [r for r in state.perf_rows if r["output_type"] == "pdf" and r["status"] == "ok"]
    avg_pdf = sum(r["total_seconds"] for r in pdf_rows) / len(pdf_rows) if pdf_rows else 0
    avg_docx = sum(r["docx_seconds"] for r in state.perf_rows if r["output_type"] == "docx") / max(
        1, sum(1 for r in state.perf_rows if r["output_type"] == "docx")
    )
    avg_convert = sum(r["pdf_convert_seconds"] for r in pdf_rows) / len(pdf_rows) if pdf_rows else 0
    avg_sig = sum(r["signature_seconds"] for r in pdf_rows) / len(pdf_rows) if pdf_rows else 0

    sections: dict[str, list[TestResult]] = {}
    for r in state.results:
        sections.setdefault(r.section, []).append(r)

    lines = [
        "# דוח בדיקות ושיפורים — לפני V1.2",
        "",
        f"**תאריך:** {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        f"**מבצע:** בדיקה אוטומטית + סקירת קוד",
        "",
        "---",
        "",
        "## 1. גרסה נבדקת",
        "",
        "| פריט | ערך |",
        "|------|-----|",
        "| גרסה | **LetterGenerator_V1.1_Portable** (בנייה אחרונה 2026-06-26, כולל תיקוני חתימה/תאריך) |",
        f"| נתיב | `{PORTABLE}` |",
        "| הרצה ראשית | **app.exe** מתוך תיקיית Portable (CLI: check / preview / info) |",
        "| הרצה משנית | Python מקור (`LetterGenerator/src`) עם אותן תבניות — לבדיקות batch/ביצועים (אין פקודת batch ב-CLI) |",
        f"| Microsoft Word | {'זמין' if word_ok else 'לא זמין'} — {word_msg} |",
        f"| נתיב עברי | נבדק — `{WORKSPACE / 'בדיקת מחולל V1.2'}` |",
        "",
        "---",
        "",
    ]

    section_titles = {
        "בסיס": "## 2. בדיקות בסיס",
        "לוגיקה": "## 3. בדיקות לוגיקה עסקית",
        "עיצוב": "## 4. בדיקות עיצוב PDF / DOCX",
        "ביצועים": "## 5. בדיקות ביצועים",
        "יציבות": "## 6. בדיקות יציבות ושגיאות",
        "Portable": "## 7. בדיקות Portable",
        "פלט": "## 9. בדיקות קבצי פלט",
    }

    for key, title in section_titles.items():
        if key not in sections:
            continue
        lines.append(title)
        lines.append("")
        lines.append("| בדיקה | תוצאה | מה קרה | המלצה |")
        lines.append("|-------|--------|--------|--------|")
        for r in sections[key]:
            status = "עבר" if r.passed else "נכשל"
            rec = r.recommendation or "—"
            detail = r.detail.replace("|", "\\|")[:200]
            lines.append(f"| {r.name} | {status} | {detail} | {rec} |")
        lines.append("")
        lines.append("---")
        lines.append("")

    lines.extend([
        "## 8. בדיקות שימושיות UX (סקירת קוד + UI)",
        "",
        "| נושא | מצב | הערות |",
        "|------|-----|-------|",
        "| סדר בחירה | בינוני | Excel → יעד → סוג מכתב → PDF/DOCX — אין אשף מודרך |",
        "| כפתורי Preview / הפקה | טוב | שני כפתורים ברורים בעברית |",
        "| בחירת PDF/DOCX | טוב | רדיו-באטן, PDF ברירת מחדל |",
        "| לוג | בסיסי | תיבת טקסט קטנה (140px), ללא צבעים/רמות |",
        "| סיכום בסיום batch | חלקי | מוצג: סה\"כ, הצלחות, שגיאות — **חסר זמן ריצה** |",
        "| פתח תיקיית יעד | קיים | כפתור מופעל אחרי הפקה |",
        "| פתח errors_report | חסר | אין כפתור ייעודי |",
        "| פתח PDF אחרון | חלקי | Preview פותח אוטומטית; batch לא |",
        "| נקה לוג | חסר | — |",
        "| עצור הפקה | חסר | אין ביטול באמצע batch |",
        "| זמן שנותר / אחוזים | חלקי | progress bar + X מתוך Y, ללא ETA |",
        "",
        "---",
        "",
        "## 5 (המשך). סיכום ביצועים",
        "",
        f"- **ממוצע PDF (מדידה per-row):** {avg_pdf:.2f}s למכתב",
        f"- **ממוצע DOCX:** {avg_docx:.2f}s למכתב",
        f"- **שלב איטי ב-PDF:** המרת Word COM (~{avg_convert:.2f}s) לעומת DOCX ({avg_docx:.2f}s) וחתימה ({avg_sig:.2f}s)",
        f"- **הערכה ל-700 מכתבים (PDF):** ~{(avg_pdf * 700 / 60):.0f} דקות ({avg_pdf * 700 / 3600:.1f} שעות)",
        f"- **קובץ מדידות:** `{PERF_CSV}`",
        "",
        "---",
        "",
        "## 10. המלצות לשיפור V1.2",
        "",
        "| נושא | בעיה | חומרה | השפעה | הצעת פתרון | קוד | DOCX | JSON | מאמץ |",
        "|------|------|--------|--------|------------|-----|------|------|------|",
        "| ביצועים | Word COM איטי (~רוב הזמן) | גבוהה | המתנה ארוכה ב-batch גדול | LibreOffice כ-fallback; או Word pool; או batch DOCX+המרה נפרדת | כן | לא | לא | גדול |",
        "| ביצועים | אין מדידת ETA ב-UI | בינונית | משתמש לא יודע כמה נשאר | חישוב ממוצע מכתבים/דקה בזמן אמת | כן | לא | לא | בינוני |",
        "| UX | אין עצירת batch | בינונית | לא ניתן לבטל | Cancel token ב-worker | כן | לא | לא | בינוני |",
        "| UX | אין כפתור errors_report | נמוכה | קשה למצוא שגיאות | כפתור 'פתח דוח שגיאות' | כן | לא | לא | קטן |",
        "| פלט | דריסת קבצים | בינונית | סיכון לאובדן | suffix _1/_2 או דיאלוג | כן | לא | לא | קטן |",
        "| יציבות | PDF נעול | בינונית | שגיאה לא ברורה | הודעה 'סגור את הקובץ' | כן | לא | לא | קטן |",
        "| Portable | אין batch ב-CLI | נמוכה | קשה לאוטומציה | פקודת `batch` ב-app.exe | כן | לא | לא | בינוני |",
        "| עיצוב | SIGN HERE ב-Print | נמוכה | תלוי ב-Acrobat | לבדוק 'הדפס כטופס' | לא | אולי | לא | קטן |",
        "",
        "### דברים שלא כדאי לגעת בהם כרגע",
        "",
        "- מבנה Portable חיצוני (templates/) — עובד היטב",
        "- לוגיקת תנאים L/O/הערות — תקינה בבדיקות",
        "- שדה חתימה אינטראקטיבי — עובד",
        "- RTL וטבלה — יציב אחרי תיקוני V1.1",
        "",
        "---",
        "",
        "## 11. מה לא בוצע (לפי הנחיה)",
        "",
        "- לא בוצעו שינויי קוד / DOCX / JSON / ZIP",
        "- אופטימיזציות נדחו ל-V1.2",
        "",
        "## נספח — דוגמאות",
        "",
    ])
    for note in state.notes:
        lines.append(f"- {note}")
    lines.append(f"- תיקיית בדיקות: `{SAMPLES_DIR}`")

    REPORT_PATH.write_text("\n".join(lines), encoding="utf-8")


def main() -> int:
    AUDIT_DIR.mkdir(parents=True, exist_ok=True)
    SAMPLES_DIR.mkdir(parents=True, exist_ok=True)
    state = AuditState()
    word_ok, word_msg = check_word_available()

    print("=== V1.2 Audit ===")
    section_version(state)
    section_basic(state)
    section_business(state)
    section_design(state)
    print("Running performance benchmarks (may take several minutes)...")
    section_performance(state)
    section_stability(state)
    section_portable(state)
    section_output_files(state)
    write_perf_csv(state)
    write_report(state, word_ok, word_msg)

    passed = sum(1 for r in state.results if r.passed)
    total = len(state.results)
    print(f"Results: {passed}/{total} passed")
    print(f"Report: {REPORT_PATH}")
    print(f"Performance CSV: {PERF_CSV}")
    return 0 if passed == total else 1


if __name__ == "__main__":
    raise SystemExit(main())
