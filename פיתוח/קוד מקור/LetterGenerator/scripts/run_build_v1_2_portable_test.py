"""Post-build verification for LetterGenerator_V1.2_Portable + ZIP."""

from __future__ import annotations

import csv
import json
import os
import shutil
import subprocess
import sys
import time
import zipfile
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
WORKSPACE = ROOT.parent
RELEASE_DIR = WORKSPACE / "LetterGenerator_V1.2_Portable"
ZIP_PATH = WORKSPACE / "LetterGenerator_V1.2_Portable.zip"
REPORT_PATH = WORKSPACE / "cursor" / "דוח-Build-V1.2-Final-Portable.md"
TEST_ROOT = WORKSPACE / "cursor" / "build_tests" / "v1_2"
TEST_DATA = TEST_ROOT / "data"
TEST_OUTPUT = TEST_ROOT / "output"
CONFIG = RELEASE_DIR / "templates" / "תחשיב זכויות אישי.json"
DOCX_NAME = "תחשיב זכויות אישי.docx"
JSON_NAME = "תחשיב זכויות אישי.json"
RUNTIME_JSON_SUFFIX = " V12_PORTABLE_JSON"
DOCX_MARKER = "V12_PORTABLE_DOCX"

ALLOWED_RELEASE_TOP = frozenset({"app.exe", "README.txt", "templates", "output"})
ALLOWED_OUTPUT = frozenset({"README.txt"})

sys.path.insert(0, str(ROOT))
from src.letter_generator import generate_letters  # noqa: E402
from src.signature_field import verify_signature_field  # noqa: E402

_DISPATCH_COUNT = 0


def _patch_word_dispatch_counter() -> None:
    global _DISPATCH_COUNT
    import win32com.client

    original = win32com.client.Dispatch

    def counted(prog_id, *args, **kwargs):
        global _DISPATCH_COUNT
        if "word" in str(prog_id).lower():
            _DISPATCH_COUNT += 1
        return original(prog_id, *args, **kwargs)

    win32com.client.Dispatch = counted  # type: ignore[assignment]


def _winword_count() -> int:
    out = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/NH"],
        capture_output=True,
        text=True,
        timeout=15,
    )
    return len([ln for ln in out.stdout.splitlines() if "WINWORD.EXE" in ln.upper()])


def _run_exe(cwd: Path, args: list[str], *, timeout: int = 300) -> subprocess.CompletedProcess:
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


def _prepare_test_dirs() -> None:
    if TEST_OUTPUT.exists():
        shutil.rmtree(TEST_OUTPUT)
    TEST_DATA.mkdir(parents=True, exist_ok=True)
    TEST_OUTPUT.mkdir(parents=True, exist_ok=True)


def _make_batch_excel(n: int, path: Path, *, bad_p_row: int | None = None) -> Path:
    rows = []
    for i in range(n):
        cols = {chr(65 + j): "" for j in range(20)}
        p_val = "לא-מספר" if bad_p_row is not None and i == bad_p_row else 1400 + i
        cols.update({
            "C": 50001 + i, "E": "כהן", "F": "ישראל",
            "H": 10, "I": 1000, "J": 100, "L": 500, "M": 1600,
            "O": 200, "P": p_val, "S": f"12-345-{678900+i:06d}", "R": "", "T": "",
        })
        rows.append(cols)
    path.parent.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(rows).to_excel(path, index=False)
    return path


def _preview(portable: Path, excel: Path, out: Path, fmt: str) -> Path:
    proc = _run_exe(
        portable,
        ["preview", str(excel), "--row", "1", "--output", str(out), "--format", fmt],
        timeout=300,
    )
    out_s, err = _decode(proc)
    if proc.returncode != 0:
        raise RuntimeError(f"preview {fmt} failed: {err or out_s}")
    p = Path(out_s.strip())
    if not p.is_file():
        raise RuntimeError(f"missing output: {out_s}")
    return p


def _sanitize_release() -> list[str]:
    """Remove test artifacts; return list of removed paths (relative to release)."""
    removed: list[str] = []
    if not RELEASE_DIR.is_dir():
        return removed

    for path in list(RELEASE_DIR.iterdir()):
        if path.name in ALLOWED_RELEASE_TOP:
            continue
        rel = path.name
        if path.is_dir():
            shutil.rmtree(path)
        else:
            path.unlink()
        removed.append(rel)

    output_dir = RELEASE_DIR / "output"
    if output_dir.is_dir():
        for path in list(output_dir.iterdir()):
            if path.name in ALLOWED_OUTPUT:
                continue
            rel = f"output/{path.name}"
            if path.is_dir():
                shutil.rmtree(path)
            else:
                path.unlink()
            removed.append(rel)

    return removed


def _assert_release_clean() -> None:
    if not RELEASE_DIR.is_dir():
        raise RuntimeError(f"Release missing: {RELEASE_DIR}")

    top = {p.name for p in RELEASE_DIR.iterdir()}
    extra = top - ALLOWED_RELEASE_TOP
    if extra:
        raise RuntimeError(f"Unexpected top-level items in release: {sorted(extra)}")

    output_dir = RELEASE_DIR / "output"
    if not output_dir.is_dir():
        raise RuntimeError("Missing output/ in release")

    out_items = {p.name for p in output_dir.iterdir()}
    bad_out = out_items - ALLOWED_OUTPUT
    if bad_out:
        raise RuntimeError(f"output/ is not clean: {sorted(bad_out)}")

    templates = RELEASE_DIR / "templates"
    tpl_files = sorted(p.name for p in templates.iterdir() if p.is_file()) if templates.is_dir() else []
    for name in tpl_files:
        low = name.lower()
        if not (low.endswith(".docx") or low.endswith(".json")):
            raise RuntimeError(f"Unexpected template file: {name}")


def test_batch_pdf(n: int) -> dict:
    global _DISPATCH_COUNT
    _DISPATCH_COUNT = 0
    excel = _make_batch_excel(n, TEST_DATA / f"batch_{n}.xlsx")
    out = TEST_OUTPUT / f"batch_{n}"
    if out.exists():
        shutil.rmtree(out)
    out.mkdir(parents=True)
    t0 = time.perf_counter()
    result = generate_letters(excel, CONFIG, out, output_format="pdf", pdf_preferred="word")
    elapsed = time.perf_counter() - t0
    return {
        "id": 1 if n == 50 else 2,
        "name": f"PDF batch {n} שורות",
        "ok": result.success == n and len(result.errors) == 0,
        "success": result.success,
        "total": n,
        "seconds": round(elapsed, 1),
        "word_dispatch": _DISPATCH_COUNT,
    }


def test_word_single_dispatch() -> dict:
    excel = _make_batch_excel(20, TEST_DATA / "dispatch_20.xlsx")
    out = TEST_OUTPUT / "dispatch_20"
    if out.exists():
        shutil.rmtree(out)
    out.mkdir(parents=True)
    global _DISPATCH_COUNT
    _DISPATCH_COUNT = 0
    generate_letters(excel, CONFIG, out, output_format="pdf", pdf_preferred="word")
    return {
        "id": 3,
        "name": "Word נפתח פעם אחת ל-batch",
        "ok": _DISPATCH_COUNT == 1,
        "word_dispatch": _DISPATCH_COUNT,
    }


def test_winword_closed() -> dict:
    time.sleep(5)
    count = _winword_count()
    return {
        "id": 4,
        "name": "אין WINWORD אחרי 5 שניות",
        "ok": count == 0,
        "winword_count": count,
    }


def test_pdf_mode(portable: Path) -> dict:
    excel = _make_batch_excel(1, TEST_DATA / "pdf_one.xlsx")
    out = TEST_OUTPUT / "pdf_mode"
    out.mkdir(parents=True, exist_ok=True)
    pdf = _preview(portable, excel, out, "pdf")
    return {
        "id": 5,
        "name": "PDF mode (preview)",
        "ok": pdf.suffix.lower() == ".pdf",
        "file": pdf.name,
    }


def test_docx_mode(portable: Path) -> dict:
    excel = _make_batch_excel(1, TEST_DATA / "docx_one.xlsx")
    out = TEST_OUTPUT / "docx_mode"
    out.mkdir(parents=True, exist_ok=True)
    docx = _preview(portable, excel, out, "docx")
    return {
        "id": 6,
        "name": "DOCX mode (preview)",
        "ok": docx.suffix.lower() == ".docx",
        "file": docx.name,
    }


def test_signature(portable: Path) -> dict:
    excel = _make_batch_excel(1, TEST_DATA / "sig.xlsx")
    out = TEST_OUTPUT / "sig"
    out.mkdir(parents=True, exist_ok=True)
    pdf = _preview(portable, excel, out, "pdf")
    info = verify_signature_field(pdf, "MemberSignature")
    return {
        "id": 7,
        "name": "חתימה אינטראקטיבית",
        "ok": info.get("interactive") is True,
        "info": info,
    }


def test_json_runtime(portable: Path) -> dict:
    config = portable / "templates" / JSON_NAME
    original = config.read_text(encoding="utf-8")
    out = TEST_OUTPUT / "json_rt"
    out.mkdir(parents=True, exist_ok=True)
    excel = _make_batch_excel(1, TEST_DATA / "json.xlsx")
    pdf = None
    try:
        data = json.loads(original)
        data["template_name"] = data["template_name"] + RUNTIME_JSON_SUFFIX
        config.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        pdf = _preview(portable, excel, out, "pdf")
        ok = RUNTIME_JSON_SUFFIX.strip() in pdf.name
    finally:
        config.write_text(original, encoding="utf-8")
    return {"id": 8, "name": "JSON runtime", "ok": ok, "file": pdf.name if pdf and ok else ""}


def test_docx_runtime(portable: Path) -> dict:
    docx_path = portable / "templates" / DOCX_NAME
    backup = docx_path.with_suffix(".docx.bak")
    shutil.copy2(docx_path, backup)
    out = TEST_OUTPUT / "docx_rt"
    out.mkdir(parents=True, exist_ok=True)
    excel = _make_batch_excel(1, TEST_DATA / "docx_rt.xlsx")
    ok = False
    try:
        with zipfile.ZipFile(docx_path, "r") as zin, zipfile.ZipFile(
            docx_path.with_suffix(".docx.tmp"), "w"
        ) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    data = data.replace(
                        "תחשיב זכויות אישי".encode("utf-8"),
                        DOCX_MARKER.encode("utf-8"),
                    )
                zout.writestr(item, data)
        docx_path.with_suffix(".docx.tmp").replace(docx_path)
        pdf = _preview(portable, excel, out, "pdf")
        import fitz

        text = fitz.open(str(pdf))[-1].get_text()
        ok = DOCX_MARKER in text
    finally:
        shutil.copy2(backup, docx_path)
        backup.unlink(missing_ok=True)
    return {"id": 8, "name": "DOCX runtime", "ok": ok}


def test_one_page(portable: Path) -> dict:
    import fitz

    excel = _make_batch_excel(1, TEST_DATA / "one_page.xlsx")
    out = TEST_OUTPUT / "one_page"
    out.mkdir(parents=True, exist_ok=True)
    pdf = _preview(portable, excel, out, "pdf")
    pages = fitz.open(str(pdf)).page_count
    return {
        "id": 10,
        "name": "דוגמה רגילה — עמוד אחד",
        "ok": pages == 1,
        "pages": pages,
    }


def test_rtl_body(portable: Path) -> dict:
    import re
    import zipfile
    import fitz

    tpl = portable / "templates" / DOCX_NAME
    with zipfile.ZipFile(tpl) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    m = re.search(r"<w:p[ >].*?עבור:.*?</w:p>", xml, re.DOTALL)
    docx_ok = bool(m and 'w:jc w:val="left"' in m.group(0) and "w:bidi" in m.group(0))

    excel = _make_batch_excel(1, TEST_DATA / "rtl.xlsx")
    out = TEST_OUTPUT / "rtl"
    out.mkdir(parents=True, exist_ok=True)
    pdf = _preview(portable, excel, out, "pdf")
    page = fitz.open(str(pdf))[0]
    pw = page.rect.width
    hits = page.search_for("עבור:")
    pdf_ok = bool(hits and (pw - hits[0].x1) < 45)

    return {
        "id": 11,
        "name": "גוף מיושר לימין (DOCX + PDF)",
        "ok": docx_ok and pdf_ok,
        "docx_jc_bidi": docx_ok,
        "pdf_right_margin_pt": round(pw - hits[0].x1, 1) if hits else None,
    }


def test_lo_conditions(portable: Path) -> dict:
    import fitz

    excel = _make_batch_excel(1, TEST_DATA / "lo.xlsx")
    out = TEST_OUTPUT / "lo"
    out.mkdir(parents=True, exist_ok=True)
    pdf = _preview(portable, excel, out, "pdf")
    text = fitz.open(str(pdf))[0].get_text()
    ok = (
        "תוספת בגין פטירה" in text
        and "קיזוז בגין מענק עידוד עבודה" in text
        and "אחרי קיזוז" in text
    )
    return {"id": 12, "name": "תנאי L/O — שורות מותנות", "ok": ok}


def test_date_field(portable: Path) -> dict:
    from pypdf import PdfReader

    excel = _make_batch_excel(1, TEST_DATA / "date.xlsx")
    out = TEST_OUTPUT / "date"
    out.mkdir(parents=True, exist_ok=True)
    pdf = _preview(portable, excel, out, "pdf")
    fields = PdfReader(str(pdf)).get_fields() or {}
    ok = "SignDateEntry" in fields
    return {"id": 13, "name": "שדה תאריך", "ok": ok}


def test_editable_table(portable: Path) -> dict:
    import fitz

    docx_path = portable / "templates" / DOCX_NAME
    backup = docx_path.with_suffix(".docx.bak")
    shutil.copy2(docx_path, backup)
    marker = "EDITABLE_TABLE_V12"
    out = TEST_OUTPUT / "editable_tbl"
    out.mkdir(parents=True, exist_ok=True)
    excel = _make_batch_excel(1, TEST_DATA / "editable_tbl.xlsx")
    ok = False
    try:
        with zipfile.ZipFile(docx_path, "r") as zin, zipfile.ZipFile(
            docx_path.with_suffix(".docx.tmp"), "w"
        ) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    data = data.replace(
                        "מספר שנים לצורך התחשיב".encode("utf-8"),
                        marker.encode("utf-8"),
                    )
                zout.writestr(item, data)
        docx_path.with_suffix(".docx.tmp").replace(docx_path)
        pdf = _preview(portable, excel, out, "pdf")
        ok = marker in fitz.open(str(pdf))[0].get_text()
    finally:
        shutil.copy2(backup, docx_path)
        backup.unlink(missing_ok=True)
    return {"id": 14, "name": "עריכת טקסט בטבלה → PDF בלי rebuild", "ok": ok}


def test_template_writable(portable: Path) -> dict:
    import os
    import stat

    tpl = portable / "templates" / DOCX_NAME
    readonly = False
    if os.name == "nt":
        readonly = bool(os.stat(tpl).st_file_attributes & stat.FILE_ATTRIBUTE_READONLY)
    return {
        "id": 15,
        "name": "תבנית לא Read Only",
        "ok": not readonly and tpl.is_file(),
    }


def test_app_check(portable: Path) -> dict:
    proc = _run_exe(portable, ["check"], timeout=60)
    out_s, err = _decode(proc)
    return {
        "id": 16,
        "name": "app.exe נפתח (check)",
        "ok": proc.returncode == 0,
        "stdout": out_s.strip()[:200],
        "stderr": err.strip()[:200],
    }


def test_invalid_p() -> dict:
    excel = _make_batch_excel(3, TEST_DATA / "bad_p.xlsx", bad_p_row=1)
    out = TEST_OUTPUT / "bad_p"
    if out.exists():
        shutil.rmtree(out)
    out.mkdir(parents=True)
    result = generate_letters(excel, CONFIG, out, output_format="pdf", pdf_preferred="word")
    report = out / "errors_report.csv"
    rows = list(csv.DictReader(report.open(encoding="utf-8-sig"))) if report.is_file() else []
    ok = result.success == 2 and len(result.errors) == 1 and report.is_file() and len(rows) == 1
    return {
        "id": 9,
        "name": "P לא מספרי — שגיאה + errors_report",
        "ok": ok,
        "success": result.success,
        "errors": result.errors,
        "report_rows": rows,
    }


def create_zip() -> Path:
    removed = _sanitize_release()
    if removed:
        print("Sanitized release (removed):", removed)
    _assert_release_clean()

    if ZIP_PATH.is_file():
        ZIP_PATH.unlink()
    with zipfile.ZipFile(ZIP_PATH, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in RELEASE_DIR.rglob("*"):
            if path.is_file():
                zf.write(path, path.relative_to(RELEASE_DIR.parent))
    return ZIP_PATH


def write_report(results: list[dict]) -> None:
    passed = sum(1 for r in results if r.get("ok"))
    lines = [
        "# דוח Build V1.2 — Portable סופי",
        "",
        f"**תאריך:** {time.strftime('%Y-%m-%d %H:%M')}",
        f"**תיקייה:** `{RELEASE_DIR}`",
        f"**ZIP:** `{ZIP_PATH}`",
        f"**בדיקות חיצוניות:** `{TEST_ROOT}`",
        "",
        f"## סיכום: {passed}/{len(results)} בדיקות עברו",
        "",
        "| # | בדיקה | תוצאה |",
        "|---|--------|--------|",
    ]
    for r in results:
        lines.append(f"| {r.get('id', '—')} | {r['name']} | {'עבר' if r.get('ok') else 'נכשל'} |")
    lines.append("")
    for r in results:
        lines.append(f"### {r['name']}")
        for k, v in r.items():
            if k not in ("name", "id", "ok"):
                lines.append(f"- {k}: {v}")
        lines.append("")
    REPORT_PATH.write_text("\n".join(lines), encoding="utf-8")


def main() -> int:
    if not RELEASE_DIR.is_dir():
        print(f"ERROR: {RELEASE_DIR} missing — run build_v1_2_portable.py first")
        return 1

    _prepare_test_dirs()
    _patch_word_dispatch_counter()
    results: list[dict] = []

    print("Test 1-2: PDF batch 50/100...")
    results.append(test_batch_pdf(50))
    results.append(test_batch_pdf(100))

    print("Test 3: Word dispatch...")
    results.append(test_word_single_dispatch())

    print("Test 4: WINWORD cleanup...")
    results.append(test_winword_closed())

    print("Test 5-8: portable exe...")
    results.append(test_pdf_mode(RELEASE_DIR))
    results.append(test_docx_mode(RELEASE_DIR))
    results.append(test_signature(RELEASE_DIR))
    results.append(test_json_runtime(RELEASE_DIR))
    results.append(test_docx_runtime(RELEASE_DIR))

    print("Test 9: invalid P...")
    results.append(test_invalid_p())

    print("Test 10-16: RTL, one page, L/O, date, editable, writable, check...")
    results.append(test_one_page(RELEASE_DIR))
    results.append(test_rtl_body(RELEASE_DIR))
    results.append(test_lo_conditions(RELEASE_DIR))
    results.append(test_date_field(RELEASE_DIR))
    results.append(test_editable_table(RELEASE_DIR))
    results.append(test_template_writable(RELEASE_DIR))
    results.append(test_app_check(RELEASE_DIR))

    print("Sanitizing release + creating ZIP...")
    create_zip()

    write_report(results)
    failed = [r for r in results if not r.get("ok")]
    if failed:
        print("FAILED:", [r["name"] for r in failed])
        return 1
    print(f"All tests passed. ZIP: {ZIP_PATH}")
    print(f"Report: {REPORT_PATH}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
