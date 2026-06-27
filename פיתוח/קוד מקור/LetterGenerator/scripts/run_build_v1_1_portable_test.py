"""Post-build verification for LetterGenerator_V1.1_Portable + zip + report."""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
RELEASE_DIR = ROOT.parent / "LetterGenerator_V1.1_Portable"
ZIP_PATH = ROOT.parent / "LetterGenerator_V1.1_Portable.zip"
REPORT_PATH = ROOT.parent / "cursor" / "דוח-Build-V1.1-Portable.md"
DEBUG_EXE = ROOT / "dist" / "app_debug.exe"
SAMPLE_EXCEL = ROOT / "samples" / "sample_data.xlsx"

JSON_NAME = "תחשיב זכויות אישי.json"
DOCX_NAME = "תחשיב זכויות אישי.docx"
RUNTIME_JSON_SUFFIX = " V11_PORTABLE_JSON"
DOCX_MARKER = "V11_PORTABLE_DOCX"
HEBREW_FOLDER = "בדיקת מחולל V1.1"

REQUIRED = (
    "app.exe",
    f"templates/{JSON_NAME}",
    f"templates/{DOCX_NAME}",
    "output",
    "README.txt",
)

sys.path.insert(0, str(ROOT))
from src.signature_field import verify_signature_field  # noqa: E402


def _dir_size(path: Path) -> int:
    return sum(p.stat().st_size for p in path.rglob("*") if p.is_file())


def _fmt_size(n: int) -> str:
    if n >= 1024 * 1024:
        return f"{n / (1024 * 1024):.1f} MB"
    if n >= 1024:
        return f"{n / 1024:.1f} KB"
    return f"{n} B"


def _read_pdf_text(pdf_path: Path) -> str:
    import fitz

    parts: list[str] = []
    with fitz.open(str(pdf_path)) as doc:
        for page in doc:
            parts.append(page.get_text())
    return "\n".join(parts)


def _replace_in_docx_xml(docx_path: Path, old: str, new: str) -> None:
    tmp = docx_path.with_suffix(".docx.tmp")
    with zipfile.ZipFile(docx_path, "r") as zin, zipfile.ZipFile(tmp, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml" and old.encode("utf-8") in data:
                data = data.replace(old.encode("utf-8"), new.encode("utf-8"))
            zout.writestr(item, data)
    tmp.replace(docx_path)


def _run_exe(
    exe: Path,
    cwd: Path,
    args: list[str],
    *,
    timeout: int = 240,
) -> subprocess.CompletedProcess:
    return subprocess.run(
        [str(exe), *args],
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


def _make_row_excel(values: dict, path: Path) -> Path:
    import pandas as pd

    cols = {chr(65 + i): "" for i in range(20)}
    mapping = {
        "C": "MEMBER_CODE",
        "E": "LAST_NAME",
        "F": "FIRST_NAME",
        "H": "H",
        "I": "I",
        "J": "J",
        "L": "L",
        "M": "M",
        "O": "O",
        "P": "P",
        "R": "R",
        "S": "BANK_ACCOUNT",
        "T": "T",
    }
    for col, key in mapping.items():
        if key in values:
            cols[col] = values[key]
    path.parent.mkdir(parents=True, exist_ok=True)
    pd.DataFrame([cols]).to_excel(path, index=False)
    return path


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


def _preview(
    portable: Path,
    excel: Path,
    out_dir: Path,
    *,
    fmt: str = "pdf",
) -> Path:
    proc = _run_exe(
        portable / "app.exe",
        portable,
        [
            "preview",
            str(excel),
            "--row",
            "1",
            "--output",
            str(out_dir),
            "--format",
            fmt,
        ],
    )
    out, err = _decode(proc)
    if proc.returncode != 0:
        raise RuntimeError(f"preview failed ({fmt}): {err or out}")
    result = Path(out.strip())
    if not result.is_file():
        raise RuntimeError(f"output missing: {out.strip()}")
    return result


def check_structure() -> dict:
    missing = [rel for rel in REQUIRED if not (RELEASE_DIR / rel.replace("/", os.sep)).exists()]
    app_exe = RELEASE_DIR / "app.exe"
    return {
        "id": 0,
        "name": "מבנה תיקיית Portable",
        "ok": RELEASE_DIR.is_dir() and not missing,
        "missing": missing,
        "path": str(RELEASE_DIR),
    }


def check_no_hardcoded_paths() -> dict:
    pattern = re.compile(r"[Cc]:\\Users\\dfusb")
    hits = [
        str(p.relative_to(ROOT))
        for p in (ROOT / "src").rglob("*.py")
        if pattern.search(p.read_text(encoding="utf-8"))
    ]
    return {
        "id": 0,
        "name": "ללא נתיבים קשיחים ב-src",
        "ok": not hits,
        "hits": hits,
    }


def check_double_click() -> dict:
    exe = RELEASE_DIR / "app.exe"
    proc = subprocess.Popen(
        [str(exe)],
        cwd=RELEASE_DIR,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    time.sleep(4)
    running = proc.poll() is None
    if running:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            proc.kill()
    return {
        "id": 1,
        "name": "פתיחת app.exe בדאבל-קליק",
        "ok": running,
    }


def _setup_portable_copy() -> Path:
    stage = Path(tempfile.mkdtemp(prefix="v11_portable_test_"))
    portable = stage / "LetterGenerator_V1.1_Portable"
    shutil.copytree(RELEASE_DIR, portable)
    shutil.copy2(SAMPLE_EXCEL, portable / "sample_data.xlsx")
    return portable


def check_pdf_mode(portable: Path) -> dict:
    out = portable / "output" / "_pdf_mode"
    out.mkdir(parents=True, exist_ok=True)
    excel = _make_row_excel(_base_row(), portable / "test_rows" / "base.xlsx")
    _preview(portable, excel, out, fmt="pdf")
    pdfs = list(out.glob("*.pdf"))
    docxs = list(out.glob("*.docx"))
    return {
        "id": 2,
        "name": "מצב PDF — PDF בלבד",
        "ok": len(pdfs) == 1 and len(docxs) == 0,
        "files": [p.name for p in pdfs + docxs],
    }


def check_docx_mode(portable: Path) -> dict:
    out = portable / "output" / "_docx_mode"
    out.mkdir(parents=True, exist_ok=True)
    excel = _make_row_excel(_base_row(), portable / "test_rows" / "base2.xlsx")
    path = _preview(portable, excel, out, fmt="docx")
    pdfs = list(out.glob("*.pdf"))
    docxs = list(out.glob("*.docx"))
    return {
        "id": 3,
        "name": "מצב DOCX — Word בלבד",
        "ok": len(docxs) == 1 and len(pdfs) == 0 and path.suffix.lower() == ".docx",
        "files": [p.name for p in pdfs + docxs],
    }


def check_preview_both(portable: Path) -> dict:
    excel = _make_row_excel(_base_row(), portable / "test_rows" / "preview.xlsx")
    pdf_out = portable / "output" / "_preview_pdf"
    docx_out = portable / "output" / "_preview_docx"
    pdf_out.mkdir(parents=True, exist_ok=True)
    docx_out.mkdir(parents=True, exist_ok=True)
    pdf_path = _preview(portable, excel, pdf_out, fmt="pdf")
    docx_path = _preview(portable, excel, docx_out, fmt="docx")
    return {
        "id": 4,
        "name": "Preview בשני המצבים",
        "ok": pdf_path.suffix.lower() == ".pdf" and docx_path.suffix.lower() == ".docx",
        "pdf": pdf_path.name,
        "docx": docx_path.name,
    }


def check_l0(portable: Path) -> dict:
    out = portable / "output" / "_l0"
    out.mkdir(parents=True, exist_ok=True)
    excel = _make_row_excel(_base_row(L=0, O=200), portable / "test_rows" / "l0.xlsx")
    pdf = _preview(portable, excel, out, fmt="pdf")
    text = _read_pdf_text(pdf)
    return {
        "id": 5,
        "name": "L=0 מסתיר שורת תוספת פטירה",
        "ok": "תוספת בגין פטירה" not in text,
    }


def check_o0(portable: Path) -> dict:
    out = portable / "output" / "_o0"
    out.mkdir(parents=True, exist_ok=True)
    excel = _make_row_excel(_base_row(L=500, O=0, P=1600), portable / "test_rows" / "o0.xlsx")
    pdf = _preview(portable, excel, out, fmt="pdf")
    text = _read_pdf_text(pdf).replace("״", '"')
    return {
        "id": 6,
        "name": "O=0 מסתיר מענק עידוד וסה״כ אחרי קיזוז",
        "ok": "קיזוז בגין מענק עידוד עבודה" not in text
        and 'סה"כ אחרי קיזוז' not in text,
    }


def check_no_notes(portable: Path) -> dict:
    out = portable / "output" / "_no_notes"
    out.mkdir(parents=True, exist_ok=True)
    excel = _make_row_excel(
        _base_row(L=0, O=0, R="", T=""),
        portable / "test_rows" / "no_notes.xlsx",
    )
    pdf = _preview(portable, excel, out, fmt="pdf")
    text = _read_pdf_text(pdf)
    return {
        "id": 7,
        "name": "ללא הערות — ללא כותרת הערות והבהרות",
        "ok": "הערות והבהרות" not in text,
    }


def check_signature(portable: Path) -> dict:
    out = portable / "output" / "_sig"
    out.mkdir(parents=True, exist_ok=True)
    excel = _make_row_excel(_base_row(), portable / "test_rows" / "sig.xlsx")
    pdf = _preview(portable, excel, out, fmt="pdf")
    info = verify_signature_field(pdf, "MemberSignature")
    return {
        "id": 8,
        "name": "שדה חתימה אינטראקטיבי",
        "ok": info.get("interactive") is True,
        "info": info,
    }


def check_json_runtime(portable: Path) -> dict:
    config = portable / "templates" / JSON_NAME
    original = config.read_text(encoding="utf-8")
    out = portable / "output" / "_json_rt"
    out.mkdir(parents=True, exist_ok=True)
    excel = _make_row_excel(_base_row(), portable / "test_rows" / "json.xlsx")
    try:
        data = json.loads(original)
        data["template_name"] = data["template_name"] + RUNTIME_JSON_SUFFIX
        config.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        pdf = _preview(portable, excel, out, fmt="pdf")
        ok = RUNTIME_JSON_SUFFIX.strip() in pdf.name
        return {"id": 9, "name": "JSON runtime test אחרי build", "ok": ok, "pdf": pdf.name}
    finally:
        config.write_text(original, encoding="utf-8")


def check_docx_runtime(portable: Path) -> dict:
    docx = portable / "templates" / DOCX_NAME
    backup = docx.with_suffix(".docx.bak")
    shutil.copy2(docx, backup)
    out = portable / "output" / "_docx_rt"
    out.mkdir(parents=True, exist_ok=True)
    excel = _make_row_excel(_base_row(), portable / "test_rows" / "docx.xlsx")
    try:
        _replace_in_docx_xml(docx, "תחשיב זכויות אישי", DOCX_MARKER)
        pdf = _preview(portable, excel, out, fmt="pdf")
        text = _read_pdf_text(pdf)
        return {
            "id": 10,
            "name": "DOCX runtime test אחרי build",
            "ok": DOCX_MARKER in text,
            "pdf": pdf.name,
        }
    finally:
        shutil.copy2(backup, docx)
        backup.unlink(missing_ok=True)


def check_relocated(portable: Path) -> dict:
    parent = Path(tempfile.mkdtemp(prefix="lg_reloc_v11_"))
    dest = parent / "LetterGenerator_V1.1_Portable"
    shutil.copytree(portable, dest)
    proc = _run_exe(dest / "app.exe", dest, ["check"])
    out, err = _decode(proc)
    preview = _run_exe(
        dest / "app.exe",
        dest,
        ["preview", str(dest / "sample_data.xlsx"), "--row", "1", "--format", "pdf"],
    )
    pout, _ = _decode(preview)
    return {
        "id": 11,
        "name": "העתקת תיקייה למיקום חדש",
        "ok": proc.returncode == 0 and preview.returncode == 0 and Path(pout.strip()).is_file(),
        "location": str(dest),
        "stdout": out.strip(),
    }


def check_hebrew_path() -> dict:
    parent = Path(tempfile.gettempdir()) / HEBREW_FOLDER
    if parent.exists():
        shutil.rmtree(parent)
    parent.mkdir(parents=True)
    dest = parent / "LetterGenerator_V1.1_Portable"
    shutil.copytree(RELEASE_DIR, dest)
    shutil.copy2(SAMPLE_EXCEL, dest / "sample_data.xlsx")
    proc = _run_exe(dest / "app.exe", dest, ["check"])
    out, _ = _decode(proc)
    preview = _run_exe(
        dest / "app.exe",
        dest,
        ["preview", str(dest / "sample_data.xlsx"), "--row", "1", "--format", "pdf"],
    )
    pout, perr = _decode(preview)
    pdf_ok = preview.returncode == 0 and Path(pout.strip()).is_file()
    return {
        "id": 12,
        "name": "בדיקה מנתיב עברי",
        "ok": proc.returncode == 0 and pdf_ok,
        "location": str(dest),
        "stdout": out.strip(),
    }


def check_debug_smoke() -> dict:
    if not DEBUG_EXE.is_file():
        return {"name": "app_debug.exe smoke", "ok": True, "skipped": "no debug exe"}
    with tempfile.TemporaryDirectory(prefix="lg_dbg_") as tmp:
        stage = Path(tmp) / "portable"
        shutil.copytree(RELEASE_DIR, stage)
        shutil.copy2(DEBUG_EXE, stage / "app_debug.exe")
        shutil.copy2(SAMPLE_EXCEL, stage / "sample_data.xlsx")
        check = _run_exe(stage / "app_debug.exe", stage, ["check"])
        if check.returncode != 0:
            out, err = _decode(check)
            return {"name": "app_debug.exe smoke", "ok": False, "stderr": err, "stdout": out}
        preview = _run_exe(
            stage / "app_debug.exe",
            stage,
            ["preview", str(stage / "sample_data.xlsx"), "--row", "1", "--format", "pdf"],
        )
        out, err = _decode(preview)
        pdf = Path(out.strip())
        return {"name": "app_debug.exe smoke", "ok": preview.returncode == 0 and pdf.is_file()}


def create_zip() -> dict:
    if ZIP_PATH.is_file():
        ZIP_PATH.unlink()
    with zipfile.ZipFile(ZIP_PATH, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in RELEASE_DIR.rglob("*"):
            if path.is_file():
                zf.write(path, path.relative_to(RELEASE_DIR.parent))
    return {
        "name": "ZIP",
        "path": str(ZIP_PATH),
        "size": _fmt_size(ZIP_PATH.stat().st_size),
        "ok": ZIP_PATH.is_file(),
    }


def write_report(results: list[dict], zip_info: dict, sizes: dict) -> None:
    numbered = [r for r in results if r.get("id")]
    passed = sum(1 for r in numbered if r.get("ok"))
    lines = [
        "# דוח Build V1.1 — Portable",
        "",
        f"**תאריך:** {time.strftime('%Y-%m-%d %H:%M')}",
        f"**תיקייה:** `{RELEASE_DIR}`",
        f"**ZIP:** `{ZIP_PATH}`",
        "",
        "## סיכום",
        "",
        f"בדיקות 1–12: **{passed}/12 עברו**",
        "",
        "| # | בדיקה | תוצאה |",
        "|---|-------|--------|",
    ]
    for r in sorted(numbered, key=lambda x: x["id"]):
        lines.append(f"| {r['id']} | {r['name']} | {'עבר' if r.get('ok') else 'נכשל'} |")

    extra = [r for r in results if not r.get("id")]
    if extra:
        lines.extend(["", "## בדיקות נוספות", ""])
        for r in extra:
            lines.append(f"- **{r['name']}:** {'עבר' if r.get('ok') else 'נכשל'}")

    lines.extend(
        [
            "",
            "## מבנה Portable",
            "",
            "```",
            "LetterGenerator_V1.1_Portable/",
            "├── app.exe",
            "├── templates/",
            f"│   ├── {DOCX_NAME}",
            f"│   └── {JSON_NAME}",
            "├── output/",
            "└── README.txt",
            "```",
            "",
            "## עקרונות שנשמרו",
            "",
            "- אין Installer",
            "- DOCX ו-JSON חיצוניים ב-templates/",
            "- שינוי תבנית אחרי build ללא rebuild",
            "- ללא נתיבים קשיחים ב-src",
            "",
            "## גדלים",
            "",
            f"- `app.exe`: {_fmt_size(sizes['app_exe'])} ({sizes['app_exe']:,} bytes)",
            f"- תיקיית Portable: {_fmt_size(sizes['release_total'])}",
            f"- `{ZIP_PATH.name}`: {zip_info.get('size', '—')}",
            "",
            "## פירוט",
            "",
        ]
    )
    for r in results:
        lines.append(f"### {r.get('name', 'test')}")
        for k, v in r.items():
            if k not in ("name", "id", "ok"):
                lines.append(f"- {k}: {v}")
        lines.append("")

    REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
    REPORT_PATH.write_text("\n".join(lines), encoding="utf-8")
    print(f"Report: {REPORT_PATH}")


def main() -> int:
    if not RELEASE_DIR.is_dir():
        print(f"ERROR: {RELEASE_DIR} not found. Run scripts/build_v1_1_portable.py first.")
        return 1

    portable = _setup_portable_copy()
    results: list[dict] = [
        check_structure(),
        check_no_hardcoded_paths(),
        check_double_click(),
        check_pdf_mode(portable),
        check_docx_mode(portable),
        check_preview_both(portable),
        check_l0(portable),
        check_o0(portable),
        check_no_notes(portable),
        check_signature(portable),
        check_json_runtime(portable),
        check_docx_runtime(portable),
        check_relocated(portable),
        check_hebrew_path(),
        check_debug_smoke(),
    ]

    app_exe = RELEASE_DIR / "app.exe"
    sizes = {
        "app_exe": app_exe.stat().st_size if app_exe.is_file() else 0,
        "release_total": _dir_size(RELEASE_DIR),
    }
    zip_info = create_zip()
    write_report(results, zip_info, sizes)

    failed = [r for r in results if not r.get("ok")]
    if failed:
        print("FAILED:", [r.get("name") for r in failed])
        return 1
    print("All V1.1 portable build checks passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
