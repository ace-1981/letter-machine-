"""Full portable build verification and final report."""

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
RELEASE_DIR = ROOT.parent / "LetterGenerator_release"
DEBUG_EXE = ROOT / "dist" / "app_debug.exe"
SAMPLE_EXCEL = ROOT / "samples" / "sample_data.xlsx"
REPORT_PATH = ROOT.parent / "cursor" / "דוח-Build-Portable.md"
BUILD_LOG = ROOT.parent / "cursor" / "build_log.txt"

JSON_NAME = "תחשיב זכויות אישי.json"
DOCX_NAME = "תחשיב זכויות אישי.docx"
RUNTIME_JSON_SUFFIX = " FINAL_JSON_TEST"
DOCX_MARKER = "FINAL_DOCX_MARKER"
HEBREW_FOLDER = "בדיקת מחולל מכתבים"

REQUIRED = (
    "app.exe",
    f"templates/{JSON_NAME}",
    f"templates/{DOCX_NAME}",
    "output",
    "README.txt",
)


def _dir_size(path: Path) -> int:
    total = 0
    for p in path.rglob("*"):
        if p.is_file():
            total += p.stat().st_size
    return total


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
    env: dict | None = None,
    timeout: int = 180,
) -> subprocess.CompletedProcess:
    full_env = {**os.environ, **(env or {}), "PYTHONUTF8": "1"}
    return subprocess.run(
        [str(exe), *args],
        capture_output=True,
        cwd=cwd,
        timeout=timeout,
        env=full_env,
    )


def _decode_output(proc: subprocess.CompletedProcess) -> tuple[str, str]:
    return (
        proc.stdout.decode("utf-8", errors="replace"),
        proc.stderr.decode("utf-8", errors="replace"),
    )


def check_release_structure() -> dict:
    missing = []
    present = {}
    for rel in REQUIRED:
        p = RELEASE_DIR / rel.replace("/", os.sep)
        if p.exists():
            present[rel] = p.stat().st_size if p.is_file() else "dir"
        else:
            missing.append(rel)
    return {
        "name": "Release folder structure",
        "folder_exists": RELEASE_DIR.is_dir(),
        "path": str(RELEASE_DIR),
        "present": present,
        "missing": missing,
        "ok": RELEASE_DIR.is_dir() and not missing,
    }


def check_no_hardcoded_dev_paths() -> dict:
    pattern = re.compile(r"[Cc]:\\Users\\dfusb")
    hits: list[str] = []
    for py in (ROOT / "src").rglob("*.py"):
        if pattern.search(py.read_text(encoding="utf-8")):
            hits.append(str(py.relative_to(ROOT)))
    return {
        "name": "No hardcoded C:\\Users\\dfusb paths in runtime src",
        "hits": hits,
        "ok": not hits,
    }


def check_debug_smoke() -> dict:
    if not DEBUG_EXE.is_file():
        return {"name": "app_debug.exe smoke test", "ok": False, "detail": "app_debug.exe missing"}

    with tempfile.TemporaryDirectory(prefix="lg_smoke_") as tmp:
        stage = Path(tmp) / "portable"
        shutil.copytree(RELEASE_DIR, stage)
        shutil.copy2(DEBUG_EXE, stage / "app_debug.exe")
        shutil.copy2(SAMPLE_EXCEL, stage / "sample_data.xlsx")

        check = _run_exe(stage / "app_debug.exe", stage, ["check"])
        out, err = _decode_output(check)
        if check.returncode != 0:
            return {
                "name": "app_debug.exe smoke test",
                "ok": False,
                "step": "check",
                "stdout": out,
                "stderr": err,
            }

        preview = _run_exe(
            stage / "app_debug.exe",
            stage,
            ["preview", str(stage / "sample_data.xlsx"), "--row", "1"],
        )
        out, err = _decode_output(preview)
        if preview.returncode != 0:
            return {
                "name": "app_debug.exe smoke test",
                "ok": False,
                "step": "preview",
                "stdout": out,
                "stderr": err,
            }
        pdf = Path(out.strip())
        return {
            "name": "app_debug.exe smoke test",
            "ok": pdf.is_file(),
            "pdf": pdf.name if pdf.is_file() else out.strip(),
        }


def check_app_double_click() -> dict:
    exe = RELEASE_DIR / "app.exe"
    if not exe.is_file():
        return {"name": "app.exe double-click launch", "ok": False, "detail": "app.exe missing"}

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
        "name": "app.exe double-click launch",
        "ok": running,
        "pid_started": running,
        "note": "GUI נפתח ונשאר פעיל — נסגר אוטומטית לאחר הבדיקה",
    }


def _copy_release(dest_parent: Path, folder_name: str) -> Path:
    dest = dest_parent / folder_name
    if dest.exists():
        shutil.rmtree(dest)
    shutil.copytree(RELEASE_DIR, dest)
    shutil.copy2(SAMPLE_EXCEL, dest / "sample_data.xlsx")
    return dest


def test_json_runtime(portable: Path) -> dict:
    config = portable / "templates" / JSON_NAME
    original = config.read_text(encoding="utf-8")
    out_dir = portable / "output" / "_final_json_test"
    try:
        data = json.loads(original)
        data["template_name"] = data["template_name"] + RUNTIME_JSON_SUFFIX
        config.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        proc = _run_exe(
            portable / "app.exe",
            portable,
            ["preview", str(portable / "sample_data.xlsx"), "--row", "1", "--output", str(out_dir)],
        )
        out, err = _decode_output(proc)
        if proc.returncode != 0:
            return {"name": "Runtime JSON (template_name)", "ok": False, "stderr": err, "stdout": out}
        pdf = Path(out.strip())
        ok = RUNTIME_JSON_SUFFIX.strip() in pdf.name
        return {"name": "Runtime JSON (template_name)", "ok": ok, "pdf": pdf.name}
    finally:
        config.write_text(original, encoding="utf-8")


def test_docx_runtime(portable: Path) -> dict:
    docx = portable / "templates" / DOCX_NAME
    backup = docx.with_suffix(".docx.bak")
    shutil.copy2(docx, backup)
    out_dir = portable / "output" / "_final_docx_test"
    try:
        _replace_in_docx_xml(docx, "תחשיב זכויות אישי", DOCX_MARKER)
        proc = _run_exe(
            portable / "app.exe",
            portable,
            ["preview", str(portable / "sample_data.xlsx"), "--row", "1", "--output", str(out_dir)],
        )
        out, err = _decode_output(proc)
        if proc.returncode != 0:
            return {"name": "Runtime DOCX (text change)", "ok": False, "stderr": err, "stdout": out}
        pdf = Path(out.strip())
        text = _read_pdf_text(pdf)
        ok = DOCX_MARKER in text
        return {"name": "Runtime DOCX (text change)", "ok": ok, "pdf": pdf.name}
    finally:
        shutil.copy2(backup, docx)
        backup.unlink(missing_ok=True)


def test_relocated_ascii_path() -> dict:
    parent = Path(tempfile.mkdtemp(prefix="lg_reloc_ascii_"))
    portable = _copy_release(parent, "LetterGenerator_release")
    proc = _run_exe(portable / "app.exe", portable, ["check"])
    out, err = _decode_output(proc)
    json_r = test_json_runtime(portable)
    return {
        "name": "Relocated folder (ASCII temp path)",
        "location": str(portable),
        "startup_ok": proc.returncode == 0,
        "json_runtime_ok": json_r.get("ok"),
        "ok": proc.returncode == 0 and json_r.get("ok"),
        "stdout": out.strip(),
    }


def test_hebrew_path() -> dict:
    parent = Path(tempfile.gettempdir()) / HEBREW_FOLDER
    if parent.exists():
        shutil.rmtree(parent)
    parent.mkdir(parents=True)
    portable = _copy_release(parent, "LetterGenerator_release")
    proc = _run_exe(portable / "app.exe", portable, ["check"])
    out, err = _decode_output(proc)
    preview = _run_exe(
        portable / "app.exe",
        portable,
        ["preview", str(portable / "sample_data.xlsx"), "--row", "1"],
    )
    pout, perr = _decode_output(preview)
    pdf_ok = preview.returncode == 0 and Path(pout.strip()).is_file()
    return {
        "name": f"Hebrew path ({HEBREW_FOLDER})",
        "location": str(portable),
        "startup_ok": proc.returncode == 0,
        "preview_ok": pdf_ok,
        "ok": proc.returncode == 0 and pdf_ok,
        "stdout": out.strip(),
    }


def test_word_unavailable() -> dict:
    """Test error handling via app.py (same code as app.exe; env hooks in pdf_converter)."""
    shutil.copy2(SAMPLE_EXCEL, RELEASE_DIR / "sample_data.xlsx")
    proc = subprocess.run(
        [
            sys.executable,
            str(ROOT / "app.py"),
            "preview",
            str(RELEASE_DIR / "sample_data.xlsx"),
            "--row",
            "1",
        ],
        capture_output=True,
        cwd=ROOT,
        timeout=60,
        env={
            **os.environ,
            "LETTER_GEN_APP_ROOT": str(RELEASE_DIR),
            "LETTER_GEN_TEST_NO_WORD": "1",
            "LETTER_GEN_TEST_NO_LIBREOFFICE": "1",
            "PYTHONUTF8": "1",
        },
    )
    out = proc.stdout.decode("utf-8", errors="replace")
    err = proc.stderr.decode("utf-8", errors="replace")
    combined = f"{out}\n{err}"
    clear = proc.returncode != 0 and (
        "Word" in combined
        or "PDF converter" in combined
        or "not available" in combined
        or "No PDF converter" in combined
    )
    return {
        "name": "Word unavailable — clear error, no crash",
        "exit_code": proc.returncode,
        "method": "python app.py (same runtime code as app.exe)",
        "message_snippet": combined.strip()[:500],
        "ok": clear,
    }


def write_report(results: list[dict], sizes: dict) -> None:
    def ok(name: str) -> bool:
        return any(r.get("name") == name and r.get("ok") for r in results)

    structure = next(r for r in results if r["name"] == "Release folder structure")
    hardcoded = next(
        (r for r in results if "hardcoded" in r.get("name", "").lower()),
        {"ok": False},
    )

    lines = [
        "# דוח Build סופי — Portable App",
        "",
        "## 1. תיקיית הפצה",
        "",
        f"**נוצרה:** {'כן' if structure.get('folder_exists') else 'לא'}",
        f"**נתיב:** `{RELEASE_DIR}`",
        "",
        "## 2. קבצים בתיקייה",
        "",
        "| קובץ/תיקייה | קיים |",
        "|-------------|------|",
    ]
    for rel in REQUIRED:
        exists = rel not in structure.get("missing", [])
        lines.append(f"| `{rel}` | {'כן' if exists else 'לא'} |")

    lines.extend(
        [
            "",
            "## 3. app_debug.exe + smoke test",
            "",
            f"**עבר:** {'כן' if ok('app_debug.exe smoke test') else 'לא'}",
            "",
            "## 4. app.exe — דאבל-קליק",
            "",
            f"**נפתח בהצלחה:** {'כן' if ok('app.exe double-click launch') else 'לא'}",
            "",
            "## 5. שינוי JSON אחרי build (template_name → שם PDF)",
            "",
            f"**משפיע:** {'כן' if ok('Runtime JSON (template_name)') else 'לא'}",
            "",
            "## 6. שינוי DOCX אחרי build (טקסט → תוכן PDF)",
            "",
            f"**משפיע:** {'כן' if ok('Runtime DOCX (text change)') else 'לא'}",
            "",
            "## 7. העתקה למיקום חדש",
            "",
            f"**עובד:** {'כן' if ok('Relocated folder (ASCII temp path)') else 'לא'}",
            "",
            "## 8. נתיב עברי",
            "",
            f"**תיקייה:** `{HEBREW_FOLDER}`",
            f"**עובד:** {'כן' if ok(f'Hebrew path ({HEBREW_FOLDER})') else 'לא'}",
            "",
            "## 9. נתיבים קשיחים ל-C:\\Users\\dfusb",
            "",
            f"**לא נמצאו ב-src:** {'כן' if hardcoded.get('ok') else 'לא'}",
            "",
            "## 10. Word לא זמין",
            "",
            f"**הודעת שגיאה ברורה (לא קריסה):** {'כן' if ok('Word unavailable — clear error, no crash') else 'לא'}",
            "",
            "## 11. גדלים",
            "",
            f"- `app.exe`: {_fmt_size(sizes['app_exe'])} ({sizes['app_exe']:,} bytes)",
            f"- תיקיית `LetterGenerator_release` כולה: {_fmt_size(sizes['release_total'])} ({sizes['release_total']:,} bytes)",
            "",
            "## הפעלה",
            "",
            "1. העתיקו את כל `LetterGenerator_release` למיקום כלשהו (כולל נתיב עברי).",
            "2. דאבל-קליק על `app.exe`.",
            "3. בחרו Excel ותיקיית יעד (`output/` כברירת מחדל).",
            "",
            "## קבצים חובה ליד app.exe",
            "",
            "- `app.exe`",
            f"- `templates/{JSON_NAME}`",
            f"- `templates/{DOCX_NAME}`",
            "- `output/`",
            "- `README.txt`",
            "",
            "## Microsoft Word",
            "",
            "כן — Word מותקן נדרש להמרת DOCX ל-PDF.",
            "",
            "## פירוט בדיקות",
            "",
        ]
    )
    for r in results:
        lines.append(f"### {r['name']}")
        for k, v in r.items():
            if k not in ("name",):
                lines.append(f"- {k}: {v}")
        lines.append("")

    REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
    REPORT_PATH.write_text("\n".join(lines), encoding="utf-8")
    print(f"Report: {REPORT_PATH}")


def main() -> int:
    if not RELEASE_DIR.is_dir():
        print("ERROR: LetterGenerator_release not found. Run build_release.py first.")
        return 1

    portable = _copy_release(Path(tempfile.mkdtemp(prefix="lg_final_")), "LetterGenerator_release")

    results = [
        check_release_structure(),
        check_no_hardcoded_dev_paths(),
        check_debug_smoke(),
        check_app_double_click(),
        test_json_runtime(portable),
        test_docx_runtime(portable),
        test_relocated_ascii_path(),
        test_hebrew_path(),
        test_word_unavailable(),
    ]

    app_exe = RELEASE_DIR / "app.exe"
    sizes = {
        "app_exe": app_exe.stat().st_size if app_exe.is_file() else 0,
        "release_total": _dir_size(RELEASE_DIR),
    }

    write_report(results, sizes)

    failed = [r for r in results if not r.get("ok")]
    if failed:
        print("FAILED:", [r["name"] for r in failed])
        return 1
    print("All final build checks passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
