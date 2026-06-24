"""Portable release tests: app.exe, relocated folder, runtime JSON/DOCX."""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
RELEASE_DIR = ROOT.parent / "LetterGenerator_release"
SAMPLE_EXCEL = ROOT / "samples" / "sample_data.xlsx"
REPORT_PATH = ROOT.parent / "cursor" / "דוח-Build-Portable.md"

JSON_NAME = "תחשיב זכויות אישי.json"
DOCX_NAME = "תחשיב זכויות אישי.docx"
RUNTIME_JSON_SUFFIX = " PORTABLE_JSON_TEST"
DOCX_MARKER = "PORTABLE_DOCX_MARKER"


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


def _run_preview(exe: Path, portable: Path, excel: Path, output: Path) -> Path:
    proc = subprocess.run(
        [str(exe), "preview", str(excel), "--row", "1", "--output", str(output)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        cwd=portable,
        timeout=180,
    )
    if proc.returncode != 0:
        raise RuntimeError(f"preview failed:\n{proc.stdout}\n{proc.stderr}")
    pdf = Path(proc.stdout.strip())
    if not pdf.is_file():
        raise RuntimeError(f"PDF missing: {pdf}")
    return pdf


def copy_portable_to_temp() -> Path:
    if not RELEASE_DIR.is_dir():
        raise FileNotFoundError(f"Release missing: {RELEASE_DIR}")
    dest = Path(tempfile.mkdtemp(prefix="lg_portable_reloc_"))
    shutil.copytree(RELEASE_DIR, dest / "LetterGenerator_release")
    portable = dest / "LetterGenerator_release"
    shutil.copy2(SAMPLE_EXCEL, portable / "sample_data.xlsx")
    return portable


def test_relocated_check(portable: Path) -> dict:
    exe = portable / "app.exe"
    proc = subprocess.run(
        [str(exe), "check"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        cwd=portable,
    )
    return {
        "name": "Relocated folder — startup check",
        "location": str(portable),
        "ok": proc.returncode == 0,
        "stdout": proc.stdout.strip(),
    }


def test_json_via_exe(portable: Path) -> dict:
    exe = portable / "app.exe"
    config = portable / "templates" / JSON_NAME
    original = config.read_text(encoding="utf-8")
    out_dir = portable / "output" / "_json_test"
    try:
        data = json.loads(original)
        data["template_name"] = data["template_name"] + RUNTIME_JSON_SUFFIX
        config.write_text(
            json.dumps(data, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        pdf = _run_preview(exe, portable, portable / "sample_data.xlsx", out_dir)
        ok = RUNTIME_JSON_SUFFIX.strip() in pdf.name
        return {
            "name": "Runtime JSON via app.exe",
            "pdf": pdf.name,
            "ok": ok,
        }
    finally:
        config.write_text(original, encoding="utf-8")


def test_docx_via_exe(portable: Path) -> dict:
    exe = portable / "app.exe"
    docx = portable / "templates" / DOCX_NAME
    backup = docx.with_suffix(".docx.bak")
    shutil.copy2(docx, backup)
    out_dir = portable / "output" / "_docx_test"
    title = "תחשיב זכויות אישי"
    try:
        _replace_in_docx_xml(docx, title, DOCX_MARKER)
        pdf = _run_preview(exe, portable, portable / "sample_data.xlsx", out_dir)
        text = _read_pdf_text(pdf)
        ok = DOCX_MARKER in text or "PORTABLE_DOCX" in text
        return {
            "name": "Runtime DOCX via app.exe preview",
            "pdf": pdf.name,
            "ok": ok,
        }
    finally:
        shutil.copy2(backup, docx)
        backup.unlink(missing_ok=True)


def write_report(portable: Path, results: list[dict]) -> None:
    json_ok = any(r.get("name") == "Runtime JSON via app.exe" and r.get("ok") for r in results)
    docx_ok = any(
        r.get("name") == "Runtime DOCX via app.exe preview" and r.get("ok") for r in results
    )
    reloc_ok = any(
        r.get("name") == "Relocated folder — startup check" and r.get("ok") for r in results
    )

    lines = [
        "# דוח Build — Portable App",
        "",
        "## תיקיית הפצה",
        "",
        f"`{RELEASE_DIR}`",
        "",
        "## הפעלה",
        "",
        "1. העתיקו את כל תיקיית `LetterGenerator_release` למיקום כלשהו.",
        "2. הפעילו `app.exe` בדאבל-קליק.",
        "3. בחרו Excel ותיקיית יעד, והפיקו או השתמשו בתצוגה מקדימה.",
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
        "כן — Word מותקן נדרש במחשב להמרת DOCX ל-PDF.",
        "",
        "## תוצאות בדיקות",
        "",
        f"| בדיקה | תוצאה |",
        f"|-------|--------|",
        f"| עובד מתיקייה מועתקת (לא תיקיית פיתוח) | {'כן' if reloc_ok else 'לא'} |",
        f"| מיקום בדיקה | `{portable}` |",
        f"| שינוי JSON אחרי build משפיע | {'כן' if json_ok else 'לא'} |",
        f"| שינוי DOCX אחרי build משפיע | {'כן' if docx_ok else 'לא'} |",
        "",
        "## פירוט",
        "",
    ]
    for r in results:
        lines.append(f"### {r['name']}")
        for k, v in r.items():
            if k != "name":
                lines.append(f"- {k}: {v}")
        lines.append("")

    REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
    REPORT_PATH.write_text("\n".join(lines), encoding="utf-8")
    print(f"Report: {REPORT_PATH}")


def main() -> int:
    portable = copy_portable_to_temp()
    print(f"Testing relocated copy: {portable}")
    results = [
        test_relocated_check(portable),
        test_json_via_exe(portable),
        test_docx_via_exe(portable),
    ]
    write_report(portable, results)

    failed = [r for r in results if not r.get("ok")]
    if failed:
        print("FAILED:", failed)
        return 1
    print("All portable release tests passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
