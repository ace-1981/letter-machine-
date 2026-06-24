"""Runtime tests: external JSON/DOCX in LetterGenerator_release affect output without rebuild."""

from __future__ import annotations

import json
import os
import shutil
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from src.letter_generator import generate_single_letter

RELEASE_DIR = ROOT.parent / "LetterGenerator_release"
EXCEL = ROOT / "samples" / "sample_data.xlsx"
JSON_NAME = "תחשיב זכויות אישי.json"
DOCX_NAME = "תחשיב זכויות אישי.docx"
RUNTIME_JSON_SUFFIX = " RUNTIME_JSON_TEST"
DOCX_MARKER = "RT_DOCX_MARKER_9X7"


def _release_paths() -> tuple[Path, Path, Path, Path]:
    if not RELEASE_DIR.is_dir():
        raise FileNotFoundError(f"Release folder missing: {RELEASE_DIR}")
    exe = RELEASE_DIR / "app.exe"
    if not exe.is_file():
        raise FileNotFoundError(f"app.exe missing: {exe}")
    templates = RELEASE_DIR / "templates"
    config = templates / JSON_NAME
    docx = templates / DOCX_NAME
    return exe, templates, config, docx


def _generate_preview(config_path: Path, output_dir: Path) -> Path:
    os.environ["LETTER_GEN_APP_ROOT"] = str(RELEASE_DIR)
    output_dir.mkdir(parents=True, exist_ok=True)
    for old in output_dir.glob("*.pdf"):
        old.unlink()
    result = generate_single_letter(
        excel_path=EXCEL,
        config_path=config_path,
        output_dir=output_dir,
        row_index=0,
        pdf_preferred="word",
        keep_docx=False,
    )
    return Path(result["pdf"])


def _read_pdf_text(pdf_path: Path) -> str:
    import fitz

    parts: list[str] = []
    with fitz.open(str(pdf_path)) as doc:
        for page in doc:
            parts.append(page.get_text())
    return "\n".join(parts)


def test_json_runtime() -> dict:
    _, _, config_path, _ = _release_paths()
    original = config_path.read_text(encoding="utf-8")
    backup = original
    out_dir = RELEASE_DIR / "output" / "_runtime_json_test"
    try:
        data = json.loads(original)
        baseline_name = data["template_name"]
        data["template_name"] = baseline_name + RUNTIME_JSON_SUFFIX
        config_path.write_text(
            json.dumps(data, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )

        pdf_path = _generate_preview(config_path, out_dir)
        ok = RUNTIME_JSON_SUFFIX.strip() in pdf_path.name
        return {
            "name": "JSON runtime",
            "pdf": pdf_path.name,
            "ok": ok,
            "detail": f"Expected suffix in filename: {RUNTIME_JSON_SUFFIX.strip()}",
        }
    finally:
        config_path.write_text(backup, encoding="utf-8")


def _replace_in_docx_xml(docx_path: Path, old: str, new: str) -> None:
    tmp = docx_path.with_suffix(".docx.tmp")
    with zipfile.ZipFile(docx_path, "r") as zin, zipfile.ZipFile(tmp, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml" and old.encode("utf-8") in data:
                data = data.replace(old.encode("utf-8"), new.encode("utf-8"))
            zout.writestr(item, data)
    tmp.replace(docx_path)


def test_docx_runtime() -> dict:
    _, _, config_path, docx_path = _release_paths()
    backup_docx = docx_path.with_suffix(".docx.bak")
    shutil.copy2(docx_path, backup_docx)
    out_dir = RELEASE_DIR / "output" / "_runtime_docx_test"
    title_fragment = "תחשיב זכויות אישי"
    try:
        _replace_in_docx_xml(docx_path, title_fragment, DOCX_MARKER)
        pdf_path = _generate_preview(config_path, out_dir)
        text = _read_pdf_text(pdf_path)
        ok = DOCX_MARKER in text or "RT_DOCX" in text
        return {
            "name": "DOCX runtime",
            "pdf": pdf_path.name,
            "ok": ok,
            "detail": f"Expected marker in PDF text: {DOCX_MARKER}",
        }
    finally:
        if backup_docx.is_file():
            shutil.copy2(backup_docx, docx_path)
            backup_docx.unlink()


def test_startup_paths() -> dict:
    os.environ["LETTER_GEN_APP_ROOT"] = str(RELEASE_DIR)
    from src.app_paths import get_app_root, get_templates_dir
    from src.startup_check import check_startup_templates

    templates = get_templates_dir()
    json_ok = (templates / JSON_NAME).is_file()
    docx_ok = (templates / DOCX_NAME).is_file()
    errors = check_startup_templates()
    return {
        "name": "Startup paths",
        "app_root": str(get_app_root()),
        "templates_dir": str(templates),
        "json_ok": json_ok,
        "docx_ok": docx_ok,
        "startup_errors": errors,
        "ok": json_ok and docx_ok and not errors,
    }


def write_report(results: list[dict], path: Path) -> None:
    lines = [
        "# דוח Build — תבניות חיצוניות (External Templates)",
        "",
        f"**תיקיית Release:** `{RELEASE_DIR}`",
        "",
        "## סיכום",
        "",
        "| בדיקה | תוצאה |",
        "|-------|--------|",
    ]
    for r in results:
        if "startup_errors" in r:
            lines.append(
                f"| JSON נטען מתיקיית templates | {'כן' if r['json_ok'] else 'לא'} |"
            )
            lines.append(
                f"| DOCX נטען מתיקיית templates | {'כן' if r['docx_ok'] else 'לא'} |"
            )
            lines.append(
                f"| בדיקת Startup (ללא שגיאות) | {'כן' if r['ok'] else 'לא'} |"
            )
        else:
            status = "עבר" if r.get("ok") else "נכשל"
            lines.append(f"| {r['name']} | {status} |")

    lines.extend(
        [
            "",
            "## שינוי JSON אחרי build",
            "",
            f"- {'משפיע על שם קובץ PDF' if any(r.get('name') == 'JSON runtime' and r.get('ok') for r in results) else 'לא אומת'}",
            "",
            "## שינוי DOCX אחרי build",
            "",
            f"- {'משפיע על תוכן PDF' if any(r.get('name') == 'DOCX runtime' and r.get('ok') for r in results) else 'לא אומת'}",
            "",
            "## קבצים לשמירה ליד app.exe",
            "",
            "- `app.exe`",
            f"- `templates/{JSON_NAME}`",
            f"- `templates/{DOCX_NAME}`",
            "- `output/` (תיקיית פלט)",
            "- `README.txt`",
            "",
            "## פירוט בדיקות",
            "",
        ]
    )
    for r in results:
        lines.append(f"### {r.get('name', 'test')}")
        for k, v in r.items():
            if k != "name":
                lines.append(f"- {k}: {v}")
        lines.append("")

    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines), encoding="utf-8")
    print(f"Report: {path}")


def main() -> int:
    results: list[dict] = []
    results.append(test_startup_paths())
    results.append(test_json_runtime())
    results.append(test_docx_runtime())

    report = ROOT.parent / "cursor" / "דוח-Build-External-Templates.md"
    write_report(results, report)

    failed = [r for r in results if not r.get("ok")]
    if failed:
        print("FAILED:", failed)
        return 1
    print("All runtime external template tests passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
