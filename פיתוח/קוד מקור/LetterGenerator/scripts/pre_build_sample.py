"""Pre-build smoke: regenerate template, verify, produce Word/PDF samples."""

from __future__ import annotations

import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SAMPLE_OUT = ROOT / "samples" / "pre_build_sample"
EXCEL = ROOT / "samples" / "sample_data.xlsx"
CONFIG = ROOT / "templates" / "תחשיב זכויות אישי.json"


def _kill_word() -> None:
    if sys.platform != "win32":
        return
    subprocess.run(
        ["taskkill", "/IM", "WINWORD.EXE", "/F"],
        capture_output=True,
        check=False,
    )
    time.sleep(1)


def _clear_sample_outputs() -> bool:
    """Return False if an existing sample PDF is locked open."""
    if not SAMPLE_OUT.is_dir():
        return True
    ok = True
    for path in SAMPLE_OUT.glob("*"):
        if not path.is_file():
            continue
        for attempt in range(3):
            try:
                path.unlink()
                break
            except PermissionError:
                _kill_word()
                time.sleep(1)
                if attempt == 2:
                    ok = False
    return ok


def _run(cmd: list[str], label: str) -> int:
    print(f"\n=== {label} ===")
    proc = subprocess.run(cmd, cwd=ROOT, env={**dict(**__import__("os").environ), "PYTHONUTF8": "1"})
    print(f"exit {proc.returncode}")
    return proc.returncode


def main() -> int:
    steps: list[tuple[str, list[str]]] = [
        ("Regenerate DOCX template", [sys.executable, str(ROOT / "scripts" / "create_template_docx.py")]),
        ("Portable design check", [sys.executable, str(ROOT / "scripts" / "verify_portable_design.py")]),
        ("Editable template check", [sys.executable, str(ROOT / "scripts" / "verify_editable_template.py")]),
    ]
    for label, cmd in steps:
        if _run(cmd, label) != 0:
            print(f"\nFAILED at: {label}")
            return 1

    sys.path.insert(0, str(ROOT))
    from src.letter_generator import generate_single_letter

    SAMPLE_OUT.mkdir(parents=True, exist_ok=True)
    _kill_word()
    if not _clear_sample_outputs():
        from datetime import datetime

        run_dir = SAMPLE_OUT / datetime.now().strftime("run_%Y%m%d_%H%M%S")
        run_dir.mkdir(parents=True, exist_ok=True)
        print(f"Warning: prior sample files locked — writing to {run_dir}")
        sample_out = run_dir
    else:
        sample_out = SAMPLE_OUT
    print("\n=== Generate sample PDF (Word) ===")
    pdf_result = generate_single_letter(
        EXCEL, CONFIG, sample_out, output_format="pdf", pdf_preferred="word"
    )
    pdf_path = Path(pdf_result["pdf"])
    print("PDF:", pdf_path)

    print("\n=== Generate sample DOCX ===")
    docx_result = generate_single_letter(EXCEL, CONFIG, sample_out, output_format="docx")
    docx_path = Path(docx_result["docx"])
    print("DOCX:", docx_path)

    import fitz

    doc = fitz.open(str(pdf_path))
    pages = doc.page_count
    page = doc[0]
    hits = page.search_for("סעיף")
    pw = page.rect.width
    right_margin = min(pw - h.x1 for h in hits) if hits else -1
    doc.close()

    print(f"pages: {pages}")
    print(f"table header right margin (pt): {right_margin:.1f}")

    report = ROOT.parent / "cursor" / "דוח-Pre-Build-Sample.md"
    report.write_text(
        "\n".join([
            "# דוח Pre-Build — דוגמה לפני build",
            "",
            f"**PDF:** `{pdf_path}`",
            f"**DOCX:** `{docx_path}`",
            f"**עמודים:** {pages}",
            f"**שוליים ימין לטבלה (סעיף):** {right_margin:.1f} pt",
            "",
            "בדיקות שעברו:",
            "- create_template_docx",
            "- verify_portable_design",
            "- verify_editable_template",
            "- PDF via Word",
            "- DOCX output",
        ]),
        encoding="utf-8",
    )
    print("\nReport:", report)
    print("\nAll pre-build checks passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
