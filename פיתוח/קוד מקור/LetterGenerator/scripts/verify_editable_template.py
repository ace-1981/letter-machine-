"""Verify templates/תחשיב זכויות אישי.docx is editable and drives PDF output."""

from __future__ import annotations

import os
import stat
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from scripts.proof_editable_template import (  # noqa: E402
    ORIGINAL_TOTAL,
    TEMPLATE,
    TEST_TOTAL,
    proof_spacing_edit,
    proof_text_edit,
    _pdf_text,
    _preview_pdf,
)

PHRASES = (
    "מספר שנים לצורך התחשיב",
    "סכום בגין שנת ותק אחת",
    "סה\"כ בגין ותק",
    "תוספת בגין פטירה לכל זכאי",
    "סה\"כ זכאות",
    "קיזוז בגין מענק עידוד עבודה",
    "מענק עידוד עבודה",
)


def inspect_docx(path: Path) -> dict:
    info: dict = {}
    if os.name == "nt":
        info["readonly_os"] = bool(os.stat(path).st_file_attributes & stat.FILE_ATTRIBUTE_READONLY)
    else:
        info["readonly_os"] = False
    with zipfile.ZipFile(path) as z:
        settings = z.read("word/settings.xml").decode("utf-8")
        doc = z.read("word/document.xml").decode("utf-8")
    info["documentProtection"] = "documentProtection" in settings
    info["writeProtection"] = "writeProtection" in settings
    info["readOnlyRecommended"] = "readOnlyRecommended" in settings
    info["calc_table_placeholder"] = "calc_table" in doc
    info["phrases"] = {p: p in doc for p in PHRASES}
    return info


def main() -> int:
    if not TEMPLATE.is_file():
        print("ERROR: template missing", TEMPLATE)
        return 1

    print("=== DOCX inspection ===")
    info = inspect_docx(TEMPLATE)
    for k, v in info.items():
        if k != "phrases":
            print(f"{k}: {v}")
    for p, ok in info["phrases"].items():
        print(f"phrase[{p}]: {'OK' if ok else 'MISSING'}")

    print("\n=== Proof 1: text edit without rebuild ===")
    ok_text = proof_text_edit()
    print("text edit proof:", ok_text)

    print("\n=== Proof 2: restore original text ===")
    text2 = _pdf_text(_preview_pdf())
    ok_restore = ORIGINAL_TOTAL in text2 and "זכאותTEST" not in text2.replace(" ", "")
    print("original in PDF:", ORIGINAL_TOTAL in text2)
    print("TEST removed:", "זכאותTEST" not in text2.replace(" ", ""))

    print("\n=== Proof 3: spacing edit without rebuild ===")
    ok_spacing = proof_spacing_edit()
    print("spacing edit proof:", ok_spacing)

    failed = []
    if info.get("readonly_os") or info.get("documentProtection") or info.get("writeProtection"):
        failed.append("read-only/protected")
    if info.get("calc_table_placeholder"):
        failed.append("calc_table placeholder still in DOCX")
    if not all(info["phrases"].values()):
        failed.append("missing table phrases in DOCX")
    if not ok_text:
        failed.append("text edit proof failed")
    if not ok_restore:
        failed.append("restore proof failed")
    if not ok_spacing:
        failed.append("spacing edit proof failed")

    if failed:
        print("\nFAILED:", failed)
        return 1
    print("\nAll editable-template checks passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
