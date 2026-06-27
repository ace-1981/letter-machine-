"""V1.1 regression tests before new app.exe build."""

from __future__ import annotations

import json
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from src.letter_generator import generate_single_letter
from src.output_format import OUTPUT_DOCX, OUTPUT_PDF
from src.signature_field import verify_signature_field

CONFIG = ROOT / "templates" / "תחשיב זכויות אישי.json"
TEMPLATE = ROOT / "templates" / "תחשיב זכויות אישי.docx"
SAMPLE = ROOT / "samples" / "sample_data.xlsx"
REPORT = ROOT.parent / "cursor" / "דוח-V1.1-שינויים.md"


def _read_pdf_text(pdf: Path) -> str:
    import fitz

    parts = []
    with fitz.open(str(pdf)) as doc:
        for page in doc:
            parts.append(page.get_text())
    return "\n".join(parts)


def _read_docx_text(docx: Path) -> str:
    with zipfile.ZipFile(docx, "r") as zf:
        return zf.read("word/document.xml").decode("utf-8", errors="replace")


def _row_excel(values: dict, out: Path) -> Path:
    import pandas as pd

    out.mkdir(parents=True, exist_ok=True)
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
    df = pd.DataFrame([cols])
    path = out / "row.xlsx"
    df.to_excel(path, index=False)
    return path


def _gen(row_values: dict, out: Path, *, fmt: str, row_index: int = 0) -> dict:
    excel = _row_excel(row_values, out)
    return generate_single_letter(
        excel_path=excel,
        config_path=CONFIG,
        output_dir=out,
        row_index=row_index,
        output_format=fmt,
        pdf_preferred="word",
    )


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


def test_output_formats(out: Path) -> list[dict]:
    results = []
    pdf_r = _gen(_base_row(), out / "fmt_pdf", fmt=OUTPUT_PDF)
    pdf_only = list((out / "fmt_pdf").glob("*.pdf"))
    docx_left = list((out / "fmt_pdf").glob("*.docx"))
    results.append(
        {
            "name": "PDF mode — PDF only",
            "ok": len(pdf_only) == 1 and len(docx_left) == 0,
            "files": [p.name for p in pdf_only + docx_left],
        }
    )

    docx_r = _gen(_base_row(), out / "fmt_docx", fmt=OUTPUT_DOCX)
    docx_only = list((out / "fmt_docx").glob("*.docx"))
    pdf_left = list((out / "fmt_docx").glob("*.pdf"))
    results.append(
        {
            "name": "DOCX mode — DOCX only",
            "ok": len(docx_only) == 1 and len(pdf_left) == 0,
            "files": [p.name for p in docx_only + pdf_left],
        }
    )
    results.append(
        {
            "name": "DOCX filename extension",
            "ok": str(docx_r["filename"]).endswith(".docx"),
            "filename": docx_r["filename"],
        }
    )
    return results


def test_table_and_notes(out: Path) -> list[dict]:
    results = []

    full = _gen(_base_row(L=500, O=200), out / "full", fmt=OUTPUT_PDF)
    text = _read_pdf_text(full["pdf"])
    results.append(
        {
            "name": "L>0 O>0 — death + grant rows",
            "ok": "תוספת בגין פטירה" in text
            and "קיזוז בגין מענק עידוד עבודה" in text
            and "סה\"כ אחרי קיזוז" in text.replace("״", '"'),
        }
    )
    results.append(
        {
            "name": "L>0 O>0 — notes section",
            "ok": "הערות והבהרות" in text,
        }
    )

    l0 = _gen(_base_row(L=0, O=200), out / "l0", fmt=OUTPUT_PDF)
    t0 = _read_pdf_text(l0["pdf"])
    results.append(
        {
            "name": "L=0 — no death row",
            "ok": "תוספת בגין פטירה" not in t0,
        }
    )

    o0 = _gen(_base_row(L=500, O=0, P=1600), out / "o0", fmt=OUTPUT_PDF)
    t1 = _read_pdf_text(o0["pdf"])
    results.append(
        {
            "name": "O=0 — no grant rows",
            "ok": "קיזוז בגין מענק עידוד עבודה" not in t1
            and "סה\"כ אחרי קיזוז" not in t1.replace("״", '"'),
        }
    )

    no_notes = _gen(_base_row(L=0, O=0, R="", T=""), out / "no_notes", fmt=OUTPUT_PDF)
    t2 = _read_pdf_text(no_notes["pdf"])
    results.append(
        {
            "name": "No active notes — no notes heading",
            "ok": "הערות והבהרות" not in t2,
        }
    )
    return results


def test_signature(out: Path) -> dict:
    r = _gen(_base_row(), out / "sig", fmt=OUTPUT_PDF)
    info = verify_signature_field(r["pdf"], "MemberSignature")
    return {
        "name": "Signature field interactive",
        "ok": info.get("interactive") is True,
        "info": info,
    }


def test_runtime_json_docx(out: Path) -> list[dict]:
    results = []
    config = CONFIG.read_text(encoding="utf-8")
    backup = config
    try:
        data = json.loads(config)
        data["template_name"] = data["template_name"] + " V11TEST"
        CONFIG.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        r = _gen(_base_row(), out / "json_rt", fmt=OUTPUT_PDF)
        results.append(
            {
                "name": "JSON runtime influence",
                "ok": "V11TEST" in r["filename"],
                "filename": r["filename"],
            }
        )
    finally:
        CONFIG.write_text(backup, encoding="utf-8")

    docx_backup = TEMPLATE.with_suffix(".docx.bak")
    shutil.copy2(TEMPLATE, docx_backup)
    try:
        with zipfile.ZipFile(TEMPLATE, "r") as zin, zipfile.ZipFile(
            TEMPLATE.with_suffix(".docx.tmp"), "w"
        ) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    data = data.replace(
                        "תחשיב זכויות אישי".encode("utf-8"),
                        "תחשיב זכויות אישי V11DOCX".encode("utf-8"),
                    )
                zout.writestr(item, data)
        TEMPLATE.with_suffix(".docx.tmp").replace(TEMPLATE)
        r = _gen(_base_row(), out / "docx_rt", fmt=OUTPUT_PDF)
        text = _read_pdf_text(r["pdf"])
        results.append(
            {
                "name": "DOCX runtime influence",
                "ok": "V11DOCX" in text,
            }
        )
    finally:
        shutil.copy2(docx_backup, TEMPLATE)
        docx_backup.unlink(missing_ok=True)
    return results


def write_report(all_results: list[dict]) -> None:
    passed = sum(1 for r in all_results if r.get("ok"))
    lines = [
        "# דוח V1.1 — שינויים",
        "",
        "## סיכום בדיקות",
        "",
        f"עברו: {passed}/{len(all_results)}",
        "",
        "| בדיקה | תוצאה |",
        "|-------|--------|",
    ]
    for r in all_results:
        lines.append(f"| {r['name']} | {'עבר' if r.get('ok') else 'נכשל'} |")

    lines.extend(
        [
            "",
            "## מה שונה במסך",
            "",
            "- נוספה בחירת סוג פלט: PDF (ברירת מחדל) / Word DOCX",
            "- תצוגה מקדימה פותחת את סוג הקובץ שנבחר",
            "- לוג ההפקה מציין את סוג הפלט",
            "",
            "## מה שונה בלוגיקת הטבלה",
            "",
            "- שורת \"תוספת בגין פטירה\" — רק כש-L>0",
            "- שורות מענק עידוד עבודה + סה\"כ אחרי קיזוז — רק כש-O>0",
            "- רקע מיוחד לשורות סה\"כ בלבד (ללא zebra בשורות רגילות)",
            "",
            "## מה שונה בהערות והבהרות",
            "",
            "- כותרת \"הערות והבהרות\" מוסתרת כשאין הערות פעילות",
            "- כל הערה: כותרת מודגשת + גוף רגיל",
            "",
            "## מה שונה באזור החתימה",
            "",
            "- מלבן חתימה רחב וגבוה יותר",
            "- שדה תאריך ברור: תאריך: ______________________",
            "- JSON חתימה עודכן (width/height/margins)",
            "",
            "## תוצאות PDF/DOCX",
            "",
        ]
    )
    for r in all_results:
        if "filename" in r or "files" in r or "info" in r:
            lines.append(f"- **{r['name']}:** {r}")

    lines.extend(
        [
            "",
            "## JSON / DOCX runtime",
            "",
            f"- JSON משפיע runtime: {'כן' if any(r['name']=='JSON runtime influence' and r.get('ok') for r in all_results) else 'לא'}",
            f"- DOCX משפיע runtime: {'כן' if any(r['name']=='DOCX runtime influence' and r.get('ok') for r in all_results) else 'לא'}",
            "",
            "## האם נדרש build חדש של app.exe?",
            "",
            "כן — שינויים ב-GUI ובלוגיקת הפקה דורשים `python scripts/build_release.py`.",
            "תבניות DOCX/JSON נשארות חיצוניות — ניתן לערוך בלי rebuild.",
        ]
    )
    REPORT.write_text("\n".join(lines), encoding="utf-8")
    print(f"Report: {REPORT}")


def main() -> int:
    if not TEMPLATE.is_file():
        print("Template missing — run create_template_docx.py first")
        return 1

    with tempfile.TemporaryDirectory(prefix="v11_test_") as tmp:
        out = Path(tmp)
        results: list[dict] = []
        results.extend(test_output_formats(out))
        results.extend(test_table_and_notes(out))
        results.append(test_signature(out))
        results.extend(test_runtime_json_docx(out))
        write_report(results)

        failed = [r for r in results if not r.get("ok")]
        if failed:
            print("FAILED:", [r["name"] for r in failed])
            return 1
        print("All V1.1 tests passed.")
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
