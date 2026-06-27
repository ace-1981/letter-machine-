"""POC: Excel → DOCX → PDF → signature field → named output."""

from __future__ import annotations

import json
import sys
from pathlib import Path

# Windows console Hebrew output
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from scripts.create_sample_excel import create_sample_excel
from scripts.create_template import create_template
from src.letter_generator import generate_single_letter
from src.pdf_converter import PdfConverterFactory
from src.signature_field import verify_signature_field


def main() -> int:
    templates_dir = ROOT / "templates"
    samples_dir = ROOT / "samples"
    output_dir = samples_dir / "poc_output"

    config_path = templates_dir / "תחשיב זכויות אישי.json"
    template_docx = templates_dir / "תחשיב זכויות אישי.docx"
    excel_path = samples_dir / "sample_data.xlsx"

    print("=== Letter Generator POC ===\n")

    print("1. Creating sample Excel...")
    create_sample_excel(excel_path)

    print("2. Creating DOCX template...")
    create_template(template_docx)

    print("3. PDF converter status:")
    for line in PdfConverterFactory.list_status():
        print(f"   - {line}")

    print("\n4. Running validation + generation (row 0)...")
    result = generate_single_letter(
        excel_path=excel_path,
        config_path=config_path,
        output_dir=output_dir,
        row_index=0,
        pdf_preferred="word",
        keep_docx=True,
    )

    config = json.loads(config_path.read_text(encoding="utf-8"))
    field_name = config["signature_field"]["field_name"]
    verification = verify_signature_field(result["pdf"], field_name)

    print("\n=== POC Results ===")
    print(f"DOCX:  {result['docx']}")
    print(f"PDF:   {result['pdf']}")
    print(f"Name:  {result['filename']}")
    print(f"Converter: {result['pdf_converter']}")
    if result.get("validation_warnings"):
        print("Warnings:")
        for w in result["validation_warnings"]:
            print(f"  - {w}")

    print("\nSignature field verification (programmatic):")
    print(json.dumps(verification, ensure_ascii=False, indent=2))

    report_path = ROOT.parent / "cursor" / "POC-תוצאות.md"
    _write_poc_report(report_path, result, verification, config)
    print(f"\nReport: {report_path}")

    if not verification.get("is_signature_field"):
        print("\nERROR: Signature field verification failed.")
        return 1

    print("\nPOC completed successfully.")
    print("Please open the PDF in Adobe Reader or Foxit and verify 'Fill & Sign' shows a signature field.")
    return 0


def _write_poc_report(path: Path, result: dict, verification: dict, config: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    content = f"""# POC — תוצאות

## קבצים שנוצרו

| קובץ | נתיב |
|------|------|
| Excel לדוגמה | `LetterGenerator/samples/sample_data.xlsx` |
| DOCX זמני | `{result['docx']}` |
| PDF סופי | `{result['pdf']}` |
| JSON טכני | `LetterGenerator/templates/תחשיב זכויות אישי.json` |

## שם קובץ פלט

`{result['filename']}`

פורמט: קוד חבר + שם משפחה + שם פרטי + שם המכתב (לפי האפיון).

## ממיר PDF

{result['pdf_converter']}

## אימות שדה חתימה (אוטומטי)

```json
{json.dumps(verification, ensure_ascii=False, indent=2)}
```

שדה החתימה נוצר עם **pyHanko** כשדה AcroForm מסוג `/Sig` (לא מלבן רגיל).

### בדיקה ידנית ב-Adobe Reader / Foxit

1. פתח את `{result['pdf']}`
2. ב-Adobe Reader: **Fill & Sign** (מילוי וחתימה) → אמור להופיע שדה חתימה בשם `{config['signature_field']['field_name']}`
3. ב-Foxit: **Fill & Sign** → Sign Document → בחר את שדה החתימה
4. חתום (חתימה דיגיטלית או ציור) — השדה אמור להיות אינטראקטיבי

## תלויות

ראה `LetterGenerator/requirements.txt`

## שלב הבא

לאחר אישור POC — בניית ממשק PySide6 לפי האפיון.
"""
    path.write_text(content, encoding="utf-8")


if __name__ == "__main__":
    raise SystemExit(main())
