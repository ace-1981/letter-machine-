"""Regenerate fidelity template and produce sample DOCX + PDF."""

from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

from scripts.create_fidelity_preview_image import render_pdf_preview
from scripts.create_sample_excel import create_sample_excel
from scripts.build_template_with_word import create_fidelity_template
from src.letter_generator import generate_single_letter
from src.validator import validate_all
from src.config_loader import load_template_config


def main() -> int:
    templates = ROOT / "templates"
    config_path = templates / "תחשיב זכויות אישי.json"
    template_path = templates / "תחשיב זכויות אישי.docx"
    output_dir = ROOT / "samples" / "fidelity_output"

    print("=== Visual Template Fidelity ===\n")
    print("1. Building fidelity DOCX template...")
    create_fidelity_template(template_path)

    config = load_template_config(config_path)
    print("2. Validating template...")
    excel = ROOT / "samples" / "sample_data.xlsx"
    create_sample_excel(excel)
    validation = validate_all(excel, config, template_path, output_dir, "word")
    if not validation.ok:
        print("Validation FAILED:")
        for err in validation.errors:
            print(" -", err)
        return 1
    for warn in validation.warnings:
        print(" warn:", warn)

    print("3. Generating sample DOCX + PDF...")
    result = generate_single_letter(
        excel_path=excel,
        config_path=config_path,
        output_dir=output_dir,
        row_index=0,
        pdf_preferred="word",
        keep_docx=True,
    )

    preview_image = render_pdf_preview(result["pdf"], output_dir / "preview_page1.png")

    print("\n=== Output ===")
    print("DOCX:", result["docx"])
    print("PDF:", result["pdf"])
    if preview_image:
        print("Preview image:", preview_image)

    report_path = ROOT.parent / "cursor" / "דוח-Visual-Template-Fidelity.md"
    _write_report(report_path, result, preview_image, validation)
    print("\nReport:", report_path)
    return 0


def _write_report(path: Path, result: dict, preview_image: Path | None, validation) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    lines = [
        "# דוח Visual Template Fidelity",
        "",
        "## מקור עיצובי",
        "",
        "**לא נמצא** בפרויקט קובץ DOCX/PDF מקורי לצפייה.",
        "התבנית נבנתה על בסיס:",
        "- `מסמך אפיון – מחולל מכתבים אוטומטי (.txt` — סעיפים 11–16 (מבנה, עיצוב, חלקים)",
        "- `תחשיב זכויות אישי/תחשיב_זכויות_מעוצב.json` — נוסחי הבהרות (הוטמעו ב-DOCX בלבד)",
        "",
        "> אם קיים אצלך מסמך מקורי — יש להחליף את `templates/תחשיב זכויות אישי.docx` בגרסה המעוצבת שלך.",
        "",
        "## מה היה לא תקין בעיצוב הקודם (POC)",
        "",
        "| בעיה | פירוט |",
        "|------|--------|",
        "| תבנית טכנית | נוצרה אוטומטית ללא עיצוב מסמך אמיתי |",
        "| ללא RTL מלא | לא הוגדרו bidi/RTL ברמת מסמך |",
        "| גופן ברירת מחדל | Arial ללא התאמה לעברית מקצועית |",
        "| טבלה כבדה | `Table Grid` — נראית כמו Excel |",
        "| מבנה חלקים | חסרו כותרות משנה ברורות לכל חלק |",
        "| כתב קבלה | נוסח קצר מדי, לא לפי 4 סעיפי האפיון |",
        "| אזור חתימה | שורת טקסט בלבד, ללא מלבן ברור |",
        "| שוליים / A4 | לא הוגדרו במפורש |",
        "",
        "## מה תוקן",
        "",
        "| רכיב | תיקון |",
        "|------|--------|",
        "| A4 Portrait | שוליים 2.0–2.5 ס\"מ |",
        "| RTL | bidi ברמת מסמך ופסקאות |",
        "| גופן | David 11pt, כותרת 14pt מודגשת |",
        "| חלק 1 | תאריך, שם החבר, חשבון בנק |",
        "| חלק 2 | טבלת תחשיב מודרנית — כותרת אפורה, שורות סה\"כ מודגשות |",
        "| חלק 3 | הבהרות עם כותרות משנה + תנאים ({%p if %}) |",
        "| חלק 4 | כתב קבלה ושחרור — 4 סעיפי אישור לפי האפיון |",
        "| חלק 5 | מלבן חתימה גדול + שורת חתימה |",
        "| חלק 6 | הנחיות החזרה עם רשימה מסודרת |",
        "| משתנים | כל המשתנים נשמרו ללא שינוי |",
        "| JSON | נשאר טכני בלבד |",
        "",
        "## קבצי פלט לדוגמה",
        "",
        f"- DOCX: `{result['docx']}`",
        f"- PDF: `{result['pdf']}`",
    ]
    if preview_image and preview_image.exists():
        lines.extend(
            [
                f"- תצוגה מקדימה: `{preview_image}`",
                "",
                f"![עמוד ראשון]({preview_image.as_posix()})",
            ]
        )
    else:
        lines.append("- תצוגה מקדימה: לא נוצרה (נדרש PyMuPDF)")

    lines.extend(
        [
            "",
            "## אימות משתנים",
            "",
            "כל המשתנים הנדרשים עברו validate לפני הפקה.",
            "",
            "## שלב הבא",
            "",
            "אסף מאשר שה-PDF נראה כמו הדוגמה הרצויה → רק אז ממשיכים ל-app.exe.",
        ]
    )
    path.write_text("\n".join(lines), encoding="utf-8")


if __name__ == "__main__":
    raise SystemExit(main())
