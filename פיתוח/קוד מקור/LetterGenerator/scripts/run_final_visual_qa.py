"""Final Visual QA — generate 3 PDF/DOCX/preview sets for manual approval."""

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

import pandas as pd

from scripts.build_template_with_word import create_fidelity_template
from scripts.create_fidelity_preview_image import render_pdf_preview
from src.config_loader import load_template_config
from src.letter_generator import generate_single_letter
from src.signature_field import verify_signature_field

CONFIG = ROOT / "templates" / "תחשיב זכויות אישי.json"
TEMPLATE = ROOT / "templates" / "תחשיב זכויות אישי.docx"
OUT = ROOT / "samples" / "final_visual_qa"
EXCEL_DIR = OUT / "excel"


def _row(
    code,
    last,
    first,
    bank,
    *,
    h=15,
    i=5000,
    j=75000,
    l=2000,
    m=82000,
    o=1500,
    p=80500,
    r="חוב בנייה",
    t="נקלט.ת",
) -> dict:
    return {
        "A": "",
        "B": "",
        "C": code,
        "D": "",
        "E": last,
        "F": first,
        "G": "",
        "H": h,
        "I": i,
        "J": j,
        "K": "",
        "L": l,
        "M": m,
        "N": "",
        "O": o,
        "P": p,
        "Q": "",
        "R": r,
        "S": bank,
        "T": t,
    }


def _save_excel(path: Path, rows: list[dict]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(rows).to_excel(path, index=False, engine="openpyxl")
    return path


def _pdf_checks(pdf_path: Path) -> dict:
    import fitz

    with fitz.open(str(pdf_path)) as doc:
        text = doc[0].get_text()
        pages = len(doc)
    sig = verify_signature_field(pdf_path, "MemberSignature")
    return {
        "pages": pages,
        "has_title": "תחשיב זכויות אישי" in text,
        "has_bank_header": "חשבון בנק" in text,
        "amount_format_ok": " ₪" in text and "₪ {{" not in text,
        "signature_interactive": sig.get("interactive"),
        "signature_found": sig.get("found"),
    }


def main() -> int:
    print("=== Final Visual QA / RTL Alignment Pass ===\n")
    print("1. Rebuilding template...")
    create_fidelity_template(TEMPLATE)

    scenarios = [
        {
            "id": "case_full",
            "title": "שורה רגילה — כל הערכים והתנאים",
            "excel": _save_excel(EXCEL_DIR / "case_full.xlsx", [_row(12345, "כהן", "ישראל", "12-345-678901")]),
            "row": 0,
        },
        {
            "id": "case_conditions_off",
            "title": "תנאים כבויים — L=0, O=0, R ריק",
            "excel": _save_excel(
                EXCEL_DIR / "case_conditions_off.xlsx",
                [_row(20001, "לוי", "דנה", "22-111-222333", l=0, o=0, m=80000, p=80000, r="", t="פעיל")],
            ),
            "row": 0,
        },
        {
            "id": "case_edge",
            "title": "שם ארוך, מספרים גדולים, קיזוז שלילי",
            "excel": _save_excel(
                EXCEL_DIR / "case_edge.xlsx",
                [
                    _row(
                        30099,
                        "בן אברהם הכהן לנדאו",
                        "ישראל-מנשה",
                        "99-888-777666",
                        h=42,
                        i=125000,
                        j=2500000,
                        l=50000,
                        m=2675000,
                        o=-12500,
                        p=2662500,
                        r="",
                        t="פעיל",
                    )
                ],
            ),
            "row": 0,
        },
    ]

    results = []
    for sc in scenarios:
        print(f"2. Generating {sc['id']}...")
        case_out = OUT / sc["id"]
        case_out.mkdir(parents=True, exist_ok=True)
        gen = generate_single_letter(sc["excel"], CONFIG, case_out, row_index=sc["row"], pdf_preferred="word")
        preview = render_pdf_preview(gen["pdf"], case_out / "preview_page1.png")
        checks = _pdf_checks(gen["pdf"])
        results.append(
            {
                "id": sc["id"],
                "title": sc["title"],
                "excel": str(sc["excel"]),
                "docx": str(gen["docx"]),
                "pdf": str(gen["pdf"]),
                "preview": str(preview) if preview else None,
                "checks": checks,
            }
        )
        print(f"   PDF: {gen['pdf'].name}")
        print(f"   Preview: {preview}")

    report_path = ROOT.parent / "cursor" / "דוח-Final-Visual-QA.md"
    _write_report(report_path, results)
    print(f"\nReport: {report_path}")
    print("\n*** המתן לאישור ידני לפני app.exe ***")
    return 0


def _write_report(path: Path, results: list[dict]) -> None:
    lines = [
        "# דוח Final Visual QA / RTL Alignment Pass",
        "",
        "**סטטוס:** ממתין לאישור ידני לפני `app.exe`",
        "",
        "## שינויים שבוצעו",
        "",
        "- יישור RTL מלא בכל הפסקאות",
        "- עמודת סכומים מיושרת **לימין** בכל השורות",
        "- פורמט סכום אחיד: `12,345 ₪` (כולל מינוס ו-0)",
        "- שורת חשבון בנק בכותרת העליונה",
        "- ריווח צפוף יותר במקטעים מותנים",
        "",
        "## קבצים לבדיקה ידנית",
        "",
    ]

    for r in results:
        c = r["checks"]
        lines.extend(
            [
                f"### {r['id']} — {r['title']}",
                "",
                f"- Excel: `{r['excel']}`",
                f"- DOCX: `{r['docx']}`",
                f"- PDF: `{r['pdf']}`",
                f"- Preview: `{r['preview']}`",
                "",
                "**בדיקות אוטומטיות:**",
                f"- עמודים: {c['pages']}",
                f"- כותרת: {'✅' if c['has_title'] else '❌'}",
                f"- חשבון בנק בכותרת: {'✅' if c['has_bank_header'] else '❌'}",
                f"- פורמט סכום: {'✅' if c['amount_format_ok'] else '❌'}",
                f"- חתימה אינטראקטיבית: {'✅' if c['signature_interactive'] else '❌'}",
                "",
                "**בדיקה ידנית נדרשת:**",
                "- כל הטקסט מיושר לימין",
                "- סכומים מיושרים לימין בטבלה",
                "- אין קפיצות LTR/RTL",
                "- דמיון לתמונת הדוגמה",
                "",
            ]
        )

    lines.extend(
        [
            "## רשימת אישור (למלא ידנית)",
            "",
            "- [ ] המסמך נראה כמו הדוגמה",
            "- [ ] כל הטקסט מיושר לימין",
            "- [ ] הסכומים מיושרים לימין",
            "- [ ] אין בעיות RTL",
            "- [ ] החתימה עובדת ב-Acrobat",
            "",
            "לאחר סימון כל הסעיפים — ניתן לעבור ל-`app.exe`.",
            "",
        ]
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines), encoding="utf-8")

    summary = OUT / "qa_summary.json"
    summary.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    raise SystemExit(main())
