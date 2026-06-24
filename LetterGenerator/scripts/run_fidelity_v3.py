"""Build template + sample PDF and verify digital signature field."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

from scripts.build_template_with_word import create_fidelity_template
from scripts.create_sample_excel import create_sample_excel
from scripts.create_fidelity_preview_image import render_pdf_preview
from src.config_loader import load_template_config
from src.letter_generator import generate_single_letter
from src.signature_field import verify_signature_field


def main() -> int:
    templates = ROOT / "templates"
    config_path = templates / "תחשיב זכויות אישי.json"
    template_path = templates / "תחשיב זכויות אישי.docx"
    out = ROOT / "samples" / "fidelity_output"
    excel = ROOT / "samples" / "sample_data.xlsx"

    print("=== Fidelity attempt v4 (Word-native RTL + signature) ===\n")
    print("1. Building template with Word COM...")
    create_fidelity_template(template_path)

    create_sample_excel(excel)
    print("2. Generating DOCX + PDF...")
    result = generate_single_letter(excel, config_path, out, row_index=0, pdf_preferred="word")

    config = load_template_config(config_path)
    field = config["signature_field"]["field_name"]
    verification = verify_signature_field(result["pdf"], field)
    placement = _check_signature_placement(result["pdf"], field)

    import fitz

    pages = len(fitz.open(str(result["pdf"])))

    preview = render_pdf_preview(result["pdf"], out / "preview_page1.png")
    marked = out / "preview_signature_marked.png"
    _mark_signature_on_preview(result["pdf"], marked, field)

    print("\n=== Results ===")
    print("Pages:", pages)
    print("DOCX:", result["docx"])
    print("PDF:", result["pdf"])
    print("Preview:", preview)
    print("Marked:", marked)
    print("\nSignature verification:")
    print(json.dumps(verification, ensure_ascii=False, indent=2))
    print("\nSignature placement:")
    print(json.dumps(placement, ensure_ascii=False, indent=2))

    _open_pdf_readers(result["pdf"])

    report = ROOT.parent / "cursor" / "דוח-ניסיון-v4-RTL-וחתימה.md"
    _write_report(report, result, verification, placement, pages, marked)
    print("\nReport:", report)

    ok = verification.get("is_signature_field") and placement.get("ok", False)
    return 0 if ok else 1


def _check_signature_placement(pdf_path: Path, field_name: str) -> dict:
    import fitz

    verification = verify_signature_field(pdf_path, field_name)
    rect = verification.get("rect")
    if not rect:
        return {"ok": False, "message": "No widget rect found"}

    doc = fitz.open(str(pdf_path))
    page = doc[0]
    patterns = [
        "חתימה (שדה לחתימה דיגיטלית)",
        ")חתימה (שדה לחתימה דיגיטלית",
        "שדה לחתימה דיגיטלית",
    ]
    anchor = None
    for pattern in patterns:
        hits = page.search_for(pattern)
        if hits:
            anchor = max(hits, key=lambda r: r.y0)
            break

    sig = fitz.Rect(rect)
    result = {
        "ok": False,
        "signature_rect": [float(sig.x0), float(sig.y0), float(sig.x1), float(sig.y1)],
        "anchor_rect": None,
        "below_anchor": False,
        "message": "",
    }

    if anchor is None:
        result["message"] = "Anchor text not found; using margin fallback"
        result["ok"] = sig.y0 > page.rect.height * 0.45
        doc.close()
        return result

    result["anchor_rect"] = [float(anchor.x0), float(anchor.y0), float(anchor.x1), float(anchor.y1)]
    below = sig.y0 >= anchor.y1 - 2
    horizontal_overlap = sig.x0 <= anchor.x1 + 20 and sig.x1 >= anchor.x0 - 20
    result["below_anchor"] = below
    result["horizontal_overlap"] = horizontal_overlap
    result["ok"] = below and horizontal_overlap
    result["message"] = (
        "Signature field is inside the signature box area"
        if result["ok"]
        else "Signature field is misaligned relative to anchor text"
    )
    doc.close()
    return result


def _mark_signature_on_preview(pdf_path: Path, output_path: Path, field_name: str) -> None:
    import fitz

    doc = fitz.open(str(pdf_path))
    page = doc[0]
    fields = verify_signature_field(pdf_path, field_name)
    rect = fields.get("rect")
    if rect:
        r = fitz.Rect(rect)
        page.draw_rect(r, color=(1, 0, 0), width=2)
        page.insert_text(
            fitz.Point(r.x0, max(12, r.y0 - 4)),
            "SIGN",
            fontsize=8,
            color=(1, 0.4, 0),
        )
    pix = page.get_pixmap(dpi=150)
    pix.save(str(output_path))
    doc.close()


def _open_pdf_readers(pdf_path: Path) -> None:
    viewers = [
        Path(r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"),
        Path(r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"),
        Path(r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"),
        Path(r"C:\Program Files (x86)\Foxit Software\Foxit PDF Reader\FoxitPDFReader.exe"),
    ]
    for v in viewers:
        if v.exists():
            subprocess.Popen([str(v), str(pdf_path)])
            print("Opened with:", v)
            return
    subprocess.Popen(["cmd", "/c", "start", "", str(pdf_path)], shell=True)
    print("Opened with default PDF viewer")


def _write_report(path, result, verification, placement, pages, marked) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        "\n".join(
            [
                "# ניסיון v4 — Word-native RTL + חתימה",
                "",
                "## שינויים",
                "",
                "- תיקון בניית תבנית ב-Word COM (מסגרת עמוד, טבלאות, חתימה)",
                "- תיקון קואורדינטות שדה חתימה לפי עוגן טקסט (RTL)",
                "- אימות מיקום שדה החתימה ביחס למלבן התכלת",
                "",
                f"## פלט",
                "",
                f"- עמודים: **{pages}**",
                f"- PDF: `{result['pdf']}`",
                f"- תצוגה: `preview_page1.png`",
                f"- סימון שדה חתימה: `{marked}`",
                "",
                "## אימות חתימה",
                "",
                "```json",
                json.dumps(verification, ensure_ascii=False, indent=2),
                "```",
                "",
                "## מיקום שדה",
                "",
                "```json",
                json.dumps(placement, ensure_ascii=False, indent=2),
                "```",
                "",
                "## בדיקה ידנית",
                "",
                "1. פתח את ה-PDF",
                "2. Fill & Sign / חתימה",
                "3. ודא שניתן לחתום בשדה `MemberSignature` בתוך המלבן התכלת",
                "",
            ]
        ),
        encoding="utf-8",
    )


if __name__ == "__main__":
    raise SystemExit(main())
