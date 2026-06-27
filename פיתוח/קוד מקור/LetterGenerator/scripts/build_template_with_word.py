"""
Template build: python-docx for content, optional Word COM post-process for native RTL.
"""

from __future__ import annotations

import time
from pathlib import Path

GRAY_BORDER = 191 + 191 * 256 + 191 * 65536
WD_BORDER_TOP = -1
WD_BORDER_LEFT = -2
WD_BORDER_BOTTOM = -3
WD_BORDER_RIGHT = -4
WD_LINE_STYLE_SINGLE = 1
WD_LINE_WIDTH_050PT = 4
WD_READING_ORDER_RTL = 1


def apply_word_rtl(docx_path: Path) -> None:
    import win32com.client

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = None
    try:
        doc = word.Documents.Open(str(docx_path.resolve()), ReadOnly=False)
        for paragraph in doc.Paragraphs:
            fmt = paragraph.Range.ParagraphFormat
            fmt.ReadingOrder = WD_READING_ORDER_RTL
            paragraph.Range.LanguageID = 1037

        for table in doc.Tables:
            table.Range.ParagraphFormat.ReadingOrder = WD_READING_ORDER_RTL
            try:
                table.Direction = 1
            except Exception:
                pass

        section = doc.Sections(1)
        section.Borders.AlwaysInFront = True
        section.Borders.DistanceFromTop = 4
        section.Borders.DistanceFromBottom = 4
        section.Borders.DistanceFromLeft = 4
        section.Borders.DistanceFromRight = 4
        for border_type in (WD_BORDER_TOP, WD_BORDER_LEFT, WD_BORDER_BOTTOM, WD_BORDER_RIGHT):
            border = section.Borders(border_type)
            border.LineStyle = WD_LINE_STYLE_SINGLE
            border.LineWidth = WD_LINE_WIDTH_050PT
            border.Color = GRAY_BORDER

        doc.Save()
    finally:
        if doc is not None:
            doc.Close(False)
        word.Quit()
        time.sleep(0.3)


def create_fidelity_template(output_path: Path) -> Path:
    from scripts.create_template_docx import build_with_python_docx

    build_with_python_docx(output_path)
    try:
        apply_word_rtl(output_path)
        print("Applied Word RTL post-process")
    except Exception as exc:
        print(f"Word RTL post-process skipped ({exc})")
    return output_path


if __name__ == "__main__":
    out = Path(__file__).resolve().parent.parent / "templates" / "תחשיב זכויות אישי.docx"
    print(create_fidelity_template(out))
