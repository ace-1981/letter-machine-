"""Adjust RTL paragraph jc for PDF export while keeping Word-friendly template XML."""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.oxml.ns import qn


def prepare_docx_for_pdf_export(docx_path: Path) -> Path:
    """
    Body Hebrew paragraphs use jc=left + bidi so Word shows right alignment.
    Word's PDF exporter needs jc=right + bidi for the same visual layout.
    Only top-level body paragraphs are updated (not table cells).
    """
    doc = Document(str(docx_path))
    for paragraph in doc.paragraphs:
        p_pr = paragraph._p.pPr
        if p_pr is None:
            continue
        if p_pr.find(qn("w:bidi")) is None:
            continue
        jc = p_pr.find(qn("w:jc"))
        if jc is None:
            continue
        if jc.get(qn("w:val")) == "left":
            jc.set(qn("w:val"), "right")
    doc.save(str(docx_path))
    return docx_path
