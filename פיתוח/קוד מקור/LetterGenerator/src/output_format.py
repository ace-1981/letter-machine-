"""Output format constants and normalization for PDF/DOCX generation."""

from __future__ import annotations

OUTPUT_PDF = "pdf"
OUTPUT_DOCX = "docx"

_VALID = frozenset({OUTPUT_PDF, OUTPUT_DOCX})


def normalize_output_format(value: str | None) -> str:
    if not value:
        return OUTPUT_PDF
    fmt = str(value).strip().lower()
    if fmt in _VALID:
        return fmt
    raise ValueError(f"Unsupported output format: {value!r} (expected pdf or docx)")


def format_label(value: str) -> str:
    fmt = normalize_output_format(value)
    return "PDF" if fmt == OUTPUT_PDF else "DOCX"
