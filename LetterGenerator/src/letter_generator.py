"""Batch letter generation with per-row error handling."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from src.config_loader import get_template_dir, load_template_config
from src.errors_report import write_errors_report
from src.excel_reader import read_excel, row_to_context
from src.pdf_converter import PdfConverter, PdfConverterFactory
from src.signature_field import add_signature_field
from src.template_engine import build_output_filename, render_template
from src.validator import validate_all, validate_row


@dataclass
class GenerationResult:
    total: int
    success: int
    errors: list[dict]


def generate_letters(
    excel_path: Path,
    config_path: Path,
    output_dir: Path,
    pdf_preferred: str | None = None,
    converter: PdfConverter | None = None,
    on_progress: Callable[[int, int], None] | None = None,
) -> GenerationResult:
    config = load_template_config(config_path)
    template_dir = get_template_dir(config_path)
    template_docx = template_dir / config["template_file"]

    validation = validate_all(excel_path, config, template_docx, output_dir, pdf_preferred)
    if not validation.ok:
        raise ValueError("Validation failed:\n" + "\n".join(validation.errors))

    df = read_excel(excel_path)
    pdf_converter = converter or PdfConverterFactory.create(pdf_preferred)
    sig_cfg = config["signature_field"]

    errors: list[dict] = []
    success = 0

    for row_index in range(len(df)):
        excel_row_number = row_index + 2  # 1-based with header
        try:
            _generate_single_row(
                df=df,
                row_index=row_index,
                config=config,
                template_docx=template_docx,
                output_dir=output_dir,
                pdf_converter=pdf_converter,
                sig_cfg=sig_cfg,
            )
            success += 1
        except Exception as exc:
            errors.append(
                {
                    "excel_row": excel_row_number,
                    "error": str(exc),
                }
            )
        if on_progress:
            on_progress(row_index + 1, len(df))

    if errors:
        write_errors_report(output_dir / "errors_report.csv", errors)

    return GenerationResult(total=len(df), success=success, errors=errors)


def generate_single_letter(
    excel_path: Path,
    config_path: Path,
    output_dir: Path,
    row_index: int = 0,
    pdf_preferred: str | None = None,
    keep_docx: bool = True,
) -> dict:
    """POC helper: generate one letter from one Excel row."""
    config = load_template_config(config_path)
    template_dir = get_template_dir(config_path)
    template_docx = template_dir / config["template_file"]

    validation = validate_all(excel_path, config, template_docx, output_dir, pdf_preferred)
    if not validation.ok:
        raise ValueError("Validation failed:\n" + "\n".join(validation.errors))

    df = read_excel(excel_path)
    pdf_converter = PdfConverterFactory.create(pdf_preferred)
    sig_cfg = config["signature_field"]

    paths = _generate_single_row(
        df=df,
        row_index=row_index,
        config=config,
        template_docx=template_docx,
        output_dir=output_dir,
        pdf_converter=pdf_converter,
        sig_cfg=sig_cfg,
        keep_docx=keep_docx,
    )
    paths["validation_warnings"] = validation.warnings
    paths["pdf_converter"] = pdf_converter.name
    return paths


def _generate_single_row(
    df,
    row_index: int,
    config: dict,
    template_docx: Path,
    output_dir: Path,
    pdf_converter: PdfConverter,
    sig_cfg: dict,
    keep_docx: bool = True,
) -> dict:
    context = row_to_context(df, row_index, config)
    excel_row_number = row_index + 2
    row_errors = validate_row(context, config, excel_row_number)
    if row_errors:
        raise ValueError(row_errors[0] if len(row_errors) == 1 else "; ".join(row_errors))

    filename = build_output_filename(context, config)
    pdf_path = output_dir / filename
    docx_path = output_dir / filename.replace(".pdf", ".docx")
    temp_pdf = output_dir / f"_temp_{filename}"

    render_template(template_docx, docx_path, context)
    pdf_converter.convert(docx_path, temp_pdf)

    add_signature_field(
        pdf_path=temp_pdf,
        output_path=pdf_path,
        field_name=sig_cfg["field_name"],
        width=sig_cfg["width"],
        height=sig_cfg["height"],
        margin_bottom=sig_cfg["margin_bottom"],
        margin_left=sig_cfg["margin_left"],
        page=sig_cfg.get("page", "last"),
        anchor_text=sig_cfg.get("anchor_text"),
    )
    temp_pdf.unlink(missing_ok=True)

    if not keep_docx:
        docx_path.unlink(missing_ok=True)

    return {
        "docx": docx_path,
        "pdf": pdf_path,
        "context": context,
        "filename": filename,
    }
