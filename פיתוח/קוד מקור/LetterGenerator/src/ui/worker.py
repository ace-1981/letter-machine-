"""Background worker for letter generation."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QObject, Signal

from src.letter_generator import generate_letters, generate_single_letter
from src.output_format import OUTPUT_PDF, format_label, normalize_output_format


class GenerationWorker(QObject):
    log = Signal(str)
    error_line = Signal(str)
    progress = Signal(int, int)
    finished = Signal(dict)
    failed = Signal(str)

    def __init__(
        self,
        mode: str,
        excel_path: Path,
        config_path: Path,
        output_dir: Path,
        row_index: int = 0,
        output_format: str = OUTPUT_PDF,
    ):
        super().__init__()
        self.mode = mode
        self.excel_path = excel_path
        self.config_path = config_path
        self.output_dir = output_dir
        self.row_index = row_index
        self.output_format = normalize_output_format(output_format)

    def run(self) -> None:
        try:
            fmt_label = format_label(self.output_format)
            if self.mode == "preview":
                self.log.emit(f"מתחיל תצוגה מקדימה ({fmt_label})...")
                result = generate_single_letter(
                    excel_path=self.excel_path,
                    config_path=self.config_path,
                    output_dir=self.output_dir,
                    row_index=self.row_index,
                    output_format=self.output_format,
                    pdf_preferred="word",
                )
                out_path = result.get("pdf") or result.get("docx")
                self.log.emit(f"נוצר {fmt_label}: {result['filename']}")
                self.finished.emit(
                    {
                        "mode": "preview",
                        "total": 1,
                        "success": 1,
                        "errors": 0,
                        "file": str(out_path) if out_path else "",
                        "output_format": self.output_format,
                    }
                )
                return

            self.log.emit(f"מתחיל הפקה ({fmt_label})...")
            from src.excel_reader import read_excel

            df = read_excel(self.excel_path)
            total = len(df)
            self.progress.emit(0, total)

            def on_progress(current: int, total_rows: int) -> None:
                self.progress.emit(current, total_rows)
                self.log.emit(f"מעבד שורה {current} מתוך {total_rows} ({fmt_label})...")

            batch_result = generate_letters(
                excel_path=self.excel_path,
                config_path=self.config_path,
                output_dir=self.output_dir,
                output_format=self.output_format,
                pdf_preferred="word",
                on_progress=on_progress,
            )

            for err in batch_result.errors:
                self.error_line.emit(f"שורה {err['excel_row']} - {err['error']}")

            self.log.emit(
                f"הפקה הסתיימה ({fmt_label}). סה\"כ: {batch_result.total}, "
                f"הצליחו: {batch_result.success}, שגיאות: {len(batch_result.errors)}"
            )
            self.finished.emit(
                {
                    "mode": "batch",
                    "total": batch_result.total,
                    "success": batch_result.success,
                    "errors": len(batch_result.errors),
                    "output_dir": str(self.output_dir),
                    "output_format": self.output_format,
                    "output_format_label": fmt_label,
                }
            )
        except Exception as exc:
            self.failed.emit(str(exc))
