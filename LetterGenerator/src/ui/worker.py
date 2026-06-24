"""Background worker for letter generation."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QObject, Signal

from src.letter_generator import generate_letters, generate_single_letter


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
    ):
        super().__init__()
        self.mode = mode
        self.excel_path = excel_path
        self.config_path = config_path
        self.output_dir = output_dir
        self.row_index = row_index

    def run(self) -> None:
        try:
            if self.mode == "preview":
                self.log.emit("מתחיל תצוגה מקדימה...")
                result = generate_single_letter(
                    excel_path=self.excel_path,
                    config_path=self.config_path,
                    output_dir=self.output_dir,
                    row_index=self.row_index,
                    pdf_preferred="word",
                    keep_docx=True,
                )
                self.log.emit(f"נוצר PDF: {result['filename']}")
                self.finished.emit(
                    {
                        "mode": "preview",
                        "total": 1,
                        "success": 1,
                        "errors": 0,
                        "pdf": str(result["pdf"]),
                    }
                )
                return

            self.log.emit("מתחיל הפקה...")
            from src.excel_reader import read_excel

            df = read_excel(self.excel_path)
            total = len(df)
            self.progress.emit(0, total)

            def on_progress(current: int, total_rows: int) -> None:
                self.progress.emit(current, total_rows)
                self.log.emit(f"מעבד שורה {current} מתוך {total_rows}...")

            batch_result = generate_letters(
                excel_path=self.excel_path,
                config_path=self.config_path,
                output_dir=self.output_dir,
                pdf_preferred="word",
                on_progress=on_progress,
            )

            for err in batch_result.errors:
                self.error_line.emit(f"שורה {err['excel_row']} - {err['error']}")

            self.log.emit(
                f"הפקה הסתיימה. סה\"כ: {batch_result.total}, "
                f"הצליחו: {batch_result.success}, שגיאות: {len(batch_result.errors)}"
            )
            self.finished.emit(
                {
                    "mode": "batch",
                    "total": batch_result.total,
                    "success": batch_result.success,
                    "errors": len(batch_result.errors),
                    "output_dir": str(self.output_dir),
                }
            )
        except Exception as exc:
            self.failed.emit(str(exc))
