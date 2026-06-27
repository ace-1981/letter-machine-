"""Main application window."""

from __future__ import annotations

import os
import sys
from pathlib import Path

from PySide6.QtCore import QThread, Qt
from PySide6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QRadioButton,
    QSpinBox,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from src.app_paths import get_default_output_dir, get_templates_dir
from src.config_loader import load_template_config
from src.output_format import OUTPUT_DOCX, OUTPUT_PDF
from src.startup_check import check_startup_templates
from src.ui.worker import GenerationWorker


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("מחולל מכתבים")
        self.resize(720, 640)
        self._thread: QThread | None = None
        self._worker: GenerationWorker | None = None
        self._last_output_dir: Path | None = None
        self._build_ui()
        self.output_edit.setText(str(get_default_output_dir()))
        self._load_letter_types()

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        self.excel_edit = QLineEdit()
        excel_btn = QPushButton("בחר")
        excel_btn.clicked.connect(self._pick_excel)
        layout.addWidget(QLabel("קובץ קלט (xlsx / xls):"))
        layout.addLayout(self._row(self.excel_edit, excel_btn))

        self.output_edit = QLineEdit()
        output_btn = QPushButton("בחר")
        output_btn.clicked.connect(self._pick_output)
        layout.addWidget(QLabel("תיקיית יעד:"))
        layout.addLayout(self._row(self.output_edit, output_btn))

        self.letter_combo = QComboBox()
        layout.addWidget(QLabel("סוג מכתב:"))
        layout.addWidget(self.letter_combo)

        layout.addWidget(QLabel("סוג פלט:"))
        output_type_row = QHBoxLayout()
        self.output_pdf_radio = QRadioButton("PDF")
        self.output_docx_radio = QRadioButton("Word / DOCX")
        self.output_pdf_radio.setChecked(True)
        self._output_format_group = QButtonGroup(self)
        self._output_format_group.addButton(self.output_pdf_radio)
        self._output_format_group.addButton(self.output_docx_radio)
        output_type_row.addWidget(self.output_pdf_radio)
        output_type_row.addWidget(self.output_docx_radio)
        output_type_row.addStretch()
        layout.addLayout(output_type_row)

        row_layout = QHBoxLayout()
        row_layout.addWidget(QLabel("שורה לתצוגה מקדימה:"))
        self.preview_row = QSpinBox()
        self.preview_row.setMinimum(1)
        self.preview_row.setMaximum(9999)
        self.preview_row.setValue(1)
        row_layout.addWidget(self.preview_row)
        row_layout.addStretch()
        layout.addLayout(row_layout)

        btn_layout = QHBoxLayout()
        self.preview_btn = QPushButton("תצוגה מקדימה")
        self.preview_btn.clicked.connect(self._run_preview)
        self.generate_btn = QPushButton("הפקת מכתבים")
        self.generate_btn.clicked.connect(self._run_batch)
        btn_layout.addWidget(self.preview_btn)
        btn_layout.addWidget(self.generate_btn)
        layout.addLayout(btn_layout)

        self.progress = QProgressBar()
        self.progress.setValue(0)
        layout.addWidget(self.progress)
        self.progress_label = QLabel("0 מתוך 0")
        layout.addWidget(self.progress_label)

        layout.addWidget(QLabel("לוג:"))
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setMaximumHeight(140)
        layout.addWidget(self.log_box)

        layout.addWidget(QLabel("שגיאות:"))
        self.error_box = QTextEdit()
        self.error_box.setReadOnly(True)
        self.error_box.setMaximumHeight(120)
        layout.addWidget(self.error_box)

        self.open_dir_btn = QPushButton("פתח תיקיית יעד")
        self.open_dir_btn.setEnabled(False)
        self.open_dir_btn.clicked.connect(self._open_output_dir)
        layout.addWidget(self.open_dir_btn)

    def _row(self, edit: QLineEdit, button: QPushButton) -> QHBoxLayout:
        row = QHBoxLayout()
        row.addWidget(edit)
        row.addWidget(button)
        return row

    def _load_letter_types(self) -> None:
        self.letter_combo.clear()
        for json_file in sorted(get_templates_dir().glob("*.json")):
            config = load_template_config(json_file)
            name = config.get("template_name", json_file.stem)
            self.letter_combo.addItem(name, str(json_file))

    def _pick_excel(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "בחר קובץ Excel",
            "",
            "Excel (*.xlsx *.xls)",
        )
        if path:
            self.excel_edit.setText(path)

    def _pick_output(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "בחר תיקיית יעד")
        if path:
            self.output_edit.setText(path)

    def _config_path(self) -> Path | None:
        data = self.letter_combo.currentData()
        return Path(data) if data else None

    def _output_format(self) -> str:
        return OUTPUT_DOCX if self.output_docx_radio.isChecked() else OUTPUT_PDF

    def _validate_inputs(self) -> tuple[Path, Path, Path] | None:
        excel = Path(self.excel_edit.text().strip())
        output = Path(self.output_edit.text().strip())
        config = self._config_path()

        if not excel.exists():
            QMessageBox.warning(self, "שגיאה", "קובץ Excel לא נמצא.")
            return None
        if excel.suffix.lower() not in (".xlsx", ".xls"):
            QMessageBox.warning(self, "שגיאה", "פורמט לא נתמך. נדרש xlsx או xls.")
            return None
        if not output:
            QMessageBox.warning(self, "שגיאה", "יש לבחור תיקיית יעד.")
            return None
        if config is None or not config.exists():
            QMessageBox.warning(self, "שגיאה", "סוג מכתב לא תקין.")
            return None
        return excel, output, config

    def _set_busy(self, busy: bool) -> None:
        self.preview_btn.setEnabled(not busy)
        self.generate_btn.setEnabled(not busy)

    def _run_preview(self) -> None:
        inputs = self._validate_inputs()
        if not inputs:
            return
        excel, output, config = inputs
        self.error_box.clear()
        self._start_worker(
            "preview", excel, config, output, self.preview_row.value() - 1, self._output_format()
        )

    def _run_batch(self) -> None:
        inputs = self._validate_inputs()
        if not inputs:
            return
        excel, output, config = inputs
        self.error_box.clear()
        self.progress.setValue(0)
        self._start_worker("batch", excel, config, output, 0, self._output_format())

    def _start_worker(
        self,
        mode: str,
        excel: Path,
        config: Path,
        output: Path,
        row_index: int,
        output_format: str,
    ) -> None:
        self._set_busy(True)
        self.open_dir_btn.setEnabled(False)
        self._last_output_dir = output
        output.mkdir(parents=True, exist_ok=True)

        self._thread = QThread()
        self._worker = GenerationWorker(
            mode, excel, config, output, row_index, output_format=output_format
        )
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.log.connect(self._append_log)
        self._worker.error_line.connect(self._append_error)
        self._worker.progress.connect(self._update_progress)
        self._worker.finished.connect(self._on_finished)
        self._worker.failed.connect(self._on_failed)
        self._worker.finished.connect(self._thread.quit)
        self._worker.failed.connect(self._thread.quit)
        self._thread.finished.connect(self._cleanup_thread)
        self._thread.start()

    def _append_log(self, text: str) -> None:
        self.log_box.append(text)

    def _append_error(self, text: str) -> None:
        self.error_box.append(text)

    def _update_progress(self, current: int, total: int) -> None:
        self.progress.setMaximum(max(total, 1))
        self.progress.setValue(current)
        self.progress_label.setText(f"{current} מתוך {total}")

    def _on_finished(self, result: dict) -> None:
        self._set_busy(False)
        self.open_dir_btn.setEnabled(True)
        if result.get("mode") == "preview" and result.get("file"):
            self._open_file(Path(result["file"]))
        if result.get("mode") == "batch":
            fmt = result.get("output_format_label", "PDF")
            QMessageBox.information(
                self,
                "הפקה הסתיימה",
                f"סוג פלט: {fmt}\n"
                f"סה\"כ רשומות: {result['total']}\n"
                f"הופקו בהצלחה: {result['success']}\n"
                f"שגיאות: {result['errors']}",
            )

    def _on_failed(self, message: str) -> None:
        self._set_busy(False)
        self.error_box.append(message)
        QMessageBox.critical(self, "שגיאה", message)

    def _cleanup_thread(self) -> None:
        if self._worker:
            self._worker.deleteLater()
            self._worker = None
        if self._thread:
            self._thread.deleteLater()
            self._thread = None

    def _open_output_dir(self) -> None:
        if self._last_output_dir and self._last_output_dir.exists():
            os.startfile(self._last_output_dir)

    @staticmethod
    def _open_file(path: Path) -> None:
        if path.exists():
            os.startfile(path)


def run_app() -> int:
    app = QApplication(sys.argv)
    app.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
    window = MainWindow()
    startup_errors = check_startup_templates()
    if startup_errors:
        QMessageBox.warning(
            window,
            "שגיאת הפעלה",
            "\n".join(startup_errors),
        )
        window.preview_btn.setEnabled(False)
        window.generate_btn.setEnabled(False)
    window.show()
    return app.exec()
