"""PDF conversion backends: LibreOffice headless and Microsoft Word COM."""

from __future__ import annotations

import os
import shutil
import subprocess
import time
from abc import ABC, abstractmethod
from pathlib import Path


class PdfConversionError(Exception):
    pass


class PdfConverter(ABC):
    @abstractmethod
    def is_available(self) -> bool:
        pass

    @abstractmethod
    def availability_message(self) -> str:
        pass

    @abstractmethod
    def convert(self, docx_path: Path, pdf_path: Path) -> None:
        pass

    @property
    def name(self) -> str:
        return self.__class__.__name__


class LibreOfficePdfConverter(PdfConverter):
    COMMON_PATHS = [
        Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
        Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
    ]

    def __init__(self, soffice_path: Path | None = None):
        self._soffice_path = soffice_path or self._find_soffice()

    def _find_soffice(self) -> Path | None:
        found = shutil.which("soffice")
        if found:
            return Path(found)
        for candidate in self.COMMON_PATHS:
            if candidate.exists():
                return candidate
        return None

    def is_available(self) -> bool:
        if os.environ.get("LETTER_GEN_TEST_NO_LIBREOFFICE") == "1":
            return False
        return self._soffice_path is not None and self._soffice_path.exists()

    def availability_message(self) -> str:
        if self.is_available():
            return f"LibreOffice available at {self._soffice_path}"
        return (
            "LibreOffice not found. Install LibreOffice or add soffice.exe to PATH. "
            "Common path: C:\\Program Files\\LibreOffice\\program\\soffice.exe"
        )

    def convert(self, docx_path: Path, pdf_path: Path) -> None:
        if not self.is_available():
            raise PdfConversionError(self.availability_message())

        pdf_path.parent.mkdir(parents=True, exist_ok=True)
        out_dir = pdf_path.parent
        cmd = [
            str(self._soffice_path),
            "--headless",
            "--norestore",
            "--convert-to",
            "pdf",
            "--outdir",
            str(out_dir),
            str(docx_path.resolve()),
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        if result.returncode != 0:
            raise PdfConversionError(
                f"LibreOffice conversion failed (code {result.returncode}): "
                f"{result.stderr or result.stdout}"
            )

        produced = out_dir / f"{docx_path.stem}.pdf"
        if not produced.exists():
            raise PdfConversionError(
                f"LibreOffice did not produce expected PDF: {produced}"
            )
        if produced.resolve() != pdf_path.resolve():
            produced.replace(pdf_path)


class WordComPdfConverter(PdfConverter):
    WD_EXPORT_FORMAT_PDF = 17
    WD_READING_ORDER_RTL = 1
    WD_ALIGN_RIGHT = 2
    WD_LANGUAGE_HEBREW = 1037

    def is_available(self) -> bool:
        if os.environ.get("LETTER_GEN_TEST_NO_WORD") == "1":
            return False
        try:
            import win32com.client  # noqa: F401
        except ImportError:
            return False
        try:
            import win32com.client

            word = win32com.client.Dispatch("Word.Application")
            word.Quit()
            return True
        except Exception:
            return False

    def availability_message(self) -> str:
        if self.is_available():
            return "Microsoft Word COM automation is available."
        try:
            import win32com.client  # noqa: F401
        except ImportError:
            return "pywin32 is not installed. Install with: pip install pywin32"
        return (
            "Microsoft Word is not available for COM automation. "
            "Ensure Word is installed on this machine."
        )

    def convert(self, docx_path: Path, pdf_path: Path) -> None:
        if not self.is_available():
            raise PdfConversionError(self.availability_message())

        import win32com.client

        pdf_path.parent.mkdir(parents=True, exist_ok=True)
        word = None
        doc = None
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(str(docx_path.resolve()))
            _apply_rtl_for_pdf_export(doc)
            doc.ExportAsFixedFormat(
                OutputFileName=str(pdf_path.resolve()),
                ExportFormat=self.WD_EXPORT_FORMAT_PDF,
                OpenAfterExport=False,
            )
        except Exception as exc:
            raise PdfConversionError(f"Word COM conversion failed: {exc}") from exc
        finally:
            if doc is not None:
                doc.Close(False)
            if word is not None:
                word.Quit()
            time.sleep(0.5)


def _apply_rtl_for_pdf_export(doc) -> None:
    """Force RTL reading order in Word before PDF export (preserve per-paragraph alignment)."""
    try:
        for paragraph in doc.Paragraphs:
            fmt = paragraph.Range.ParagraphFormat
            fmt.ReadingOrder = 1  # wdReadingOrderRtl
            paragraph.Range.LanguageID = 1037  # Hebrew
        for table in doc.Tables:
            table.Range.ParagraphFormat.ReadingOrder = 1
    except Exception:
        pass


class PdfConverterFactory:
    @staticmethod
    def create(preferred: str | None = None) -> PdfConverter:
        converters: list[PdfConverter] = []
        if preferred == "libreoffice":
            converters = [LibreOfficePdfConverter(), WordComPdfConverter()]
        elif preferred == "word":
            converters = [WordComPdfConverter(), LibreOfficePdfConverter()]
        else:
            converters = [WordComPdfConverter(), LibreOfficePdfConverter()]

        for converter in converters:
            if converter.is_available():
                return converter

        raise PdfConversionError(
            "No PDF converter available. "
            + LibreOfficePdfConverter().availability_message()
            + " | "
            + WordComPdfConverter().availability_message()
        )

    @staticmethod
    def list_status() -> list[str]:
        return [
            LibreOfficePdfConverter().availability_message(),
            WordComPdfConverter().availability_message(),
        ]
