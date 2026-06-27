"""PDF conversion backends: LibreOffice headless and Microsoft Word COM."""

from __future__ import annotations

import os
import shutil
import subprocess
import time
from abc import ABC, abstractmethod
from pathlib import Path

# Post-convert delay for single-shot Word COM (batch uses 0). Override: LETTER_GEN_WORD_SLEEP
_WORD_POST_SLEEP_S = float(os.environ.get("LETTER_GEN_WORD_SLEEP", "0"))
_WORD_AVAILABLE: bool | None = None


def _probe_word_available() -> bool:
    global _WORD_AVAILABLE
    if _WORD_AVAILABLE is not None:
        return _WORD_AVAILABLE
    if os.environ.get("LETTER_GEN_TEST_NO_WORD") == "1":
        _WORD_AVAILABLE = False
        return False
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        _WORD_AVAILABLE = False
        return False
    try:
        import win32com.client

        word = win32com.client.Dispatch("Word.Application")
        word.Quit()
        _WORD_AVAILABLE = True
    except Exception:
        _WORD_AVAILABLE = False
    return _WORD_AVAILABLE


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

    def close(self) -> None:
        """Release resources held by a batch session (no-op for stateless converters)."""


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
        return _probe_word_available()

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
        """Single conversion — opens and closes Word (preview / single letter)."""
        if not self.is_available():
            raise PdfConversionError(self.availability_message())

        import win32com.client

        pdf_path.parent.mkdir(parents=True, exist_ok=True)
        word = None
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            _export_docx_to_pdf(word, docx_path, pdf_path)
        except Exception as exc:
            raise PdfConversionError(f"Word COM conversion failed: {exc}") from exc
        finally:
            if word is not None:
                word.Quit()
            if _WORD_POST_SLEEP_S > 0:
                time.sleep(_WORD_POST_SLEEP_S)


class WordComBatchSession(PdfConverter):
    """Reuse one Word.Application for an entire PDF batch."""

    def __init__(self) -> None:
        self._word = None

    @property
    def name(self) -> str:
        return "WordComBatchSession"

    def is_available(self) -> bool:
        return WordComPdfConverter().is_available()

    def availability_message(self) -> str:
        return WordComPdfConverter().availability_message()

    def start(self) -> None:
        if self._word is not None:
            return
        if not self.is_available():
            raise PdfConversionError(self.availability_message())
        import win32com.client

        self._word = win32com.client.Dispatch("Word.Application")
        self._word.Visible = False

    def convert(self, docx_path: Path, pdf_path: Path) -> None:
        if self._word is None:
            self.start()
        pdf_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            _export_docx_to_pdf(self._word, docx_path, pdf_path)
        except Exception as exc:
            raise PdfConversionError(f"Word COM conversion failed: {exc}") from exc

    def close(self) -> None:
        if self._word is None:
            return
        word = self._word
        self._word = None
        try:
            while word.Documents.Count > 0:
                word.Documents(1).Close(False)
        except Exception:
            pass
        try:
            word.Quit()
        except Exception:
            pass

    def __enter__(self) -> WordComBatchSession:
        self.start()
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()


def _export_docx_to_pdf(word, docx_path: Path, pdf_path: Path) -> None:
    """Open one DOCX on an existing Word instance, export PDF, close document only."""
    doc = None
    try:
        doc = word.Documents.Open(str(docx_path.resolve()))
        _apply_rtl_for_pdf_export(doc)
        doc.ExportAsFixedFormat(
            OutputFileName=str(pdf_path.resolve()),
            ExportFormat=WordComPdfConverter.WD_EXPORT_FORMAT_PDF,
            OpenAfterExport=False,
        )
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass


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
    def create_batch(preferred: str | None = None) -> PdfConverter:
        """PDF converter for multi-row batch — Word COM reuses one Application instance."""
        single = PdfConverterFactory.create(preferred)
        if isinstance(single, WordComPdfConverter):
            return WordComBatchSession()
        return single

    @staticmethod
    def list_status() -> list[str]:
        return [
            LibreOfficePdfConverter().availability_message(),
            WordComPdfConverter().availability_message(),
        ]
