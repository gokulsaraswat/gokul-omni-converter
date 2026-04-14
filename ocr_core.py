from __future__ import annotations

import io
import os
import shutil
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

import fitz
import pytesseract
from PIL import Image


class OcrError(RuntimeError):
    """Raised when an OCR workflow cannot be completed."""


ProgressFn = Callable[[int, int], None]
LogFn = Callable[[str], None]


@dataclass(slots=True)
class OcrConfig:
    language: str = "eng"
    dpi: int = 220
    psm: int = 6
    tesseract_path: str = ""


def detect_tesseract_status(configured_path: str | Path | None = None) -> dict[str, str | bool]:
    configured = str(configured_path or "").strip()
    if configured:
        candidate = Path(configured).expanduser()
        available = candidate.exists()
        return {
            "available": available,
            "path": str(candidate),
            "source": "configured",
        }

    discovered = shutil.which("tesseract")
    return {
        "available": bool(discovered),
        "path": discovered or "",
        "source": "PATH" if discovered else "missing",
    }


def _resolved_tesseract_path(config: OcrConfig | None = None) -> str:
    status = detect_tesseract_status(config.tesseract_path if config else "")
    if not bool(status["available"]):
        raise OcrError(
            "Tesseract OCR was not found. Install Tesseract and set its path in Settings if it is not on PATH."
        )
    return str(status["path"])


@contextmanager
def _temporary_tesseract_environment(config: OcrConfig | None = None):
    resolved = _resolved_tesseract_path(config)
    pytesseract.pytesseract.tesseract_cmd = resolved

    binary_dir = str(Path(resolved).expanduser().resolve().parent)
    previous_path = os.environ.get("PATH", "")
    restore_path = False
    if binary_dir and binary_dir not in previous_path.split(os.pathsep):
        os.environ["PATH"] = binary_dir + (os.pathsep + previous_path if previous_path else "")
        restore_path = True

    try:
        yield resolved
    finally:
        if restore_path:
            os.environ["PATH"] = previous_path


def _image_to_png_bytes(image: Image.Image) -> bytes:
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    return buffer.getvalue()


def _tesseract_args(config: OcrConfig) -> str:
    return f"--oem 3 --psm {max(int(config.psm), 0)}"


def _ocr_text(image: Image.Image, config: OcrConfig) -> str:
    with _temporary_tesseract_environment(config):
        try:
            return pytesseract.image_to_string(
                image,
                lang=config.language,
                config=_tesseract_args(config),
            )
        except pytesseract.TesseractNotFoundError as exc:
            raise OcrError(
                "Tesseract OCR was not found. Install Tesseract and set its path in Settings if it is not on PATH."
            ) from exc


def _pixmap_to_ocr_pdf_bytes(pixmap: fitz.Pixmap, config: OcrConfig) -> bytes:
    with _temporary_tesseract_environment(config):
        try:
            return pixmap.pdfocr_tobytes(language=config.language)
        except Exception as exc:
            raise OcrError(f"OCR PDF generation failed: {exc}") from exc


def _emit_progress(progress: ProgressFn | None, current: int, total: int) -> None:
    if progress is not None:
        progress(int(current), max(int(total), 1))


def _emit_log(log: LogFn | None, message: str) -> None:
    if log is not None:
        log(message)


def image_to_searchable_pdf(
    image_path: str | Path,
    output_pdf: str | Path,
    *,
    config: OcrConfig | None = None,
    progress: ProgressFn | None = None,
    log: LogFn | None = None,
) -> Path:
    cfg = config or OcrConfig()
    source = Path(image_path).expanduser()
    if not source.exists():
        raise OcrError(f"Image not found: {source}")

    _emit_log(log, f"OCR image input: {source.name}")
    target = Path(output_pdf).expanduser()
    target.parent.mkdir(parents=True, exist_ok=True)

    pix = fitz.Pixmap(str(source))
    try:
        pdf_bytes = _pixmap_to_ocr_pdf_bytes(pix, cfg)
    finally:
        pix = None

    target.write_bytes(pdf_bytes)
    _emit_log(log, "Searchable PDF generated.")
    _emit_progress(progress, 1, 1)
    return target


def pdf_to_searchable_pdf(
    input_pdf: str | Path,
    output_pdf: str | Path,
    *,
    config: OcrConfig | None = None,
    progress: ProgressFn | None = None,
    log: LogFn | None = None,
) -> Path:
    cfg = config or OcrConfig()
    source = Path(input_pdf).expanduser()
    if not source.exists():
        raise OcrError(f"PDF not found: {source}")

    input_doc = fitz.open(str(source))
    output_doc = fitz.open()
    try:
        total_pages = max(input_doc.page_count, 1)
        scale = max(cfg.dpi / 72.0, 1.0)
        matrix = fitz.Matrix(scale, scale)
        for page_number, page in enumerate(input_doc, start=1):
            _emit_log(log, f"OCR page {page_number}/{total_pages}")
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            ocr_pdf = fitz.open("pdf", _pixmap_to_ocr_pdf_bytes(pix, cfg))
            try:
                output_doc.insert_pdf(ocr_pdf)
            finally:
                ocr_pdf.close()
            _emit_progress(progress, page_number, total_pages)

        target = Path(output_pdf).expanduser()
        target.parent.mkdir(parents=True, exist_ok=True)
        output_doc.save(str(target), garbage=3, deflate=True)
        return target
    finally:
        output_doc.close()
        input_doc.close()


def extract_text_with_ocr(
    source_path: str | Path,
    output_txt: str | Path,
    *,
    config: OcrConfig | None = None,
    progress: ProgressFn | None = None,
    log: LogFn | None = None,
) -> Path:
    cfg = config or OcrConfig()
    source = Path(source_path).expanduser()
    if not source.exists():
        raise OcrError(f"Source file not found: {source}")

    target = Path(output_txt).expanduser()
    target.parent.mkdir(parents=True, exist_ok=True)

    if source.suffix.lower() == ".pdf":
        document = fitz.open(str(source))
        try:
            total_pages = max(document.page_count, 1)
            lines: list[str] = []
            scale = max(cfg.dpi / 72.0, 1.0)
            matrix = fitz.Matrix(scale, scale)
            for page_number, page in enumerate(document, start=1):
                _emit_log(log, f"Extracting OCR text from page {page_number}/{total_pages}")
                pix = page.get_pixmap(matrix=matrix, alpha=False)
                image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                page_text = _ocr_text(image, cfg).strip()
                lines.append(f"[Page {page_number}]\n{page_text}")
                image.close()
                _emit_progress(progress, page_number, total_pages)
            target.write_text("\n\n".join(lines).strip(), encoding="utf-8")
            return target
        finally:
            document.close()

    image = Image.open(source).convert("RGB")
    try:
        text = _ocr_text(image, cfg).strip()
        target.write_text(text, encoding="utf-8")
        _emit_progress(progress, 1, 1)
        _emit_log(log, f"OCR text extracted from image: {source.name}")
        return target
    finally:
        image.close()
