from __future__ import annotations

import io
import os
import shutil
import sys
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

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


PYTESSERACT_WRAPPER_MESSAGE = (
    "This is pytesseract.exe, the Python wrapper. Select the real tesseract.exe from Tesseract-OCR."
)


TESSERACT_BUNDLE_DIR_NAMES = ("tesseract", "Tesseract-OCR")
TESSERACT_BINARY_NAMES = ("tesseract.exe", "tesseract")


def _runtime_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    meipass = getattr(sys, "_MEIPASS", "")
    if meipass:
        return Path(str(meipass)).resolve()
    return Path(__file__).resolve().parent


def _language_tokens(language: str) -> list[str]:
    raw = str(language or "eng").strip() or "eng"
    return [token.strip() for token in raw.replace(",", "+").split("+") if token.strip()]


def _is_pytesseract_wrapper(path: Path) -> bool:
    return path.name.strip().lower() == "pytesseract.exe"


def _candidate_tessdata_dirs(binary_path: Path) -> list[Path]:
    parent = binary_path.expanduser().resolve().parent
    candidates: list[Path] = []
    env_value = os.environ.get("TESSDATA_PREFIX", "").strip()
    if env_value:
        env_path = Path(env_value).expanduser()
        candidates.append(env_path)
        if env_path.name.lower() != "tessdata":
            candidates.append(env_path / "tessdata")
    candidates.extend(
        [
            parent / "tessdata",
            parent.parent / "tessdata",
            parent.parent / "share" / "tessdata",
            parent.parent / "share" / "tesseract-ocr" / "5" / "tessdata",
            parent.parent / "share" / "tesseract-ocr" / "4.00" / "tessdata",
            Path("/usr/share/tesseract-ocr/5/tessdata"),
            Path("/usr/share/tesseract-ocr/4.00/tessdata"),
            Path("/usr/share/tessdata"),
            Path("/usr/local/share/tessdata"),
            Path("/opt/homebrew/share/tessdata"),
        ]
    )
    unique: list[Path] = []
    seen: set[str] = set()
    for candidate in candidates:
        key = str(candidate.expanduser())
        if key in seen:
            continue
        seen.add(key)
        unique.append(candidate.expanduser())
    return unique


def _find_tessdata_dir(binary_path: Path, language: str = "eng") -> Path | None:
    tokens = _language_tokens(language)
    for candidate in _candidate_tessdata_dirs(binary_path):
        if not candidate.exists() or not candidate.is_dir():
            continue
        if all((candidate / f"{token}.traineddata").exists() for token in tokens):
            return candidate.resolve()
    return None


def _validate_tesseract_candidate(
    binary_path: str | Path,
    *,
    source: str,
    language: str = "eng",
    require_exists: bool = True,
) -> dict[str, Any]:
    path = Path(binary_path).expanduser()
    if _is_pytesseract_wrapper(path):
        return {
            "available": False,
            "path": str(path),
            "source": source,
            "tessdata": "",
            "message": PYTESSERACT_WRAPPER_MESSAGE,
        }
    if require_exists and not path.exists():
        return {
            "available": False,
            "path": str(path),
            "source": source,
            "tessdata": "",
            "message": f"Tesseract executable was not found: {path}",
        }
    if require_exists and path.is_dir():
        return {
            "available": False,
            "path": str(path),
            "source": source,
            "tessdata": "",
            "message": f"Select the tesseract executable, not the folder: {path}",
        }

    resolved = path.resolve() if path.exists() else path
    tessdata = _find_tessdata_dir(resolved, language=language)
    if tessdata is None:
        tokens = ", ".join(f"{token}.traineddata" for token in _language_tokens(language))
        return {
            "available": False,
            "path": str(resolved),
            "source": source,
            "tessdata": "",
            "message": f"Tesseract binary found, but tessdata for {tokens} was not found next to the binary or in known locations.",
        }
    return {
        "available": True,
        "path": str(resolved),
        "source": source,
        "tessdata": str(tessdata),
        "message": f"Tesseract ready via {source}: {resolved}",
    }


def _find_bundled_tesseract(language: str = "eng") -> dict[str, Any] | None:
    roots = [_runtime_root()]
    meipass = getattr(sys, "_MEIPASS", "")
    if meipass:
        roots.append(Path(str(meipass)).resolve())
    seen_roots: set[str] = set()
    for root in roots:
        root = root.expanduser().resolve()
        if str(root) in seen_roots:
            continue
        seen_roots.add(str(root))
        for folder_name in TESSERACT_BUNDLE_DIR_NAMES:
            folder = root / folder_name
            for binary_name in TESSERACT_BINARY_NAMES:
                binary = folder / binary_name
                if binary.exists():
                    return _validate_tesseract_candidate(binary, source="bundled", language=language)
    return None


def detect_tesseract_status(configured_path: str | Path | None = None, language: str = "eng") -> dict[str, Any]:
    configured = str(configured_path or "").strip()
    if configured:
        return _validate_tesseract_candidate(configured, source="configured", language=language)

    bundled = _find_bundled_tesseract(language=language)
    if bundled is not None and bool(bundled.get("available")):
        return bundled

    discovered = shutil.which("tesseract")
    if discovered:
        return _validate_tesseract_candidate(discovered, source="PATH", language=language)

    if bundled is not None:
        # A bundle candidate existed but failed validation; report its specific problem.
        return bundled

    return {
        "available": False,
        "path": "",
        "source": "missing",
        "tessdata": "",
        "message": "Tesseract OCR was not found. Install Tesseract, configure tesseract.exe, or bundle tesseract/tessdata with the app.",
    }


def _resolved_tesseract_status(config: OcrConfig | None = None) -> dict[str, Any]:
    language = config.language if config else "eng"
    status = detect_tesseract_status(config.tesseract_path if config else "", language=language)
    if not bool(status.get("available")):
        raise OcrError(str(status.get("message") or "Tesseract OCR was not found."))
    return status


@contextmanager
def _temporary_tesseract_environment(config: OcrConfig | None = None):
    status = _resolved_tesseract_status(config)
    resolved = str(status["path"])
    tessdata = str(status.get("tessdata", ""))

    previous_cmd = getattr(pytesseract.pytesseract, "tesseract_cmd", "tesseract")
    previous_path = os.environ.get("PATH", "")
    previous_tessdata = os.environ.get("TESSDATA_PREFIX")
    previous_fitz_tessdata = getattr(fitz, "TESSDATA_PREFIX", None)
    previous_fitz_module_tessdata = getattr(getattr(fitz, "fitz", None), "TESSDATA_PREFIX", None)
    previous_pixmap_global_tessdata = fitz.Pixmap.pdfocr_tobytes.__globals__.get("TESSDATA_PREFIX")

    pytesseract.pytesseract.tesseract_cmd = resolved
    binary_dir = str(Path(resolved).expanduser().resolve().parent)
    if binary_dir and binary_dir not in previous_path.split(os.pathsep):
        os.environ["PATH"] = binary_dir + (os.pathsep + previous_path if previous_path else "")
    if tessdata:
        os.environ["TESSDATA_PREFIX"] = tessdata
        try:
            fitz.TESSDATA_PREFIX = tessdata
        except Exception:
            pass
        try:
            fitz.fitz.TESSDATA_PREFIX = tessdata
        except Exception:
            pass
        try:
            fitz.Pixmap.pdfocr_tobytes.__globals__["TESSDATA_PREFIX"] = tessdata
        except Exception:
            pass

    try:
        yield status
    finally:
        pytesseract.pytesseract.tesseract_cmd = previous_cmd
        os.environ["PATH"] = previous_path
        if previous_tessdata is None:
            os.environ.pop("TESSDATA_PREFIX", None)
        else:
            os.environ["TESSDATA_PREFIX"] = previous_tessdata
        try:
            fitz.TESSDATA_PREFIX = previous_fitz_tessdata
        except Exception:
            pass
        try:
            fitz.fitz.TESSDATA_PREFIX = previous_fitz_module_tessdata
        except Exception:
            pass
        try:
            fitz.Pixmap.pdfocr_tobytes.__globals__["TESSDATA_PREFIX"] = previous_pixmap_global_tessdata
        except Exception:
            pass


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
    with _temporary_tesseract_environment(config) as status:
        try:
            tessdata = str(status.get("tessdata", ""))
            if tessdata:
                try:
                    return pixmap.pdfocr_tobytes(language=config.language, tessdata=tessdata)
                except TypeError:
                    # Older PyMuPDF versions do not expose the tessdata keyword. The
                    # temporary environment still sets TESSDATA_PREFIX for those builds.
                    return pixmap.pdfocr_tobytes(language=config.language)
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
