from __future__ import annotations

import re

_ACRONYM_MAP = {
    "api": "API",
    "csv": "CSV",
    "doc": "DOC",
    "docx": "DOCX",
    "gif": "GIF",
    "github": "GitHub",
    "html": "HTML",
    "http": "HTTP",
    "https": "HTTPS",
    "ibm": "IBM",
    "id": "ID",
    "jpg": "JPG",
    "jpeg": "JPEG",
    "json": "JSON",
    "md": "Markdown",
    "ocr": "OCR",
    "odt": "ODT",
    "pdf": "PDF",
    "png": "PNG",
    "ppt": "PPT",
    "pptx": "PPTX",
    "psm": "PSM",
    "rtf": "RTF",
    "smtp": "SMTP",
    "sql": "SQL",
    "svg": "SVG",
    "txt": "TXT",
    "ui": "UI",
    "url": "URL",
    "ux": "UX",
    "webp": "WEBP",
    "xls": "XLS",
    "xlsx": "XLSX",
    "xml": "XML",
    "zip": "ZIP",
}

_CAMEL_BOUNDARY_1 = re.compile(r"([a-z0-9])([A-Z])")
_CAMEL_BOUNDARY_2 = re.compile(r"([A-Z]+)([A-Z][a-z])")


def _split_identifier(value: str) -> list[str]:
    text = str(value or "").strip()
    if not text:
        return []
    text = text.replace("/", " / ")
    text = text.replace("-", " ")
    text = text.replace("_", " ")
    text = _CAMEL_BOUNDARY_1.sub(r"\1 \2", text)
    text = _CAMEL_BOUNDARY_2.sub(r"\1 \2", text)
    text = re.sub(r"\s+", " ", text).strip()
    return [part for part in text.split(" ") if part]


def humanize_identifier(value: str, *, fallback: str = "") -> str:
    parts = _split_identifier(value)
    if not parts:
        return fallback
    words: list[str] = []
    for part in parts:
        token = part.strip()
        if not token:
            continue
        lowered = token.lower()
        if lowered in _ACRONYM_MAP:
            words.append(_ACRONYM_MAP[lowered])
        elif token == "/":
            words.append(token)
        elif token.isupper() and len(token) <= 5:
            words.append(token)
        else:
            words.append(token.capitalize())
    return " ".join(words).strip() or fallback


def format_flag(value: object, *, true_text: str = "Yes", false_text: str = "No") -> str:
    if isinstance(value, str):
        normalized = value.strip().lower()
        if normalized in {"1", "true", "yes", "on", "enabled"}:
            return true_text
        if normalized in {"0", "false", "no", "off", "disabled", ""}:
            return false_text
    return true_text if bool(value) else false_text


def format_engine_label(value: str, *, fallback: str = "Pure Python") -> str:
    return humanize_identifier(value, fallback=fallback)
