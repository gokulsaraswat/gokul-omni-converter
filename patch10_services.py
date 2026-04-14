from __future__ import annotations

import io
import mimetypes
import os
import re
import smtplib
import webbrowser
from dataclasses import dataclass
from email.message import EmailMessage
from pathlib import Path
from typing import Iterable, Sequence
from urllib.parse import quote

import fitz
import pytesseract
from PIL import Image


class Patch10Error(RuntimeError):
    """Raised when a Patch 10 service operation fails."""


@dataclass(slots=True)
class OcrConfig:
    language: str = "eng"
    dpi: int = 220
    psm: int = 6


@dataclass(slots=True)
class SmtpConfig:
    host: str
    port: int
    username: str = ""
    password: str = ""
    sender: str = ""
    use_tls: bool = True
    use_ssl: bool = False


def ensure_dir(path: str | Path) -> Path:
    directory = Path(path)
    directory.mkdir(parents=True, exist_ok=True)
    return directory


def unique_path(path: str | Path) -> Path:
    target = Path(path)
    if not target.exists():
        return target
    stem = target.stem
    suffix = target.suffix
    parent = target.parent
    index = 1
    while True:
        candidate = parent / f"{stem}_{index}{suffix}"
        if not candidate.exists():
            return candidate
        index += 1


def safe_name(text: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", text).strip("._")
    return cleaned or "output"


def _image_to_png_bytes(image: Image.Image) -> bytes:
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    return buffer.getvalue()


def _tesseract_config(config: OcrConfig) -> str:
    return f"--oem 3 --psm {int(config.psm)}"


def _ocr_data(image: Image.Image, config: OcrConfig) -> dict[str, list]:
    try:
        return pytesseract.image_to_data(
            image,
            lang=config.language,
            config=_tesseract_config(config),
            output_type=pytesseract.Output.DICT,
        )
    except pytesseract.TesseractNotFoundError as exc:
        raise Patch10Error(
            "Tesseract was not found. Install Tesseract OCR and make sure it is on PATH."
        ) from exc


def _ocr_text(image: Image.Image, config: OcrConfig) -> str:
    try:
        return pytesseract.image_to_string(
            image,
            lang=config.language,
            config=_tesseract_config(config),
        )
    except pytesseract.TesseractNotFoundError as exc:
        raise Patch10Error(
            "Tesseract was not found. Install Tesseract OCR and make sure it is on PATH."
        ) from exc


def _add_invisible_text_layer(
    page: fitz.Page,
    image: Image.Image,
    config: OcrConfig,
) -> int:
    data = _ocr_data(image, config)
    image_width, image_height = image.size
    page_width = float(page.rect.width)
    page_height = float(page.rect.height)
    count = 0

    for idx, raw_text in enumerate(data.get("text", [])):
        text = str(raw_text or "").strip()
        if not text:
            continue
        try:
            conf = float(data["conf"][idx])
        except Exception:
            conf = 0.0
        if conf < 0:
            continue

        x = float(data["left"][idx])
        y = float(data["top"][idx])
        w = float(data["width"][idx])
        h = float(data["height"][idx])
        if w <= 0 or h <= 0:
            continue

        rect = fitz.Rect(
            x / image_width * page_width,
            y / image_height * page_height,
            (x + w) / image_width * page_width,
            (y + h) / image_height * page_height,
        )
        fontsize = max(6.0, rect.height * 0.85)
        try:
            inserted = page.insert_textbox(
                rect,
                text,
                fontsize=fontsize,
                fontname="helv",
                render_mode=3,
                overlay=True,
            )
            if inserted >= 0:
                count += 1
        except Exception:
            continue
    return count


def image_to_searchable_pdf(
    image_path: str | Path,
    output_pdf: str | Path,
    *,
    config: OcrConfig | None = None,
) -> Path:
    cfg = config or OcrConfig()
    source = Path(image_path)
    if not source.exists():
        raise Patch10Error(f"Image not found: {source}")

    image = Image.open(source).convert("RGB")
    png_bytes = _image_to_png_bytes(image)
    target = unique_path(output_pdf)
    ensure_dir(target.parent)

    document = fitz.open()
    try:
        page = document.new_page(width=image.width, height=image.height)
        page.insert_image(page.rect, stream=png_bytes)
        _add_invisible_text_layer(page, image, cfg)
        document.save(str(target), garbage=3, deflate=True)
    finally:
        document.close()
        image.close()
    return target


def pdf_to_searchable_pdf(
    input_pdf: str | Path,
    output_pdf: str | Path,
    *,
    config: OcrConfig | None = None,
) -> Path:
    cfg = config or OcrConfig()
    source = Path(input_pdf)
    if not source.exists():
        raise Patch10Error(f"PDF not found: {source}")

    input_doc = fitz.open(str(source))
    output_doc = fitz.open()
    try:
        scale = max(cfg.dpi / 72.0, 1.0)
        matrix = fitz.Matrix(scale, scale)
        for page in input_doc:
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            out_page = output_doc.new_page(width=float(page.rect.width), height=float(page.rect.height))
            out_page.insert_image(out_page.rect, stream=_image_to_png_bytes(image))
            _add_invisible_text_layer(out_page, image, cfg)
            image.close()
        target = unique_path(output_pdf)
        ensure_dir(target.parent)
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
) -> Path:
    cfg = config or OcrConfig()
    source = Path(source_path)
    target = unique_path(output_txt)
    ensure_dir(target.parent)

    if source.suffix.lower() == ".pdf":
        document = fitz.open(str(source))
        try:
            lines: list[str] = []
            for page_index, page in enumerate(document, start=1):
                pix = page.get_pixmap(matrix=fitz.Matrix(max(cfg.dpi / 72.0, 1.0), max(cfg.dpi / 72.0, 1.0)), alpha=False)
                image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                page_text = _ocr_text(image, cfg).strip()
                lines.append(f"[Page {page_index}]\n{page_text}")
                image.close()
            target.write_text("\n\n".join(lines).strip(), encoding="utf-8")
            return target
        finally:
            document.close()

    image = Image.open(source).convert("RGB")
    try:
        text = _ocr_text(image, cfg)
        target.write_text(text.strip(), encoding="utf-8")
        return target
    finally:
        image.close()


def redact_text(
    input_pdf: str | Path,
    output_pdf: str | Path,
    phrases: Sequence[str],
    *,
    fill_rgb: tuple[float, float, float] = (0, 0, 0),
) -> Path:
    source = Path(input_pdf)
    if not source.exists():
        raise Patch10Error(f"PDF not found: {source}")
    terms = [term.strip() for term in phrases if term and term.strip()]
    if not terms:
        raise Patch10Error("At least one phrase is required for redaction.")

    document = fitz.open(str(source))
    try:
        hit_count = 0
        for page in document:
            page_has_hits = False
            for phrase in terms:
                for rect in page.search_for(phrase):
                    page.add_redact_annot(rect, fill=fill_rgb)
                    hit_count += 1
                    page_has_hits = True
            if page_has_hits:
                page.apply_redactions()
        if hit_count == 0:
            raise Patch10Error("No matching text was found for the requested redaction terms.")
        target = unique_path(output_pdf)
        ensure_dir(target.parent)
        document.save(str(target), garbage=4, deflate=True)
        return target
    finally:
        document.close()


def password_protect_pdf(
    input_pdf: str | Path,
    output_pdf: str | Path,
    *,
    user_password: str,
    owner_password: str | None = None,
) -> Path:
    if not user_password:
        raise Patch10Error("A user password is required to lock the PDF.")
    source = Path(input_pdf)
    document = fitz.open(str(source))
    try:
        target = unique_path(output_pdf)
        ensure_dir(target.parent)
        document.save(
            str(target),
            garbage=3,
            deflate=True,
            encryption=fitz.PDF_ENCRYPT_AES_256,
            user_pw=user_password,
            owner_pw=owner_password or user_password,
        )
        return target
    finally:
        document.close()


def remove_pdf_password(
    input_pdf: str | Path,
    output_pdf: str | Path,
    *,
    password: str,
) -> Path:
    if not password:
        raise Patch10Error("The current password is required to unlock the PDF.")
    source = Path(input_pdf)
    document = fitz.open(str(source))
    try:
        if document.needs_pass and document.authenticate(password) <= 0:
            raise Patch10Error("The supplied password was rejected.")
        target = unique_path(output_pdf)
        ensure_dir(target.parent)
        document.save(str(target), garbage=3, deflate=True, encryption=fitz.PDF_ENCRYPT_NONE)
        return target
    finally:
        document.close()


def compress_pdf(
    input_pdf: str | Path,
    output_pdf: str | Path,
    *,
    profile: str = "balanced",
    password: str = "",
) -> Path:
    source = Path(input_pdf)
    document = fitz.open(str(source))
    try:
        if document.needs_pass:
            if not password:
                raise Patch10Error("Password required to compress this PDF.")
            if document.authenticate(password) <= 0:
                raise Patch10Error("The supplied password was rejected.")
        profile_key = (profile or "balanced").strip().lower()
        presets = {
            "safe": {"garbage": 1, "clean": False, "deflate": True},
            "balanced": {"garbage": 3, "clean": True, "deflate": True},
            "strong": {"garbage": 4, "clean": True, "deflate": True},
        }
        if profile_key not in presets:
            raise Patch10Error("Compression profile must be safe, balanced, or strong.")
        target = unique_path(output_pdf)
        ensure_dir(target.parent)
        kwargs = presets[profile_key].copy()
        keep_encryption = getattr(fitz, "PDF_ENCRYPT_KEEP", None)
        if keep_encryption is not None and source.suffix.lower() == ".pdf":
            kwargs["encryption"] = keep_encryption
        document.save(str(target), **kwargs)
        return target
    finally:
        document.close()


def create_mailto_url(
    *,
    to: Sequence[str] | None = None,
    cc: Sequence[str] | None = None,
    subject: str = "",
    body: str = "",
) -> str:
    to_value = ",".join([item.strip() for item in (to or []) if item.strip()])
    query_parts: list[str] = []
    if cc:
        cc_value = ",".join([item.strip() for item in cc if item.strip()])
        if cc_value:
            query_parts.append(f"cc={quote(cc_value)}")
    if subject:
        query_parts.append(f"subject={quote(subject)}")
    if body:
        query_parts.append(f"body={quote(body)}")
    query = "&".join(query_parts)
    return f"mailto:{quote(to_value)}?{query}" if query else f"mailto:{quote(to_value)}"


def open_mailto_draft(
    *,
    to: Sequence[str] | None = None,
    cc: Sequence[str] | None = None,
    subject: str = "",
    body: str = "",
) -> None:
    webbrowser.open(create_mailto_url(to=to, cc=cc, subject=subject, body=body))


def _attach_files(message: EmailMessage, attachments: Iterable[str | Path]) -> None:
    for attachment in attachments:
        file_path = Path(attachment)
        if not file_path.exists():
            raise Patch10Error(f"Attachment not found: {file_path}")
        mime_type, _ = mimetypes.guess_type(str(file_path))
        main_type, sub_type = (mime_type or "application/octet-stream").split("/", 1)
        with file_path.open("rb") as handle:
            message.add_attachment(handle.read(), maintype=main_type, subtype=sub_type, filename=file_path.name)


def build_eml_draft(
    output_eml: str | Path,
    *,
    sender: str,
    to: Sequence[str],
    cc: Sequence[str] | None = None,
    subject: str,
    body: str,
    attachments: Iterable[str | Path] = (),
) -> Path:
    if not sender:
        raise Patch10Error("Sender address is required for an EML draft.")
    recipients = [item.strip() for item in to if item.strip()]
    if not recipients:
        raise Patch10Error("At least one recipient is required for an EML draft.")

    message = EmailMessage()
    message["From"] = sender
    message["To"] = ", ".join(recipients)
    if cc:
        cc_values = [item.strip() for item in cc if item.strip()]
        if cc_values:
            message["Cc"] = ", ".join(cc_values)
    message["Subject"] = subject
    message.set_content(body or "")
    _attach_files(message, attachments)

    target = unique_path(output_eml)
    ensure_dir(target.parent)
    target.write_bytes(message.as_bytes())
    return target


def send_email_smtp(
    smtp: SmtpConfig,
    *,
    to: Sequence[str],
    cc: Sequence[str] | None = None,
    subject: str,
    body: str,
    attachments: Iterable[str | Path] = (),
) -> None:
    recipients = [item.strip() for item in to if item.strip()]
    cc_values = [item.strip() for item in (cc or []) if item.strip()]
    if not recipients:
        raise Patch10Error("At least one recipient is required.")
    if not smtp.sender:
        raise Patch10Error("Sender address is required for SMTP sending.")

    message = EmailMessage()
    message["From"] = smtp.sender
    message["To"] = ", ".join(recipients)
    if cc_values:
        message["Cc"] = ", ".join(cc_values)
    message["Subject"] = subject
    message.set_content(body or "")
    _attach_files(message, attachments)

    if smtp.use_ssl:
        server = smtplib.SMTP_SSL(smtp.host, smtp.port, timeout=30)
    else:
        server = smtplib.SMTP(smtp.host, smtp.port, timeout=30)
    try:
        server.ehlo()
        if smtp.use_tls and not smtp.use_ssl:
            server.starttls()
            server.ehlo()
        if smtp.username:
            server.login(smtp.username, smtp.password)
        server.send_message(message)
    finally:
        server.quit()
