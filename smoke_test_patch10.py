from __future__ import annotations

import tempfile
from pathlib import Path

import fitz
from PIL import Image, ImageDraw, ImageFont

from patch10_services import (
    OcrConfig,
    build_eml_draft,
    compress_pdf,
    extract_text_with_ocr,
    image_to_searchable_pdf,
    open_mailto_draft,
    password_protect_pdf,
    redact_text,
    remove_pdf_password,
)


def create_sample_image(path: Path) -> None:
    image = Image.new("RGB", (1400, 420), "white")
    draw = ImageDraw.Draw(image)
    try:
        font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 84)
    except Exception:
        font = ImageFont.load_default()
    draw.text((70, 120), "INVOICE 12345", fill="black", font=font)
    image.save(path)


def create_text_pdf(path: Path) -> None:
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)
    page.insert_text((72, 120), "SECRET DATA BLOCK", fontsize=28)
    page.insert_text((72, 170), "Visible paragraph for redaction test.", fontsize=14)
    doc.save(str(path))
    doc.close()


def run() -> None:
    with tempfile.TemporaryDirectory() as temp_dir:
        root = Path(temp_dir)
        image_path = root / "sample_image.png"
        create_sample_image(image_path)

        searchable_pdf = image_to_searchable_pdf(
            image_path,
            root / "sample_image_searchable.pdf",
            config=OcrConfig(language="eng", dpi=220, psm=6),
        )
        assert searchable_pdf.exists(), "OCR searchable PDF was not created."

        ocr_text_path = extract_text_with_ocr(
            image_path,
            root / "sample_image_ocr.txt",
            config=OcrConfig(language="eng", dpi=220, psm=6),
        )
        ocr_text = ocr_text_path.read_text(encoding="utf-8")
        assert "INVOICE" in ocr_text.upper(), ocr_text

        source_pdf = root / "plain_text.pdf"
        create_text_pdf(source_pdf)

        redacted_pdf = redact_text(source_pdf, root / "plain_text_redacted.pdf", ["SECRET"])
        redacted_doc = fitz.open(str(redacted_pdf))
        try:
            extracted = "\n".join(page.get_text("text") for page in redacted_doc)
            assert "SECRET" not in extracted, extracted
        finally:
            redacted_doc.close()

        locked_pdf = password_protect_pdf(source_pdf, root / "plain_text_locked.pdf", user_password="patch10", owner_password="owner10")
        locked_doc = fitz.open(str(locked_pdf))
        try:
            assert locked_doc.needs_pass, "Locked PDF should require a password."
            assert locked_doc.authenticate("patch10") > 0, "Could not open locked PDF with the configured password."
        finally:
            locked_doc.close()

        unlocked_pdf = remove_pdf_password(locked_pdf, root / "plain_text_unlocked.pdf", password="patch10")
        unlocked_doc = fitz.open(str(unlocked_pdf))
        try:
            assert not unlocked_doc.needs_pass, "Unlocked PDF should not require a password."
        finally:
            unlocked_doc.close()

        compressed_pdf = compress_pdf(source_pdf, root / "plain_text_compressed.pdf", profile="balanced")
        assert compressed_pdf.exists(), "Compressed PDF was not created."

        attachment = root / "attachment.txt"
        attachment.write_text("patch10 attachment", encoding="utf-8")
        eml_path = build_eml_draft(
            root / "draft_message.eml",
            sender="sender@example.com",
            to=["to@example.com"],
            cc=["cc@example.com"],
            subject="Patch10 Draft",
            body="Hello from Patch 10",
            attachments=[attachment],
        )
        eml_content = eml_path.read_text(encoding="utf-8", errors="replace")
        assert "Patch10 Draft" in eml_content, eml_content
        assert "attachment.txt" in eml_content, eml_content

        mailto_url = "mailto:"
        # Build only; avoid opening a browser in the smoke test environment.
        from patch10_services import create_mailto_url
        created_url = create_mailto_url(to=["to@example.com"], cc=["cc@example.com"], subject="Hello", body="Patch10")
        assert created_url.startswith(mailto_url), created_url

        print("Patch 10 smoke test completed successfully.")
        print(searchable_pdf)
        print(ocr_text_path)
        print(redacted_pdf)
        print(locked_pdf)
        print(unlocked_pdf)
        print(compressed_pdf)
        print(eml_path)


if __name__ == "__main__":
    run()
