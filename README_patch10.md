# Gokul Omni Convert Lite - Patch 10

This patch is a **self-contained feature pack** you can merge into your main branch while Patch 9 is being completed elsewhere.

## What Patch 10 adds

Patch 10 focuses on three areas:

- **OCR**
  - image to searchable PDF
  - PDF to searchable PDF
  - OCR text extraction to TXT
- **Advanced PDF operations**
  - redact matching text
  - lock PDF with password
  - unlock PDF
  - compress PDF with safe, balanced, and strong presets
- **Mail workflow**
  - create `.eml` draft files with attachments
  - open default mail client using `mailto:`
  - send email with attachments through SMTP

## Files in this patch

- `patch10_services.py` - service layer for OCR, PDF, and mail features
- `patch10_gui.py` - standalone demo UI for the Patch 10 feature set
- `smoke_test.py` - backend smoke test
- `integration_notes.md` - suggested integration plan for your main branch
- `requirements.txt` - Python dependency list for this patch
- `installer_notes.md` - notes for packaging this feature pack later

## Run the demo

```bash
python patch10_gui.py
```

Quick launch test:

```bash
python patch10_gui.py --smoke-test-ui
```

## Run the smoke test

```bash
python smoke_test.py
```

## Requirements

Install Python packages from `requirements.txt`.

Patch 10 also expects a working **Tesseract OCR** binary on PATH for OCR operations.

## Notes

- The OCR workflow creates **searchable PDFs** by rasterizing pages and adding an invisible text layer from Tesseract output.
- The PDF lock/unlock/compress functions use PyMuPDF.
- SMTP sending uses Python's standard library `smtplib`.
- The `.eml` draft flow is useful when you want a reviewable message file before sending.
