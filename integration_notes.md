# Patch 10 Integration Notes

This patch was produced as a merge-ready feature pack because the main branch source was not available in this environment.

## Recommended merge strategy

Add the following into your main branch:

- copy `patch10_services.py` into your services layer
- wire OCR menu actions or a new OCR page to these functions:
  - `image_to_searchable_pdf`
  - `pdf_to_searchable_pdf`
  - `extract_text_with_ocr`
- wire advanced PDF tools to these functions:
  - `redact_text`
  - `password_protect_pdf`
  - `remove_pdf_password`
  - `compress_pdf`
- wire the mail section to these functions:
  - `build_eml_draft`
  - `open_mailto_draft`
  - `send_email_smtp`

## Suggested UI sections in the main app

### OCR page
- input file
- output folder
- OCR language
- DPI and PSM controls
- actions for searchable PDF and OCR TXT

### Advanced PDF page
- redact phrases
- open/current password
- user password
- owner password
- compression profile
- actions for redact, lock, unlock, compress

### Mail page
- sender
- to / cc
- subject
- body
- attachment picker
- SMTP host / port / credentials
- actions for `mailto`, `.eml`, and SMTP send

## Notes on expectations

- OCR requires **Tesseract OCR** on PATH
- PDF compression is a practical size-reduction step, not a magic optimizer for every file
- `mailto:` cannot attach files directly in a portable way, which is why Patch 10 also includes `.eml` drafts and SMTP sending
