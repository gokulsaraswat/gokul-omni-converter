# Final Release Checklist — Gokul Omni Convert Lite 2.3.0

## Required checks

- [ ] `python release_validation.py`
- [ ] `python app.py --validate-release`
- [ ] `python -m py_compile app.py converter_core.py ocr_core.py image_folder_pdf.py release_validation.py`
- [ ] App opens without LibreOffice installed
- [ ] Footer shows `Gokul Omni Convert Lite | 2.3.0`
- [ ] Images -> PDF works for multiple PNG/JPG/WEBP inputs
- [ ] Recursive image folder -> one PDF works
- [ ] OCR page gives a clear Tesseract status
- [ ] About, Mail, Settings, Build Center, and Preview Center open

## Installer checks

- [ ] Build from a clean virtual environment
- [ ] Run installed app on a clean Windows machine or VM
- [ ] Confirm bundled assets load offline
- [ ] Confirm app launches even when Tesseract is not bundled
- [ ] Confirm bundled Tesseract is detected when provided
- [ ] Confirm support bundle and activity report exports work
