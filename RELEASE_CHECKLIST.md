# Release Checklist — Gokul Omni Convert Lite 2.3.0

Use this checklist before sharing an installer build.

## Source validation

- [ ] `python release_validation.py`
- [ ] `python app.py --validate-release`
- [ ] `python -m py_compile app.py converter_core.py ocr_core.py image_folder_pdf.py release_validation.py`
- [ ] `python smoke_test.py` when optional local dependencies are available

## Manual UI check

- [ ] App opens without LibreOffice installed
- [ ] Pure Python engine is the default
- [ ] Images -> PDF works for multiple PNG/JPG/WEBP inputs
- [ ] Folder -> one PDF and recursive image folder workflows work
- [ ] OCR page shows clear Tesseract status
- [ ] About, Mail, Settings, Build Center, and Preview Center open
- [ ] Dark mode and compact/responsive layouts still look correct

## Installer check

- [ ] Build from a clean virtual environment
- [ ] Run installed app on a clean machine or VM
- [ ] Confirm assets load from bundled files
- [ ] Confirm optional Tesseract bundle is detected when provided
- [ ] Confirm app still launches when Tesseract is not bundled
- [ ] Confirm support bundle and activity report exports work

## Release assets

- [ ] Update `about_profile.json`
- [ ] Replace placeholder profile/header/splash assets
- [ ] Update `installer/update_manifest.example.json` or hosted manifest
- [ ] Upload final installer to the chosen release location
