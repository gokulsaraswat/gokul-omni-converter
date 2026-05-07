# Patch 37 — Final Release Readiness

Patch 37 focuses on stabilization and packaging readiness rather than new feature bloat.

## Changed

- App version bumped to **2.3.0**.
- Added `release_validation.py` for repeatable release-readiness checks.
- Added `python app.py --validate-release` headless validation mode.
- Added PyInstaller spec and build scripts.
- Added installer build notes and a final release checklist.
- Added static installer About snapshot and update manifest example.
- Fixed a runtime bug in an earlier duplicate external job callback where an undefined `status` variable could crash organizer/history callbacks.

## Validation coverage

`release_validation.py` checks:

- Python syntax for core files
- required release/build files
- Images -> PDF conversion using generated PNG files
- recursive image folder PDF generation
- OCR/Tesseract detection status shape
- clear rejection of `pytesseract.exe`
- release version consistency across release files

## Notes

LibreOffice remains optional and user-configurable. Pure Python remains the default engine. Tesseract remains optional unless you want bundled OCR support in the installer.
