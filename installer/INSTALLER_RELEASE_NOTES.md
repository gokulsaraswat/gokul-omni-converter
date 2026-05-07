# Gokul Omni Convert Lite 2.3.0 - Installer Release Notes

Patch 37 is a release-readiness patch focused on packaging stability, validation, and safe defaults.

## Included

- Pure Python remains the default conversion engine.
- LibreOffice remains optional and user-configurable.
- Optional bundled Tesseract support is documented and included by the PyInstaller spec only when present.
- Release validation can be run with `python release_validation.py` or `python app.py --validate-release`.
- Static installer About metadata is stored in `installer/about_static.json`.
- Update manifest example is stored in `installer/update_manifest.example.json`.

## Pre-release checks

Run `python release_validation.py` before building the installer.
