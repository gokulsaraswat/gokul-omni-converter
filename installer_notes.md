# Patch 10 Installer Notes

Patch 10 is designed to be packable later with your main application.

## What to remember for packaging

- include the Python packages listed in `requirements.txt`
- make sure the packaged app can still find **Tesseract OCR**
- if you want OCR to work out of the box, bundle or separately install the Tesseract binary depending on your platform strategy
- SMTP sending may require firewall and antivirus exceptions in some environments

## Suggested packaging path later

- package the full desktop app with PyInstaller or Nuitka
- move user-editable config and profile files into a writable app-data location
- provide a first-run dependency check screen for OCR and optional mail settings
