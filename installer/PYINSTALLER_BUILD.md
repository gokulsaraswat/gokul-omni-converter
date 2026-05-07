# PyInstaller Build Notes

Patch 37 prepares **Gokul Omni Convert Lite 2.3.0** for installer builds.

## Clean build flow

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# Linux/macOS: source .venv/bin/activate
pip install -r requirements.txt
python release_validation.py
python app.py --validate-release
pyinstaller --clean --noconfirm installer/GokulOmniConvertLite.spec
```

## Optional Tesseract runtime

OCR works without bundling when the user configures a local Tesseract path or has it on PATH. To bundle OCR with the installer, place one of these folders beside `app.py` before building:

```text
tesseract/
  tesseract.exe
  *.dll
  tessdata/
    eng.traineddata
```

or:

```text
Tesseract-OCR/
  tesseract.exe
  *.dll
  tessdata/
    eng.traineddata
```

The spec includes those folders only when present, so installer builds do not fail if OCR is not bundled.
