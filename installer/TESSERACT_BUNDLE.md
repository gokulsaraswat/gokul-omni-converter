# Bundled Tesseract OCR runtime

Patch 36.5 keeps OCR optional for normal app use, but installer builds can ship a local Tesseract runtime so OCR works without asking users to configure a separate install.

## Required runtime layout

Place the Tesseract runtime beside `app.py` during development or beside the packaged executable after installation:

```text
tesseract/
  tesseract.exe
  *.dll
  tessdata/
    eng.traineddata
```

The app also supports this alternative folder name:

```text
Tesseract-OCR/
  tesseract.exe
  *.dll
  tessdata/
    eng.traineddata
```

On Linux/macOS development builds, the binary can be named `tesseract` instead of `tesseract.exe`.

## Detection priority

OCR detection now uses this priority:

1. A valid user-configured real `tesseract.exe` path.
2. A bundled runtime in `tesseract/` or `Tesseract-OCR/`.
3. A `tesseract` executable found on `PATH`.

The app rejects `pytesseract.exe` because that file is only the Python wrapper. Users must select the real Tesseract executable from a Tesseract-OCR installation.

## PyInstaller packaging note

If a PyInstaller spec or build script is used later, include the bundled runtime only when the folder exists. Do not fail the build when it is absent, because OCR should remain optional.

Example spec idea:

```python
from pathlib import Path

runtime_root = Path("tesseract")
if runtime_root.exists():
    datas += [(str(runtime_root), "tesseract")]

runtime_root = Path("Tesseract-OCR")
if runtime_root.exists():
    datas += [(str(runtime_root), "Tesseract-OCR")]
```

Keep `tessdata/eng.traineddata` inside the bundled runtime folder. Add more `*.traineddata` files only for languages you plan to support.
