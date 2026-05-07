# Patch 36.5 Summary

Patch name: `PATCH-036.5-OMNI-STABILITY-OCR-IMAGEPDF-WORKSPACE-CLEANUP`

## Summary

This patch applies the requested targeted stability fixes before Patch 37. It does not redesign the UI or remove existing features.

## Files changed

- `converter_core.py`
- `ocr_core.py`
- `app.py`
- `smoke_test.py`
- `README.md`
- `installer/TESSERACT_BUNDLE.md`
- `WORKSPACE_CLEANUP_REPORT.md`

## Main fixes

- Images -> PDF now handles PNG/JPG/JPEG/WEBP/BMP/TIFF/GIF, animated GIF first frames, transparency on white, and a ReportLab fallback for Pillow PDF/JPEG plugin issues.
- OCR detection now validates real Tesseract binaries, rejects `pytesseract.exe`, locates bundled `tesseract/` and `Tesseract-OCR/` runtime folders, validates tessdata, and wires `TESSDATA_PREFIX` for pytesseract and PyMuPDF.

## Verification

Passed:

```bash
python -m py_compile app.py converter_core.py ocr_core.py smoke_test.py
python smoke_test.py
```

LibreOffice-specific smoke coverage was skipped because LibreOffice was not installed in the sandbox.

## Git note

This sandbox copy was not a live Git repository, so no branch, tag, commit, or push was performed here. Use the same changed files on your repository branch `fix/omni-installer-ocr-images-pdf-cleanup` if you want to commit and push from your machine.
