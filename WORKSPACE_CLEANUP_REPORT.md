# Patch 36.5 Workspace Cleanup Report

Patch: `PATCH-036.5-OMNI-STABILITY-OCR-IMAGEPDF-WORKSPACE-CLEANUP`

## Main repo selected

This packaged patch was built from `gokul_omni_convert_lite_patch36.zip`, extracted as `gokul_omni_convert_lite_patch36_5`.

The execution workspace did not contain a live Git repository, so no branch, tag, commit, push, or destructive workspace cleanup was performed inside this environment.

## Duplicate / legacy folders found

Inside this sandbox, the available project material was a set of uploaded patch ZIPs and a flat extracted source snapshot. No live duplicate Git working folders were archived.

Available related archives included older patch ZIPs such as:

- `gokul_omni_convert_lite_patch1.zip` through `gokul_omni_convert_lite_patch36.zip`
- `gokul_omni_convert_lite_ui_refresh_patch.zip`
- `omni_file_converter_app.zip`

These were left untouched.

## Useful updates merged

No duplicate-folder merge was needed. The fixes were applied directly to the Patch 36 source.

## Folders archived or removed

None. The prompt required not deleting permanently, and there were no live duplicate source folders to archive in this sandbox.

## Bugs fixed

### Images -> PDF `KeyError: 'JPEG'`

`converter_core.py` now:

- registers Pillow JPEG/PDF plugin imports explicitly
- normalizes each image to loaded RGB content
- uses first frame for animated images
- composites alpha/transparency on white
- tries Pillow PDF save first
- falls back to ReportLab when Pillow PDF writing raises `KeyError('JPEG')` or another PDF/JPEG plugin-style failure
- preserves image aspect ratio in fallback output pages

### OCR/Tesseract/tessdata detection

`ocr_core.py` now:

- rejects `pytesseract.exe` with a clear message
- locates bundled Tesseract under `tesseract/` or `Tesseract-OCR/`
- validates `tessdata/<language>.traineddata`
- returns `available`, `path`, `source`, `tessdata`, and `message`
- sets `pytesseract.pytesseract.tesseract_cmd`, `PATH`, and `TESSDATA_PREFIX` during OCR jobs
- passes `tessdata=` into PyMuPDF `pdfocr_tobytes()` when supported, with a safe fallback for older PyMuPDF versions

## Verification results

The following checks were run successfully in the sandbox on 2026-05-07:

```bash
python -m py_compile app.py converter_core.py ocr_core.py smoke_test.py
```

Targeted tests were also run for:

- transparent PNG + RGB PNG -> one two-page PDF
- `pytesseract.exe` rejection
- Tesseract status shape including `available`, `path`, `source`, `tessdata`, and `message`

Full smoke test passed in this sandbox. LibreOffice-specific conversion was skipped because LibreOffice was not installed, but Pure Python paths and OCR with the available PATH Tesseract runtime passed.

## Remaining manual installer step

Before a real installer build, place the Tesseract runtime folder next to the packaged executable using one of the layouts documented in `installer/TESSERACT_BUNDLE.md`.
