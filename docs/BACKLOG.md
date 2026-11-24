# BACKLOG: gokul-omni-converter (OMN)

Rules: `RULES.md` section 4. Architecture: `docs/ARCHITECTURE.md`. Public-safe: no secrets, no employer material.
Priority: P0 broken/blocking or security, P1 important, P2 worthwhile, P3 nice to have.
Status: open, doing, blocked, done, dropped. Cross-repo: `needs:` waits on, `blocks:` is waited on by.

## Bugs
| ID | Title | Pri | Status | Needs/Blocks | Notes (where, repro, evidence) |
|---|---|---|---|---|---|
| OMN-B-003 | App startup rewrites tracked `installer/about_static.json` (`generated_at`) | P3 | open | | `app.py` `_sync_static_about_profile` runs on every launch and in `tests/test_ui_smoke.py`; leaves the git tree dirty after any run. Write only when content other than the timestamp changes, or move the file to the build step. |
| OMN-B-004 | `tests/test_ui_smoke.py` crashes printing the info-dot glyph on a cp1252 Windows console | P3 | open | | `UnicodeEncodeError` at the "hover hints present" line unless `PYTHONUTF8=1`. Fix: reconfigure stdout to UTF-8 in the test. |
| OMN-B-005 | README clone URL points at `go-omni-pro.git`, repo is `gokul-omni-converter` | P3 | open | | `README.md` macOS and Linux "run from source" blocks. |

## Improvements
| ID | Title | Pri | Status | Needs/Blocks | Notes |
|---|---|---|---|---|---|
| OMN-I-001 | Lazy-import heavy libs to cut cold start | P2 | open | | Baseline 2026-11-25: every heavy lib (fitz/pymupdf, pdfplumber, openpyxl, docx, pptx, reportlab, pypdf, pytesseract, xlrd, PIL) is loaded before the window; `import app` ~1.1 s (converter_core 0.9 s), window built ~6 s. See `docs/FEATURES.md` baseline. |
| OMN-I-002 | Frozen build is 206 MB; PyMuPDF 2.x drags in onnxruntime, numpy, pymupdf4llm, pymupdf_layout | P2 | open | | PyInstaller dist 206 MB / 1378 files (pymupdf 94 MB, numpy.libs 21 MB). Options: pin `PyMuPDF<2`, or `excludes=` in `installer/GOOmniPro.spec` after proving nothing imports them. |

## Features
| ID | Title | Pri | Status | Needs/Blocks | Notes |
|---|---|---|---|---|---|

## Ideas (unscoped, not committed to)
| ID | Idea | Pri | Status | Needs/Blocks | Notes |
|---|---|---|---|---|---|

## Tech debt
| ID | Title | Pri | Status | Needs/Blocks | Notes |
|---|---|---|---|---|---|
| OMN-D-001 | No lint config; 35 pyflakes findings | P2 | open | | `ruff check --select E9,F`: 31 unused imports, F541, F601 (`app_state.py:506` repeated `"id"`, intentional override), F821 (`app.py:831` annotation `Tooltip` not imported; harmless under `from __future__ import annotations`), F841 (`app.py:346`). Add a `ruff` config and fix the safe ones. |
| OMN-D-002 | `requirements.txt` has no upper bounds and no lock | P2 | open | | Fresh installs on 2026-11-25 pulled pypdf 7 and customtkinter 6, both of which broke the app (OMN-B-001, OMN-B-002). Add a tested constraints/lock file. |
| OMN-D-003 | `import fitz` is deprecated by PyMuPDF (FutureWarning) | P3 | open | | Move to `import pymupdf` (keep `fitz` fallback for old installs). |
| OMN-D-004 | Support customtkinter 6 | P3 | open | | With ctk 6.0.0 the UI smoke test fails from "theme dark" on (root destroyed after the first `_apply_theme`). Pinned `<6` for now (OMN-B-002). |

## Done (recent)
| ID | Title | Commit | Date |
|---|---|---|---|
| OMN-B-001 | Watermark / overlay / sign PDF tools crash on pypdf 7 (`DeprecationError` from `merge_page` on a reader page) | (this commit) | 2026-11-25 |
