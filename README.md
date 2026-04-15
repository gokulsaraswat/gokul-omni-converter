# Gokul Omni Convert Lite

A local Python desktop app for batch conversion, PDF workflows, integrated OCR, visual page organization, automation presets, and installer-ready packaging prep.


## Patch 15 highlights

Patch 15 focuses on workflow acceleration, queue control, cache hygiene, and faster navigation without removing any Patch 1–14 capability.

New in Patch 15:
- **favorite presets** with quick-launch shortcuts on the Home page
- **Quick Actions palette** for fast navigation and common actions
- new keyboard shortcuts:
  - `Ctrl+Enter` -> start conversion
  - `Ctrl+Shift+Enter` -> run PDF tool
  - `Ctrl+Shift+L` -> focus the URL box
  - `Ctrl+K` -> open Quick Actions
  - `Ctrl+,` -> open Settings
  - `F5` -> refresh dependency status
- **pause / resume** controls for online link downloads
- **link cache manager** improvements:
  - cache summary
  - keep-days policy
  - size-cap policy
  - prune action
- new **performance mode** in Settings:
  - `eco`
  - `balanced`
  - `quality`
- state persistence expanded for:
  - performance mode
  - cache policy
  - window geometry
  - last opened page
- smoke tests expanded for:
  - favorite preset persistence
  - cache stats and prune logic

The app still keeps **Pure Python** as the default engine. LibreOffice remains optional and only participates when configured and selected or when fallback rules allow it.


## Patch 14 highlights

Patch 14 focuses on advanced PDF finishing tools, session recovery, output management, and production-minded quality-of-life upgrades.

New in Patch 14:
- new PDF tools:
  - **Redact area / region**
  - **Edit PDF text (best-effort)** for extractable text only
- stronger PDF tool UI with dedicated fields for:
  - area rectangle values
  - replacement text
- new state and workflow features:
  - **recent outputs manager**
  - **failed jobs list** with retry / remove / clear actions
  - **restore last session** on startup
  - **auto-open output folder** after successful runs
  - **temporary session cleanup** on exit
  - **update checker placeholder**
- new log/export improvements:
  - export combined app logs to a text file
  - recent output tracking persisted in local app state
- smoke tests expanded for:
  - area redaction
  - best-effort text edit
  - new state persistence keys
  - text log export helper

The app still keeps **Pure Python** as the default engine. LibreOffice remains optional and is only used when configured and selected or when fallback rules allow it.

## Patch 13 highlights

Patch 13 adds a full online-links workflow and makes several conversion modes easier to find directly in the UI.

New in Patch 13:
- **Online links / URLs** panel on the Convert page
  - paste one or many HTTP/HTTPS links
  - deduplicate repeated URLs
  - fetch links into the same local conversion queue
  - retry failed links
  - cancel in-progress downloads
  - safe cache folder handling
  - per-link status table with local file path and details
  - recent-links memory stored in local app state
  - **Fetch + Start** workflow for download-and-convert in one action
- new explicit conversion modes:
  - **Markdown -> PDF**
  - **HTML -> PDF**
  - **HTML -> Markdown**
  - **PPT / PPTX / ODP -> Images**
- new Settings controls for:
  - link cache folder
  - link timeout
  - keep downloaded link files in cache
- patch-13 smoke tests now cover:
  - explicit Markdown/HTML conversions
  - presentation-to-images export
  - local HTTP URL download + conversion pipeline

The app still keeps **Pure Python** as the default engine. LibreOffice remains optional and is only used when configured and selected or when fallback rules allow it.

## Patch 12 highlights

Patch 12 focuses on startup polish, state migration safety, About customization, and packaging prep without removing any earlier conversion, OCR, organizer, automation, or PDF tooling features.

New in Patch 12:
- first-launch **splash screen** with configurable GIF path and safe fallback behavior
- bottom-right **login reminder popup** that:
  - never appears on install day
  - only becomes eligible after 3 or more days
  - stays gone forever after dismiss
  - stays gone forever after completion
- new state keys with backward-safe migration:
  - `install_date`
  - `login_popup_dismissed`
  - `login_popup_completed`
  - `login_popup_last_shown`
  - `login_popup_enabled`
  - `splash_enabled`
  - `splash_seen`
  - `splash_gif_path`
- improved **About page** with:
  - editable local profile JSON
  - company/project fields
  - feedback button
  - contribute button
  - static installer-safe About snapshot at `installer/about_static.json`
- expanded **Settings** page with splash, reminder, and optional LibreOffice path controls
- bundled placeholder splash GIF in `assets/gokul_splash.gif`
- smoke-test startup path that skips overlays so automated checks remain stable

## Engine model

The app supports three engine modes in **Settings**:
- `pure_python`
- `auto`
- `libreoffice`

### Recommended default
Use **pure_python** when you want a portable setup without depending on LibreOffice.

### Auto mode
Auto uses the richer built-in pipeline first and only falls back to LibreOffice when needed and available.

### LibreOffice mode
LibreOffice remains available as an **optional external renderer**. You can set the exact `soffice` path in Settings and test it from the UI.

## Current conversion features

- all earlier conversion modes from Patch 9 remain available

## OCR workspace

The new OCR screen supports:
- image -> searchable PDF
- PDF -> searchable PDF
- image/PDF -> OCR TXT extraction
- saved OCR language / DPI / PSM defaults
- saved optional Tesseract executable path

- many images -> 1 PDF
- images -> separate PDFs
- PDF -> images
- DOC / DOCX / ODT / RTF -> PDF
- PPT / PPTX / ODP -> PDF
- PDF -> DOCX
- PDF -> PPTX
- PDF -> HTML
- XLS / XLSX / ODS / CSV / TSV -> PDF
- PDF -> XLSX with best-effort table and text extraction
- TXT / HTML / Markdown -> PDF
- HTML -> DOCX
- Markdown -> DOCX
- Markdown -> HTML
- merge many PDFs into 1 PDF
- mixed batch conversion with **Any Supported -> PDF**

## PDF tools workspace

Available tools:
- Merge PDFs
- Split PDF by ranges
- Split PDF every N pages
- Extract pages
- Remove pages
- Reorder pages
- Add text watermark
- Add image watermark
- Edit PDF with text overlay
- Edit PDF with image overlay
- Redact searched text
- Sign PDF (visible)
- Edit metadata
- Lock PDF with password
- Unlock PDF
- Compress PDF

## Organizer workspace

Use the **Organizer** screen when you want a more visual workflow:
- inspect pages as thumbnails
- move selected pages up or down
- rotate selected pages
- duplicate selected pages
- remove selected pages
- reverse the sequence
- extract selected pages into a new PDF
- save a reorganized PDF without typing page specs manually
- export selected pages as images
- preview a page in a larger window

## Automation workspace

The new **Automation** screen adds three workflow layers:
- **Presets**
  - save current Convert settings
  - reuse them later
  - export/import as JSON
- **Watch folder**
  - poll a folder for newly arrived supported files
  - process only unseen files
  - optionally archive processed sources
- **Sharing helpers**
  - ZIP the latest outputs
  - export a run report
  - open a mail draft for the latest outputs
  - send the latest outputs directly with SMTP

## UI overview

Screens:
- Home
- Convert
- PDF Tools
- Organizer
- Automation
- History
- Settings
- About

The app includes:
- dark, light, and system theme modes
- first-launch splash support with editable GIF asset
- delayed login reminder lifecycle with local state
- sidebar navigation
- top menu bar
- footer notes window driven by `footer_notes.md`
- local recent-job history with setting reuse
- editable About profile with placeholder image and link buttons
- in-app About profile editor
- routing preview for engine selection
- mail-draft helper for the latest outputs
- direct SMTP delivery window
- build center for diagnostics and packaging shortcuts

## Installer prep

The project includes an `installer/` folder with:
- `gokul_omni_convert_lite.spec`
- `build_windows.bat`
- `build_linux.sh`
- `windows/GokulOmniConvertLite.iss`
- `windows/version_info.txt`
- `linux/gokul-omni-convert-lite.desktop`
- `macos/build_app_bundle.sh`
- `BUILDING.md`

This is a stronger preparation step for packaging the desktop app later as a distributable build.

## Pure Python fidelity notes

Pure Python has multiple quality levels depending on format:
- **Structure-aware**: DOCX, HTML, Markdown
- **Table-aware**: XLSX, XLS, CSV, TSV
- **Content-first**: PPTX
- **Text-first fallback**: RTF, ODT, FODT, ODS, ODP, TXT-like formats
- **LibreOffice required for best support**: legacy `.doc` and `.ppt`

This makes the app stronger for portable installs, but it still does **not** promise pixel-identical Office rendering.

## Requirements

### Python packages

```bash
pip install -r requirements.txt
```

### Optional external tools

- **LibreOffice**
  - optional fallback only
  - configure the exact `soffice` path from Settings if you want the external renderer
- **Pandoc**
  - improves Markdown and HTML conversion for DOCX/HTML outputs where supported
- **pdftotext**
  - improves PDF -> DOCX extraction when available

## Run the app

### Windows
```bat
python app.py
```

### macOS / Linux
```bash
python app.py
```

## Quick UI smoke test

```bash
python app.py --smoke-test-ui
```

## Backend smoke test

```bash
python smoke_test.py
```

Patch 9 smoke coverage includes:
- organizer save / extract / image export
- previous PDF tools from earlier patches
- previous pure-Python conversion coverage
- watch-folder discovery helpers
- preset export/import helpers
- ZIP bundle and run-report helpers
- settings snapshot export/import
- diagnostics export
- SMTP email message generation

## Notes and limitations

- `PDF -> DOCX` is best-effort and works best on text-heavy PDFs.
- `PDF -> XLSX` works best when tables are detectable in the source PDF.
- `PDF -> PPTX` prioritizes page fidelity by placing each PDF page on a slide as an image.
- complex layouts and scanned documents may not round-trip perfectly
- OCR depends on a working **Tesseract OCR** installation
- the visible organizer does not yet support drag-and-drop page cards
- overlay tools are page-overlay based; they are useful for labels, stamps, and approvals but they are not full Acrobat-style paragraph reflow editing
- mail draft support opens your default mail client with a prepared message; you still attach files manually from the output folder
- direct SMTP send depends on a valid mail server, credentials when required, and provider attachment-size limits
- the watch-folder engine tracks files by path, size, and modified time so changed files can be picked up as new work later

## Project files

- `app.py` - desktop GUI shell and Patch 9 delivery, build-center, and About-editor flows
- `mail_core.py` - SMTP configuration, connection tests, mailto helpers, and EML draft generation
- `ocr_core.py` - OCR engine and Tesseract integration for searchable PDFs and OCR text extraction
- `build_support.py` - diagnostics export and settings snapshot helpers
- `converter_core.py` - conversion engine and PDF tool engine
- `pure_python_renderers.py` - built-in DOCX/XLSX/PPTX/HTML/Markdown renderers
- `organizer_core.py` - organizer backend for page sequencing, save, extract, and image export
- `page_organizer.py` - organizer UI panel and preview window
- `automation_core.py` - Patch 8 watch-folder, preset import/export, report, and ZIP helper functions
- `app_state.py` - local settings, history, presets, and watch-folder state
- `ui_theme.py` - theme and widget styling helpers
- `footer_notes.md` - markdown source for the footer notes window
- `about_profile.json` - editable profile content for the About page
- `assets/gokul_profile_placeholder.png` - placeholder image for the About page
- `installer/` - packaging prep files
- `requirements.txt` - Python dependencies
- `smoke_test.py` - backend smoke test
