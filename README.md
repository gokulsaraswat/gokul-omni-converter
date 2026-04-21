# Gokul Omni Convert Lite

A local Python desktop app for batch conversion, PDF workflows, integrated OCR, visual page organization, automation presets, and installer-ready packaging prep.





## Patch 29 highlights

Patch 29 finishes this cleanup pass by tightening the three identity-heavy surfaces one more time and removing a little more visual noise.

New in Patch 29:
- leaner **header chrome**
  - even smaller GIF/logo footprint
  - nav-only header buttons
  - footer remains the only place with **About** and **Mail**
- cleaner **Home** dashboard
  - shorter quick-start copy
  - stronger metric hierarchy
  - denser quick-tools and favorites area
- cleaner **About** page
  - better typography on surface cards
  - smaller image footprint
  - fewer duplicate actions and clearer grouping for primary actions, social links, and local files
- app version bumped to **2.2.6**

Pure Python remains the default engine. LibreOffice remains optional and user-controlled.


## Patch 28 highlights

Patch 28 keeps tightening the three identity-heavy screens so they waste less space and feel more stable at common laptop sizes.

New in Patch 28:
- leaner **header chrome**
  - smaller GIF/logo footprint
  - tighter header/body/sidebar spacing
  - slimmer sidebar width and status area
- denser **Home** layout
  - smaller hero and action strip
  - less vertical stretching in Recent Jobs and Quick Tools
  - tighter metric cards and shorter status copy
- cleaner **About** layout
  - fewer duplicate actions
  - smaller profile image area
  - combined contact + social action row with less filler text
- fixed a small **responsive-layout reference bug** so the Convert page keeps its own layout controller instead of being overwritten by the PDF Tools page
- app version bumped to **2.2.5**

Pure Python remains the default engine. LibreOffice remains optional and user-controlled.


## Patch 27 highlights

Patch 27 is another targeted UX cleanup pass for the three identity-heavy screens: **Header**, **Home**, and **About**.

New in Patch 27:
- tighter **header chrome**
  - smaller GIF/logo footprint
  - slimmer header/footer spacing
  - tighter sidebar width and navigation padding
- more compact **Home** dashboard
  - shorter hero copy and faster action strip
  - tighter metric cards
  - cleaner recent-jobs and quick-tools spacing
  - compact mode/output labels to avoid tall cards
- cleaner **About** page
  - smaller profile image footprint
  - shorter header text and reduced action clutter
  - denser profile/details layout with less empty space
- app version bumped to **2.2.4**

Pure Python remains the default engine. LibreOffice remains optional and user-controlled.


## Patch 26 highlights

Patch 26 tightens the three screens that still felt the loosest: **Header**, **Home**, and **About**.

New in Patch 26:
- slimmer **header**
  - smaller GIF/logo footprint
  - denser header buttons with less wasted vertical space
  - tighter body/sidebar chrome spacing
- denser **Home** page
  - less filler copy
  - tighter hero, metric cards, and quick tool spacing
  - recent jobs and quick tools no longer stretch into empty space on taller windows
- cleaner **About** page
  - shorter copy and smaller profile image footprint
  - removed the extra wide asset card from the bottom
  - fewer, more focused profile actions
  - less empty vertical space and better responsive stacking
- app version bumped to **2.2.3**

Pure Python remains the default engine. LibreOffice remains optional and user-controlled.


## Patch 25 highlights

Patch 25 tightens the busiest identity screens so the app feels cleaner right away.

New in Patch 25:
- slimmer **header** with a smaller GIF/logo slot and buttons-only top bar
- denser **Home** layout with shorter copy, compact stat cards, and cleaner quick tools
- cleaner **About** page with a smaller image area, shorter copy, and simpler asset actions
- fixed the missing `open_url()` helper so About mail/social buttons open correctly
- app version bumped to **2.2.2**

Pure Python remains the default engine. LibreOffice remains optional and user-controlled.



## Patch 24 highlights

Patch 24 focuses on the last stretch of UI stability work: smarter responsive wrapping, a scroll-safe sidebar, cleaner label formatting, and lighter copy in busy workspaces.

New in Patch 24:
- stronger **responsive button wrapping**
  - `FlowButtonBar` now prefers the real allocated width instead of the requested width, so organizer, settings, convert, and header action rows wrap properly on tighter screens
- new **scrollable workspace sidebar**
  - the left navigation rail now stays usable on shorter displays instead of pushing items below the fold
- cleaner **label formatting**
  - internal identifiers such as `pure_python`, `top-left`, and mixed camel/snake labels now render as cleaner UI text like **Pure Python** and **Top Left**
  - run details now show clearer Yes/No flags instead of raw boolean values
- lighter **organizer wording**
  - the organizer hero and hints are shorter and less text-heavy
- app version bumped to **2.2.1**

Pure Python remains the default engine. LibreOffice remains optional and user-controlled.


## Patch 23 highlights

Patch 23 finishes the current UI phase with an installer-friendly **remote asset system** that keeps local fallbacks safe while letting you later point the app at GitHub-hosted branding files.

New in Patch 23:
- optional **remote asset config** in `remote_assets.json`
  - header GIF local path + optional remote URL
  - splash GIF local path + optional remote URL
  - About image optional remote URL
  - optional About profile JSON URL for pulling editable profile data from a hosted source
- safe **cached asset loading**
  - remote assets are downloaded into a local cache
  - bundled local files stay as the fallback path
  - the app still works offline after packaging
- upgraded **Settings** file section
  - remote assets enable toggle
  - header GIF path and URL controls
  - splash GIF URL controls
  - About image URL controls
  - About profile JSON URL controls
  - asset cache open / clear actions
  - timeout and refresh-hour controls
- upgraded **About** actions
  - refresh remote assets
  - open remote asset config
- better **workspace / support / diagnostics** packaging
  - `remote_assets.json` now travels with bundles
  - resolved header/splash/About assets are included when exported
- installer prep updated
  - PyInstaller spec now includes `remote_assets.json`
  - packaged builds keep local fallbacks while allowing later remote refresh
- app version bumped to **2.2.0**

Pure Python remains the default engine. LibreOffice remains optional and user-controlled.


## Patch 22 highlights

Patch 22 focuses on stronger responsiveness, cleaner interaction polish, and more visible hover states while keeping every Patch 1–21 feature intact.

New in Patch 22:
- broader **responsive layout logic**
  - Home metric cards now reflow on tighter widths
  - Home lower panels stack cleanly when space is limited
  - Convert page input and options panels switch between side-by-side and stacked layouts
  - About page profile and details sections stack smoothly on narrower windows
  - Settings cards now collapse from a 2x2 grid into a cleaner single-column flow
- stronger **button hover polish**
  - subtle staged hover feedback with soft-hover, hover, and pressed styles
  - clearer border and background changes in dark and light themes
  - consistent hover handling for header, footer, nav, primary, and standard buttons
- more **responsive action bars**
  - Convert input actions now wrap instead of forcing a tall rigid column
  - Home dependency actions wrap cleanly
  - History recent-output and failed-job actions wrap cleanly
  - OCR hero actions wrap on narrow widths
  - About profile actions wrap instead of overflowing
  - Settings utility action groups such as soffice, cache, splash, backup, and support actions now wrap safely
- improved **dynamic wrapping**
  - descriptive labels with wraplength now adapt to available width instead of staying fixed to wide-screen values
- polished **tooltips**
  - tooltip colors now follow the active theme for a more integrated look
- app version bumped to **2.1.3**

The app still keeps **Pure Python** as the default engine. LibreOffice remains optional and only participates when configured and selected or when fallback rules allow it.


## Patch 21 highlights

Patch 21 focuses on responsiveness, layout cleanup, and a lighter chrome without removing any earlier Patch 1–20 capability.

New in Patch 21:
- cleaner **responsive shell**
  - compact top header with a replaceable animated GIF/logo slot
  - smaller footer that now only shows `Gokul Omni Convert Lite | 2.1.3`, **About**, and **Mail**
- broader **scroll support** across the main pages
  - Home
  - Convert
  - PDF Tools
  - OCR
  - Automation
  - History
  - Settings
  - About
- improved responsive action layout with wrapping button bars for:
  - header controls
  - Home hero actions
  - online link controls
  - OCR hero actions
  - About action/social buttons
  - organizer hero and toolbar actions
- more visible hover feedback
  - buttons now react more clearly on hover with subtle background and border changes
  - pointer cursor on interactive controls
- organizer UI cleanup
  - responsive wrapped toolbar instead of a single overflow-prone row
  - cleaner button naming like **Select All**, **Move Up**, **Move Down**, **Rotate Left**, **Rotate Right**
- smaller default window minimum to behave better on tighter screens
- bundled header GIF placeholder at `assets/gokul_header.gif` so you can swap branding later

The app still keeps **Pure Python** as the default engine. LibreOffice remains optional and only participates when configured and selected or when fallback rules allow it.


## Patch 20 highlights

Patch 20 adds a dedicated **Preview Center** without removing any earlier Patch 1–19 capability.

New in Patch 20:
- new **Preview Center** window for inspecting files before or after conversion
  - preview selected Convert inputs
  - preview recent outputs
  - preview files stored with a selected History job
  - add extra files directly inside the Preview Center
- file-aware preview rendering:
  - real page preview for **PDF**
  - image preview for common image formats
  - structured summary previews for **DOCX**, **XLS/XLSX/CSV/TSV**, **PPTX**, **TXT**, **Markdown**, and **HTML**
- page controls inside Preview Center:
  - previous / next file
  - previous / next PDF page
  - zoom presets
  - open file / open folder shortcuts
- easier access across the UI:
  - **Preview selected** button on the Convert page
  - **Preview** action for Recent Outputs
  - **Preview selected outputs** action for History jobs
  - menu and command-palette entries
  - `Ctrl+Shift+P` shortcut
- smoke tests expanded for:
  - preview rendering for PDF, image, text, DOCX, sheet, and presentation inputs
  - preview fallback behavior for missing files

The app still keeps **Pure Python** as the default engine. LibreOffice remains optional and only participates when configured and selected or when fallback rules allow it.


## Patch 19 highlights

Patch 19 upgrades the visual organizer without removing any earlier Patch 1–18 capability.

New in Patch 19:
- **drag-and-drop page cards** in the Organizer screen for direct visual reordering
- **Undo / Redo** history for organizer actions such as reorder, rotate, duplicate, remove, reverse, and layout loads
- **layout snapshot export/import**:
  - save the current organizer order/rotation/selection as JSON
  - reload the same layout later for the matching PDF
- improved organizer usability:
  - drop-target highlighting while dragging
  - focused organizer shortcuts for `Ctrl+Z`, `Ctrl+Y`, `Ctrl+A`, and `Delete`
  - safer reset of drag state on reload and after changes
- smoke tests expanded for:
  - drag-style reorder helper logic
  - organizer layout payload save/load round-trip

The app still keeps **Pure Python** as the default engine. LibreOffice remains optional and only participates when configured and selected or when fallback rules allow it.


## Patch 18 highlights

Patch 18 focuses on accessibility, resilience, and developer-friendly recovery tools without removing any Patch 1–17 capability.

New in Patch 18:
- new **accessibility controls** in Settings:
  - UI scale presets (`90%`, `100%`, `110%`, `125%`, `140%`)
  - **high contrast** toggle
  - **reduced motion** toggle for startup splash and login reminder
- new **automatic state backup** system:
  - timestamped JSON backups before saves
  - configurable keep count
  - manual **Create backup now**
  - **Restore latest backup**
  - **Open backup folder**
- new **keyboard shortcut guide**:
  - in-app Markdown viewer for `keyboard_shortcuts.md`
  - Help menu entry and `F1` shortcut
  - customizable by editing the local Markdown file
- **Build Center** expanded again with:
  - backup actions
  - shortcut-guide entry
  - accessibility and recovery summary
- support/workspace exports now include:
  - shortcut guide file
  - latest state backup when available
- state persistence expanded for:
  - UI scale
  - high contrast
  - reduced motion
  - state-backup preferences
  - last backup path
- smoke tests expanded for:
  - state backup creation
  - backup recovery from corrupted state JSON
  - shortcut guide presence

The app still keeps **Pure Python** as the default engine. LibreOffice remains optional and only participates when configured and selected or when fallback rules allow it.


## Patch 17 highlights

Patch 17 focuses on support readiness, compact-mode UX cleanup, startup control, and richer reporting without removing any Patch 1–16 capability.

New in Patch 17:
- new **compact UI mode** in Settings with denser spacing and tighter tables/buttons
- new **start page** setting so launch can prefer Home, Convert, PDF Tools, OCR, Organizer, Automation, History, Settings, or About
- new **activity report export**:
  - polished HTML summary of recent jobs, outputs, failed retries, and dependency signals
  - available from File menu, Build Center, Settings, Quick Actions, and CLI
- new **support bundle export**:
  - ZIP package with diagnostics JSON
  - state snapshot
  - app logs
  - activity report
  - footer notes
  - About profile
  - installer assets
  - optional profile image / splash asset when available
- **Build Center** expanded with:
  - export activity report
  - export support bundle
  - open app state folder
  - compact/start-page summary
- **History page** improved with:
  - instant filter/search
  - export selected run report
  - export activity report shortcut
- new **headless CLI hooks**:
  - `python app.py --export-activity-report out.html`
  - `python app.py --export-support-bundle out.zip`
- state persistence expanded for:
  - compact UI preference
  - preferred start page
  - default activity report folder
  - default support bundle folder
- smoke tests expanded for:
  - activity report rendering
  - support bundle creation

The app still keeps **Pure Python** as the default engine. LibreOffice remains optional and only participates when configured and selected or when fallback rules allow it.


## Patch 16 highlights

Patch 16 focuses on release workflow polish, workspace portability, and a real manifest-driven update check without removing any Patch 1–15 capability.

New in Patch 16:
- **real update checker** backed by a local JSON file or remote HTTP/HTTPS manifest
- bundled **installer/update_manifest.example.json** template for future release feeds
- new **Build Center** actions for:
  - export workspace bundle
  - import workspace bundle
  - choose manifest file
  - check for updates
- new **workspace bundle** export/import helpers that package:
  - app state
  - footer notes
  - About profile
  - installer metadata
  - selected profile/splash assets when available
- new **headless CLI hooks**:
  - `python app.py --check-updates`
  - `python app.py --export-workspace out.zip`
  - `python app.py --import-workspace bundle.zip --workspace-target ./restore_here`
- state persistence expanded for:
  - update manifest source
  - last update result
  - workspace bundle destination
- smoke tests expanded for:
  - update manifest parsing
  - workspace bundle export/import

The app still keeps **Pure Python** as the default engine. LibreOffice remains optional and only participates when configured and selected or when fallback rules allow it.


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

Headless release helpers:
- `python app.py --check-updates`
- `python app.py --check-updates installer/update_manifest.example.json`
- `python app.py --export-workspace gokul_workspace.zip`
- `python app.py --import-workspace gokul_workspace.zip --workspace-target ./restore_here`


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
- overlay tools are page-overlay based; they are useful for labels, stamps, and approvals but they are not full Acrobat-style paragraph reflow editing
- mail draft support opens your default mail client with a prepared message; you still attach files manually from the output folder
- direct SMTP send depends on a valid mail server, credentials when required, and provider attachment-size limits
- the watch-folder engine tracks files by path, size, and modified time so changed files can be picked up as new work later

## Project files

- `app.py` - desktop GUI shell, conversion workspaces, Preview Center integration, and Patch 20 workflow polish
- `mail_core.py` - SMTP configuration, connection tests, mailto helpers, and EML draft generation
- `ocr_core.py` - OCR engine and Tesseract integration for searchable PDFs and OCR text extraction
- `build_support.py` - diagnostics export and settings snapshot helpers
- `converter_core.py` - conversion engine and PDF tool engine
- `pure_python_renderers.py` - built-in DOCX/XLSX/PPTX/HTML/Markdown renderers
- `organizer_core.py` - organizer backend for page sequencing, save, extract, and image export
- `page_organizer.py` - organizer UI panel with drag-and-drop, layout snapshots, undo/redo, and preview window
- `automation_core.py` - Patch 8 watch-folder, preset import/export, report, and ZIP helper functions
- `app_state.py` - local settings, history, presets, and watch-folder state
- `ui_theme.py` - theme and widget styling helpers
- `ui_text.py` - UI text humanization helpers for readable labels and flags
- `preview_support.py` - file-aware preview rendering helpers for PDF, image, text, DOCX, sheet, and presentation inputs
- `preview_ui.py` - Preview Center window, page navigation, zoom, and file inspector UI
- `footer_notes.md` - markdown source for the footer notes window
- `about_profile.json` - editable profile content for the About page
- `assets/gokul_profile_placeholder.png` - placeholder image for the About page
- `installer/` - packaging prep files
- `requirements.txt` - Python dependencies
- `smoke_test.py` - backend smoke test
