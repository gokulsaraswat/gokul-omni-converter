# Graph Report - gokul-omni-converter  (2026-11-24)

## Corpus Check
- 95 files · ~121,309 words
- Verdict: corpus is large enough that graph structure adds value.
- Unclassified: 10 file(s) not represented in the graph (top: (none) 3, .bat 2, .tex 1)

## Summary
- 1982 nodes · 4563 edges · 142 communities (89 shown, 53 thin omitted)
- Extraction: 98% EXTRACTED · 2% INFERRED · 0% AMBIGUOUS · INFERRED: 79 edges (avg confidence: 0.89)
- Token cost: 0 input · 0 output

## Graph Freshness
- Built from commit: `e6d1292d`
- Run `git rev-parse HEAD` and compare to check if the graph is stale.
- Run `graphify update .` after code changes (no API cost).

## Community Hubs (Navigation)
- PageOrganizerPanel
- ArchivePage
- GOOmniProApp
- release_validation.py
- rename_organizer_core.py
- pure_python_renderers.py
- WORKLOG
- register_drop_target
- AppStateStore
- ocr_core.py
- ScrollablePage
- PreviewCenterWindow
- 1.4 Mermaid — Cancellation Sequence Diagrams
- Cancellation — Temporal Workflow
- UIKit
- ._log_automation
- converter_core.py
- smoke_test.py
- ._register
- REPO SECTIONS
- ._persist_state
- open_path
- preview_support.py
- RenameOrganizePage
- automation_core.py
- render
- release_support.py
- 5.20.3 Trip Data Model Changes
- link_ingest.py
- mail_core.py
- ._image_folder_worker_run
- Animator
- ._refresh_about_profile
- build_support.py
- Any
- .__init__
- asset_support.py
- FirstLaunchSplashWindow
- visit
- load_about_profile
- MarkdownNotesWindow
- LLD — Project Repository Initial Setup & Architecture
- Misc
- ScrollArea
- test_ui_smoke.py
- Ideas: One Platform, Many Apps
- packages
- smtp
- REPO SECTIONS
- ui_theme.py
- WORKLOG: gokul-omni-converter
- settings.py
- ._apply_state_snapshot
- GradientBackdrop
- GO Omni Pro Activity Report
- Kickoff prompt: gokul-omni-converter
- kit.py
- ButtonBar
- app.py
- .radio
- diagnostics.json
- CommandPalette
- Supply Cancellation Execute — Success / Failure Event Flow
- Repo Kit (every repo must have these)
- PDF Conversion
- paths
- Tooltip
- dependencies
- PDF Split Every N Output Part 1
- automation.py
- CX Cancellation Flow
- gsclaude/sync-architecture.sh
- system
- 🤖 System Architecture & Co-Worker Protocol
- GO Omni Pro app icon 256px — purple/indigo rounded-square badge with circular convert arrow and document sheet icon
- AGENTS.md
- Patch 16 profile image — solid blue square, test profile asset for patch 16 release smoke test
- app
- Gokul Omni Convert header branding GIF — dark background pill badge with blue circle and bold GOKUL / OMNI CONVERT text
- build_linux_release.sh
- Smoke test output: PDF-to-image conversion result — page 1 rendered as PNG showing tabular data (Sample PDF for extraction)
- Organizer Extract - Selected Pages PDF
- Patch 36.5 stability test output — plain PNG (white background with 'plain png' label text, tests opaque rendering)
- PDF Image Overlay Output
- PDF Lock Output
- PDF Redact Area Output
- PDF Image Watermark Output
- Presentations-to-images smoke test output — page 1 of sample_presentation rendered as PNG showing slide title and content
- start_linux.sh
- ui/__init__.py
- pages/__init__.py
- Any to PDF - Combined Output
- HTML to PDF Explicit - Sample Output
- Image Folder Per Folder - Root Images PDF
- Link Cache - Sample PDF
- Links HTML to PDF - Sample Output
- MD to PDF Explicit - Sample Output
- Organizer export image page 1 — multi-page PDF page 1 exported as PNG (landscape orientation, content rotated)
- Organizer export image page 2 — multi-page PDF page 2 exported as PNG
- PDF Edit Text Output (Replace)
- PDF Extract Output
- PDF Merge Output
- PDF Remove Pages Output
- PDF Reorder Pages Output
- PDF Visible Signature Output
- Sample Spreadsheet PDF (Pure Python export, 2-page with data table and summary)
- Sample Text-to-PDF output (plain text lines converted to PDF)
- Prompt Playbook — CLAUDE.md Workflow
- Keyboard Shortcuts
- Prompt Playbook — CLAUDE.md Workflow
- 🤖 System Architecture & Co-Worker Protocol
- COLLABORATORS.md
- Release Checklist — GO Omni Pro 2.3.0
- PLANNING-CHAT-PROMPT.md
- GO Omni Pro Notes
- sync-architecture.sh
- convert.py
- workflow_support.py
- home.py
- html_to_md/sample.md
- link_cache/sample.md
- patch16_release/footer_notes.md
- imported_bundle/footer_notes.md
- BUILDING.md
- apply_ttk_theme
- CONTEXT: read this first in every new session
- Any
- CTkFrame
- Misc
- StringVar
- Widget
- ._show_page
- 6.6 Refund Quote Entity
- re
- Prompt Playbook — CLAUDE.md Workflow
- pdf_tools.py
- 🤖 System Architecture & Co-Worker Protocol
- Repo Kit (every repo must have these)
- Agent/sync-architecture.sh
- COLLABORATOR-DESIGN-CHAT-PROMPT.md
- prompts/PLANNING-CHAT-PROMPT.md

## God Nodes (most connected - your core abstractions)
1. `GOOmniProApp` - 259 edges
2. `ConversionError` - 62 edges
3. `PageOrganizerPanel` - 57 edges
4. `run()` - 51 edges
5. `AppStateStore` - 43 edges
6. `ensure_directory()` - 40 edges
7. `ThemePalette` - 37 edges
8. `ArchivePage` - 36 edges
9. `process_pdf_tool()` - 32 edges
10. `process_batch()` - 30 edges

## Surprising Connections (you probably didn't know these)
- `PageOrganizerPanel` --uses--> `ThemePalette`  [INFERRED]
  page_organizer.py → ui_theme.py
- `ArchivePage` --uses--> `ConversionError`  [INFERRED]
  ui/pages/archive.py → converter_core.py
- `PreviewCenterWindow` --uses--> `PreviewResult`  [INFERRED]
  preview_ui.py → preview_support.py
- `PreviewCenterWindow` --uses--> `ThemePalette`  [INFERRED]
  preview_ui.py → ui_theme.py
- `PreviewWindow` --uses--> `ThemePalette`  [INFERRED]
  page_organizer.py → ui_theme.py

## Import Cycles
- None detected.

## Communities (142 total, 53 thin omitted)

### Community 0 - "PageOrganizerPanel"
Cohesion: 0.05
Nodes (41): Document, OrganizedPage, build_default_sequence(), duplicate_positions(), export_pages_as_images(), extract_selected_pdf(), move_positions_down(), move_positions_to_index() (+33 more)

### Community 1 - "ArchivePage"
Cohesion: 0.06
Nodes (34): ArchiveEntry, ArchiveSummary, batch_compress(), _collect_folder_files(), _compression(), create_archive(), extract_archive(), format_size() (+26 more)

### Community 2 - "GOOmniProApp"
Cohesion: 0.05
Nodes (4): GOOmniProApp, OcrConfig, StringVar, Widget

### Community 3 - "release_validation.py"
Cohesion: 0.17
Nodes (30): main(), build_image_folder_pdfs(), discover_image_files(), _emit_log(), _emit_progress(), group_images_by_folder(), ImageFolderPdfConfig, _is_hidden() (+22 more)

### Community 4 - "rename_organizer_core.py"
Cohesion: 0.07
Nodes (44): GO Omni Pro, apply_case(), apply_plan(), _bucket_for(), describe_manifest(), _file_dates(), latest_manifest(), _log() (+36 more)

### Community 5 - "pure_python_renderers.py"
Cohesion: 0.07
Nodes (43): _Cell, csv, docx, docx_document, docx_oxml_table, docx_oxml_text_paragraph, docx_shared, docx_table (+35 more)

### Community 6 - "WORKLOG"
Cohesion: 0.07
Nodes (28): Agreed order, Architecture (agreed 2026-10-31), Cleanup scan: `/cleanscan` (Claude slash command), Commit + push everything: `/gitmd` (Claude slash command), Current focus, Decisions, Dropped, Far later (+20 more)

### Community 7 - "register_drop_target"
Cohesion: 0.11
Nodes (20): tkinterdnd2, dnd_available(), DropZone, init_root(), parse_drop_paths(), Misc, Path, Drag-and-drop file input, with a working fallback when tkdnd is absent.… (+12 more)

### Community 8 - "AppStateStore"
Cohesion: 0.08
Nodes (26): AppStateStore, _default_state_dir(), migrate_legacy_state_dir(), Any, Path, Copy a pre-2.4.0 state directory to the new location, once. Copies rather than…, Per-platform state location, following each OS's own convention., normalize_preset_record() (+18 more)

### Community 9 - "ocr_core.py"
Cohesion: 0.14
Nodes (35): contextlib, io, _candidate_tessdata_dirs(), detect_tesseract_status(), _emit_log(), _emit_progress(), extract_text_with_ocr(), _find_bundled_tesseract() (+27 more)

### Community 10 - "ScrollablePage"
Cohesion: 0.09
Nodes (12): visit(), AnimatedGifLabel, bind_responsive_wrap(), FlowButtonBar, Misc, Path, Widget, A simple responsive page shell with a vertical scrollbar. The inner frame… (+4 more)

### Community 11 - "PreviewCenterWindow"
Cohesion: 0.19
Nodes (3): PreviewCenterWindow, Misc, Path

### Community 12 - "1.4 Mermaid — Cancellation Sequence Diagrams"
Cohesion: 0.05
Nodes (37): 1.1 Scope, 1.2 Actors & Boundaries, 1.3 Functional Breakdown (Low-Level), 1.4.1 Discovery, 1.4.2 Quote, 1.4.3 Execute, 1.4.4 Callback & Workflow Completion, 1.4.5 Status Poll (+29 more)

### Community 13 - "Cancellation — Temporal Workflow"
Cohesion: 0.07
Nodes (29): 1.1 Purpose, 1.2 Key design decisions, 1.3 Workflow settings, 1. Overview, 2.1 Three-phase flow, 2.2 Multi-pod signal routing, 2. Architecture, 3.1 Step sequence (+21 more)

### Community 14 - "UIKit"
Cohesion: 0.12
Nodes (14): CTkLabel, CTkFrame, Widget factory bound to one root window and one palette at a time., Push the palette into CustomTkinter's and ttk's global state., Repaint every widget this kit created for the new palette., Shared backdrop bitmap for (bucketed) size + palette. Returns (photo, key).…, A bordered surface panel. When ``title`` is given the card exposes ``.body`` --…, Small ⓘ carrying help text on hover, instead of a wall of static prose. (+6 more)

### Community 16 - "converter_core.py"
Cohesion: 0.06
Nodes (122): CompletedProcess, _anchor_point(), _anchor_rect(), apply_image_overlay(), apply_image_watermark(), _apply_overlay_to_pages(), apply_text_overlay(), apply_text_watermark() (+114 more)

### Community 17 - "smoke_test.py"
Cohesion: 0.15
Nodes (27): add_fingerprints(), BatchConfig, dependency_status(), PdfToolConfig, functools, http_server, extract_urls(), pypdf (+19 more)

### Community 18 - "._register"
Cohesion: 0.11
Nodes (15): Combobox, CTkButton, CTkCheckBox, CTkEntry, CTkSegmentedButton, CTkSlider, CTkSwitch, Entry (+7 more)

### Community 19 - "REPO SECTIONS"
Cohesion: 0.14
Nodes (13): 1. platform (first), 2. admin, 3. gokulmobilize, 4. bestportofolio, 5. Progress (GSpace), 6. Markdown-studio  (TO BE UPDATED: scope still open, revisit after the first 5 repos), 7. checkers-gym-management  (TO BE UPDATED: scope still open, revisit after admin works), 8. gokul-omni-converter  (TO BE UPDATED: standalone, not part of the web platform) (+5 more)

### Community 21 - "open_path"
Cohesion: 0.07
Nodes (6): compact_display_path(), open_path(), Path, Create (and register) the container for one page. ``padding`` and ``style`` are…, export_text_file(), CTkFrame

### Community 22 - "preview_support.py"
Cohesion: 0.20
Nodes (27): ImageDraw, ImageFont, pptx, _doc_preview(), _docx_excerpt(), _font(), _generic_summary(), _html_to_text() (+19 more)

### Community 23 - "RenameOrganizePage"
Cohesion: 0.14
Nodes (3): Path, Only show the inputs the selected rule actually uses., RenameOrganizePage

### Community 24 - "automation_core.py"
Cohesion: 0.19
Nodes (10): bundle_paths_as_zip(), discover_watch_candidates(), export_presets_to_json(), fingerprint_file(), import_presets_from_json(), move_files_to_archive(), Path, WatchFolderConfig (+2 more)

### Community 25 - "render"
Cohesion: 0.16
Nodes (19): _arc_box(), _gradient(), _hex_to_rgb(), main(), _masked(), Image, Rasterise the GO Omni Pro mark into the icon files the installers need. The…, Draw the mark at ``size`` pixels, supersampled then downscaled for smooth edges. (+11 more)

### Community 26 - "release_support.py"
Cohesion: 0.29
Nodes (12): build_example_update_manifest(), check_for_updates(), compare_versions(), export_workspace_bundle(), import_workspace_bundle(), load_update_manifest(), Any, Path (+4 more)

### Community 27 - "5.20.3 Trip Data Model Changes"
Cohesion: 0.07
Nodes (26): 5.20.1 Overview, 5.20.2.1 Overview, 5.20.2.2 Segment entries, 5.20.2.3 Ancillary entries, 5.20.2.4 Sample hold response (abbreviated), 5.20.2.5 Key principles, 5.20.2 Hold Response Changes, 5.20.3.1 Overview (+18 more)

### Community 28 - "link_ingest.py"
Cohesion: 0.19
Nodes (17): CancelFn, hashlib, cache_root_from_setting(), clear_cache_dir(), download_many_urls(), emit(), wait_while_paused(), _guess_extension_from_content_type() (+9 more)

### Community 29 - "mail_core.py"
Cohesion: 0.14
Nodes (17): email_message, EmailMessage, build_email_message(), build_eml_draft(), _connect(), create_mailto_url(), open_mailto_draft(), _parse_recipients() (+9 more)

### Community 30 - "._image_folder_worker_run"
Cohesion: 0.16
Nodes (5): log(), progress(), BatchConfig, ImageFolderPdfConfig, PdfToolConfig

### Community 31 - "Animator"
Cohesion: 0.16
Nodes (12): key(), Animator, apply(), frame(), ease_out(), StringVar, cubic-bezier(0.4, 0, 0.2, 1)-ish: fast start, gentle settle., Frame-based tweening on Tk's event loop. Every animation checks… (+4 more)

### Community 33 - "build_support.py"
Cohesion: 0.23
Nodes (15): main(), collect_installer_assets(), collect_package_versions(), export_diagnostics_report(), export_json(), export_state_snapshot(), export_support_bundle(), import_state_snapshot() (+7 more)

### Community 34 - "Any"
Cohesion: 0.18
Nodes (8): LogBox, ProgressBar, Any, Text-only options (state, font, wrap) are forwarded to the inner Text., CTkProgressBar that also speaks the ttk.Progressbar API. The conversion/OCR/PDF…, Ease toward ``value`` (0..1) so batch progress glides rather than jumps., Scrolling multi-line text area. Also avoids CTkTextbox, which embeds the same…, A ``tk.Text`` plus a ttk scrollbar, proxying the Text API. Used wherever the…

### Community 35 - ".__init__"
Cohesion: 0.11
Nodes (3): Set the window/taskbar icon from the app mark, per-platform., default_merged_name(), outputs_pdf()

### Community 36 - "asset_support.py"
Cohesion: 0.38
Nodes (14): asset_cache_root(), cache_summary(), cached_asset_path(), clear_asset_cache(), download_binary_asset(), download_text_file(), is_cache_fresh(), load_asset_config() (+6 more)

### Community 37 - "FirstLaunchSplashWindow"
Cohesion: 0.20
Nodes (4): FirstLaunchSplashWindow, LoginReminderToast, Misc, Path

### Community 38 - "visit"
Cohesion: 0.33
Nodes (8): cancel_pending(), visit(), on_enter(), on_leave(), on_press(), on_release(), widget_contains_pointer(), Button

### Community 39 - "load_about_profile"
Cohesion: 0.42
Nodes (8): load_about_profile(), _normalize_links(), Path, Read the profile, dropping any retired fields left over from older versions., Write only the supported fields back out, so retired keys never reappear., resolve_profile_image(), save_about_profile(), Any

### Community 40 - "MarkdownNotesWindow"
Cohesion: 0.11
Nodes (7): AboutProfileEditorWindow, BuildCenterWindow, MarkdownNotesWindow, Editor for the four profile fields plus the social links. Company, job title,…, SMTPSettingsWindow, Misc, ThemePalette

### Community 41 - "LLD — Project Repository Initial Setup & Architecture"
Cohesion: 0.08
Nodes (24): 1.1 Purpose, 1.2 Tech Stack, 1.3 High-Level System Context, 1. Executive Summary & Tech Stack, 2.1 Gradle Multi-Module Layout, 2.2 Package Tree (ASCII), 2.3 Module Responsibilities, 2. Directory & Module Structure (+16 more)

### Community 42 - "Misc"
Cohesion: 0.20
Nodes (6): attach_tooltip(), Misc, Treeview, CTk has no tree widget; this ``ttk.Treeview`` is styled by apply_ttk_theme., Hover hint. Deliberately plain ``tk`` so it can float above CTk widgets., Tooltip

### Community 43 - "ScrollArea"
Cohesion: 0.21
Nodes (4): A vertically scrolling page body. Deliberately *not* CTkScrollableFrame: in…, Vertically scrolling container that hosts CustomTkinter children. Mirrors the…, Coalesce Configure bursts into one geometry pass. Resizing the canvas window…, ScrollArea

### Community 44 - "test_ui_smoke.py"
Cohesion: 0.15
Nodes (11): faulthandler, os, sys, main(), count_info_dots(), Headless UI smoke test: build every page, switch every theme, check the…, _watchdog(), time (+3 more)

### Community 45 - "Ideas: One Platform, Many Apps"
Cohesion: 0.13
Nodes (14): 10. Open questions, 11. Excellidraw opensource, 1. Vision, 2. Current state (checked in repos), 3. Framework, 4. Content & CMS, 5. Roles (super admin → app admins), 6. Shared packages (so "every component works everywhere") (+6 more)

### Community 46 - "packages"
Cohesion: 0.18
Nodes (11): packages, openpyxl, pdfplumber, Pillow, PyMuPDF, pypdf, pytesseract, python-docx (+3 more)

### Community 47 - "smtp"
Cohesion: 0.18
Nodes (11): smtp, default_to, host, password, port, save_password, sender, timeout_seconds (+3 more)

### Community 48 - "REPO SECTIONS"
Cohesion: 0.14
Nodes (13): 1. platform (first), 2. admin, 3. gokulmobilize, 4. bestportofolio, 5. Progress (GSpace), 6. Markdown-studio  (TO BE UPDATED: scope still open, revisit after the first 5 repos), 7. checkers-gym-management  (TO BE UPDATED: scope still open, revisit after admin works), 8. gokul-omni-converter  (TO BE UPDATED: standalone, not part of the web platform) (+5 more)

### Community 49 - "ui_theme.py"
Cohesion: 0.12
Nodes (29): dataclasses, appearance_mode(), apply_ctk_theme(), apply_menu_theme(), apply_text_widget_theme(), apply_treeview_tag_colors(), button_colors(), card_colors() (+21 more)

### Community 50 - "WORKLOG: gokul-omni-converter"
Cohesion: 0.29
Nodes (6): Blockers / requests to other repos, Current state, Done, In progress / next, Reminders, WORKLOG: gokul-omni-converter

### Community 51 - "settings.py"
Cohesion: 0.31
Nodes (8): build(), _build_appearance(), _build_engine(), _build_files(), _build_quick_actions(), Settings page: appearance, engine, local files and folders, and quick actions., An entry that shows a path but cannot be edited., _readonly()

### Community 53 - "GradientBackdrop"
Cohesion: 0.28
Nodes (4): Canvas, GradientBackdrop, Paints the palette's gradient onto a tk.Canvas and keeps it sized. Tk has no…, Force a repaint, e.g. after the palette changed.

### Community 54 - "GO Omni Pro Activity Report"
Cohesion: 0.28
Nodes (9): LibreOffice Engine, Pure Python Engine, Smoke Test Output, DOCX to PDF via LibreOffice - Sample Output, DOCX to PDF via Pure Python - Sample Output, HTML to PDF Pure Python - Sample Output, Markdown to PDF Pure Python - Sample Output, BUILDING.md - Patch 16 Release Build Notes (+1 more)

### Community 55 - "Kickoff prompt: gokul-omni-converter"
Cohesion: 0.33
Nodes (5): 8. gokul-omni-converter  (TO BE UPDATED: standalone, not part of the web platform), COLLABORATOR ADD-ON, Kickoff prompt: gokul-omni-converter, REPO SECTION, SHARED BLOCK

### Community 57 - "kit.py"
Cohesion: 0.18
Nodes (10): ConversionError, CustomTkinter widgets avoided due to redraw loop, Design tokens via ui_theme for dependency inversion, Layered architecture dependency rule, apply(), lerp_color(), CustomTkinter widget kit for GO Omni Pro. Every screen builds from…, Blend two hex colours. Used by the animation helpers. (+2 more)

### Community 58 - "ButtonBar"
Cohesion: 0.31
Nodes (4): ButtonBar, Widget, Wrapping row of buttons, API-compatible with the old FlowButtonBar. ``add``…, Wrapping action row. Replaces the ttk-era FlowButtonBar.

### Community 59 - "app.py"
Cohesion: 0.16
Nodes (15): argparse, customtkinter, fitz, pathlib, pil, queue, subprocess, threading (+7 more)

### Community 60 - ".radio"
Cohesion: 0.22
Nodes (8): CTkRadioButton, Button, CheckBox, RadioButton, Stops CustomTkinter's redraw from recursing into itself. Several CTk widgets…, Adds ttk's ``state([...])`` to a CustomTkinter widget. _update_mode_controls…, SafeDrawMixin, TtkStateMixin

### Community 61 - "diagnostics.json"
Cohesion: 0.25
Nodes (7): created_at, current_context, last_outputs, selected_files, extra, smoke, installer_assets

### Community 63 - "Supply Cancellation Execute — Success / Failure Event Flow"
Cohesion: 0.18
Nodes (10): 1. Overview, 2. Actors, 3. Success Path, 4.1 Agent Action (Manual Queue), 4. Failure Path, 5. Flow Diagram, 6. Notes, ASCII (+2 more)

### Community 64 - "Repo Kit (every repo must have these)"
Cohesion: 0.33
Nodes (5): Definition of done (any step), Deploy-ready rules (every repo), Master docs (live in `platform/docs/`), Repo Kit (every repo must have these), Shared constants (from `@gokulsaraswat/ui`, same on every app)

### Community 65 - "PDF Conversion"
Cohesion: 0.29
Nodes (7): PDF Conversion, Image Folder All - Combined PDF, Image Folder Per Folder - Nested PDF, Patch 36.5 Stability - PNGs to PDF, PDF to HTML - Multi-page Sample Output, PDF Tool Compress - Compressed Output, PDF Tool Edit Metadata - Metadata Sample

### Community 66 - "paths"
Cohesion: 0.29
Nodes (7): paths, about_profile_path, asset_config_path, installer_dir, notes_path, output_dir, state_path

### Community 68 - "dependencies"
Cohesion: 0.33
Nodes (6): dependencies, LibreOffice, Pandoc, pdftotext, python-pptx, xlrd

### Community 69 - "PDF Split Every N Output Part 1"
Cohesion: 0.33
Nodes (6): PDF Split Every N Output Part 1, PDF Split Every N Output Part 2, PDF Split Every N Output Part 3, PDF Split Ranges Output Part 1, PDF Split Ranges Output Part 2, PDF Split Ranges Output Part 3

### Community 70 - "automation.py"
Cohesion: 0.53
Nodes (5): build(), _build_lower(), _build_presets(), _build_watch(), Automation page: saved presets, the watch folder, and output bundling.

### Community 71 - "CX Cancellation Flow"
Cohesion: 0.22
Nodes (8): Actors, CX Cancellation Flow, Flowchart, Notes, Open Issue — `quote_id` on Execute, Phase 1 — Cancel Discovery (via Trips), Phase 2 — Refund Discovery (Quote), Phase 3 — Cancellation Execute

### Community 72 - "gsclaude/sync-architecture.sh"
Cohesion: 0.70
Nodes (3): edge(), node(), sync-architecture.sh script

### Community 73 - "system"
Cohesion: 0.40
Nodes (5): system, cwd, platform, python, python_executable

### Community 74 - "🤖 System Architecture & Co-Worker Protocol"
Cohesion: 0.20
Nodes (9): 🔄 Auto-Sync: System Architecture & Graphify, ⏱️ Communication Constraints (Zero Waste Mode), Context, 📐 Engineering Guardrails, 🎨 Frontend Architecture, UI/UX & Motion Standard, 🛠️ Project Execution Modes (Personal vs. Organization), STANDING RULES (from Gokul, apply in every session), 🤖 System Architecture & Co-Worker Protocol (+1 more)

### Community 75 - "GO Omni Pro app icon 256px — purple/indigo rounded-square badge with circular convert arrow and document sheet icon"
Cohesion: 0.83
Nodes (4): GO Omni Pro app icon 256px — purple/indigo rounded-square badge with circular convert arrow and document sheet icon, GO Omni Pro app icon 48px — small version of the purple document-convert badge, GO Omni Pro app icon 512px — large high-res version of the purple document-convert badge with circular arrow, GO Omni Pro SVG vector mark — indigo-to-violet gradient badge with circular convert arrow and document sheet, scalable source for all PNG sizes

### Community 77 - "Patch 16 profile image — solid blue square, test profile asset for patch 16 release smoke test"
Cohesion: 0.67
Nodes (3): Profile placeholder image — dark navy card with rounded corners, blue glow border, generic silhouette avatar, 'Profile Placeholder' label, Patch 16 imported bundle profile image — solid blue square, likely a test profile asset, Patch 16 profile image — solid blue square, test profile asset for patch 16 release smoke test

### Community 78 - "app"
Cohesion: 0.67
Nodes (3): app, name, version

### Community 108 - "Prompt Playbook — CLAUDE.md Workflow"
Cohesion: 0.22
Nodes (8): 1. Starting a New Folder / Project, 2. Starting with an Existing Project, 3. The "After-Command" Chaser, 4. Start / Debug Prompt Template, 5. Session Wrap-Up Prompt Template, 6. The Context Recovery Prompt (Anti-Hallucination Trigger), 7. Pending Work — GO Omni Pro (resume prompt), Prompt Playbook — CLAUDE.md Workflow

### Community 109 - "Keyboard Shortcuts"
Cohesion: 0.25
Nodes (7): Core actions, Keyboard Shortcuts, Navigation, Organizer, Recovery and support, Release validation, Tips

### Community 110 - "Prompt Playbook — CLAUDE.md Workflow"
Cohesion: 0.25
Nodes (7): 1. Starting a New Folder / Project, 2. Starting with an Existing Project, 3. The "After-Command" Chaser, 4. Start / Debug Prompt Template, 5. Session Wrap-Up Prompt Template, 6. The Context Recovery Prompt (Anti-Hallucination Trigger), Prompt Playbook — CLAUDE.md Workflow

### Community 111 - "🤖 System Architecture & Co-Worker Protocol"
Cohesion: 0.29
Nodes (6): ⏱️ Communication Constraints (Zero Waste Mode), 📐 Engineering Guardrails, 🎨 Frontend Architecture, UI/UX & Motion Standard, 🛠️ Project Execution Modes (Personal vs. Organization), 🤖 System Architecture & Co-Worker Protocol, 📉 Token-Optimized Communication Mode

### Community 113 - "Release Checklist — GO Omni Pro 2.3.0"
Cohesion: 0.33
Nodes (5): Installer check, Manual UI check, Release assets, Release Checklist — GO Omni Pro 2.3.0, Source validation

### Community 115 - "GO Omni Pro Notes"
Cohesion: 0.40
Nodes (4): GO Omni Pro Notes, Next patch ideas, Patch 8 highlights, Suggested uses

### Community 116 - "sync-architecture.sh"
Cohesion: 0.70
Nodes (3): edge(), node(), sync-architecture.sh script

### Community 117 - "convert.py"
Cohesion: 0.36
Nodes (9): build(), _build_center(), _build_image_folder_card(), _build_inputs(), _build_links(), _build_log_card(), _build_mode_card(), _build_options() (+1 more)

### Community 118 - "workflow_support.py"
Cohesion: 0.46
Nodes (6): directory_stats(), DirectoryStats, format_bytes(), prune_directory(), Path, summarize_directory()

### Community 125 - "apply_ttk_theme"
Cohesion: 0.23
Nodes (12): Style, apply_ttk_theme(), _configure_fonts(), font_spec(), Misc, Return the first font in the stack that this Tk install actually has., A (family, size, weight) tuple usable anywhere Tk accepts a font., Scaled spacing token. Unknown names fall back to the medium step. (+4 more)

### Community 126 - "CONTEXT: read this first in every new session"
Cohesion: 0.20
Nodes (9): CONTEXT: read this first in every new session, Future changes (expect these; do not block them), How repos interact (the contracts), Repo map, STANDING RULES (from Gokul, apply in every session), This repo: gokul-omni-converter, Vision, Work modes (ask Gokul which one if unclear) (+1 more)

### Community 133 - "6.6 Refund Quote Entity"
Cohesion: 0.22
Nodes (8): 6.6 Refund Quote Entity, 6.6 Refund Quote Entity (Redis), Design principles, Field reference, Penalty fields, Quote API vs Redis (example), Redis JSON, Relationship to §6.5 Snapshot

### Community 134 - "re"
Cohesion: 0.46
Nodes (6): re, validate_ui_text_helpers(), format_engine_label(), format_flag(), humanize_identifier(), _split_identifier()

### Community 135 - "Prompt Playbook — CLAUDE.md Workflow"
Cohesion: 0.25
Nodes (7): 1. Starting a New Folder / Project, 2. Starting with an Existing Project, 3. The "After-Command" Chaser, 4. Start / Debug Prompt Template, 5. Session Wrap-Up Prompt Template, 6. The Context Recovery Prompt (Anti-Hallucination Trigger), Prompt Playbook — CLAUDE.md Workflow

### Community 136 - "pdf_tools.py"
Cohesion: 0.39
Nodes (6): build(), _build_center(), _build_log_card(), _build_options(), _build_tool_card(), PDF Tools page. The options panel exposes every field any tool might need;…

### Community 137 - "🤖 System Architecture & Co-Worker Protocol"
Cohesion: 0.29
Nodes (6): ⏱️ Communication Constraints (Zero Waste Mode), 📐 Engineering Guardrails, 🎨 Frontend Architecture, UI/UX & Motion Standard, 🛠️ Project Execution Modes (Personal vs. Organization), 🤖 System Architecture & Co-Worker Protocol, 📉 Token-Optimized Communication Mode

### Community 138 - "Repo Kit (every repo must have these)"
Cohesion: 0.33
Nodes (5): Definition of done (any step), Deploy-ready rules (every repo), Master docs (live in `platform/docs/`), Repo Kit (every repo must have these), Shared constants (from `@gokulsaraswat/ui`, same on every app)

### Community 139 - "Agent/sync-architecture.sh"
Cohesion: 0.70
Nodes (3): edge(), node(), sync-architecture.sh script

## Knowledge Gaps
- **347 isolated node(s):** `1. Starting a New Folder / Project`, `2. Starting with an Existing Project`, `3. The "After-Command" Chaser`, `4. Start / Debug Prompt Template`, `5. Session Wrap-Up Prompt Template` (+342 more)
  These have ≤1 connection - possible missing edges or undocumented components. (Counts symbols only; 684 node(s) total have ≤1 connection when file, concept and rationale nodes are included.)
- **53 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `GOOmniProApp` connect `GOOmniProApp` to `._refresh_about_profile`, `build_support.py`, `.__init__`, `._show_page`, `visit`, `re`, `MarkdownNotesWindow`, `ScrollablePage`, `._log_automation`, `._apply_state_snapshot`, `open_path`, `._persist_state`, `._poll_worker_queue`, `automation_core.py`, `app.py`, `mail_core.py`, `._image_folder_worker_run`?**
  _High betweenness centrality (0.135) - this node is a cross-community bridge._
- **Why does `UIKit` connect `UIKit` to `Any`, `.__init__`, `Misc`, `ScrollArea`, `ui_theme.py`, `._register`, `GradientBackdrop`, `kit.py`, `ButtonBar`, `.radio`?**
  _High betweenness centrality (0.037) - this node is a cross-community bridge._
- **Why does `PageOrganizerPanel` connect `PageOrganizerPanel` to `ui_theme.py`, `app.py`?**
  _High betweenness centrality (0.036) - this node is a cross-community bridge._
- **Are the 4 inferred relationships involving `ConversionError` (e.g. with `test_archive()` and `test_rename()`) actually correct?**
  _`ConversionError` has 4 INFERRED edges - model-reasoned connections that need verification._
- **What connects `1. Starting a New Folder / Project`, `2. Starting with an Existing Project`, `3. The "After-Command" Chaser` to the rest of the system?**
  _347 weakly-connected nodes found - possible documentation gaps or missing edges._
- **Should `PageOrganizerPanel` be split into smaller, more focused modules?**
  _Cohesion score 0.05207920792079208 - nodes in this community are weakly interconnected._
- **Should `ArchivePage` be split into smaller, more focused modules?**
  _Cohesion score 0.05754527162977867 - nodes in this community are weakly interconnected._