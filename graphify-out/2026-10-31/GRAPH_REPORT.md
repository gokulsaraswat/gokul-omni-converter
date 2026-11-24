# Graph Report - gokul-omni-converter  (2026-10-31)

## Corpus Check
- 107 files · ~106,033 words
- Verdict: corpus is large enough that graph structure adds value.
- Unclassified: 16 file(s) not represented in the graph (top: (none) 3, .zip 3, .bat 2)

## Summary
- 1742 nodes · 4449 edges · 122 communities (76 shown, 46 thin omitted)
- Extraction: 98% EXTRACTED · 2% INFERRED · 0% AMBIGUOUS · INFERRED: 101 edges (avg confidence: 0.91)
- Token cost: 0 input · 0 output

## Graph Freshness
- Built from commit: `86b1f89a`
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
- process_pdf_tool
- AppStateStore
- ocr_core.py
- ScrollablePage
- ConversionError
- _apply_overlay_to_pages
- ._show_page
- .frame
- ._log_automation
- converter_core.py
- smoke_test.py
- UIKit
- REPO SECTIONS
- open_path
- preview_support.py
- RenameOrganizePage
- automation_core.py
- render
- image_folder_pdf.py
- test_cores.py
- link_ingest.py
- mail_core.py
- Path
- Animator
- ._refresh_about_profile
- build_support.py
- Any
- .__init__
- asset_support.py
- FirstLaunchSplashWindow
- apply_text_widget_theme
- load_about_profile
- ThemePalette
- main
- Misc
- ScrollArea
- .radio
- _StructuredHTMLParser
- packages
- smtp
- format_engine_label
- kit.py
- WORKLOG: gokul-omni-converter
- settings.py
- ._update_login_popup_state
- GradientBackdrop
- GO Omni Pro Activity Report
- Kickoff prompt: gokul-omni-converter
- bind_responsive_wrap
- ButtonBar
- app.py
- _BasicHTMLTextExtractor
- diagnostics.json
- CommandPalette
- pdf_tools.py
- Repo Kit (every repo must have these)
- PDF Conversion
- paths
- Tooltip
- dependencies
- PDF Split Every N Output Part 1
- automation.py
- FlowButtonBar
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
- html_to_md/sample.md
- link_cache/sample.md
- patch16_release/footer_notes.md
- imported_bundle/footer_notes.md
- BUILDING.md

## God Nodes (most connected - your core abstractions)
1. `GOOmniProApp` - 277 edges
2. `ConversionError` - 62 edges
3. `PageOrganizerPanel` - 59 edges
4. `ThemePalette` - 52 edges
5. `run()` - 51 edges
6. `AppStateStore` - 45 edges
7. `ensure_directory()` - 40 edges
8. `ArchivePage` - 36 edges
9. `process_pdf_tool()` - 33 edges
10. `process_batch()` - 31 edges

## Surprising Connections (you probably didn't know these)
- `GOOmniProApp` --uses--> `PageOrganizerPanel`  [INFERRED]
  app.py → page_organizer.py
- `PageOrganizerPanel` --uses--> `ThemePalette`  [INFERRED]
  page_organizer.py → ui_theme.py
- `GOOmniProApp` --uses--> `AnimatedGifLabel`  [INFERRED]
  app.py → responsive_ui.py
- `GOOmniProApp` --uses--> `ScrollablePage`  [INFERRED]
  app.py → responsive_ui.py
- `GOOmniProApp` --uses--> `QuickAction`  [INFERRED]
  app.py → workflow_ui.py

## Import Cycles
- None detected.

## Communities (122 total, 46 thin omitted)

### Community 0 - "PageOrganizerPanel"
Cohesion: 0.06
Nodes (40): Document, OrganizedPage, build_default_sequence(), duplicate_positions(), export_pages_as_images(), extract_selected_pdf(), move_positions_down(), move_positions_to_index() (+32 more)

### Community 1 - "ArchivePage"
Cohesion: 0.05
Nodes (34): ArchiveEntry, ArchiveSummary, batch_compress(), _collect_folder_files(), _compression(), create_archive(), extract_archive(), format_size() (+26 more)

### Community 2 - "GOOmniProApp"
Cohesion: 0.05
Nodes (5): GOOmniProApp, CTkFrame, StringVar, Widget, Create (and register) the container for one page. ``padding`` and ``style`` are…

### Community 3 - "release_validation.py"
Cohesion: 0.18
Nodes (18): log(), progress(), main(), ImageFolderPdfConfig, json, py_compile, _compile_files(), _fail() (+10 more)

### Community 4 - "rename_organizer_core.py"
Cohesion: 0.09
Nodes (35): apply_case(), apply_plan(), _bucket_for(), describe_manifest(), _file_dates(), latest_manifest(), _log(), manifest_dir() (+27 more)

### Community 5 - "pure_python_renderers.py"
Cohesion: 0.08
Nodes (40): _Cell, markdown_to_pdf(), csv, docx, docx_document, docx_oxml_table, docx_oxml_text_paragraph, docx_shared (+32 more)

### Community 6 - "WORKLOG"
Cohesion: 0.08
Nodes (23): Agreed order, Architecture (agreed 2026-10-31), Cleanup scan: `/cleanscan` (Claude slash command), Commit + push everything: `/gitmd` (Claude slash command), Current focus, Decisions, Dropped, Far later (+15 more)

### Community 7 - "process_pdf_tool"
Cohesion: 0.25
Nodes (15): extract_pdf_pages(), _log(), page_indices_to_label(), parse_page_groups(), parse_page_spec(), process_pdf_tool(), _progress(), LogFn (+7 more)

### Community 8 - "AppStateStore"
Cohesion: 0.08
Nodes (23): AppStateStore, _default_state_dir(), migrate_legacy_state_dir(), Any, Path, Copy a pre-2.4.0 state directory to the new location, once. Copies rather than…, Per-platform state location, following each OS's own convention., datetime (+15 more)

### Community 9 - "ocr_core.py"
Cohesion: 0.14
Nodes (36): contextlib, fitz, io, _candidate_tessdata_dirs(), detect_tesseract_status(), _emit_log(), _emit_progress(), extract_text_with_ocr() (+28 more)

### Community 10 - "ScrollablePage"
Cohesion: 0.15
Nodes (6): AnimatedGifLabel, Misc, Path, A simple responsive page shell with a vertical scrollbar. The inner frame…, Lightweight animated GIF label with a safe text fallback., ScrollablePage

### Community 11 - "ConversionError"
Cohesion: 0.19
Nodes (17): _coerce_rect_value(), ConversionError, convert_images_to_single_pdf(), convert_input_to_pdf_for_mode(), _draw_images_to_pdf_with_reportlab(), file_to_pdf(), _image_to_pdf_ready_rgb(), looks_like_text() (+9 more)

### Community 12 - "_apply_overlay_to_pages"
Cohesion: 0.24
Nodes (16): _anchor_point(), _anchor_rect(), apply_image_overlay(), apply_image_watermark(), _apply_overlay_to_pages(), apply_text_overlay(), apply_text_watermark(), apply_visible_signature() (+8 more)

### Community 14 - ".frame"
Cohesion: 0.16
Nodes (9): CTkLabel, CTkFrame, A bordered surface panel. When ``title`` is given the card exposes ``.body`` --…, Small ⓘ carrying help text on hover, instead of a wall of static prose., A label with an optional ⓘ beside it., A card whose border warms toward the accent while the pointer is over it., Label on the left, control on the right, optional hint underneath. Returns the…, Shown where a list would be, when the list is empty. Always tells the user the… (+1 more)

### Community 16 - "converter_core.py"
Cohesion: 0.10
Nodes (37): _autosize_columns(), _choose_blank_slide_layout(), collect_files_from_folder(), dedupe_paths(), document_to_plain_text(), extract_html_title(), extract_text_from_html_string(), _extract_text_from_odf_document() (+29 more)

### Community 17 - "smoke_test.py"
Cohesion: 0.14
Nodes (29): add_fingerprints(), BatchConfig, dependency_status(), PdfToolConfig, functools, http_server, cache_root_from_setting(), extract_urls() (+21 more)

### Community 18 - "UIKit"
Cohesion: 0.12
Nodes (14): CTkButton, CTkCheckBox, CTkEntry, CTkSegmentedButton, CTkSlider, CTkSwitch, attach_tooltip(), Entry (+6 more)

### Community 19 - "REPO SECTIONS"
Cohesion: 0.14
Nodes (13): 1. platform (first), 2. admin, 3. gokulmobilize, 4. bestportofolio, 5. Progress (GSpace), 6. Markdown-studio  (TO BE UPDATED: scope still open, revisit after the first 5 repos), 7. checkers-gym-management  (TO BE UPDATED: scope still open, revisit after admin works), 8. gokul-omni-converter  (TO BE UPDATED: standalone, not part of the web platform) (+5 more)

### Community 21 - "open_path"
Cohesion: 0.07
Nodes (4): compact_display_path(), open_path(), Path, export_text_file()

### Community 22 - "preview_support.py"
Cohesion: 0.10
Nodes (30): ImageDraw, ImageFont, pptx, _doc_preview(), _docx_excerpt(), _font(), _generic_summary(), _html_to_text() (+22 more)

### Community 23 - "RenameOrganizePage"
Cohesion: 0.05
Nodes (36): cancel_pending(), visit(), on_enter(), on_leave(), on_press(), on_release(), widget_contains_pointer(), dnd_available() (+28 more)

### Community 24 - "automation_core.py"
Cohesion: 0.20
Nodes (11): bundle_paths_as_zip(), discover_watch_candidates(), export_presets_to_json(), fingerprint_file(), import_presets_from_json(), move_files_to_archive(), normalize_preset_record(), PresetRecord (+3 more)

### Community 25 - "render"
Cohesion: 0.08
Nodes (30): faulthandler, _arc_box(), _gradient(), _hex_to_rgb(), main(), _masked(), Image, Rasterise the GO Omni Pro mark into the icon files the installers need. The… (+22 more)

### Community 26 - "image_folder_pdf.py"
Cohesion: 0.34
Nodes (13): build_image_folder_pdfs(), discover_image_files(), _emit_log(), _emit_progress(), group_images_by_folder(), _is_hidden(), _natural_key(), LogFn (+5 more)

### Community 27 - "test_cores.py"
Cohesion: 0.38
Nodes (9): shutil, tempfile, check(), main(), make_tree(), Path, Headless tests for archive_core and rename_organizer_core. These two modules…, test_archive() (+1 more)

### Community 28 - "link_ingest.py"
Cohesion: 0.15
Nodes (17): CancelFn, hashlib, clear_cache_dir(), download_many_urls(), emit(), wait_while_paused(), _guess_extension_from_content_type(), LinkDownloadResult (+9 more)

### Community 29 - "mail_core.py"
Cohesion: 0.13
Nodes (17): email_message, EmailMessage, build_email_message(), build_eml_draft(), _connect(), create_mailto_url(), open_mailto_draft(), _parse_recipients() (+9 more)

### Community 30 - "Path"
Cohesion: 0.19
Nodes (34): CompletedProcess, build_conversion_route_preview(), compress_pdf(), copy_as_pdf(), describe_input_route_for_mode(), describe_pdf_output_route(), _docx_apply_runs(), edit_pdf_metadata() (+26 more)

### Community 31 - "Animator"
Cohesion: 0.15
Nodes (13): key(), Animator, apply(), apply(), frame(), ease_out(), StringVar, cubic-bezier(0.4, 0, 0.2, 1)-ish: fast start, gentle settle. (+5 more)

### Community 33 - "build_support.py"
Cohesion: 0.23
Nodes (14): collect_installer_assets(), collect_package_versions(), export_diagnostics_report(), export_json(), export_state_snapshot(), export_support_bundle(), import_state_snapshot(), Any (+6 more)

### Community 34 - "Any"
Cohesion: 0.14
Nodes (10): Combobox, ProgressBar, Any, Menu, Treeview, CTkProgressBar that also speaks the ttk.Progressbar API. The conversion/OCR/PDF…, Ease toward ``value`` (0..1) so batch progress glides rather than jumps., A vertically scrolling page body. Deliberately *not* CTkScrollableFrame: in… (+2 more)

### Community 35 - ".__init__"
Cohesion: 0.12
Nodes (3): Set the window/taskbar icon from the app mark, per-platform., default_merged_name(), outputs_pdf()

### Community 36 - "asset_support.py"
Cohesion: 0.38
Nodes (14): asset_cache_root(), cache_summary(), cached_asset_path(), clear_asset_cache(), download_binary_asset(), download_text_file(), is_cache_fresh(), load_asset_config() (+6 more)

### Community 37 - "FirstLaunchSplashWindow"
Cohesion: 0.20
Nodes (4): FirstLaunchSplashWindow, LoginReminderToast, Misc, Path

### Community 38 - "apply_text_widget_theme"
Cohesion: 0.22
Nodes (7): Listbox, Text, CTk has no list widget, so this is a ``tk.Listbox`` kept in palette., Plain ``tk.Text`` for rich/tagged content CTkTextbox cannot express., apply_text_widget_theme(), Listbox, Text

### Community 39 - "load_about_profile"
Cohesion: 0.22
Nodes (10): load_about_profile(), _normalize_links(), Any, Path, Read the profile, dropping any retired fields left over from older versions., Write only the supported fields back out, so retired keys never reappear., resolve_profile_image(), save_about_profile() (+2 more)

### Community 40 - "ThemePalette"
Cohesion: 0.10
Nodes (10): BuildCenterWindow, MarkdownNotesWindow, Misc, SMTPSettingsWindow, PreviewWindow, apply_menu_theme(), detect_system_theme(), Menu (+2 more)

### Community 41 - "main"
Cohesion: 0.30
Nodes (13): main(), build_example_update_manifest(), check_for_updates(), compare_versions(), export_workspace_bundle(), import_workspace_bundle(), load_update_manifest(), Any (+5 more)

### Community 42 - "Misc"
Cohesion: 0.16
Nodes (7): LogBox, Misc, Text-only options (state, font, wrap) are forwarded to the inner Text., Scrolling multi-line text area. Also avoids CTkTextbox, which embeds the same…, Hover hint. Deliberately plain ``tk`` so it can float above CTk widgets., A ``tk.Text`` plus a ttk scrollbar, proxying the Text API. Used wherever the…, Tooltip

### Community 43 - "ScrollArea"
Cohesion: 0.25
Nodes (3): Vertically scrolling container that hosts CustomTkinter children. Mirrors the…, Coalesce Configure bursts into one geometry pass. Resizing the canvas window…, ScrollArea

### Community 44 - ".radio"
Cohesion: 0.22
Nodes (8): CTkRadioButton, Button, CheckBox, RadioButton, Stops CustomTkinter's redraw from recursing into itself. Several CTk widgets…, Adds ttk's ``state([...])`` to a CustomTkinter widget. _update_mode_controls…, SafeDrawMixin, TtkStateMixin

### Community 45 - "_StructuredHTMLParser"
Cohesion: 0.36
Nodes (4): _html_to_blocks(), HtmlBlock, HTMLParser, _StructuredHTMLParser

### Community 46 - "packages"
Cohesion: 0.18
Nodes (11): packages, openpyxl, pdfplumber, Pillow, PyMuPDF, pypdf, pytesseract, python-docx (+3 more)

### Community 47 - "smtp"
Cohesion: 0.18
Nodes (11): smtp, default_to, host, password, port, save_password, sender, timeout_seconds (+3 more)

### Community 48 - "format_engine_label"
Cohesion: 0.57
Nodes (5): validate_ui_text_helpers(), format_engine_label(), format_flag(), humanize_identifier(), _split_identifier()

### Community 49 - "kit.py"
Cohesion: 0.08
Nodes (40): ConversionError, CustomTkinter widgets avoided due to redraw loop, Design tokens via ui_theme for dependency inversion, Layered architecture dependency rule, Style, lerp_color(), CustomTkinter widget kit for GO Omni Pro. Every screen builds from…, Blend two hex colours. Used by the animation helpers. (+32 more)

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

### Community 57 - "bind_responsive_wrap"
Cohesion: 0.40
Nodes (3): visit(), bind_responsive_wrap(), Shrink wraplength when containers narrow, while preserving the original wide-…

### Community 58 - "ButtonBar"
Cohesion: 0.32
Nodes (4): ButtonBar, Widget, Wrapping row of buttons, API-compatible with the old FlowButtonBar. ``add``…, Wrapping action row. Replaces the ttk-era FlowButtonBar.

### Community 59 - "app.py"
Cohesion: 0.10
Nodes (26): argparse, customtkinter, dataclasses, pathlib, pil, queue, subprocess, threading (+18 more)

### Community 61 - "diagnostics.json"
Cohesion: 0.25
Nodes (7): created_at, current_context, last_outputs, selected_files, extra, smoke, installer_assets

### Community 63 - "pdf_tools.py"
Cohesion: 0.39
Nodes (6): build(), _build_center(), _build_log_card(), _build_options(), _build_tool_card(), PDF Tools page. The options panel exposes every field any tool might need;…

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

### Community 71 - "FlowButtonBar"
Cohesion: 0.47
Nodes (3): FlowButtonBar, Widget, A small wrapping button bar for responsive toolbars and action rows.

### Community 72 - "gsclaude/sync-architecture.sh"
Cohesion: 0.70
Nodes (3): edge(), node(), sync-architecture.sh script

### Community 73 - "system"
Cohesion: 0.40
Nodes (5): system, cwd, platform, python, python_executable

### Community 74 - "🤖 System Architecture & Co-Worker Protocol"
Cohesion: 0.25
Nodes (7): 🔄 Auto-Sync: System Architecture & Graphify, ⏱️ Communication Constraints (Zero Waste Mode), 📐 Engineering Guardrails, 🎨 Frontend Architecture, UI/UX & Motion Standard, 🛠️ Project Execution Modes (Personal vs. Organization), 🤖 System Architecture & Co-Worker Protocol, 📉 Token-Optimized Communication Mode

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
Cohesion: 0.25
Nodes (7): 1. Starting a New Folder / Project, 2. Starting with an Existing Project, 3. The "After-Command" Chaser, 4. Start / Debug Prompt Template, 5. Session Wrap-Up Prompt Template, 6. The Context Recovery Prompt (Anti-Hallucination Trigger), Prompt Playbook — CLAUDE.md Workflow

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

## Knowledge Gaps
- **179 isolated node(s):** `AGENTS.md: gokul-omni-converter`, `COLLABORATORS: gokul-omni-converter`, `Current state`, `Done`, `In progress / next` (+174 more)
  These have ≤1 connection - possible missing edges or undocumented components. (Counts symbols only; 490 node(s) total have ≤1 connection when file, concept and rationale nodes are included.)
- **46 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `GOOmniProApp` connect `GOOmniProApp` to `PageOrganizerPanel`, `release_validation.py`, `AppStateStore`, `ocr_core.py`, `ScrollablePage`, `._show_page`, `._log_automation`, `smoke_test.py`, `UIKit`, `._persist_state`, `open_path`, `preview_support.py`, `RenameOrganizePage`, `automation_core.py`, `link_ingest.py`, `mail_core.py`, `Animator`, `._refresh_about_profile`, `.__init__`, `FirstLaunchSplashWindow`, `ThemePalette`, `main`, `format_engine_label`, `._update_login_popup_state`, `._start_ocr_job`, `bind_responsive_wrap`, `app.py`, `CommandPalette`?**
  _High betweenness centrality (0.215) - this node is a cross-community bridge._
- **Why does `UIKit` connect `UIKit` to `GOOmniProApp`, `.__init__`, `Any`, `apply_text_widget_theme`, `ThemePalette`, `Misc`, `.radio`, `.frame`, `kit.py`, `GradientBackdrop`, `ButtonBar`, `app.py`?**
  _High betweenness centrality (0.094) - this node is a cross-community bridge._
- **Why does `PageOrganizerPanel` connect `PageOrganizerPanel` to `ThemePalette`, `GOOmniProApp`, `app.py`?**
  _High betweenness centrality (0.080) - this node is a cross-community bridge._
- **Are the 18 inferred relationships involving `GOOmniProApp` (e.g. with `AppStateStore` and `WatchFolderConfig`) actually correct?**
  _`GOOmniProApp` has 18 INFERRED edges - model-reasoned connections that need verification._
- **Are the 4 inferred relationships involving `ConversionError` (e.g. with `test_archive()` and `test_rename()`) actually correct?**
  _`ConversionError` has 4 INFERRED edges - model-reasoned connections that need verification._
- **Are the 2 inferred relationships involving `PageOrganizerPanel` (e.g. with `GOOmniProApp` and `ThemePalette`) actually correct?**
  _`PageOrganizerPanel` has 2 INFERRED edges - model-reasoned connections that need verification._
- **Are the 12 inferred relationships involving `ThemePalette` (e.g. with `AboutProfileEditorWindow` and `BuildCenterWindow`) actually correct?**
  _`ThemePalette` has 12 INFERRED edges - model-reasoned connections that need verification._