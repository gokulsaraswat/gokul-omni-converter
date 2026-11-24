# Graph Report - gokul-omni-converter  (2026-11-24)

## Corpus Check
- 78 files · ~101,349 words
- Verdict: corpus is large enough that graph structure adds value.
- Unclassified: 9 file(s) not represented in the graph (top: (none) 3, .bat 2, .ico 1)

## Summary
- 1762 nodes · 4356 edges · 132 communities (77 shown, 55 thin omitted)
- Extraction: 98% EXTRACTED · 2% INFERRED · 0% AMBIGUOUS · INFERRED: 79 edges (avg confidence: 0.89)
- Token cost: 0 input · 0 output

## Graph Freshness
- Built from commit: `ac3d22e9`
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
- converter_core.py
- ConversionError
- ._show_page
- Misc
- ._log_automation
- Path
- run
- UIKit
- REPO SECTIONS
- ._persist_state
- open_path
- preview_support.py
- RenameOrganizePage
- automation_core.py
- render
- PreviewCenterWindow
- test_cores.py
- link_ingest.py
- mail_core.py
- ._resolve_link_cache_dir
- Animator
- ._refresh_about_profile
- build_support.py
- Any
- .__init__
- smoke_test.py
- FirstLaunchSplashWindow
- visit
- load_about_profile
- ._current_palette
- _apply_overlay_to_pages
- Tooltip
- ScrollArea
- .radio
- html_to_docx
- packages
- smtp
- pure_python_file_to_pdf
- kit.py
- WORKLOG: gokul-omni-converter
- settings.py
- ._apply_state_snapshot
- GradientBackdrop
- GO Omni Pro Activity Report
- Kickoff prompt: gokul-omni-converter
- ._start_ocr_job
- PreviewWindow
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
- datetime
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
- resolve_flow_layout_width
- workflow_support.py
- home.py
- html_to_md/sample.md
- link_cache/sample.md
- patch16_release/footer_notes.md
- imported_bundle/footer_notes.md
- BUILDING.md
- apply_ttk_theme
- .treeview
- Any
- CTkFrame
- Misc
- StringVar
- Widget

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
- `frame()` --indirect_call--> `key()`  [INFERRED]
  ui/kit.py → rename_organizer_core.py
- `PageOrganizerPanel` --uses--> `ThemePalette`  [INFERRED]
  page_organizer.py → ui_theme.py
- `PreviewWindow` --uses--> `ThemePalette`  [INFERRED]
  page_organizer.py → ui_theme.py
- `ArchivePage` --uses--> `ConversionError`  [INFERRED]
  ui/pages/archive.py → converter_core.py
- `test_archive()` --uses--> `ConversionError`  [INFERRED]
  tests/test_cores.py → converter_core.py

## Import Cycles
- None detected.

## Communities (132 total, 55 thin omitted)

### Community 0 - "PageOrganizerPanel"
Cohesion: 0.06
Nodes (40): Document, OrganizedPage, build_default_sequence(), duplicate_positions(), export_pages_as_images(), extract_selected_pdf(), move_positions_down(), move_positions_to_index() (+32 more)

### Community 1 - "ArchivePage"
Cohesion: 0.05
Nodes (34): ArchiveEntry, ArchiveSummary, batch_compress(), _collect_folder_files(), _compression(), create_archive(), extract_archive(), format_size() (+26 more)

### Community 2 - "GOOmniProApp"
Cohesion: 0.05
Nodes (5): GOOmniProApp, Create (and register) the container for one page. ``padding`` and ``style`` are…, CTkFrame, StringVar, Widget

### Community 3 - "release_validation.py"
Cohesion: 0.13
Nodes (35): main(), build_image_folder_pdfs(), discover_image_files(), _emit_log(), _emit_progress(), group_images_by_folder(), ImageFolderPdfConfig, _is_hidden() (+27 more)

### Community 4 - "rename_organizer_core.py"
Cohesion: 0.09
Nodes (36): apply_case(), apply_plan(), _bucket_for(), describe_manifest(), _file_dates(), latest_manifest(), _log(), manifest_dir() (+28 more)

### Community 5 - "pure_python_renderers.py"
Cohesion: 0.07
Nodes (44): _Cell, markdown_to_pdf(), csv, docx, docx_document, docx_oxml_table, docx_oxml_text_paragraph, docx_shared (+36 more)

### Community 6 - "WORKLOG"
Cohesion: 0.07
Nodes (28): Agreed order, Architecture (agreed 2026-10-31), Cleanup scan: `/cleanscan` (Claude slash command), Commit + push everything: `/gitmd` (Claude slash command), Current focus, Decisions, Dropped, Far later (+20 more)

### Community 7 - "register_drop_target"
Cohesion: 0.09
Nodes (29): tkinterdnd2, dnd_available(), DropZone, init_root(), parse_drop_paths(), Misc, Path, Drag-and-drop file input, with a working fallback when tkdnd is absent.… (+21 more)

### Community 8 - "AppStateStore"
Cohesion: 0.08
Nodes (17): AppStateStore, _default_state_dir(), migrate_legacy_state_dir(), Any, Path, Copy a pre-2.4.0 state directory to the new location, once. Copies rather than…, Per-platform state location, following each OS's own convention., normalize_preset_record() (+9 more)

### Community 9 - "ocr_core.py"
Cohesion: 0.14
Nodes (36): contextlib, fitz, io, _candidate_tessdata_dirs(), detect_tesseract_status(), _emit_log(), _emit_progress(), extract_text_with_ocr() (+28 more)

### Community 10 - "ScrollablePage"
Cohesion: 0.09
Nodes (12): visit(), AnimatedGifLabel, bind_responsive_wrap(), FlowButtonBar, Misc, Path, Widget, A simple responsive page shell with a vertical scrollbar. The inner frame… (+4 more)

### Community 11 - "converter_core.py"
Cohesion: 0.12
Nodes (27): _autosize_columns(), build_conversion_route_preview(), _choose_blank_slide_layout(), describe_input_route_for_mode(), describe_pdf_output_route(), document_to_plain_text(), _extract_text_from_odf_document(), _extract_text_from_xml_string() (+19 more)

### Community 12 - "ConversionError"
Cohesion: 0.19
Nodes (17): _coerce_rect_value(), ConversionError, extract_pdf_pages(), page_indices_to_label(), parse_page_groups(), _parse_page_number(), parse_page_spec(), parse_rect_spec() (+9 more)

### Community 14 - "Misc"
Cohesion: 0.17
Nodes (11): CTkLabel, CTkFrame, Misc, A bordered surface panel. When ``title`` is given the card exposes ``.body`` --…, Small ⓘ carrying help text on hover, instead of a wall of static prose., A label with an optional ⓘ beside it., A card whose border warms toward the accent while the pointer is over it., Label on the left, control on the right, optional hint underneath. Returns the… (+3 more)

### Community 16 - "Path"
Cohesion: 0.21
Nodes (30): collect_files_from_folder(), compress_pdf(), copy_as_pdf(), dedupe_paths(), _docx_apply_runs(), edit_pdf_metadata(), edit_pdf_text_best_effort(), ensure_directory() (+22 more)

### Community 17 - "run"
Cohesion: 0.23
Nodes (16): BatchConfig, dependency_status(), PdfToolConfig, build_conversion_jobs(), build_pdf_tool_jobs(), create_sample_files(), Path, run() (+8 more)

### Community 18 - "UIKit"
Cohesion: 0.09
Nodes (21): Combobox, CTkButton, CTkCheckBox, CTkEntry, CTkSegmentedButton, CTkSlider, CTkSwitch, attach_tooltip() (+13 more)

### Community 19 - "REPO SECTIONS"
Cohesion: 0.14
Nodes (13): 1. platform (first), 2. admin, 3. gokulmobilize, 4. bestportofolio, 5. Progress (GSpace), 6. Markdown-studio  (TO BE UPDATED: scope still open, revisit after the first 5 repos), 7. checkers-gym-management  (TO BE UPDATED: scope still open, revisit after admin works), 8. gokul-omni-converter  (TO BE UPDATED: standalone, not part of the web platform) (+5 more)

### Community 21 - "open_path"
Cohesion: 0.08
Nodes (3): compact_display_path(), open_path(), Path

### Community 22 - "preview_support.py"
Cohesion: 0.20
Nodes (27): ImageDraw, ImageFont, pptx, _doc_preview(), _docx_excerpt(), _font(), _generic_summary(), _html_to_text() (+19 more)

### Community 23 - "RenameOrganizePage"
Cohesion: 0.14
Nodes (4): build(), Path, Only show the inputs the selected rule actually uses., RenameOrganizePage

### Community 24 - "automation_core.py"
Cohesion: 0.18
Nodes (11): add_fingerprints(), bundle_paths_as_zip(), discover_watch_candidates(), export_presets_to_json(), fingerprint_file(), import_presets_from_json(), move_files_to_archive(), Path (+3 more)

### Community 25 - "render"
Cohesion: 0.12
Nodes (22): _arc_box(), _gradient(), _hex_to_rgb(), main(), _masked(), Image, Rasterise the GO Omni Pro mark into the icon files the installers need. The…, Draw the mark at ``size`` pixels, supersampled then downscaled for smooth edges. (+14 more)

### Community 26 - "PreviewCenterWindow"
Cohesion: 0.19
Nodes (3): PreviewCenterWindow, Misc, Path

### Community 27 - "test_cores.py"
Cohesion: 0.44
Nodes (8): tempfile, check(), main(), make_tree(), Path, Headless tests for archive_core and rename_organizer_core. These two modules…, test_archive(), test_rename()

### Community 28 - "link_ingest.py"
Cohesion: 0.16
Nodes (20): CancelFn, hashlib, cache_root_from_setting(), clear_cache_dir(), download_many_urls(), emit(), wait_while_paused(), extract_urls() (+12 more)

### Community 29 - "mail_core.py"
Cohesion: 0.17
Nodes (15): email_message, EmailMessage, build_email_message(), build_eml_draft(), _connect(), create_mailto_url(), open_mailto_draft(), _parse_recipients() (+7 more)

### Community 30 - "._resolve_link_cache_dir"
Cohesion: 0.12
Nodes (5): log(), progress(), BatchConfig, ImageFolderPdfConfig, PdfToolConfig

### Community 31 - "Animator"
Cohesion: 0.16
Nodes (12): Animator, apply(), apply(), frame(), ease_out(), StringVar, cubic-bezier(0.4, 0, 0.2, 1)-ish: fast start, gentle settle., Frame-based tweening on Tk's event loop. Every animation checks… (+4 more)

### Community 33 - "build_support.py"
Cohesion: 0.18
Nodes (15): collect_installer_assets(), collect_package_versions(), export_diagnostics_report(), export_json(), export_state_snapshot(), export_support_bundle(), export_text_file(), import_state_snapshot() (+7 more)

### Community 34 - "Any"
Cohesion: 0.13
Nodes (10): InfoDot, LogBox, ProgressBar, Any, Text-only options (state, font, wrap) are forwarded to the inner Text., CTkProgressBar that also speaks the ttk.Progressbar API. The conversion/OCR/PDF…, Ease toward ``value`` (0..1) so batch progress glides rather than jumps., A small ⓘ that reveals its explanation on hover. Used instead of printing a… (+2 more)

### Community 35 - ".__init__"
Cohesion: 0.11
Nodes (3): Set the window/taskbar icon from the app mark, per-platform., default_merged_name(), outputs_pdf()

### Community 36 - "smoke_test.py"
Cohesion: 0.15
Nodes (32): main(), asset_cache_root(), cache_summary(), cached_asset_path(), clear_asset_cache(), download_binary_asset(), download_text_file(), is_cache_fresh() (+24 more)

### Community 37 - "FirstLaunchSplashWindow"
Cohesion: 0.20
Nodes (4): FirstLaunchSplashWindow, LoginReminderToast, Misc, Path

### Community 38 - "visit"
Cohesion: 0.33
Nodes (8): cancel_pending(), visit(), on_enter(), on_leave(), on_press(), on_release(), widget_contains_pointer(), Button

### Community 39 - "load_about_profile"
Cohesion: 0.42
Nodes (8): load_about_profile(), _normalize_links(), Path, Read the profile, dropping any retired fields left over from older versions., Write only the supported fields back out, so retired keys never reappear., resolve_profile_image(), save_about_profile(), Any

### Community 40 - "._current_palette"
Cohesion: 0.10
Nodes (7): AboutProfileEditorWindow, BuildCenterWindow, MarkdownNotesWindow, Editor for the four profile fields plus the social links. Company, job title,…, SMTPSettingsWindow, Misc, ThemePalette

### Community 41 - "_apply_overlay_to_pages"
Cohesion: 0.24
Nodes (16): _anchor_point(), _anchor_rect(), apply_image_overlay(), apply_image_watermark(), _apply_overlay_to_pages(), apply_text_overlay(), apply_text_watermark(), apply_visible_signature() (+8 more)

### Community 43 - "ScrollArea"
Cohesion: 0.21
Nodes (4): A vertically scrolling page body. Deliberately *not* CTkScrollableFrame: in…, Vertically scrolling container that hosts CustomTkinter children. Mirrors the…, Coalesce Configure bursts into one geometry pass. Resizing the canvas window…, ScrollArea

### Community 44 - ".radio"
Cohesion: 0.22
Nodes (8): CTkRadioButton, Button, Entry, RadioButton, Stops CustomTkinter's redraw from recursing into itself. Several CTk widgets…, Adds ttk's ``state([...])`` to a CustomTkinter widget. _update_mode_controls…, SafeDrawMixin, TtkStateMixin

### Community 45 - "html_to_docx"
Cohesion: 0.25
Nodes (15): CompletedProcess, extract_html_title(), extract_pdf_pages_as_text(), extract_text_from_html_string(), find_command(), _html_blocks_to_markdown_lines(), html_to_docx(), html_to_markdown() (+7 more)

### Community 46 - "packages"
Cohesion: 0.18
Nodes (11): packages, openpyxl, pdfplumber, Pillow, PyMuPDF, pypdf, pytesseract, python-docx (+3 more)

### Community 47 - "smtp"
Cohesion: 0.18
Nodes (11): smtp, default_to, host, password, port, save_password, sender, timeout_seconds (+3 more)

### Community 48 - "pure_python_file_to_pdf"
Cohesion: 0.21
Nodes (14): convert_images_to_single_pdf(), convert_input_to_pdf_for_mode(), _draw_images_to_pdf_with_reportlab(), file_to_pdf(), _image_to_pdf_ready_rgb(), looks_like_text(), _markdown_inline_to_html(), office_like_to_pdf() (+6 more)

### Community 49 - "kit.py"
Cohesion: 0.10
Nodes (34): ConversionError, CustomTkinter widgets avoided due to redraw loop, Design tokens via ui_theme for dependency inversion, Layered architecture dependency rule, lerp_color(), CustomTkinter widget kit for GO Omni Pro. Every screen builds from…, Blend two hex colours. Used by the animation helpers., appearance_mode() (+26 more)

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

### Community 58 - "ButtonBar"
Cohesion: 0.27
Nodes (4): ButtonBar, Widget, Wrapping row of buttons, API-compatible with the old FlowButtonBar. ``add``…, Wrapping action row. Replaces the ttk-era FlowButtonBar.

### Community 59 - "app.py"
Cohesion: 0.11
Nodes (25): argparse, customtkinter, dataclasses, faulthandler, json, os, pathlib, pil (+17 more)

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

### Community 71 - "datetime"
Cohesion: 0.72
Nodes (8): datetime, days_since_install(), ensure_install_date(), parse_datetime(), Any, should_show_first_launch_splash(), should_show_login_popup(), summarize_login_popup_state()

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

### Community 117 - "resolve_flow_layout_width"
Cohesion: 0.50
Nodes (3): Prefer the real allocated width so wrapping still works on narrow screens., resolve_flow_layout_width(), validate_flow_wrap_helper()

### Community 118 - "workflow_support.py"
Cohesion: 0.42
Nodes (7): run_patch15_workflow_smoke(), directory_stats(), DirectoryStats, format_bytes(), prune_directory(), Path, summarize_directory()

### Community 125 - "apply_ttk_theme"
Cohesion: 0.23
Nodes (12): Style, apply_ttk_theme(), _configure_fonts(), font_spec(), Misc, Return the first font in the stack that this Tk install actually has., A (family, size, weight) tuple usable anywhere Tk accepts a font., Scaled spacing token. Unknown names fall back to the medium step. (+4 more)

## Knowledge Gaps
- **185 isolated node(s):** `Planning Chat: resume prompt`, `URGENT (Gokul only)`, `Where each repo stands`, `Requests waiting for the `platform` repo (collected from repo logs; platform chat applies, planning chat syncs)`, `GitHub / sync status (2026-11-14)` (+180 more)
  These have ≤1 connection - possible missing edges or undocumented components. (Counts symbols only; 506 node(s) total have ≤1 connection when file, concept and rationale nodes are included.)
- **55 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `GOOmniProApp` connect `GOOmniProApp` to `release_validation.py`, `AppStateStore`, `ScrollablePage`, `._show_page`, `._log_automation`, `._persist_state`, `open_path`, `automation_core.py`, `mail_core.py`, `._resolve_link_cache_dir`, `._refresh_about_profile`, `build_support.py`, `.__init__`, `smoke_test.py`, `visit`, `._current_palette`, `._apply_state_snapshot`, `._start_ocr_job`, `app.py`?**
  _High betweenness centrality (0.215) - this node is a cross-community bridge._
- **Why does `UIKit` connect `UIKit` to `Any`, `.__init__`, `Tooltip`, `ScrollArea`, `.radio`, `Misc`, `kit.py`, `GradientBackdrop`, `ButtonBar`, `.treeview`?**
  _High betweenness centrality (0.054) - this node is a cross-community bridge._
- **Why does `ArchivePage` connect `ArchivePage` to `app.py`, `ConversionError`, `register_drop_target`?**
  _High betweenness centrality (0.051) - this node is a cross-community bridge._
- **Are the 4 inferred relationships involving `ConversionError` (e.g. with `test_archive()` and `test_rename()`) actually correct?**
  _`ConversionError` has 4 INFERRED edges - model-reasoned connections that need verification._
- **What connects `Planning Chat: resume prompt`, `URGENT (Gokul only)`, `Where each repo stands` to the rest of the system?**
  _185 weakly-connected nodes found - possible documentation gaps or missing edges._
- **Should `PageOrganizerPanel` be split into smaller, more focused modules?**
  _Cohesion score 0.055412371134020616 - nodes in this community are weakly interconnected._
- **Should `ArchivePage` be split into smaller, more focused modules?**
  _Cohesion score 0.05403348554033485 - nodes in this community are weakly interconnected._