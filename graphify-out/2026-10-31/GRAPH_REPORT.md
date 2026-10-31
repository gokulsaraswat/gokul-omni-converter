# Graph Report - gokul-omni-converter  (2026-10-31)

## Corpus Check
- 99 files · ~97,207 words
- Verdict: corpus is large enough that graph structure adds value.
- Unclassified: 16 file(s) not represented in the graph (top: (none) 3, .zip 3, .bat 2)

## Summary
- 1679 nodes · 4394 edges · 120 communities (74 shown, 46 thin omitted)
- Extraction: 98% EXTRACTED · 2% INFERRED · 0% AMBIGUOUS · INFERRED: 101 edges (avg confidence: 0.91)
- Token cost: 0 input · 0 output

## Graph Freshness
- Built from commit: `3387436c`
- Run `git rev-parse HEAD` and compare to check if the graph is stale.
- Run `graphify update .` after code changes (no API cost).

## Community Hubs (Navigation)
- PageOrganizerPanel
- archive_core.py
- GOOmniProApp
- release_validation.py
- rename_organizer_core.py
- pure_python_renderers.py
- ArchivePage
- ConversionError
- AppStateStore
- ocr_core.py
- ScrollablePage
- dataclasses
- converter_core.py
- .__init__
- Misc
- ._log_automation
- Path
- smoke_test.py
- UIKit
- .persist
- open_path
- preview_support.py
- RenameOrganizePage
- app.py
- render
- register_drop_target
- PreviewCenterWindow
- link_ingest.py
- mail_core.py
- unique_path
- Animator
- ._refresh_about_profile
- build_support.py
- Any
- asset_support.py
- FirstLaunchSplashWindow
- ._register
- load_about_profile
- ThemePalette
- main
- Tooltip
- ScrollArea
- .radio
- _StructuredHTMLParser
- packages
- smtp
- visit
- kit.py
- convert.py
- settings.py
- ._update_login_popup_state
- GradientBackdrop
- GO Omni Pro Activity Report
- datetime
- bind_responsive_wrap
- ButtonBar
- pathlib
- _BasicHTMLTextExtractor
- diagnostics.json
- CommandPalette
- pdf_tools.py
- PDF Conversion
- paths
- Tooltip
- dependencies
- PDF Split Every N Output Part 1
- automation.py
- ocr.py
- gsclaude/sync-architecture.sh
- system
- 🤖 System Architecture & Co-Worker Protocol
- GO Omni Pro app icon 256px — purple/indigo rounded-square badge with circular convert arrow and document sheet icon
- simple_markdown_to_html_document
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
- format_size
- Release Checklist — GO Omni Pro 2.3.0
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
- `frame()` --indirect_call--> `key()`  [INFERRED]
  ui/kit.py → rename_organizer_core.py
- `AboutProfileEditorWindow` --uses--> `ThemePalette`  [INFERRED]
  app.py → ui_theme.py
- `GOOmniProApp` --uses--> `AppStateStore`  [INFERRED]
  app.py → app_state.py
- `GOOmniProApp` --uses--> `WatchFolderConfig`  [INFERRED]
  app.py → automation_core.py
- `GOOmniProApp` --uses--> `BatchConfig`  [INFERRED]
  app.py → converter_core.py

## Import Cycles
- None detected.

## Communities (120 total, 46 thin omitted)

### Community 0 - "PageOrganizerPanel"
Cohesion: 0.05
Nodes (41): Document, fitz, OrganizedPage, build_default_sequence(), duplicate_positions(), export_pages_as_images(), extract_selected_pdf(), move_positions_down() (+33 more)

### Community 1 - "archive_core.py"
Cohesion: 0.11
Nodes (27): ArchiveEntry, ArchiveSummary, batch_compress(), _collect_folder_files(), _compression(), create_archive(), extract_archive(), _log() (+19 more)

### Community 2 - "GOOmniProApp"
Cohesion: 0.05
Nodes (5): GOOmniProApp, CTkFrame, StringVar, Widget, Create (and register) the container for one page. ``padding`` and ``style`` are…

### Community 3 - "release_validation.py"
Cohesion: 0.11
Nodes (31): log(), progress(), main(), build_image_folder_pdfs(), discover_image_files(), _emit_log(), _emit_progress(), group_images_by_folder() (+23 more)

### Community 4 - "rename_organizer_core.py"
Cohesion: 0.09
Nodes (36): apply_case(), apply_plan(), _bucket_for(), describe_manifest(), _file_dates(), latest_manifest(), _log(), manifest_dir() (+28 more)

### Community 5 - "pure_python_renderers.py"
Cohesion: 0.08
Nodes (39): _Cell, csv, docx, docx_document, docx_oxml_table, docx_oxml_text_paragraph, docx_shared, docx_table (+31 more)

### Community 6 - "ArchivePage"
Cohesion: 0.16
Nodes (3): ArchivePage, build(), Thread-safe log write.

### Community 7 - "ConversionError"
Cohesion: 0.13
Nodes (37): _choose_blank_slide_layout(), _coerce_rect_value(), compress_pdf(), ConversionError, edit_pdf_metadata(), ensure_directory(), extract_pdf_pages(), _fit_size_within_box() (+29 more)

### Community 8 - "AppStateStore"
Cohesion: 0.08
Nodes (17): AppStateStore, _default_state_dir(), migrate_legacy_state_dir(), Any, Path, Copy a pre-2.4.0 state directory to the new location, once. Copies rather than…, Per-platform state location, following each OS's own convention., normalize_preset_record() (+9 more)

### Community 9 - "ocr_core.py"
Cohesion: 0.14
Nodes (35): contextlib, io, _candidate_tessdata_dirs(), detect_tesseract_status(), _emit_log(), _emit_progress(), extract_text_with_ocr(), _find_bundled_tesseract() (+27 more)

### Community 10 - "ScrollablePage"
Cohesion: 0.11
Nodes (9): AnimatedGifLabel, FlowButtonBar, Misc, Path, Widget, A simple responsive page shell with a vertical scrollbar. The inner frame…, A small wrapping button bar for responsive toolbars and action rows., Lightweight animated GIF label with a safe text fallback. (+1 more)

### Community 11 - "dataclasses"
Cohesion: 0.29
Nodes (8): dataclasses, directory_stats(), DirectoryStats, format_bytes(), prune_directory(), Path, summarize_directory(), QuickAction

### Community 12 - "converter_core.py"
Cohesion: 0.12
Nodes (32): _anchor_point(), _anchor_rect(), apply_image_overlay(), apply_image_watermark(), _apply_overlay_to_pages(), apply_text_overlay(), apply_text_watermark(), apply_visible_signature() (+24 more)

### Community 13 - ".__init__"
Cohesion: 0.08
Nodes (3): compact_display_path(), Set the window/taskbar icon from the app mark, per-platform., default_merged_name()

### Community 14 - "Misc"
Cohesion: 0.17
Nodes (11): CTkLabel, CTkFrame, Misc, A bordered surface panel. When ``title`` is given the card exposes ``.body`` --…, Small ⓘ carrying help text on hover, instead of a wall of static prose., A label with an optional ⓘ beside it., A card whose border warms toward the accent while the pointer is over it., Label on the left, control on the right, optional hint underneath. Returns the… (+3 more)

### Community 16 - "Path"
Cohesion: 0.16
Nodes (30): build_conversion_route_preview(), convert_images_to_single_pdf(), convert_input_to_pdf_for_mode(), copy_as_pdf(), describe_input_route_for_mode(), describe_pdf_output_route(), document_to_plain_text(), _draw_images_to_pdf_with_reportlab() (+22 more)

### Community 17 - "smoke_test.py"
Cohesion: 0.14
Nodes (28): add_fingerprints(), BatchConfig, dependency_status(), PdfToolConfig, functools, http_server, cache_root_from_setting(), extract_urls() (+20 more)

### Community 18 - "UIKit"
Cohesion: 0.12
Nodes (14): CTkButton, CTkCheckBox, CTkEntry, CTkSegmentedButton, CTkSlider, CTkSwitch, attach_tooltip(), Entry (+6 more)

### Community 22 - "preview_support.py"
Cohesion: 0.20
Nodes (27): ImageDraw, ImageFont, pptx, _doc_preview(), _docx_excerpt(), _font(), _generic_summary(), _html_to_text() (+19 more)

### Community 23 - "RenameOrganizePage"
Cohesion: 0.14
Nodes (4): build(), Path, Only show the inputs the selected rule actually uses., RenameOrganizePage

### Community 24 - "app.py"
Cohesion: 0.12
Nodes (23): open_url(), argparse, bundle_paths_as_zip(), discover_watch_candidates(), export_presets_to_json(), fingerprint_file(), import_presets_from_json(), move_files_to_archive() (+15 more)

### Community 25 - "render"
Cohesion: 0.12
Nodes (22): _arc_box(), _gradient(), _hex_to_rgb(), main(), _masked(), Image, Rasterise the GO Omni Pro mark into the icon files the installers need. The…, Draw the mark at ``size`` pixels, supersampled then downscaled for smooth edges. (+14 more)

### Community 26 - "register_drop_target"
Cohesion: 0.17
Nodes (11): dnd_available(), DropZone, parse_drop_paths(), Misc, Path, A dashed panel that takes dropped files and doubles as a browse button., True once tkdnd has actually been loaded into a live Tk interpreter., Split a tkdnd drop payload into paths. tkdnd hands over a Tcl list: space-… (+3 more)

### Community 27 - "PreviewCenterWindow"
Cohesion: 0.19
Nodes (3): PreviewCenterWindow, Misc, Path

### Community 28 - "link_ingest.py"
Cohesion: 0.11
Nodes (29): CancelFn, hashlib, clear_cache_dir(), download_many_urls(), emit(), wait_while_paused(), _guess_extension_from_content_type(), LinkDownloadResult (+21 more)

### Community 29 - "mail_core.py"
Cohesion: 0.17
Nodes (17): email_message, EmailMessage, build_email_message(), build_eml_draft(), _connect(), create_mailto_url(), open_mailto_draft(), _parse_recipients() (+9 more)

### Community 30 - "unique_path"
Cohesion: 0.20
Nodes (19): CompletedProcess, _docx_apply_runs(), extract_html_title(), extract_pdf_pages_as_text(), extract_text_from_html_string(), find_command(), _html_blocks_to_markdown_lines(), html_to_docx() (+11 more)

### Community 31 - "Animator"
Cohesion: 0.16
Nodes (12): Animator, apply(), apply(), frame(), ease_out(), StringVar, cubic-bezier(0.4, 0, 0.2, 1)-ish: fast start, gentle settle., Frame-based tweening on Tk's event loop. Every animation checks… (+4 more)

### Community 33 - "build_support.py"
Cohesion: 0.24
Nodes (14): collect_installer_assets(), collect_package_versions(), export_diagnostics_report(), export_json(), export_state_snapshot(), export_support_bundle(), export_text_file(), import_state_snapshot() (+6 more)

### Community 34 - "Any"
Cohesion: 0.18
Nodes (7): InfoDot, ProgressBar, Any, Text-only options (state, font, wrap) are forwarded to the inner Text., CTkProgressBar that also speaks the ttk.Progressbar API. The conversion/OCR/PDF…, Ease toward ``value`` (0..1) so batch progress glides rather than jumps., A small ⓘ that reveals its explanation on hover. Used instead of printing a…

### Community 36 - "asset_support.py"
Cohesion: 0.38
Nodes (14): asset_cache_root(), cache_summary(), cached_asset_path(), clear_asset_cache(), download_binary_asset(), download_text_file(), is_cache_fresh(), load_asset_config() (+6 more)

### Community 37 - "FirstLaunchSplashWindow"
Cohesion: 0.20
Nodes (4): FirstLaunchSplashWindow, LoginReminderToast, Misc, Path

### Community 38 - "._register"
Cohesion: 0.12
Nodes (10): Combobox, Listbox, Menu, Text, Treeview, Read-only dropdown. Deliberately ttk rather than CTkOptionMenu: in…, Scrolling multi-line text area. Also avoids CTkTextbox, which embeds the same…, CTk has no list widget, so this is a ``tk.Listbox`` kept in palette. (+2 more)

### Community 39 - "load_about_profile"
Cohesion: 0.22
Nodes (10): load_about_profile(), _normalize_links(), Any, Path, Read the profile, dropping any retired fields left over from older versions., Write only the supported fields back out, so retired keys never reappear., resolve_profile_image(), save_about_profile() (+2 more)

### Community 40 - "ThemePalette"
Cohesion: 0.10
Nodes (11): BuildCenterWindow, MarkdownNotesWindow, Misc, SMTPSettingsWindow, PreviewWindow, LogBox, A ``tk.Text`` plus a ttk scrollbar, proxying the Text API. Used wherever the…, apply_text_widget_theme() (+3 more)

### Community 41 - "main"
Cohesion: 0.27
Nodes (14): main(), build_example_update_manifest(), check_for_updates(), compare_versions(), export_workspace_bundle(), import_workspace_bundle(), load_update_manifest(), Any (+6 more)

### Community 43 - "ScrollArea"
Cohesion: 0.21
Nodes (4): A vertically scrolling page body. Deliberately *not* CTkScrollableFrame: in…, Vertically scrolling container that hosts CustomTkinter children. Mirrors the…, Coalesce Configure bursts into one geometry pass. Resizing the canvas window…, ScrollArea

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

### Community 48 - "visit"
Cohesion: 0.38
Nodes (7): cancel_pending(), visit(), on_enter(), on_leave(), on_press(), on_release(), widget_contains_pointer()

### Community 49 - "kit.py"
Cohesion: 0.07
Nodes (44): ConversionError, CustomTkinter widgets avoided due to redraw loop, Design tokens via ui_theme for dependency inversion, Layered architecture dependency rule, Style, subprocess, lerp_color(), CustomTkinter widget kit for GO Omni Pro. Every screen builds from… (+36 more)

### Community 50 - "convert.py"
Cohesion: 0.36
Nodes (9): build(), _build_center(), _build_image_folder_card(), _build_inputs(), _build_links(), _build_log_card(), _build_mode_card(), _build_options() (+1 more)

### Community 51 - "settings.py"
Cohesion: 0.31
Nodes (8): build(), _build_appearance(), _build_engine(), _build_files(), _build_quick_actions(), Settings page: appearance, engine, local files and folders, and quick actions., An entry that shows a path but cannot be edited., _readonly()

### Community 53 - "GradientBackdrop"
Cohesion: 0.28
Nodes (4): Canvas, GradientBackdrop, Paints the palette's gradient onto a tk.Canvas and keeps it sized. Tk has no…, Force a repaint, e.g. after the palette changed.

### Community 54 - "GO Omni Pro Activity Report"
Cohesion: 0.28
Nodes (9): LibreOffice Engine, Pure Python Engine, Smoke Test Output, DOCX to PDF via LibreOffice - Sample Output, DOCX to PDF via Pure Python - Sample Output, HTML to PDF Pure Python - Sample Output, Markdown to PDF Pure Python - Sample Output, BUILDING.md - Patch 16 Release Build Notes (+1 more)

### Community 55 - "datetime"
Cohesion: 0.72
Nodes (8): datetime, days_since_install(), ensure_install_date(), parse_datetime(), Any, should_show_first_launch_splash(), should_show_login_popup(), summarize_login_popup_state()

### Community 57 - "bind_responsive_wrap"
Cohesion: 0.40
Nodes (3): visit(), bind_responsive_wrap(), Shrink wraplength when containers narrow, while preserving the original wide-…

### Community 58 - "ButtonBar"
Cohesion: 0.28
Nodes (4): ButtonBar, Widget, Wrapping row of buttons, API-compatible with the old FlowButtonBar. ``add``…, Wrapping action row. Replaces the ttk-era FlowButtonBar.

### Community 59 - "pathlib"
Cohesion: 0.13
Nodes (15): customtkinter, pathlib, pil, threading, tkinter, tkinter_scrolledtext, tkinterdnd2, typing (+7 more)

### Community 61 - "diagnostics.json"
Cohesion: 0.25
Nodes (7): created_at, current_context, last_outputs, selected_files, extra, smoke, installer_assets

### Community 63 - "pdf_tools.py"
Cohesion: 0.39
Nodes (6): build(), _build_center(), _build_log_card(), _build_options(), _build_tool_card(), PDF Tools page. The options panel exposes every field any tool might need;…

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

### Community 71 - "ocr.py"
Cohesion: 0.53
Nodes (5): build(), _build_log(), _build_settings(), _build_source(), OCR page. Written for someone who does not know what OCR, DPI, or a page-…

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

### Community 76 - "simple_markdown_to_html_document"
Cohesion: 0.67
Nodes (3): _markdown_inline_to_html(), simple_markdown_to_html_document(), flush_paragraph()

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

### Community 112 - "format_size"
Cohesion: 0.28
Nodes (3): format_size(), Human-readable byte size, e.g. '1.4 MB'., Path

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
- **132 isolated node(s):** `1. Starting a New Folder / Project`, `2. Starting with an Existing Project`, `3. The "After-Command" Chaser`, `4. Start / Debug Prompt Template`, `5. Session Wrap-Up Prompt Template` (+127 more)
  These have ≤1 connection - possible missing edges or undocumented components. (Counts symbols only; 435 node(s) total have ≤1 connection when file, concept and rationale nodes are included.)
- **46 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `GOOmniProApp` connect `GOOmniProApp` to `PageOrganizerPanel`, `release_validation.py`, `AppStateStore`, `ocr_core.py`, `ScrollablePage`, `dataclasses`, `.__init__`, `._log_automation`, `smoke_test.py`, `UIKit`, `._persist_state`, `open_path`, `app.py`, `PreviewCenterWindow`, `link_ingest.py`, `mail_core.py`, `Animator`, `._refresh_about_profile`, `._poll_worker_queue`, `FirstLaunchSplashWindow`, `ThemePalette`, `main`, `visit`, `._update_login_popup_state`, `._start_ocr_job`, `bind_responsive_wrap`, `CommandPalette`?**
  _High betweenness centrality (0.261) - this node is a cross-community bridge._
- **Why does `PageOrganizerPanel` connect `PageOrganizerPanel` to `app.py`, `ThemePalette`, `GOOmniProApp`, `pathlib`?**
  _High betweenness centrality (0.074) - this node is a cross-community bridge._
- **Why does `UIKit` connect `UIKit` to `GOOmniProApp`, `Any`, `._register`, `ThemePalette`, `Tooltip`, `ScrollArea`, `.radio`, `.__init__`, `Misc`, `kit.py`, `GradientBackdrop`, `app.py`, `ButtonBar`?**
  _High betweenness centrality (0.072) - this node is a cross-community bridge._
- **Are the 18 inferred relationships involving `GOOmniProApp` (e.g. with `AppStateStore` and `WatchFolderConfig`) actually correct?**
  _`GOOmniProApp` has 18 INFERRED edges - model-reasoned connections that need verification._
- **Are the 4 inferred relationships involving `ConversionError` (e.g. with `test_archive()` and `test_rename()`) actually correct?**
  _`ConversionError` has 4 INFERRED edges - model-reasoned connections that need verification._
- **Are the 2 inferred relationships involving `PageOrganizerPanel` (e.g. with `GOOmniProApp` and `ThemePalette`) actually correct?**
  _`PageOrganizerPanel` has 2 INFERRED edges - model-reasoned connections that need verification._
- **Are the 12 inferred relationships involving `ThemePalette` (e.g. with `AboutProfileEditorWindow` and `BuildCenterWindow`) actually correct?**
  _`ThemePalette` has 12 INFERRED edges - model-reasoned connections that need verification._