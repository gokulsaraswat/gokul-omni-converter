# Graph Report - gokul-omni-converter  (2026-10-01)

## Corpus Check
- 159 files · ~94,543 words
- Verdict: corpus is large enough that graph structure adds value.
- Unclassified: 15 file(s) not represented in the graph (top: .zip 3, (none) 2, .bat 2)

## Summary
- 1613 nodes · 4412 edges · 108 communities (68 shown, 40 thin omitted)
- Extraction: 98% EXTRACTED · 2% INFERRED · 0% AMBIGUOUS · INFERRED: 103 edges (avg confidence: 0.91)
- Token cost: 0 input · 0 output

## Community Hubs (Navigation)
- PDF & Document Processing
- Archive Tools
- App Shell & UI Wiring
- Background Workers
- Rename & Organize Core
- Office Format Converters
- Automation & Presets
- Conversion Engine Core
- App State & Persistence
- OCR Pipeline
- Responsive UI Layout
- SMTP & Mail Settings
- Converter Utilities
- File Management Handlers
- CustomTkinter Widget Kit
- File Add & Browse Handlers
- PDF Route & Image Conversion
- Automation Config & Batch
- CTk Primitive Widgets
- Profile & State Persistence
- Module Group 21
- Module Group 22
- Module Group 23
- Module Group 24
- Module Group 25
- Module Group 26
- Module Group 27
- Module Group 28
- Module Group 29
- Module Group 30
- Module Group 31
- Module Group 32
- Module Group 33
- Module Group 34
- Module Group 36
- Module Group 37
- Module Group 38
- Module Group 39
- Module Group 40
- Module Group 41
- Module Group 42
- Module Group 43
- Module Group 44
- Module Group 45
- Module Group 46
- Module Group 47
- Module Group 48
- Module Group 49
- Module Group 50
- Module Group 51
- Module Group 52
- Module Group 53
- Module Group 54
- Module Group 55
- Module Group 56
- Module Group 57
- Module Group 58
- Module Group 60
- Module Group 61
- Module Group 62
- Module Group 63
- Module Group 64
- Module Group 65
- Module Group 66
- Module Group 67
- Module Group 68
- Module Group 69
- Module Group 70
- Module Group 71
- Module Group 72
- Module Group 73
- Module Group 74
- Module Group 75
- Module Group 76
- Module Group 77
- Module Group 78
- Module Group 79
- Module Group 80
- Module Group 81
- Module Group 82
- Module Group 83
- Module Group 84
- Module Group 85
- Module Group 86
- Module Group 87
- Module Group 88
- Module Group 89
- Module Group 90
- Module Group 91
- Module Group 92
- Module Group 93
- Module Group 94
- Module Group 95
- Module Group 96
- Module Group 97
- Module Group 98
- Module Group 99
- Module Group 100
- Module Group 101
- Module Group 102
- Module Group 103
- Module Group 104
- Module Group 105
- Module Group 106
- Module Group 107

## God Nodes (most connected - your core abstractions)
1. `GOOmniProApp` - 277 edges
2. `ConversionError` - 62 edges
3. `PageOrganizerPanel` - 61 edges
4. `run()` - 54 edges
5. `ThemePalette` - 52 edges
6. `AppStateStore` - 45 edges
7. `ensure_directory()` - 40 edges
8. `ArchivePage` - 36 edges
9. `process_pdf_tool()` - 33 edges
10. `process_batch()` - 31 edges

## Surprising Connections (you probably didn't know these)
- `MarkdownNotesWindow` --uses--> `ThemePalette`  [INFERRED]
  app.py → ui_theme.py
- `BuildCenterWindow` --uses--> `ThemePalette`  [INFERRED]
  app.py → ui_theme.py
- `AboutProfileEditorWindow` --uses--> `ThemePalette`  [INFERRED]
  app.py → ui_theme.py
- `GOOmniProApp` --uses--> `AppStateStore`  [INFERRED]
  app.py → app_state.py
- `GOOmniProApp` --uses--> `WatchFolderConfig`  [INFERRED]
  app.py → automation_core.py

## Import Cycles
- None detected.

## Communities (108 total, 40 thin omitted)

### Community 0 - "PDF & Document Processing"
Cohesion: 0.05
Nodes (42): Document, fitz, build_default_sequence(), duplicate_positions(), export_pages_as_images(), extract_selected_pdf(), move_positions_down(), move_positions_to_index() (+34 more)

### Community 1 - "Archive Tools"
Cohesion: 0.05
Nodes (35): ArchiveEntry, ArchiveSummary, batch_compress(), _collect_folder_files(), _compression(), create_archive(), extract_archive(), format_size() (+27 more)

### Community 2 - "App Shell & UI Wiring"
Cohesion: 0.05
Nodes (5): GOOmniProApp, CTkFrame, StringVar, Widget, Create (and register) the container for one page. ``padding`` and ``style`` are…

### Community 3 - "Background Workers"
Cohesion: 0.11
Nodes (31): log(), progress(), main(), build_image_folder_pdfs(), discover_image_files(), _emit_log(), _emit_progress(), group_images_by_folder() (+23 more)

### Community 4 - "Rename & Organize Core"
Cohesion: 0.09
Nodes (35): apply_case(), apply_plan(), _bucket_for(), describe_manifest(), _file_dates(), latest_manifest(), _log(), manifest_dir() (+27 more)

### Community 5 - "Office Format Converters"
Cohesion: 0.08
Nodes (39): _Cell, csv, docx, docx_document, docx_oxml_table, docx_oxml_text_paragraph, docx_shared, docx_table (+31 more)

### Community 6 - "Automation & Presets"
Cohesion: 0.10
Nodes (8): bundle_paths_as_zip(), discover_watch_candidates(), fingerprint_file(), move_files_to_archive(), normalize_watch_config(), Path, WatchFolderConfig, write_run_report()

### Community 7 - "Conversion Engine Core"
Cohesion: 0.13
Nodes (37): _choose_blank_slide_layout(), _coerce_rect_value(), compress_pdf(), ConversionError, edit_pdf_metadata(), ensure_directory(), extract_pdf_pages(), _fit_size_within_box() (+29 more)

### Community 8 - "App State & Persistence"
Cohesion: 0.13
Nodes (4): AppStateStore, Any, add_fingerprints(), normalize_preset_record()

### Community 9 - "OCR Pipeline"
Cohesion: 0.14
Nodes (35): contextlib, io, _candidate_tessdata_dirs(), detect_tesseract_status(), _emit_log(), _emit_progress(), extract_text_with_ocr(), _find_bundled_tesseract() (+27 more)

### Community 10 - "Responsive UI Layout"
Cohesion: 0.09
Nodes (12): visit(), AnimatedGifLabel, bind_responsive_wrap(), FlowButtonBar, Misc, Path, Widget, A simple responsive page shell with a vertical scrollbar. The inner frame… (+4 more)

### Community 11 - "SMTP & Mail Settings"
Cohesion: 0.11
Nodes (28): SMTPSettingsWindow, lerp_color(), Blend two hex colours. Used by the animation helpers., Repaint every widget this kit created for the new palette., appearance_mode(), apply_ctk_theme(), apply_menu_theme(), apply_treeview_tag_colors() (+20 more)

### Community 12 - "Converter Utilities"
Cohesion: 0.12
Nodes (32): _anchor_point(), _anchor_rect(), apply_image_overlay(), apply_image_watermark(), _apply_overlay_to_pages(), apply_text_overlay(), apply_text_watermark(), apply_visible_signature() (+24 more)

### Community 13 - "File Management Handlers"
Cohesion: 0.11
Nodes (4): compact_display_path(), Set the window/taskbar icon from the app mark, per-platform., default_merged_name(), outputs_pdf()

### Community 14 - "CustomTkinter Widget Kit"
Cohesion: 0.12
Nodes (14): CTkLabel, CTkFrame, Widget factory bound to one root window and one palette at a time., Push the palette into CustomTkinter's and ttk's global state., Shared backdrop bitmap for (bucketed) size + palette. Returns (photo, key).…, A bordered surface panel. When ``title`` is given the card exposes ``.body`` --…, Small ⓘ carrying help text on hover, instead of a wall of static prose., A label with an optional ⓘ beside it. (+6 more)

### Community 16 - "PDF Route & Image Conversion"
Cohesion: 0.16
Nodes (30): build_conversion_route_preview(), convert_images_to_single_pdf(), convert_input_to_pdf_for_mode(), copy_as_pdf(), describe_input_route_for_mode(), describe_pdf_output_route(), document_to_plain_text(), _draw_images_to_pdf_with_reportlab() (+22 more)

### Community 17 - "Automation Config & Batch"
Cohesion: 0.15
Nodes (27): export_presets_to_json(), import_presets_from_json(), BatchConfig, dependency_status(), PdfToolConfig, functools, http_server, cache_root_from_setting() (+19 more)

### Community 18 - "CTk Primitive Widgets"
Cohesion: 0.11
Nodes (17): CTkButton, CTkCheckBox, CTkEntry, CTkRadioButton, CTkSegmentedButton, CTkSlider, CTkSwitch, attach_tooltip() (+9 more)

### Community 19 - "Profile & State Persistence"
Cohesion: 0.16
Nodes (19): _default_state_dir(), migrate_legacy_state_dir(), Path, Copy a pre-2.4.0 state directory to the new location, once. Copies rather than…, Per-platform state location, following each OS's own convention., PresetRecord, dataclasses, json (+11 more)

### Community 22 - "Module Group 22"
Cohesion: 0.20
Nodes (27): ImageDraw, ImageFont, pptx, _doc_preview(), _docx_excerpt(), _font(), _generic_summary(), _html_to_text() (+19 more)

### Community 23 - "Module Group 23"
Cohesion: 0.14
Nodes (3): Path, Only show the inputs the selected rule actually uses., RenameOrganizePage

### Community 24 - "Module Group 24"
Cohesion: 0.10
Nodes (18): open_url(), argparse, customtkinter, faulthandler, os, queue, subprocess, sys (+10 more)

### Community 25 - "Module Group 25"
Cohesion: 0.12
Nodes (22): _arc_box(), _gradient(), _hex_to_rgb(), main(), _masked(), Image, Rasterise the GO Omni Pro mark into the icon files the installers need. The…, Draw the mark at ``size`` pixels, supersampled then downscaled for smooth edges. (+14 more)

### Community 26 - "Module Group 26"
Cohesion: 0.12
Nodes (17): tkinterdnd2, dnd_available(), DropZone, init_root(), parse_drop_paths(), Misc, Path, Drag-and-drop file input, with a working fallback when tkdnd is absent.… (+9 more)

### Community 27 - "Module Group 27"
Cohesion: 0.19
Nodes (3): PreviewCenterWindow, Misc, Path

### Community 28 - "Module Group 28"
Cohesion: 0.16
Nodes (19): CancelFn, hashlib, clear_cache_dir(), download_many_urls(), emit(), wait_while_paused(), extract_urls(), _guess_extension_from_content_type() (+11 more)

### Community 29 - "Module Group 29"
Cohesion: 0.18
Nodes (16): email_message, EmailMessage, build_email_message(), build_eml_draft(), _connect(), create_mailto_url(), open_mailto_draft(), _parse_recipients() (+8 more)

### Community 30 - "Module Group 30"
Cohesion: 0.20
Nodes (19): CompletedProcess, _docx_apply_runs(), extract_html_title(), extract_pdf_pages_as_text(), extract_text_from_html_string(), find_command(), _html_blocks_to_markdown_lines(), html_to_docx() (+11 more)

### Community 31 - "Module Group 31"
Cohesion: 0.15
Nodes (13): key(), Animator, apply(), apply(), frame(), ease_out(), StringVar, cubic-bezier(0.4, 0, 0.2, 1)-ish: fast start, gentle settle. (+5 more)

### Community 33 - "Module Group 33"
Cohesion: 0.24
Nodes (15): main(), collect_installer_assets(), collect_package_versions(), export_diagnostics_report(), export_json(), export_state_snapshot(), export_support_bundle(), export_text_file() (+7 more)

### Community 34 - "Module Group 34"
Cohesion: 0.17
Nodes (8): LogBox, ProgressBar, Any, Text-only options (state, font, wrap) are forwarded to the inner Text., CTkProgressBar that also speaks the ttk.Progressbar API. The conversion/OCR/PDF…, Ease toward ``value`` (0..1) so batch progress glides rather than jumps., Scrolling multi-line text area. Also avoids CTkTextbox, which embeds the same…, A ``tk.Text`` plus a ttk scrollbar, proxying the Text API. Used wherever the…

### Community 36 - "Module Group 36"
Cohesion: 0.38
Nodes (14): asset_cache_root(), cache_summary(), cached_asset_path(), clear_asset_cache(), download_binary_asset(), download_text_file(), is_cache_fresh(), load_asset_config() (+6 more)

### Community 37 - "Module Group 37"
Cohesion: 0.19
Nodes (4): FirstLaunchSplashWindow, LoginReminderToast, Misc, Path

### Community 38 - "Module Group 38"
Cohesion: 0.13
Nodes (9): Combobox, Listbox, Menu, Text, Treeview, Read-only dropdown. Deliberately ttk rather than CTkOptionMenu: in…, CTk has no list widget, so this is a ``tk.Listbox`` kept in palette., CTk has no tree widget; this ``ttk.Treeview`` is styled by apply_ttk_theme. (+1 more)

### Community 39 - "Module Group 39"
Cohesion: 0.21
Nodes (10): load_about_profile(), _normalize_links(), Any, Path, Read the profile, dropping any retired fields left over from older versions., Write only the supported fields back out, so retired keys never reappear., resolve_profile_image(), save_about_profile() (+2 more)

### Community 40 - "Module Group 40"
Cohesion: 0.19
Nodes (5): MarkdownNotesWindow, Misc, apply_text_widget_theme(), Listbox, Text

### Community 41 - "Module Group 41"
Cohesion: 0.29
Nodes (12): build_example_update_manifest(), check_for_updates(), compare_versions(), export_workspace_bundle(), import_workspace_bundle(), load_update_manifest(), Any, Path (+4 more)

### Community 42 - "Module Group 42"
Cohesion: 0.21
Nodes (5): InfoDot, Misc, A small ⓘ that reveals its explanation on hover. Used instead of printing a…, Hover hint. Deliberately plain ``tk`` so it can float above CTk widgets., Tooltip

### Community 43 - "Module Group 43"
Cohesion: 0.21
Nodes (4): A vertically scrolling page body. Deliberately *not* CTkScrollableFrame: in…, Vertically scrolling container that hosts CustomTkinter children. Mirrors the…, Coalesce Configure bursts into one geometry pass. Resizing the canvas window…, ScrollArea

### Community 44 - "Module Group 44"
Cohesion: 0.22
Nodes (8): ConversionError, CustomTkinter widgets avoided due to redraw loop, Design tokens via ui_theme for dependency inversion, Layered architecture dependency rule, CustomTkinter widget kit for GO Omni Pro. Every screen builds from…, ui/pages/*.py, Scaled spacing token. Unknown names fall back to the medium step., space()

### Community 45 - "Module Group 45"
Cohesion: 0.36
Nodes (4): _html_to_blocks(), HtmlBlock, HTMLParser, _StructuredHTMLParser

### Community 46 - "Module Group 46"
Cohesion: 0.18
Nodes (11): packages, openpyxl, pdfplumber, Pillow, PyMuPDF, pypdf, pytesseract, python-docx (+3 more)

### Community 47 - "Module Group 47"
Cohesion: 0.18
Nodes (11): smtp, default_to, host, password, port, save_password, sender, timeout_seconds (+3 more)

### Community 48 - "Module Group 48"
Cohesion: 0.38
Nodes (7): cancel_pending(), visit(), on_enter(), on_leave(), on_press(), on_release(), widget_contains_pointer()

### Community 49 - "Module Group 49"
Cohesion: 0.29
Nodes (10): Style, apply_ttk_theme(), _configure_fonts(), font_spec(), Misc, Return the first font in the stack that this Tk install actually has., A (family, size, weight) tuple usable anywhere Tk accepts a font., resolve_font_family() (+2 more)

### Community 50 - "Module Group 50"
Cohesion: 0.36
Nodes (9): build(), _build_center(), _build_image_folder_card(), _build_inputs(), _build_links(), _build_log_card(), _build_mode_card(), _build_options() (+1 more)

### Community 51 - "Module Group 51"
Cohesion: 0.31
Nodes (8): build(), _build_appearance(), _build_engine(), _build_files(), _build_quick_actions(), Settings page: appearance, engine, local files and folders, and quick actions., An entry that shows a path but cannot be edited., _readonly()

### Community 53 - "Module Group 53"
Cohesion: 0.28
Nodes (4): Canvas, GradientBackdrop, Paints the palette's gradient onto a tk.Canvas and keeps it sized. Tk has no…, Force a repaint, e.g. after the palette changed.

### Community 54 - "Module Group 54"
Cohesion: 0.28
Nodes (9): LibreOffice Engine, Pure Python Engine, Smoke Test Output, DOCX to PDF via LibreOffice - Sample Output, DOCX to PDF via Pure Python - Sample Output, HTML to PDF Pure Python - Sample Output, Markdown to PDF Pure Python - Sample Output, BUILDING.md - Patch 16 Release Build Notes (+1 more)

### Community 55 - "Module Group 55"
Cohesion: 0.72
Nodes (8): datetime, days_since_install(), ensure_install_date(), parse_datetime(), Any, should_show_first_launch_splash(), should_show_login_popup(), summarize_login_popup_state()

### Community 56 - "Module Group 56"
Cohesion: 0.42
Nodes (7): run_patch15_workflow_smoke(), directory_stats(), DirectoryStats, format_bytes(), prune_directory(), Path, summarize_directory()

### Community 57 - "Module Group 57"
Cohesion: 0.44
Nodes (8): tempfile, check(), main(), make_tree(), Path, Headless tests for archive_core and rename_organizer_core. These two modules…, test_archive(), test_rename()

### Community 58 - "Module Group 58"
Cohesion: 0.31
Nodes (4): ButtonBar, Widget, Wrapping row of buttons, API-compatible with the old FlowButtonBar. ``add``…, Wrapping action row. Replaces the ttk-era FlowButtonBar.

### Community 61 - "Module Group 61"
Cohesion: 0.25
Nodes (7): created_at, current_context, last_outputs, selected_files, extra, smoke, installer_assets

### Community 63 - "Module Group 63"
Cohesion: 0.39
Nodes (6): build(), _build_center(), _build_log_card(), _build_options(), _build_tool_card(), PDF Tools page. The options panel exposes every field any tool might need;…

### Community 64 - "Module Group 64"
Cohesion: 0.57
Nodes (5): validate_ui_text_helpers(), format_engine_label(), format_flag(), humanize_identifier(), _split_identifier()

### Community 65 - "Module Group 65"
Cohesion: 0.29
Nodes (7): PDF Conversion, Image Folder All - Combined PDF, Image Folder Per Folder - Nested PDF, Patch 36.5 Stability - PNGs to PDF, PDF to HTML - Multi-page Sample Output, PDF Tool Compress - Compressed Output, PDF Tool Edit Metadata - Metadata Sample

### Community 66 - "Module Group 66"
Cohesion: 0.29
Nodes (7): paths, about_profile_path, asset_config_path, installer_dir, notes_path, output_dir, state_path

### Community 68 - "Module Group 68"
Cohesion: 0.33
Nodes (6): dependencies, LibreOffice, Pandoc, pdftotext, python-pptx, xlrd

### Community 69 - "Module Group 69"
Cohesion: 0.33
Nodes (6): PDF Split Every N Output Part 1, PDF Split Every N Output Part 2, PDF Split Every N Output Part 3, PDF Split Ranges Output Part 1, PDF Split Ranges Output Part 2, PDF Split Ranges Output Part 3

### Community 70 - "Module Group 70"
Cohesion: 0.53
Nodes (5): build(), _build_lower(), _build_presets(), _build_watch(), Automation page: saved presets, the watch folder, and output bundling.

### Community 71 - "Module Group 71"
Cohesion: 0.53
Nodes (5): build(), _build_log(), _build_settings(), _build_source(), OCR page. Written for someone who does not know what OCR, DPI, or a page-…

### Community 72 - "Module Group 72"
Cohesion: 0.70
Nodes (3): edge(), node(), sync-architecture.sh script

### Community 73 - "Module Group 73"
Cohesion: 0.40
Nodes (5): system, cwd, platform, python, python_executable

### Community 75 - "Module Group 75"
Cohesion: 0.83
Nodes (4): GO Omni Pro app icon 256px — purple/indigo rounded-square badge with circular convert arrow and document sheet icon, GO Omni Pro app icon 48px — small version of the purple document-convert badge, GO Omni Pro app icon 512px — large high-res version of the purple document-convert badge with circular arrow, GO Omni Pro SVG vector mark — indigo-to-violet gradient badge with circular convert arrow and document sheet, scalable source for all PNG sizes

### Community 76 - "Module Group 76"
Cohesion: 0.67
Nodes (3): _markdown_inline_to_html(), simple_markdown_to_html_document(), flush_paragraph()

### Community 77 - "Module Group 77"
Cohesion: 0.67
Nodes (3): Profile placeholder image — dark navy card with rounded corners, blue glow border, generic silhouette avatar, 'Profile Placeholder' label, Patch 16 imported bundle profile image — solid blue square, likely a test profile asset, Patch 16 profile image — solid blue square, test profile asset for patch 16 release smoke test

### Community 78 - "Module Group 78"
Cohesion: 0.67
Nodes (3): app, name, version

## Knowledge Gaps
- **91 isolated node(s):** `build_linux_release.sh script`, `created_at`, `name`, `version`, `platform` (+86 more)
  These have ≤1 connection - possible missing edges or undocumented components. (Counts symbols only; 381 node(s) total have ≤1 connection when file, concept and rationale nodes are included.)
- **40 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `GOOmniProApp` connect `App Shell & UI Wiring` to `PDF & Document Processing`, `Background Workers`, `Automation & Presets`, `App State & Persistence`, `OCR Pipeline`, `Responsive UI Layout`, `SMTP & Mail Settings`, `File Management Handlers`, `CustomTkinter Widget Kit`, `File Add & Browse Handlers`, `Automation Config & Batch`, `Module Group 20`, `Module Group 21`, `Module Group 24`, `Module Group 27`, `Module Group 29`, `Module Group 31`, `Module Group 32`, `Module Group 33`, `Module Group 35`, `Module Group 37`, `Module Group 48`, `Module Group 52`, `Module Group 59`, `Module Group 62`, `Module Group 64`?**
  _High betweenness centrality (0.254) - this node is a cross-community bridge._
- **Why does `UIKit` connect `CustomTkinter Widget Kit` to `App Shell & UI Wiring`, `Module Group 34`, `Module Group 38`, `Module Group 42`, `SMTP & Mail Settings`, `Module Group 44`, `File Management Handlers`, `Module Group 43`, `CTk Primitive Widgets`, `Module Group 53`, `Module Group 24`, `Module Group 58`?**
  _High betweenness centrality (0.081) - this node is a cross-community bridge._
- **Why does `PageOrganizerPanel` connect `PDF & Document Processing` to `Module Group 24`, `App Shell & UI Wiring`, `SMTP & Mail Settings`?**
  _High betweenness centrality (0.061) - this node is a cross-community bridge._
- **Are the 18 inferred relationships involving `GOOmniProApp` (e.g. with `AppStateStore` and `WatchFolderConfig`) actually correct?**
  _`GOOmniProApp` has 18 INFERRED edges - model-reasoned connections that need verification._
- **Are the 4 inferred relationships involving `ConversionError` (e.g. with `test_archive()` and `test_rename()`) actually correct?**
  _`ConversionError` has 4 INFERRED edges - model-reasoned connections that need verification._
- **Are the 4 inferred relationships involving `PageOrganizerPanel` (e.g. with `GOOmniProApp` and `OrganizedPage`) actually correct?**
  _`PageOrganizerPanel` has 4 INFERRED edges - model-reasoned connections that need verification._
- **What connects `build_linux_release.sh script`, `created_at`, `name` to the rest of the system?**
  _91 weakly-connected nodes found - possible documentation gaps or missing edges._