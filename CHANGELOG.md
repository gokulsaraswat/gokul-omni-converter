# Changelog

All notable changes to GO Omni Pro. Format loosely follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [2.4.0] — 2026-08-17

Product polish release. Renamed, restyled, and given two new tool families.

### Added
- **Archive tools** (`archive_core.py` + Archive page): create `.zip` from files or a
  folder, inspect archive contents before extracting, extract to a chosen folder, and
  batch-compress sibling folders into one archive each. Pure stdlib `zipfile`, with
  zip-slip protection on extract.
- **Batch rename & organizer** (`rename_organizer_core.py` + Rename & Organize page):
  pattern rename with `{n}`/`{name}`/`{ext}`/`{date}`/`{parent}` tokens, find & replace,
  regex replace, case transforms, live old→new preview, collision detection, and
  auto-organize a folder into subfolders by file type, extension, or date.
- **Undo for destructive batches**: every rename/organize run writes a JSON manifest to
  the app state directory; the last operation can be reverted from the UI.
- CustomTkinter design system in `ui_theme.py` — palettes, spacing scale, radii, and a
  cross-platform font stack, all derived from the existing `ThemePalette` values.
- macOS is now a first-class platform: Homebrew/MacPorts binary discovery for
  LibreOffice and Tesseract, `defaults read` dark-mode detection, and
  `~/Library/Application Support` state paths.
- `LICENSE` (MIT), `CHANGELOG.md`, and `.gitignore`, none of which previously existed.

### Changed
- **Renamed to "GO Omni Pro"** across every user-visible surface: window title, About
  page, README, PyInstaller spec, Inno Setup script, `.desktop` entry, and shortcuts.
- **Full CustomTkinter migration.** All eleven pages, the app shell, and every dialog
  (About editor, SMTP, Build Center, notes/shortcuts viewer) were rewritten against the
  new design system, as was the page organizer panel. Pages now live in one module each
  under `ui/pages/` instead of a single 7,600-line `app.py`.
  `ttk.Treeview`, `ttk.Combobox`, `ttk.Scrollbar`, and `tk.Listbox` are still used where
  CustomTkinter has no equivalent; all are themed from the same palette. The splash,
  login toast, command palette, and preview window remain ttk and are themed to match.
- New app mark at `assets/logo/go_omni_pro_mark.svg` drives the Windows `.ico`, the
  Linux/macOS `.png`, the window icon, and the About page.
- About profile is now name + email + three links only. Company, job title, handle, and
  bio fields were removed from the data model and the editor UI.
- Dropped `dataclass(slots=True)` throughout, lowering the floor from Python 3.10 to 3.9.
- `build_windows_release.bat` / `build_linux_release.sh` now emit installer artifacts
  (`Setup.exe`, `.deb`) rather than a bare PyInstaller folder.

### Design
- **New visual language**, Apple-Weather-inspired: a deep navy→violet gradient
  backdrop with layered "glass" cards on top. Tk cannot composite, so a real
  `backdrop-filter` blur is impossible — instead the gradient is painted for real and
  the card colours are the *exact* result of compositing white over it at the alphas a
  CSS glass panel would use (8% fill, 16% border). On a fixed backdrop the pixels are
  identical. Pill-shaped buttons, frosted rounded inputs, thin large hero numerals, and
  20–28px card padding complete the look. Every text pairing across all four palettes
  clears WCAG AA (4.5:1); high-contrast modes drop the gradient entirely.
- **Explanatory prose moved onto ⓘ hover hints.** Twenty-one of them now carry the help
  that used to be printed on screen — image filenames, page-range syntax, performance
  modes, remote-asset caveats. The OCR page was rewritten around this: three
  plain-language actions ("Make a searchable PDF from an image", "Just give me the
  text") with the jargon (DPI, PSM, Tesseract) tucked behind ⓘ.
- **Motion.** Pages slide up on arrival, stat tiles roll up to their value, and progress
  bars ease toward their target instead of jumping. All of it is skipped when *Reduced
  motion* is on — animation is never the only signal.
- Removed the development-credit line from the About page (it remains in README and this
  changelog), dropped the redundant "Current status" sidebar heading, and rebuilt About
  around a hero, three at-a-glance stat tiles, and grouped actions.

### Fixed
- **Redraw loops that could hang the window.** Several widgets called
  `update_idletasks()` from inside a redraw, which flushes the `<Configure>` events that
  trigger that same redraw. A theme, font, or scale change could set off a cascade that
  never drained. Fixed in `FlowButtonBar.relayout`, `PageOrganizerPanel._render_page_cards`,
  and `_apply_responsive_layouts`; `CTkScrollableFrame`, `CTkTextbox`, and `CTkOptionMenu`
  were replaced with `ui.kit` equivalents for the same reason.
- `_install_responsive_wraps` and `_install_hover_feedback` re-walked the widget tree on
  every theme change, stacking duplicate `<Configure>` and hover bindings each time.
- `_apply_responsive_layouts` could oscillate forever across a layout breakpoint; it now
  no-ops unless a breakpoint actually changed.
- The Automation and Settings pages both wrote `settings_content_frame`, so the responsive
  breakpoints depended on page build order.
- **Pages did not scroll.** `_create_page_frame` parented page content to the scroll
  area's outer frame instead of the canvas's inner frame, so content sat beside the
  scrollable region rather than inside it. Scrollbars now also hide when content fits.
- **Quick Actions hung the app.** The old command palette called `grab_set()`; on macOS
  that app-modal grab blocks the root window's event loop outright. It is replaced by a
  non-modal palette (`ui/command_palette.py`) with type-to-filter, arrow-key navigation,
  and Escape-to-close. No widget in the app grabs input any more.
- UI-scale changes were applied twice — once by this design system and again by
  CustomTkinter's global `set_widget_scaling`, whose full-app rescale was also very slow.
- `ScrollArea` recomputed geometry and toggled its scrollbar from inside a `<Configure>`
  handler, which generated further Configure events; it is now debounced and idempotent.

### Removed
- All Oracle placeholder data from `about_profile.json` and `installer/about_static.json`.
- Dead modules and docs: `patch10_gui.py`, `patch10_services.py`, `smoke_test_patch10.py`,
  `release_check.py`, `packaging_core.py`, five root-level patch reports, five
  `installer/*.md` notes, the checked-in `smoke_test_output_example/` tree, and two
  unreferenced image assets.

### Migration
`app_state.py` gained `archive` and `rename_organizer` sections and moved its state
directory on macOS. Existing `~/.gokul_omni_convert_lite/app_state.json` files are read
and migrated in place on first launch; presets, history, and settings are preserved.

## [2.3.0] and earlier

Developed as numbered patches (2.2.x → 2.3.0) building up the conversion engine, PDF
tools, OCR, page organizer, automation presets, link ingest, preview centre, workflow
history, and release tooling. Per-patch reports were folded into this file and deleted.
