# Installer and packaging prep for Gokul Omni Convert Lite

Patch 12 keeps the project installer-ready while separating editable in-app profile data from static packaging-time About data.

## Patch 23 additions

- new `remote_assets.json` lets the app keep local fallback assets while optionally refreshing branding from remote GitHub-style URLs later
- Settings now supports:
  - header GIF path + URL
  - splash GIF path + URL
  - About image URL
  - About profile JSON URL pull
  - remote asset cache directory / clear actions
  - timeout + refresh-hour controls
- exported workspace bundles and support bundles now carry `remote_assets.json`
- the PyInstaller spec now includes:
  - `remote_assets.json`
  - `keyboard_shortcuts.md`
- packaging direction stays **offline-safe first**
  - installer builds must still ship local fallback assets
  - remote assets are optional overlays, not mandatory runtime dependencies

## What is included


## Patch 17 additions

- new support/export flows can now be generated without opening the GUI:
  - `python app.py --export-activity-report out.html`
  - `python app.py --export-support-bundle out.zip`
- Build Center now bundles diagnostics, state snapshots, logs, and HTML activity reports for support handoff
- compact UI and preferred start page are now persisted in app state for installer-ready migrations

## Patch 16 additions

- `update_manifest.example.json` - starter JSON manifest for the built-in update checker
- workspace bundles can now be exported/imported from the Build Center or by using the headless CLI:
  - `python app.py --export-workspace release_workspace.zip`
  - `python app.py --import-workspace release_workspace.zip --workspace-target ./restore_here`
- `python app.py --check-updates` uses the saved manifest path or the bundled example manifest


- `gokul_omni_convert_lite.spec` - PyInstaller spec for the desktop app payload
- `build_windows.bat` - helper script for a Windows PyInstaller build
- `build_linux.sh` - helper script for a Linux PyInstaller build
- `windows/GokulOmniConvertLite.iss` - Inno Setup starter installer script
- `windows/version_info.txt` - Windows version metadata for the packaged executable
- `linux/gokul-omni-convert-lite.desktop` - starter desktop entry file
- `macos/build_app_bundle.sh` - starter macOS build helper
- `about_static.json` - static installer-safe About snapshot generated from the editable profile

## Packaging direction

The desktop app should keep reading editable local assets such as:

- `about_profile.json`
- `assets/gokul_profile_placeholder.png`
- `assets/gokul_splash.gif`
- `footer_notes.md`

The installer, however, should prefer static/non-editable About information taken from:

- `installer/about_static.json`

That keeps the bundled installer metadata stable even if the in-app editable profile changes later on disk.

## What the build bundles

The PyInstaller spec includes the desktop app and its local data files:

- `footer_notes.md`
- `about_profile.json`
- `assets/`
- `installer/about_static.json`

## Typical Windows build flow

1. Install Python dependencies from `requirements.txt`.
2. Install PyInstaller.
3. Run:

```bat
installer\build_windows.bat
```

4. After the `dist/GokulOmniConvertLite/` folder exists, open the Inno Setup script:

```text
installer\windows\GokulOmniConvertLite.iss
```

Use that script as the starter installer project if you want a branded setup executable.

## Typical Linux build flow

1. Install Python dependencies from `requirements.txt`.
2. Install PyInstaller.
3. Run:

```bash
./installer/build_linux.sh
```

4. Use `installer/linux/gokul-omni-convert-lite.desktop` as the starter desktop launcher file for packaging or desktop integration.

## Typical macOS starter flow

```bash
./installer/macos/build_app_bundle.sh
```

This is only a starter helper around the same PyInstaller spec; you can later wrap the generated app into a signed installer flow.

## Notes

- Optional tools like LibreOffice are **not** bundled automatically. The app still supports pure Python mode without LibreOffice.
- The `soffice` path remains user-defined in the app and should never be treated as a basic runtime requirement.
- The splash GIF is configurable in-app. If the configured asset is missing, startup falls back safely without blocking the app.
- The login reminder is runtime state only; it is not installer metadata.
- If you later add a real app icon, place it in `assets/` and update the spec if needed.
- Build Center actions inside the app can still export diagnostics and state snapshots before packaging.


## Patch 18 notes

- Include `keyboard_shortcuts.md` in packaged resources so the in-app shortcut guide works offline.
- State backups live under the local app-state `backups/` folder and are not required at install time.
- High-contrast, reduced-motion, and UI-scale preferences are stored in local app state and remain editable after install.


## Patch 20 notes

- Include organizer layout JSON files in support bundles or workspace exports when you want to preserve visual reorder plans.
- The organizer now supports drag-and-drop card reordering plus local undo/redo history; no extra runtime dependency is required.
