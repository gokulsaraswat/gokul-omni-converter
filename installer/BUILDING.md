# Installer and packaging prep for Gokul Omni Convert Lite

Patch 12 keeps the project installer-ready while separating editable in-app profile data from static packaging-time About data.

## What is included

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
