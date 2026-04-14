# Installer and packaging prep for Gokul Omni Convert Lite

Patch 9 expands the packaging folder so the app can move more cleanly toward real distributable builds later.

## What is included

- `gokul_omni_convert_lite.spec` - PyInstaller spec for the desktop app payload
- `build_windows.bat` - helper script for a Windows PyInstaller build
- `build_linux.sh` - helper script for a Linux PyInstaller build
- `windows/GokulOmniConvertLite.iss` - Inno Setup starter installer script
- `windows/version_info.txt` - Windows version metadata for the packaged executable
- `linux/gokul-omni-convert-lite.desktop` - starter desktop entry file
- `macos/build_app_bundle.sh` - starter macOS build helper

## What the build bundles

The PyInstaller spec includes the desktop app and its local data files:

- `footer_notes.md`
- `about_profile.json`
- `assets/`

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
- If you later add a real app icon, place it in `assets/` and update the spec if needed.
- Patch 9 also adds **Build Center** actions inside the app to export diagnostics and state snapshots before packaging.
- The Windows version metadata file is included so the executable can carry friendlier product information later.
