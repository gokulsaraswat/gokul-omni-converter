# Installer build notes

This folder contains installer-oriented assets and static metadata for **Gokul Omni Convert Lite**.

## Static About
- `about_static.json` is the installer-safe snapshot of the About profile.
- The running app still reads `about_profile.json` for local editing.

## Bundled assets
- `assets/gokul_header.gif`
- `assets/gokul_splash.gif`
- `assets/gokul_profile_placeholder.png`

## Packaging direction
- Keep Pure Python as the default engine.
- LibreOffice stays optional and user-configurable through Settings.
- Bundle local assets so the app works offline after install.
- Remote asset refresh from GitHub stays optional.
