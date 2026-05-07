# Patch 37 Release Report

Patch 37 prepares **Gokul Omni Convert Lite 2.3.0** for installer packaging and external sharing.

## Files added or refreshed

- `release_validation.py`
- `final_release_check.py`
- `release_check.py`
- `installer/GokulOmniConvertLite.spec`
- `installer/BUILDING.md`
- `installer/PYINSTALLER_BUILD.md`
- `installer/FINAL_RELEASE_CHECKLIST.md`
- `installer/INSTALLER_RELEASE_NOTES.md`
- `installer/about_static.json`
- `installer/update_manifest.example.json`
- `RELEASE_CHECKLIST.md`
- `PATCH_37_RELEASE_NOTES.md`
- `build_windows_release.bat`
- `build_linux_release.sh`

## Files updated

- `app.py` version bump to 2.3.0
- `README.md` Patch 37 notes
- `keyboard_shortcuts.md` release-check hints
- `requirements.txt` adds `pyinstaller` for packaging builds

## Notes

This patch intentionally avoids a UI redesign. It focuses on validation, release documentation, and build scaffolding.

Pure Python remains the default engine. LibreOffice remains optional and user-controlled.
