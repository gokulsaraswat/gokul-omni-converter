# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

project_dir = Path(SPEC).resolve().parent.parent
icon_candidate = project_dir / "assets" / "gokul_omni_convert_lite.ico"
icon_value = str(icon_candidate) if icon_candidate.exists() else None
version_candidate = project_dir / "installer" / "windows" / "version_info.txt"
version_value = str(version_candidate) if version_candidate.exists() else None

datas = [
    (str(project_dir / "footer_notes.md"), "."),
    (str(project_dir / "keyboard_shortcuts.md"), "."),
    (str(project_dir / "about_profile.json"), "."),
    (str(project_dir / "remote_assets.json"), "."),
    (str(project_dir / "assets"), "assets"),
    (str(project_dir / "installer" / "about_static.json"), "installer"),
]

hiddenimports = [
    "PIL._tkinter_finder",
]

block_cipher = None

a = Analysis(
    [str(project_dir / "app.py")],
    pathex=[str(project_dir)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="GokulOmniConvertLite",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    icon=icon_value,
    version=version_value,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="GokulOmniConvertLite",
)
