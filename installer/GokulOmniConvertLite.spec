# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for Gokul Omni Convert Lite.

Run from the project root:
    pyinstaller --clean --noconfirm installer/GokulOmniConvertLite.spec

Optional OCR bundle:
    Put a Tesseract runtime in ./tesseract or ./Tesseract-OCR before building.
    The spec includes it only when the folder exists, so normal builds do not fail.
"""

from pathlib import Path

block_cipher = None
project_root = Path.cwd()


def add_file(datas, relative_path: str, destination: str = "."):
    path = project_root / relative_path
    if path.exists() and path.is_file():
        datas.append((str(path), destination))


def add_tree(datas, relative_dir: str):
    root = project_root / relative_dir
    if not root.exists() or not root.is_dir():
        return
    for path in root.rglob("*"):
        if path.is_file():
            destination = str(path.parent.relative_to(project_root))
            datas.append((str(path), destination))


datas = []
for file_name in (
    "about_profile.json",
    "footer_notes.md",
    "keyboard_shortcuts.md",
    "remote_assets.json",
    "requirements.txt",
):
    add_file(datas, file_name, ".")

for folder_name in ("assets", "installer"):
    add_tree(datas, folder_name)

# Optional bundled OCR runtime. Missing folders are ignored on purpose.
for tesseract_folder in ("tesseract", "Tesseract-OCR"):
    add_tree(datas, tesseract_folder)

hiddenimports = [
    "PIL._tkinter_finder",
    "PIL.JpegImagePlugin",
    "PIL.PdfImagePlugin",
    "fitz",
    "pypdf",
    "docx",
    "openpyxl",
    "pptx",
    "pdfplumber",
    "pytesseract",
]

a = Analysis(
    [str(project_root / "app.py")],
    pathex=[str(project_root)],
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
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
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
