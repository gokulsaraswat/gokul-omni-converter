<div align="center">

<img src="assets/logo/go_omni_pro_256.png" alt="GO Omni Pro" width="128">

# GO Omni Pro

**Convert, organise, and tidy up your files — entirely on your own machine.**

![Version](https://img.shields.io/badge/version-2.4.0-6d28d9)
![Python](https://img.shields.io/badge/python-3.9%2B-4f46e5)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-8b5cf6)
![License](https://img.shields.io/badge/license-MIT-orange)

</div>

---

GO Omni Pro is an offline desktop app for the file chores that usually push you to a
sketchy web uploader: converting documents, splitting and merging PDFs, running OCR,
zipping things up, and renaming a folder full of files. Nothing is uploaded anywhere —
there is no account, no telemetry, and no network call unless you explicitly paste a URL
to fetch.

## What it does

| | |
|---|---|
| **Convert** | PDF ⇄ images, Office documents → PDF, PDF → DOCX/PPTX/XLSX/HTML, Markdown, HTML, plain text. Pure-Python engine by default, with LibreOffice as an optional fallback. |
| **PDF tools** | Merge, split, extract, remove, reorder, rotate, watermark, redact, sign, compress, edit metadata, lock/unlock. |
| **OCR** | Scans and screenshots → searchable PDFs or plain text, via Tesseract. |
| **Organizer** | Visual page-level reordering for a single PDF. |
| **Archive** | Create and extract `.zip` files, inspect contents before unpacking, batch-compress folders. |
| **Rename & Organize** | Pattern renaming with live preview, find/replace, regex, case changes, and auto-sorting a folder into subfolders — all undoable. |
| **Automation** | Saved presets and a watch folder that converts new files as they land. |

## Install

### Windows

Download `GO-Omni-Pro-Setup-2.4.0.exe` from the release page and run it. It installs to
your user profile (no admin prompt), adds a Start Menu entry, and registers a normal
uninstaller under *Apps & features*.

### macOS

Run from source (see below). A signed `.app` bundle is not yet published.

```bash
brew install python-tk           # Tk 8.6+, which macOS's system Python lacks
git clone https://github.com/gokulsaraswat/go-omni-pro.git
cd go-omni-pro
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

### Linux

```bash
sudo dpkg -i go-omni-pro_2.4.0_amd64.deb
```

Then launch **GO Omni Pro** from your applications menu. To run from source instead:

```bash
sudo apt install python3-tk python3-venv     # Debian/Ubuntu
git clone https://github.com/gokulsaraswat/go-omni-pro.git
cd go-omni-pro
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

### Optional extras

Both are auto-detected on all three platforms; the app works without them.

| Tool | Gives you | Install |
|---|---|---|
| **Tesseract** | OCR | `brew install tesseract` · `apt install tesseract-ocr` · [Windows installer](https://github.com/UB-Mannheim/tesseract/wiki) |
| **LibreOffice** | Fallback converter for tricky Office files | [libreoffice.org](https://www.libreoffice.org/download/) |

## Building the installers

```bash
python installer/generate_icons.py    # regenerate .ico/.png from the SVG master
```

| Platform | Command | Needs | Output |
|---|---|---|---|
| Windows | `build_windows_release.bat` | PyInstaller + [Inno Setup 6](https://jrsoftware.org/isinfo.php) | `installer/out/GO-Omni-Pro-Setup-2.4.0.exe` |
| Linux | `./build_linux_release.sh` | PyInstaller + `dpkg-deb` | `dist/go-omni-pro_2.4.0_amd64.deb` |

> Both scripts were written on macOS and have **not** been executed. Run each on its own
> platform and confirm the artifact installs before publishing a release.

## Your data

Settings, presets, history, and undo manifests live in one folder on your machine:

| | |
|---|---|
| Windows | `%APPDATA%\GO Omni Pro\` |
| macOS | `~/Library/Application Support/GO Omni Pro/` |
| Linux | `~/.config/go-omni-pro/` |

Upgrading from 2.3.0 or earlier copies your existing `~/.gokul_omni_convert_lite/` state
across automatically on first launch; the old folder is left untouched.

## Contributing

Issues and pull requests are welcome. `python smoke_test.py` runs the manual smoke suite.

## About

Built by **Gokul Saraswat** — [gokulsaraswat.com](https://gokulsaraswat.com) ·
[GitHub](https://github.com/gokulsaraswat) ·
[LinkedIn](https://linkedin.com/in/gokulsaraswat) ·
[gokulsaraswat07@gmail.com](mailto:gokulsaraswat07@gmail.com)

Built with development assistance from Claude (Anthropic).

Licensed under the [MIT License](LICENSE). Changes are listed in [CHANGELOG.md](CHANGELOG.md).
