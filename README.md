# Gokul Omni Convert Lite

![Version](https://img.shields.io/badge/version-2.3.0-blue)
![Python](https://img.shields.io/badge/python-3.10%2B-green)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20Linux-lightgrey)
![License](https://img.shields.io/badge/license-MIT-orange)

> A powerful, privacy-focused desktop conversion suite built in Python for batch document conversion, PDF workflows, OCR processing, visual page organization, automation presets, diagnostics, release validation, and installer-ready deployment.

---

# Overview

**Gokul Omni Convert Lite** is a fully local desktop application designed to simplify document and media conversion workflows without requiring cloud services.

The application provides:

* Batch document conversion
* PDF creation and manipulation
* OCR extraction
* Visual page organization
* Automation presets
* Workspace packaging
* Diagnostics and troubleshooting tools
* Release validation utilities
* Installer-ready deployment support

All processing occurs locally on the user's machine.

No files are uploaded to external servers.

---

# Key Features

## Document Conversion

Convert files between multiple formats using a pure Python engine.

Supported workflows include:

* Images → PDF
* PDF → Images
* Image Folder → PDF
* Batch file conversion
* Multi-file processing
* Recursive folder discovery

---

## PDF Workflows

Comprehensive PDF toolkit including:

### PDF Creation

* Single PDF from images
* Folder-based PDF generation
* Batch PDF generation
* ReportLab fallback generation

### PDF Utilities

* Merge PDFs
* Split PDFs
* Reorder pages
* Rotate pages
* Organize page sequences

### Visual PDF Organization

* Drag-and-drop page arrangement
* Page previews
* Thumbnail navigation
* Export organized PDF

---

## OCR Support

Optional OCR integration using Tesseract.

Capabilities include:

* Extract text from images
* OCR PDF pages
* Batch OCR processing
* Searchable PDF preparation
* Runtime OCR detection

Supported deployment methods:

* Installed Tesseract
* PATH-based Tesseract
* Bundled Tesseract runtime

Automatic validation prevents accidental selection of:

❌ `pytesseract.exe`

and guides users toward:

✅ `tesseract.exe`

---

## Image Folder → PDF Workflow

A dedicated workflow introduced in Patch 36.

Features:

* Recursive folder scanning
* Combined PDF output
* One PDF per folder mode
* Preview before execution
* Queue loading support
* Natural file sorting
* Time-based sorting

Sorting options:

* Natural filename order
* Filename A–Z
* Modified date
* Creation date

---

## Automation Presets

Save frequently used conversion settings and execute them instantly.

Preset capabilities:

* Reusable workflows
* Favorite presets
* One-click execution
* Export/import support
* Automation-ready configuration

---

## Remote Asset System

Introduced in Patch 23.

Allows future branding updates without rebuilding installers.

Supported remote assets:

* Header GIF
* Splash GIF
* About image
* Profile metadata JSON

Features:

* Local fallback protection
* Offline operation
* Cache management
* Configurable refresh intervals
* Runtime updates

---

## Diagnostics & Support Tools

Built-in troubleshooting utilities:

* Runtime diagnostics
* Dependency verification
* OCR validation
* Asset validation
* Packaging diagnostics
* Workspace inspection
* Support bundle generation

---

## Release Validation

Patch 37 introduces production-grade validation workflows.

### Headless Validation

```bash
python release_validation.py
```

Validates:

* Application structure
* Runtime dependencies
* Installer assets
* Configuration files
* Packaging readiness

---

### Application Validation Mode

```bash
python app.py --validate-release
```

Runs release validation directly through the application runtime.

---

# Architecture

```text
┌───────────────────────────────────────────┐
│                 UI Layer                  │
│      Tkinter Desktop Application          │
└───────────────────────────────────────────┘
                     │
                     ▼
┌───────────────────────────────────────────┐
│            Workflow Engine                │
│ Conversion │ OCR │ PDF │ Presets │ Assets │
└───────────────────────────────────────────┘
                     │
                     ▼
┌───────────────────────────────────────────┐
│            Processing Layer               │
│ Pillow │ ReportLab │ PyPDF │ OCR │ Utils │
└───────────────────────────────────────────┘
                     │
                     ▼
┌───────────────────────────────────────────┐
│             Storage Layer                 │
│ Config │ Cache │ Presets │ Workspace │ Logs│
└───────────────────────────────────────────┘
```

---

# Technology Stack

## Core

* Python 3.10+
* Tkinter

## PDF Processing

* PyPDF
* ReportLab

## Image Processing

* Pillow

## OCR

* Tesseract OCR (Optional)
* pytesseract

## Packaging

* PyInstaller

## Utilities

* pathlib
* threading
* logging
* json
* shutil
* subprocess

---

# Installation

## Clone Repository

```bash
git clone <repository-url>
cd Gokul-Omni-Convert-Lite
```

---

## Create Virtual Environment

### Windows

```bash
python -m venv venv
venv\Scripts\activate
```

### Linux

```bash
python3 -m venv venv
source venv/bin/activate
```

---

## Install Dependencies

```bash
pip install -r requirements.txt
```

---

## Run Application

```bash
python app.py
```

---

# Optional OCR Setup

## Install Tesseract

Download and install Tesseract OCR.

After installation:

```text
Settings → OCR Configuration
```

Select:

```text
tesseract.exe
```

Do NOT select:

```text
pytesseract.exe
```

The application validates this automatically.

---

# Project Structure

```text
GokulOmniConvertLite/
│
├── app.py
├── release_validation.py
├── requirements.txt
│
├── installer/
│   ├── GokulOmniConvertLite.spec
│   ├── build_windows.bat
│   ├── build_linux.sh
│   ├── RELEASE_NOTES.md
│   ├── RELEASE_CHECKLIST.md
│   └── TESSERACT_BUNDLE.md
│
├── assets/
│   ├── header/
│   ├── splash/
│   └── about/
│
├── cache/
│
├── presets/
│
├── logs/
│
├── workspace/
│
├── support/
│
├── diagnostics/
│
└── remote_assets.json
```

---

# Building Installers

## Windows

```bash
installer\build_windows.bat
```

---

## Linux

```bash
bash installer/build_linux.sh
```

---

## Manual Build

```bash
pyinstaller installer/GokulOmniConvertLite.spec
```

Generated output:

```text
dist/
```

---

# Release Validation Checklist

Before publishing:

### Application

* [ ] Application launches successfully
* [ ] No startup exceptions
* [ ] Version displayed correctly
* [ ] Settings load correctly

### Conversion

* [ ] Images → PDF
* [ ] PDF → Images
* [ ] Batch conversion
* [ ] Image Folder → PDF

### OCR

* [ ] OCR runtime detected
* [ ] OCR extraction successful
* [ ] Missing OCR handled gracefully

### Packaging

* [ ] Release validation passes
* [ ] Installer builds successfully
* [ ] Assets included
* [ ] Presets included
* [ ] Remote asset config included

### Distribution

* [ ] Installer tested
* [ ] Clean machine tested
* [ ] Release notes updated

---

# Patch History

## Version 2.3.0 (Patch 37)

### Added

* Release validation framework
* App-integrated validation mode
* PyInstaller specification
* Build launchers
* Installer release notes
* Release checklist
* About snapshot
* Update manifest example

### Fixed

* Duplicate external callback safety issue

---

## Version 2.2.6 (Patch 36.5)

### Fixed

* Pillow JPEG PDF export issue
* Alpha channel image handling
* OCR runtime detection
* Tesseract validation improvements

### Added

* TESSERACT_BUNDLE.md
* WORKSPACE_CLEANUP_REPORT.md

---

## Version 2.2.5 (Patch 36)

### Added

* Image Folder → PDF workflow
* Recursive folder discovery
* Folder PDF generation
* Preview workflow
* Queue integration

---

## Version 2.2.4 (Patch 35)

### Added

* UI density improvements
* Smaller header
* Cleaner home screen
* Compact About page

---

## Version 2.2.3 (Patch 34)

### Added

* Responsive layout refinements
* Home metric enhancements
* About page improvements

### Fixed

* Missing open_url helper

---

## Version 2.2.2 (Patch 33)

### Added

* Slim header
* Compact Home experience
* Cleaner About screen

---

## Version 2.2.1 (Patch 24)

### Added

* Scrollable sidebar
* Responsive button wrapping
* Improved label formatting

---

## Version 2.2.0 (Patch 23)

### Added

* Remote asset system
* Asset caching
* Remote profile support
* Offline-safe branding updates

---

# Security & Privacy

✅ Fully local processing

✅ No cloud dependency

✅ No file uploads

✅ Offline capable

✅ User-controlled OCR

✅ User-controlled LibreOffice integration

---

# Performance Goals

* Fast batch processing
* Low memory overhead
* Large folder support
* Installer-friendly deployment
* Offline-first operation

---

# Future Roadmap

### Planned Enhancements

* Auto-update service
* Cloud sync integrations
* Additional conversion formats
* Advanced OCR pipelines
* GPU-assisted image processing
* Workflow scheduling
* Plugin architecture
* Enterprise deployment profiles

---

# Author

**Gokul Saraswat**

Software Engineer | Backend Developer | Java & Python Enthusiast

Specializations:

* Java Development
* Microservices
* Distributed Systems
* Cloud Technologies
* Automation Tools
* Desktop Applications

---

# License

MIT License

Copyright (c) 2026 Gokul Saraswat

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files to deal in the Software without restriction.

---

## Final Release Status

**Gokul Omni Convert Lite 2.3.0 (Patch 37)** is considered **release-ready**.

✔ Release Validation Framework

✔ Installer Packaging Support

✔ OCR Runtime Detection

✔ Remote Asset System

✔ Image Folder PDF Workflow

✔ PDF Utilities

✔ Automation Presets

✔ Offline-First Architecture

✔ Production Distribution Ready
