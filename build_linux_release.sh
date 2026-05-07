#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"
echo "Running release validation..."
python release_validation.py
python app.py --validate-release
echo "Building with PyInstaller..."
pyinstaller --clean --noconfirm installer/GokulOmniConvertLite.spec
