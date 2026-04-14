#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")/../.."
python -m PyInstaller --clean --noconfirm installer/gokul_omni_convert_lite.spec
