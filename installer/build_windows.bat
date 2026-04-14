@echo off
setlocal
cd /d %~dp0\..
python -m PyInstaller --clean --noconfirm installer\gokul_omni_convert_lite.spec
