@echo off
setlocal
cd /d "%~dp0"
echo Running release validation...
python release_validation.py || exit /b 1
python app.py --validate-release || exit /b 1
echo Building with PyInstaller...
pyinstaller --clean --noconfirm installer\GokulOmniConvertLite.spec
endlocal
