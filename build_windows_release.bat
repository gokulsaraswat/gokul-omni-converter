@echo off
REM Build GO Omni Pro for Windows and package it as Setup.exe.
REM
REM NOTE: not executed during development (authored on macOS, where Inno Setup
REM is unavailable). Verify on a Windows machine before shipping.
REM
REM Requires: python, pyinstaller, and Inno Setup 6 (iscc.exe on PATH).
setlocal
cd /d "%~dp0"

echo ==^> Generating icons
python installer\generate_icons.py || exit /b 1

echo ==^> Running release validation
python release_validation.py || exit /b 1
python app.py --validate-release || exit /b 1

echo ==^> Building with PyInstaller
pyinstaller --clean --noconfirm installer\GOOmniPro.spec || exit /b 1

where iscc >nul 2>&1
if errorlevel 1 (
  echo Inno Setup ^(iscc.exe^) not found on PATH.
  echo The PyInstaller build is in dist\, but no Setup.exe was produced.
  echo Install Inno Setup 6 from https://jrsoftware.org/isinfo.php and re-run.
  exit /b 1
)

echo ==^> Building installer
iscc installer\GOOmniPro.iss || exit /b 1
echo Done: installer\out\GO-Omni-Pro-Setup-2.4.0.exe
endlocal
