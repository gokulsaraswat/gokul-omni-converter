@echo off
setlocal
cd /d %~dp0\..

python -m PyInstaller --clean --noconfirm installer\gokul_omni_convert_lite.spec
if errorlevel 1 exit /b 1

set ISCC_PATH=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe
if not exist "%ISCC_PATH%" set ISCC_PATH=%ProgramFiles%\Inno Setup 6\ISCC.exe
if not exist "%ISCC_PATH%" (
    echo Inno Setup compiler was not found. Build the dist folder first or install Inno Setup 6.
    exit /b 0
)

"%ISCC_PATH%" installer\GokulOmniConvertLite.iss
