; Inno Setup script for GO Omni Pro (Windows).
;
; Build order (from the project root, on Windows):
;   1. python installer\generate_icons.py
;   2. pyinstaller --clean --noconfirm installer\GOOmniPro.spec
;   3. iscc installer\GOOmniPro.iss
;
; Step 3 needs Inno Setup 6 (https://jrsoftware.org/isinfo.php) on PATH as `iscc`.
; The result is installer\out\GO-Omni-Pro-Setup-<version>.exe.
;
; NOTE: this script has not been compiled or run — it was authored on macOS,
; where Inno Setup is unavailable. Verify on a Windows machine before shipping.

#define AppName        "GO Omni Pro"
#define AppVersion     "2.4.0"
#define AppPublisher   "Gokul Saraswat"
#define AppURL         "https://gokulsaraswat.com"
#define AppExeName     "GO Omni Pro.exe"
#define SourceDir      "..\dist\GO Omni Pro"

[Setup]
AppId={{7A1F2C64-9B3D-4E5A-9C21-0F6B8D3E4A77}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
; Per-user install by default, so no UAC prompt is needed.
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=out
OutputBaseFilename=GO-Omni-Pro-Setup-{#AppVersion}
SetupIconFile=..\assets\logo\go_omni_pro.ico
UninstallDisplayIcon={app}\{#AppExeName}
UninstallDisplayName={#AppName}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional shortcuts:"

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{group}\Uninstall {#AppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExeName}"; Description: "Launch {#AppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; PyInstaller leaves __pycache__ behind; remove it so the folder uninstalls clean.
Type: filesandordirs; Name: "{app}\__pycache__"
