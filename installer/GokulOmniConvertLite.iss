; Gokul Omni Convert Lite installer template for Inno Setup 6
#define AppName "Gokul Omni Convert Lite"
#define AppVersion "0.10.0 Patch 9"
#define AppPublisher "Gokul Saraswat"
#define AppExeName "GokulOmniConvertLite.exe"
#define SourceDir "..\dist\GokulOmniConvertLite"

[Setup]
AppId={{A7D0FB7A-2B6A-48F3-AE20-7B0B807D692B}}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
OutputDir=..\release_output
OutputBaseFilename=GokulOmniConvertLite_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
SetupIconFile=..\assets\gokul_omni_convert_lite.ico
UninstallDisplayIcon={app}\{#AppExeName}

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{group}\Uninstall {#AppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional icons:"

[Run]
Filename: "{app}\{#AppExeName}"; Description: "Launch {#AppName}"; Flags: nowait postinstall skipifsilent
