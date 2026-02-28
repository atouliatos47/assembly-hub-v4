#define MyAppName "Assembly Hub"
#define MyAppVersion "1.0"
#define MyAppPublisher "Clamason Industries"
#define MyAppExeName "AssemblyHub.exe"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\AssemblyHub
DefaultGroupName={#MyAppName}
OutputDir=installer_output
OutputBaseFilename=AssemblyHub_Setup
SetupIconFile=AssemblyHub.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional icons:"
Name: "startupicon"; Description: "Start automatically when Windows boots"; GroupDescription: "Startup:"

[Files]
; Main application folder
Source: "dist\AssemblyHub\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; Icons
Source: "AssemblyHub.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "public\dashboard\icon-192.png"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Start Menu
Name: "{group}\Assembly Hub"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\AssemblyHub.ico"
Name: "{group}\Uninstall Assembly Hub"; Filename: "{uninstallexe}"

; Desktop
Name: "{autodesktop}\Assembly Hub"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\AssemblyHub.ico"; Tasks: desktopicon

; Startup
Name: "{autostartup}\Assembly Hub"; Filename: "{app}\{#MyAppExeName}"; Tasks: startupicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch Assembly Hub now"; Flags: nowait postinstall skipifsilent

[Dirs]
Name: "{app}\uploads"
Name: "{app}\public"
