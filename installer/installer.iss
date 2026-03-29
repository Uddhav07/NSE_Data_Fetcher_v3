; Inno Setup Script for NSE Data Fetcher v3
; Requires Inno Setup 6+ — https://jrsoftware.org/isinfo.php
;
; Build: iscc installer\installer.iss
;   (or open in Inno Setup GUI and click Compile)

#define MyAppName      "NSE Data Fetcher"
#define MyAppVersion   "3.1.0"
#define MyAppPublisher "Uddhav"
#define MyAppURL       "https://github.com/Uddhav07/NSE_Data_Fetcher_v3"
#define MyAppExeName   "NSE Data Fetcher.exe"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}/releases
DefaultDirName={userpf}\{#MyAppName}
DefaultGroupName={#MyAppName}
; Per-user install — no admin privileges required
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=..\dist\installer
OutputBaseFilename=NSE_Data_Fetcher_Setup
SetupIconFile=..\assets\icon.ico
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Single exe from PyInstaller one-file build
Source: "..\dist\NSE Data Fetcher.exe"; DestDir: "{app}"; Flags: ignoreversion

; Bundle icon for the GUI
Source: "..\assets\icon.ico"; DestDir: "{app}\assets"; Flags: ignoreversion

; Ensure default config is present (don't overwrite user's config on upgrade)
Source: "..\config\config.json"; DestDir: "{app}\config"; Flags: onlyifdoesntexist

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Dirs]
; Create writable directories that the app expects
Name: "{app}\data"
Name: "{app}\logs"
Name: "{app}\archive"
Name: "{app}\config"

[Run]
; Optionally launch app after install
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Clean up logs and cache on uninstall (but NOT data/ — user may want to keep Excel files)
Type: filesandordirs; Name: "{app}\logs"
Type: filesandordirs; Name: "{app}\archive"
Type: filesandordirs; Name: "{app}\__pycache__"
