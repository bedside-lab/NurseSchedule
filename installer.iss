; Inno Setup script -- builds a real Windows installer (Setup.exe) around the
; PyInstaller onedir build in dist\NurseScheduleApp\. Produces Start Menu +
; optional Desktop shortcuts, a proper "제거/Uninstall" entry under
; Windows 설정 > 앱, and a taskbar/shortcut icon.
;
; Compiled in CI via: iscc installer.iss  (see .github/workflows/build.yml)

#define MyAppName "간호사 근무표 자동생성"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Bedside Lab"
#define MyAppURL "https://bedsidelab.blogspot.com/"
#define MyAppExeName "NurseScheduleApp.exe"

[Setup]
; unique GUID for this app -- keep this fixed across versions so upgrades
; overwrite in place instead of installing side-by-side
AppId={{B6E1E9C1-6C2E-4C7B-9C2A-2F0B7B9B9A11}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
; installs per-user under %LocalAppData%, so no admin rights are required --
; important for hospital-managed computers where staff often aren't admins
DefaultDirName={localappdata}\Programs\NurseScheduleApp
PrivilegesRequired=lowest
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputBaseFilename=NurseScheduleApp_Setup
Compression=lzma
SolidCompression=yes
SetupIconFile=assets\icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"

[Tasks]
Name: "desktopicon"; Description: "바탕화면에 바로가기 만들기"; GroupDescription: "추가 아이콘:"

[Files]
Source: "dist\NurseScheduleApp\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
