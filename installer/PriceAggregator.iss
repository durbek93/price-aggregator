; Inno Setup 6 — компиляция: открыть этот файл в Inno Setup → Build → Compile (F9).
; Сначала соберите exe: pyinstaller price_aggregator.spec (файл должен лежать в ..\dist\)

#define MyAppName "Price Aggregator"
#define MyAppExeName "PriceAggregator.exe"
#define MyAppVersion "1.1.0"
#define MyAppPublisher "Price Aggregator"
#define MyAppURL "https://example.com"

[Setup]
AppId={{E4C9B8A7-6D5F-4E3C-2B1A-0F9E8D7C6B5A}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=..\installer_output
OutputBaseFilename=PriceAggregatorSetup
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
; Рекомендуемый идентификатор в Inno Setup 6+ (вместо устаревшего x64)
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "..\dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
