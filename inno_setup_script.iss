#define MyAppName "Sistema de Evaluación Docente"
#define MyAppVersion "1.0"
#define MyAppPublisher "UNIBE"
#define MyAppExeName "EvaluacionDocente.exe"

[Setup]
AppId={{8BE498B5-92FF-4E06-B011-84593A3638E3}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={commonpf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputBaseFilename=EvaluacionDocente_Setup
OutputDir=installer
Compression=lzma
SolidCompression=yes
DisableDirPage=no
PrivilegesRequired=admin

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Dirs]
Name: "{commonappdata}\EvaluacionDocente\logs"; Permissions: everyone-full

[Files]
Source: "dist\EvaluacionDocente\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent