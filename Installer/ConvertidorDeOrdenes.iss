; Requiere Inno Setup 6.x
; Este instalador copia el publish completo (incluye DB\Empresas.xlsx y PrestacionesMap.csv)

#define MyAppName "ConvertidorDeOrdenes"
#define MyAppPublisher "Seres Salud"
#ifndef MyAppVersion
	#define MyAppVersion "1.0.0"
#endif
#define MyAppExeName "ConvertidorDeOrdenes.Desktop.exe"

[Setup]
AppId={{C41F3AC4-3A04-4C33-A2B3-0EBA7B9F4F5A}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppPublisher}\{#MyAppName}
DefaultGroupName={#MyAppPublisher}\{#MyAppName}
DisableProgramGroupPage=yes
OutputBaseFilename={#MyAppName}-Setup
SetupIconFile=..\icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "Crear ícono en el Escritorio"; GroupDescription: "Accesos directos:"; Flags: unchecked

[Files]
; El publish se genera con scripts\publish.ps1 en artifacts\publish
Source: "..\artifacts\publish\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir {#MyAppName}"; Flags: nowait postinstall skipifsilent
