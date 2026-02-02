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
Source: "..\icon.ico"; DestDir: "{app}"; Flags: ignoreversion
; Instalador offline de Microsoft Edge WebView2 Runtime (x64)
Source: "MicrosoftEdgeWebView2RuntimeInstallerX64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall; Check: not IsWebView2RuntimeInstalled

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\icon.ico"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon; IconFilename: "{app}\icon.ico"

[Run]
; Instalar WebView2 Runtime si no está presente
Filename: "{tmp}\MicrosoftEdgeWebView2RuntimeInstallerX64.exe"; Parameters: "/silent /install"; StatusMsg: "Instalando Microsoft Edge WebView2 Runtime..."; Flags: waituntilterminated; Check: not IsWebView2RuntimeInstalled
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir {#MyAppName}"; Flags: nowait postinstall skipifsilent

[Code]

function IsWebView2RuntimeInstalled: Boolean;
var
	Val: string;
begin
	{ x64 per-machine }
	if RegQueryStringValue(HKLM,
		'SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}',
		'pv', Val) then
	begin
		Result := (Val <> '') and (Val <> '0.0.0.0');
		if Result then
			Exit;
	end;

	{ x64 per-user }
	if RegQueryStringValue(HKCU,
		'Software\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}',
		'pv', Val) then
	begin
		Result := (Val <> '') and (Val <> '0.0.0.0');
		Exit;
	end;

	Result := False;
end;
