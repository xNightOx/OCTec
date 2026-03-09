; ------------------------------------------------------------
; Instalador OCTec com Inno Setup — Serviço via NSSM
; ------------------------------------------------------------

#define MyAppName            "OCTec"
#define MyAppVersion         "0.2.1"
#define MyAppPublisher       "Tecprinters"
#define MyAppURL             "https://www.tecprinters.com.br"
#define MyAppExeName         "OCTec.exe"           ; <-- ajuste se seu EXE tiver outro nome
#define PyInstallerOutputDir "dist\OCTec"          ; <-- ajuste se sua saída for outra
#define MyAppIconCompileTime "resources\app_icon.ico"

[Setup]
AppId={{A593AEE1-8D94-4082-B854-4FDA01EF3FB4}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputDir=installer
OutputBaseFilename=OCTecInstaller
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
SetupIconFile={#MyAppIconCompileTime}
PrivilegesRequired=admin
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "english";             MessagesFile: "compiler:Default.isl"
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "desktopicon";     Description: "{cm:CreateDesktopIcon}";       GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}";   GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Dirs]
Name: "{app}\logs"

[Files]
; EXE principal (PyInstaller)
Source: "{#PyInstallerOutputDir}\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#PyInstallerOutputDir}\_internal\*"; DestDir: "{app}\_internal"; Flags: recursesubdirs createallsubdirs ignoreversion

; VC++ Redistributable offline (2015-2022 x64)
Source: "vc_redist.x64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall nocompression; Check: not IsVCInstalled

; NSSM e scripts — note o prefixo 'source\'
Source: "source\scripts\bin\nssm.exe";           DestDir: "{app}\bin";     Flags: ignoreversion
Source: "source\scripts\install_service.bat";    DestDir: "{app}\scripts"; Flags: ignoreversion
Source: "source\scripts\uninstall_service.bat";  DestDir: "{app}\scripts"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\_internal\resources\app_icon.ico"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon; IconFilename: "{app}\_internal\resources\app_icon.ico"
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: quicklaunchicon; IconFilename: "{app}\_internal\resources\app_icon.ico"

[Run]
; 1) VC++
Filename: "{tmp}\vc_redist.x64.exe"; Parameters: "/install /quiet /norestart"; StatusMsg: "{cm:InstallingVC}"; Flags: skipifdoesntexist; Check: not IsVCInstalled

; 2) Instala e inicia o serviço via NSSM (BAT que você já colocou em source\scripts\)
Filename: "{cmd}"; Parameters: "/C ""{app}\scripts\install_service.bat"""; Flags: runhidden
Filename: "{sys}\sc.exe"; Parameters: "description OCTecService ""OCTec — motor de OCR/monitoramento (sem UI)"""; Flags: runhidden

; 3) (Opcional) abrir o tray para o usuário atual ao final
Filename: "{app}\{#MyAppExeName}"; Parameters: "--tray"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifdoesntexist

[UninstallRun]
; Remoção do serviço
Filename: "{cmd}"; Parameters: "/C ""{app}\scripts\uninstall_service.bat"""; Flags: runhidden

[CustomMessages]
InstallingVC=Instalando Microsoft Visual C++ Redistributable…

[Code]
function IsVCInstalled: Boolean;
var inst: Cardinal;
begin
  Result := False;
  if RegQueryDWordValue(HKLM64, 'SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64', 'Installed', inst) then
    if inst = 1 then Result := True;
  if not Result then
    if RegQueryDWordValue(HKLM, 'SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\x64', 'Installed', inst) then
      if inst = 1 then Result := True;
end;
