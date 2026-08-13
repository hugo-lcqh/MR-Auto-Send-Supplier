#define MyAppName "MR Auto Send Supplier"
#define MyAppVersion "1.3.0"
#define MyAppPublisher "Hugo Aka Le Chi Quoc Hung"

[Setup]
AppId={{8D183A16-6C3F-4EF1-B296-502FB04CF3C6}
AppName=MR Auto Send Supplier
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\MR Auto Send Supplier
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=..\Release
OutputBaseFilename=MR-Auto-Setup-v{#MyAppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
SetupLogging=yes
CloseApplications=yes
VersionInfoVersion={#MyAppVersion}.0
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription={#MyAppName} Setup
UninstallDisplayName={#MyAppName}
InfoBeforeFile=ReleaseNotes.txt

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Messages]
WizardInfoBefore=What's New in MR Auto Send Supplier
InfoBeforeLabel=Please review the fixes and improvements in version {#MyAppVersion} before continuing.

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"

[Files]
Source: "..\Scripts\MR-Launcher.ps1"; DestDir: "{app}\Scripts"; Flags: ignoreversion
Source: "..\Scripts\MR-Outlook.ps1"; DestDir: "{app}\Scripts"; Flags: ignoreversion
Source: "..\Input\Template.xlsx"; DestDir: "{app}\Input"; Flags: ignoreversion
Source: "..\Config\suppliers.example.csv"; DestDir: "{app}\Config"; DestName: "suppliers.csv"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "..\Config\reply_folder.txt"; DestDir: "{app}\Config"; Flags: onlyifdoesntexist uninsneveruninstall

[Dirs]
Name: "{app}\Config"; Flags: uninsneveruninstall
Name: "{app}\Input"; Flags: uninsneveruninstall
Name: "{app}\MR_Out"; Flags: uninsneveruninstall

[Icons]
Name: "{group}\MR Auto Send Supplier"; Filename: "{sys}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ""{app}\Scripts\MR-Launcher.ps1"""; WorkingDir: "{app}"
Name: "{group}\Open MR Auto Folder"; Filename: "{sys}\explorer.exe"; Parameters: """{app}"""
Name: "{autodesktop}\MR Auto Send Supplier"; Filename: "{sys}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ""{app}\Scripts\MR-Launcher.ps1"""; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{sys}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ""{app}\Scripts\MR-Launcher.ps1"""; WorkingDir: "{app}"; Description: "Launch MR Auto Send Supplier"; Flags: nowait postinstall skipifsilent

[Code]
function OfficeAutomationRegistered(ProgId: String): Boolean;
begin
  Result := RegKeyExists(HKCR32, ProgId + '\CLSID');
  if IsWin64 then
    Result := Result or RegKeyExists(HKCR64, ProgId + '\CLSID');
end;

function InitializeSetup(): Boolean;
begin
  Result := False;

  if not OfficeAutomationRegistered('Excel.Application') then
  begin
    MsgBox('Microsoft Excel desktop is required. Install Excel, then run this setup again.', mbError, MB_OK);
    Exit;
  end;

  if not OfficeAutomationRegistered('Outlook.Application') then
  begin
    MsgBox('Classic Outlook desktop is required. New Outlook does not support the Outlook automation used by MR Auto.', mbError, MB_OK);
    Exit;
  end;

  Result := True;
end;
