; Inno Setup script for Tally Converter
; Build onedir first: compile_onedir.bat
; Then compile this script with Inno Setup (iscc TallyConverter_setup.iss)

#define MyAppName "Tally Converter"
#define MyAppExe "TallyConverter.exe"
#define MyAppBuild "dist\TallyConverter"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion=1.0.1
DefaultDirName={autopf}\TallyConverter
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=installer_output
OutputBaseFilename=TallyConverter_Setup
SetupIconFile=tallyconverter.ico
UninstallDisplayIcon={app}\{#MyAppExe}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Copy entire onedir build (exe + TallyConverterFiles or _internal)
Source: "{#MyAppBuild}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExe}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExe}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExe}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: dirifempty; Name: "{app}"

[Code]
const
  UninstallKey = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}_is1';

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  UninstallString: string;
  ResultCode: Integer;
begin
  Result := '';
  if RegQueryStringValue(HKLM, UninstallKey, 'UninstallString', UninstallString) then
    if UninstallString <> '' then
      Exec(RemoveQuotes(UninstallString), '/VERYSILENT', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;
