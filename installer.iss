[Setup]
AppName=Carbon Zones
AppVersion=1.0.7
AppPublisher=Sultansmooth
AppPublisherURL=https://github.com/Sultansmooth/CarbonFences
DefaultDirName={autopf}\Carbon Zones
DefaultGroupName=Carbon Zones
UninstallDisplayIcon={app}\CarbonZones.exe
OutputDir=installer_output
OutputBaseFilename=CarbonZones-Setup
Compression=lzma2
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
SetupIconFile=CarbonZones\icon.ico
WizardStyle=modern
PrivilegesRequired=lowest
CloseApplications=force
CloseApplicationsFilter=CarbonZones.exe
AppMutex=Carbon_Zones

[Files]
Source: "publish\CarbonZones\CarbonZones.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "publish\CarbonZones\CarbonZones.dll.config"; DestDir: "{app}"; Flags: ignoreversion
Source: "publish\CarbonZones\CarbonZones.pdb"; DestDir: "{app}"; Flags: ignoreversion
Source: "publish\CarbonZones\D3DCompiler_47_cor3.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "publish\CarbonZones\icon.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "publish\CarbonZones\PenImc_cor3.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "publish\CarbonZones\PresentationNative_cor3.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "publish\CarbonZones\vcruntime140_cor3.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "publish\CarbonZones\wpfgfx_cor3.dll"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Carbon Zones"; Filename: "{app}\CarbonZones.exe"; IconFilename: "{app}\icon.ico"
Name: "{group}\Uninstall Carbon Zones"; Filename: "{uninstallexe}"
Name: "{autodesktop}\Carbon Zones"; Filename: "{app}\CarbonZones.exe"; IconFilename: "{app}\icon.ico"; Tasks: desktopicon
Name: "{userstartup}\Carbon Zones"; Filename: "{app}\CarbonZones.exe"; IconFilename: "{app}\icon.ico"; Tasks: startupicon

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"
Name: "startupicon"; Description: "Start Carbon Zones when Windows starts"; GroupDescription: "Startup:"; Flags: unchecked

[Run]
Filename: "{app}\CarbonZones.exe"; Description: "Launch Carbon Zones"; Flags: nowait postinstall skipifsilent

[Code]
function InitializeSetup(): Boolean;
var
  ResultCode: Integer;
begin
  Exec('taskkill', '/f /im CarbonZones.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Sleep(500);
  Result := True;
end;

function InitializeUninstall(): Boolean;
var
  ResultCode: Integer;
begin
  Exec('taskkill', '/f /im CarbonZones.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Sleep(500);
  Result := True;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  StagingDir: String;
  DesktopDir: String;
  FindRec: TFindRec;
begin
  if CurUninstallStep = usPostUninstall then
  begin
    { Move staged files back to desktop }
    StagingDir := ExpandConstant('{localappdata}\CarbonZones\__staged');
    DesktopDir := ExpandConstant('{userdesktop}');
    if DirExists(StagingDir) then
    begin
      if FindFirst(StagingDir + '\*', FindRec) then
      try
        repeat
          if (FindRec.Name <> '.') and (FindRec.Name <> '..') then
            RenameFile(StagingDir + '\' + FindRec.Name, DesktopDir + '\' + FindRec.Name);
        until not FindNext(FindRec);
      finally
        FindClose(FindRec);
      end;
      RemoveDir(StagingDir);
    end;
    { Clean up CarbonZones data directory }
    DelTree(ExpandConstant('{localappdata}\CarbonZones'), True, True, True);
  end;
end;
