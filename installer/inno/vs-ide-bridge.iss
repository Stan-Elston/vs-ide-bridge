#define MyAppName "VS IDE Bridge"
#define MyAppFolderName "VsIdeBridge"
#define MyAppPublisher "RenegadeRiff86"
#define MyAppURL "https://github.com/RenegadeRiff86/Visual-Studio-MCP"
#define MyAppVersion "3.0.4"
#define ServiceName "VsIdeBridgeService"
#define VsixId "RenegadeRiff86.VsIdeBridge"
#define LegacyVsixId "StanElston.VsIdeBridge"
#define Configuration "Release"

[Setup]
AppId={{F0B67A29-5A6A-4A0F-AD99-9F8A907A2A2E}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppFolderName}
DisableProgramGroupPage=yes
LicenseFile=..\..\LICENSE
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=..\output
OutputBaseFilename=vs-ide-bridge-setup-{#MyAppVersion}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\service\VsIdeBridgeService.exe
SetupLogging=yes
CloseApplications=force
CloseApplicationsFilter=VsIdeBridgeService.exe

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "service"; Description: "Install Windows service (automatic start)"
Name: "startservice"; Description: "Start service after install"; Check: WizardIsTaskSelected('service')
Name: "enablehttp"; Description: "Enable local MCP HTTP endpoints on localhost (recommended)"

[Dirs]
Name: "{commonappdata}\VsIdeBridge\state"; Permissions: users-modify; Flags: uninsneveruninstall
Name: "{commonappdata}\VsIdeBridge\logs"; Permissions: users-modify; Flags: uninsneveruninstall
Name: "{app}\logs"; Permissions: users-modify; Flags: uninsneveruninstall

[Registry]
Root: HKLM; Subkey: "SOFTWARE\VsIdeBridge"; ValueType: string; ValueName: "InstallPath"; ValueData: "{app}"; Flags: uninsdeletekey

[InstallDelete]
Type: filesandordirs; Name: "{app}\cli"

[Files]
Source: "..\..\src\VsIdeBridgeService\bin\{#Configuration}\net8.0-windows\*"; DestDir: "{app}\service"; Flags: recursesubdirs createallsubdirs ignoreversion restartreplace uninsrestartdelete
Source: "..\..\src\VsIdeBridgeLauncher\bin\{#Configuration}\*"; DestDir: "{app}\service"; Flags: recursesubdirs createallsubdirs ignoreversion restartreplace uninsrestartdelete
Source: "..\..\src\VsIdeBridge\bin\{#Configuration}\net472\VsIdeBridge.vsix"; DestDir: "{app}\vsix"; Flags: ignoreversion
Source: "..\..\src\VsIdeBridgeInstaller\bin\{#Configuration}\net8.0-windows\python-runtime\*"; DestDir: "{app}\python\managed-runtime"; Flags: recursesubdirs createallsubdirs ignoreversion uninsrestartdelete; Check: ShouldInstallManagedPython
Source: "..\..\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\..\LICENSE"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\..\THIRD_PARTY_NOTICES.md"; DestDir: "{app}"; Flags: ignoreversion

[UninstallRun]
Filename: "{sys}\sc.exe"; Parameters: "stop ""{#ServiceName}"""; Flags: runhidden waituntilterminated; RunOnceId: "{#ServiceName}-stop"; StatusMsg: "Stopping VS IDE Bridge service..."
Filename: "{sys}\sc.exe"; Parameters: "delete ""{#ServiceName}"""; Flags: runhidden waituntilterminated; RunOnceId: "{#ServiceName}-delete"; StatusMsg: "Removing VS IDE Bridge service..."
Filename: "{code:GetVsixInstallerPath}"; Parameters: "/quiet /shutdownprocesses /logFile:""{log}\vsix-uninstall.log"" /uninstall:{#VsixId}"; Flags: waituntilterminated; Check: HasVsixInstaller; RunOnceId: "{#VsixId}-uninstall"; StatusMsg: "Removing VS IDE Bridge extension from Visual Studio..."
Filename: "{code:GetVsixInstallerPath}"; Parameters: "/quiet /shutdownprocesses /logFile:""{log}\vsix-uninstall-legacy.log"" /uninstall:{#LegacyVsixId}"; Flags: waituntilterminated; Check: HasVsixInstaller; RunOnceId: "{#LegacyVsixId}-uninstall"; StatusMsg: "Removing previous VS IDE Bridge extension identity..."

[Code]
const
  InstallerLineBreak = #13#10;
  InstallerDoubleLineBreak = #13#10#13#10;
  PostInstallPageTitle = 'Configuring VS IDE Bridge';
  PostInstallPageDescription = 'Running Windows service and Visual Studio extension setup.';
  PythonSupportPageTitle = 'Python Runtime Support';
  PythonSupportPageDescription = 'Choose how VS IDE Bridge should provision Python support.';
  PythonProvisioningManaged = 'managed';
  PythonProvisioningSkip = 'skip';
  // No /shutdownprocesses on the install path: instead of force-closing the user's
  // Visual Studio, EnsureVisualStudioClosedForVsix() warns and waits for them to close
  // it (force-close only on explicit opt-in or in silent/unattended installs).
  VsixInstallerLogArgumentPrefix = '/quiet /logFile:';

var
  CachedVsixInstallerPath: string;
  PostInstallProgressPage: TOutputProgressWizardPage;
  PostInstallCompleted: Boolean;
  { Set when the user declines to close Visual Studio, so the VS extension install is
    skipped and the finish page can say so instead of claiming success. }
  VsixInstallSkipped: Boolean;
  PythonSupportPage: TWizardPage;
  ManagedPythonRadioButton: TRadioButton;
  SkipManagedPythonRadioButton: TRadioButton;

function EscapeJsonString(const Value: string): string;
begin
  Result := Value;
  StringChangeEx(Result, '\', '\\', True);
  StringChangeEx(Result, '"', '\"', True);
end;

function GetSelectedPythonProvisioningMode(): string;
begin
  if (ManagedPythonRadioButton = nil) or ManagedPythonRadioButton.Checked then
    Result := PythonProvisioningManaged
  else
    Result := PythonProvisioningSkip;
end;

function ShouldInstallManagedPython(): Boolean;
begin
  Result := GetSelectedPythonProvisioningMode() = PythonProvisioningManaged;
end;

function GetManagedPythonInterpreterPath(): string;
begin
  Result := ExpandConstant('{app}\python\managed-runtime\python.exe');
end;

function GetManagedPythonRuntimeVersion(): string;
begin
  Result := '';
  GetVersionNumbersString(GetManagedPythonInterpreterPath(), Result);
end;

function GetPythonRuntimeConfigPath(): string;
begin
  Result := ExpandConstant('{app}\config\python-runtime.json');
end;

function GetHttpEnabledFlagPath(): string;
begin
  Result := ExpandConstant('{commonappdata}\VsIdeBridge\state\http-enabled.flag');
end;

function GetStreamableHttpEnabledFlagPath(): string;
begin
  Result := ExpandConstant('{commonappdata}\VsIdeBridge\state\streamable-http-enabled.flag');
end;

procedure RemoveManagedPythonRuntimeIfPresent();
begin
  DelTree(ExpandConstant('{app}\python\managed-runtime'), True, True, True);
end;

procedure WritePythonRuntimeConfig();
var
  ConfigDirectory: string;
  ConfigPath: string;
  JsonText: string;
  ManagedPythonPath: string;
  ManagedPythonVersion: string;
begin
  ConfigPath := GetPythonRuntimeConfigPath();
  ConfigDirectory := ExtractFileDir(ConfigPath);
  if ConfigDirectory <> '' then
    ForceDirectories(ConfigDirectory);

  if ShouldInstallManagedPython() then
  begin
    ManagedPythonPath := EscapeJsonString(GetManagedPythonInterpreterPath());
    ManagedPythonVersion := EscapeJsonString(GetManagedPythonRuntimeVersion());
    JsonText :=
      '{' + InstallerLineBreak +
      '  "provisioningMode": "' + PythonProvisioningManaged + '",' + InstallerLineBreak +
      '  "managedRuntimeVersion": "' + ManagedPythonVersion + '",' + InstallerLineBreak +
      '  "managedEnvironmentPath": "' + ManagedPythonPath + '",' + InstallerLineBreak +
      '  "managedBaseInterpreterPath": "' + ManagedPythonPath + '"' + InstallerLineBreak +
      '}';
  end
  else
  begin
    RemoveManagedPythonRuntimeIfPresent();
    JsonText :=
      '{' + InstallerLineBreak +
      '  "provisioningMode": "' + PythonProvisioningSkip + '"' + InstallerLineBreak +
      '}';
  end;

  SaveStringToFile(ConfigPath, JsonText, False);
end;

procedure SetHttpFlagFile(FlagPath: string; Enabled: Boolean);
var
  StateDirectory: string;
begin
  StateDirectory := ExtractFileDir(FlagPath);
  if StateDirectory <> '' then
    ForceDirectories(StateDirectory);

  if Enabled then
    SaveStringToFile(FlagPath, '', False)
  else if FileExists(FlagPath) then
    DeleteFile(FlagPath);
end;

procedure WriteHttpServerConfig();
var
  Enabled: Boolean;
begin
  Enabled := WizardIsTaskSelected('enablehttp');
  SetHttpFlagFile(GetHttpEnabledFlagPath(), Enabled);
  SetHttpFlagFile(GetStreamableHttpEnabledFlagPath(), Enabled);
end;

procedure InitializePythonSupportPage();
var
  ManagedDescription: TNewStaticText;
  SkipDescription: TNewStaticText;
begin
  PythonSupportPage := CreateCustomPage(
    wpSelectTasks,
    PythonSupportPageTitle,
    PythonSupportPageDescription);

  ManagedPythonRadioButton := TRadioButton.Create(PythonSupportPage.Surface);
  ManagedPythonRadioButton.Parent := PythonSupportPage.Surface;
  ManagedPythonRadioButton.Left := ScaleX(0);
  ManagedPythonRadioButton.Top := ScaleY(8);
  ManagedPythonRadioButton.Width := PythonSupportPage.SurfaceWidth;
  ManagedPythonRadioButton.Caption := 'Bridge-managed CPython environment (Recommended)';
  ManagedPythonRadioButton.Checked := True;

  ManagedDescription := TNewStaticText.Create(PythonSupportPage.Surface);
  ManagedDescription.Parent := PythonSupportPage.Surface;
  ManagedDescription.Left := ScaleX(20);
  ManagedDescription.Top := ManagedPythonRadioButton.Top + ScaleY(20);
  ManagedDescription.Width := PythonSupportPage.SurfaceWidth - ScaleX(20);
  ManagedDescription.Height := ScaleY(32);
  ManagedDescription.AutoSize := False;
  ManagedDescription.WordWrap := True;
  ManagedDescription.Caption := 'Install a bridge-owned CPython runtime under the bridge install directory. This keeps bridge modules separate from your working Python environments.';

  SkipManagedPythonRadioButton := TRadioButton.Create(PythonSupportPage.Surface);
  SkipManagedPythonRadioButton.Parent := PythonSupportPage.Surface;
  SkipManagedPythonRadioButton.Left := ScaleX(0);
  SkipManagedPythonRadioButton.Top := ManagedDescription.Top + ManagedDescription.Height + ScaleY(16);
  SkipManagedPythonRadioButton.Width := PythonSupportPage.SurfaceWidth;
  SkipManagedPythonRadioButton.Caption := 'Skip bridge-managed Python';

  SkipDescription := TNewStaticText.Create(PythonSupportPage.Surface);
  SkipDescription.Parent := PythonSupportPage.Surface;
  SkipDescription.Left := ScaleX(20);
  SkipDescription.Top := SkipManagedPythonRadioButton.Top + ScaleY(20);
  SkipDescription.Width := PythonSupportPage.SurfaceWidth - ScaleX(20);
  SkipDescription.Height := ScaleY(32);
  SkipDescription.AutoSize := False;
  SkipDescription.WordWrap := True;
  SkipDescription.Caption := 'Install the bridge without a managed Python runtime. You can still attach an existing interpreter or environment later.';
end;

procedure InitializeWizard();
begin
  PostInstallProgressPage := CreateOutputProgressPage(
    PostInstallPageTitle,
    PostInstallPageDescription);
  InitializePythonSupportPage();
end;

function AppendPath(const Base, Relative: string): string;
begin
  if Base = '' then
    Result := Relative
  else if Base[Length(Base)] = '\' then
    Result := Base + Relative
  else
    Result := Base + '\' + Relative;
end;

function GetVswherePath(): string;
var
  Candidate: string;
begin
  Candidate := ExpandConstant('{pf32}\Microsoft Visual Studio\Installer\vswhere.exe');
  if FileExists(Candidate) then
  begin
    Result := Candidate;
    Exit;
  end;

  Candidate := ExpandConstant('{pf}\Microsoft Visual Studio\Installer\vswhere.exe');
  if FileExists(Candidate) then
  begin
    Result := Candidate;
    Exit;
  end;

  Result := '';
end;

function RunVswhereQuery(const Arguments: string; var Lines: TArrayOfString): Boolean;
var
  VswherePath: string;
  OutputPath: string;
  CommandLine: string;
  ExitCode: Integer;
begin
  Result := False;
  SetArrayLength(Lines, 0);

  VswherePath := GetVswherePath();
  if VswherePath = '' then
  begin
    Log('vswhere.exe was not found in the Visual Studio Installer directory.');
    Exit;
  end;

  OutputPath := ExpandConstant('{tmp}\vsidebridge-vswhere.txt');
  DeleteFile(OutputPath);
  CommandLine := '/C ""' + VswherePath + '" ' + Arguments + ' > "' + OutputPath + '""';

  Log('Running vswhere: "' + VswherePath + '" ' + Arguments);
  if not Exec(ExpandConstant('{sys}\cmd.exe'), CommandLine, '', SW_HIDE, ewWaitUntilTerminated, ExitCode) then
  begin
    Log(Format('Failed to launch vswhere. Result code %d.', [ExitCode]));
    Exit;
  end;

  if ExitCode <> 0 then
  begin
    Log(Format('vswhere exited with code %d.', [ExitCode]));
    Exit;
  end;

  Result := LoadStringsFromFile(OutputPath, Lines);
  if not Result then
    Log('vswhere output file could not be read: ' + OutputPath);

  DeleteFile(OutputPath);
end;

function FirstExistingPath(const Lines: TArrayOfString): string;
var
  I: Integer;
  Candidate: string;
begin
  for I := 0 to GetArrayLength(Lines) - 1 do
  begin
    Candidate := Trim(Lines[I]);
    if (Candidate <> '') and FileExists(Candidate) then
    begin
      Result := Candidate;
      Exit;
    end;
  end;

  Result := '';
end;

function FirstExistingVsixInstallerFromInstallPaths(const Lines: TArrayOfString): string;
var
  I: Integer;
  InstallPath: string;
  Candidate: string;
begin
  for I := 0 to GetArrayLength(Lines) - 1 do
  begin
    InstallPath := Trim(Lines[I]);
    if InstallPath <> '' then
    begin
      Candidate := AppendPath(InstallPath, 'Common7\IDE\VSIXInstaller.exe');
      if FileExists(Candidate) then
      begin
        Result := Candidate;
        Exit;
      end;
    end;
  end;

  Result := '';
end;

function ResolveVsixInstallerFromVswhereFind(const Arguments: string): string;
var
  Lines: TArrayOfString;
begin
  if RunVswhereQuery(Arguments, Lines) then
    Result := FirstExistingPath(Lines)
  else
    Result := '';
end;

function ResolveVsixInstallerFromVswhereInstallPaths(const Arguments: string): string;
var
  Lines: TArrayOfString;
begin
  if RunVswhereQuery(Arguments, Lines) then
    Result := FirstExistingVsixInstallerFromInstallPaths(Lines)
  else
    Result := '';
end;

function ResolveVsixInstallerFromVisualStudioRoot(const Root: string): string;
var
  VersionRec: TFindRec;
  EditionRec: TFindRec;
  VersionPath: string;
  EditionPath: string;
  Candidate: string;
begin
  Result := '';
  if not DirExists(Root) then
    Exit;

  if FindFirst(AppendPath(Root, '*'), VersionRec) then
  begin
    try
      repeat
        VersionPath := AppendPath(Root, VersionRec.Name);
        if (VersionRec.Name <> '.') and (VersionRec.Name <> '..') and DirExists(VersionPath) then
        begin
          Candidate := AppendPath(VersionPath, 'Common7\IDE\VSIXInstaller.exe');
          if FileExists(Candidate) then
          begin
            Result := Candidate;
            Exit;
          end;

          if FindFirst(AppendPath(VersionPath, '*'), EditionRec) then
          begin
            try
              repeat
                EditionPath := AppendPath(VersionPath, EditionRec.Name);
                if (EditionRec.Name <> '.') and (EditionRec.Name <> '..') and DirExists(EditionPath) then
                begin
                  Candidate := AppendPath(EditionPath, 'Common7\IDE\VSIXInstaller.exe');
                  if FileExists(Candidate) then
                  begin
                    Result := Candidate;
                    Exit;
                  end;
                end;
              until not FindNext(EditionRec);
            finally
              FindClose(EditionRec);
            end;
          end;
        end;
      until not FindNext(VersionRec);
    finally
      FindClose(VersionRec);
    end;
  end;
end;

function ResolveVsixInstallerFromLegacyProgramFilesRoot(const Root: string): string;
var
  FindRec: TFindRec;
  VisualStudioPath: string;
  Candidate: string;
begin
  Result := '';
  if not DirExists(Root) then
    Exit;

  if FindFirst(AppendPath(Root, 'Microsoft Visual Studio*'), FindRec) then
  begin
    try
      repeat
        VisualStudioPath := AppendPath(Root, FindRec.Name);
        if (FindRec.Name <> '.') and (FindRec.Name <> '..') and DirExists(VisualStudioPath) then
        begin
          Candidate := AppendPath(VisualStudioPath, 'Common7\IDE\VSIXInstaller.exe');
          if FileExists(Candidate) then
          begin
            Result := Candidate;
            Exit;
          end;
        end;
      until not FindNext(FindRec);
    finally
      FindClose(FindRec);
    end;
  end;
end;

function ResolveVsixInstallerFromProgramFilesScan(): string;
var
  I: Integer;
  Roots: array[0..1] of string;
  Candidate: string;
begin
  Roots[0] := ExpandConstant('{pf}\Microsoft Visual Studio');
  Roots[1] := ExpandConstant('{pf32}\Microsoft Visual Studio');

  for I := 0 to 1 do
  begin
    Candidate := ResolveVsixInstallerFromVisualStudioRoot(Roots[I]);
    if Candidate <> '' then
    begin
      Result := Candidate;
      Exit;
    end;
  end;

  Roots[0] := ExpandConstant('{pf}');
  Roots[1] := ExpandConstant('{pf32}');

  for I := 0 to 1 do
  begin
    Candidate := ResolveVsixInstallerFromLegacyProgramFilesRoot(Roots[I]);
    if Candidate <> '' then
    begin
      Result := Candidate;
      Exit;
    end;
  end;

  Result := '';
end;

function CacheVsixInstallerPath(const Candidate: string): string;
begin
  CachedVsixInstallerPath := Candidate;
  Log('Resolved VSIXInstaller.exe: ' + Candidate);
  Result := Candidate;
end;

function ResolveVsixInstallerPath(): string;
var
  Candidate: string;
begin
  if CachedVsixInstallerPath <> '' then
  begin
    Result := CachedVsixInstallerPath;
    Exit;
  end;

  Candidate := ResolveVsixInstallerFromVswhereFind(
    '-latest -prerelease -products * -requires Microsoft.VisualStudio.Component.CoreEditor -find Common7\IDE\VSIXInstaller.exe');
  if Candidate <> '' then
  begin
    Result := CacheVsixInstallerPath(Candidate);
    Exit;
  end;

  Candidate := ResolveVsixInstallerFromVswhereFind(
    '-all -prerelease -products * -requires Microsoft.VisualStudio.Component.CoreEditor -find Common7\IDE\VSIXInstaller.exe');
  if Candidate <> '' then
  begin
    Result := CacheVsixInstallerPath(Candidate);
    Exit;
  end;

  Candidate := ResolveVsixInstallerFromVswhereInstallPaths(
    '-legacy -all -property installationPath');
  if Candidate <> '' then
  begin
    Result := CacheVsixInstallerPath(Candidate);
    Exit;
  end;

  Candidate := ResolveVsixInstallerFromProgramFilesScan();
  if Candidate <> '' then
  begin
    Result := CacheVsixInstallerPath(Candidate);
    Exit;
  end;

  Result := '';
end;

function GetVsixInstallerPath(Param: string): string;
begin
  Result := ResolveVsixInstallerPath();
end;

function HasVsixInstaller(): Boolean;
begin
  Result := ResolveVsixInstallerPath() <> '';
end;

function GetServiceName(): string;
begin
  Result := '{#ServiceName}';
end;

function GetServiceBinaryPath(): string;
begin
  Result := ExpandConstant('{app}\service\VsIdeBridgeService.exe');
end;

function GetServiceCommandParameters(const Verb: string): string;
begin
  Result := Format('%s "%s"', [Verb, GetServiceName()]);
end;

function GetServiceCreateParameters(): string;
begin
  Result := Format('create "%s" binPath= "%s" start= auto DisplayName= "VS IDE Bridge Service"', [GetServiceName(), GetServiceBinaryPath()]);
end;

function GetServiceDescriptionParameters(): string;
begin
  Result := Format('description "%s" "VS IDE Bridge service host (automatic start, idle auto-stop)."', [GetServiceName()]);
end;

function GetServiceFailureParameters(): string;
begin
  // Restart after 3s, 10s, 30s — reset failure count after 0s (never).
  // This ensures the service restarts after an idle auto-stop.
  Result := Format('failure "%s" reset= 0 actions= restart/3000/restart/10000/restart/30000', [GetServiceName()]);
end;

function GetServiceFailureFlagParameters(): string;
begin
  // Trigger restart actions even on a clean exit (code 0), not just on crash.
  // Without this, the idle auto-stop (which calls Stop() with code 0) would
  // not trigger the recovery actions above.
  Result := Format('failureflag "%s" 1', [GetServiceName()]);
end;

function GetVsixLogFileArgument(const LogFileName: string): string;
begin
  // {log} is the Setup log FILE, so '{log}\name.log' was an invalid file-as-directory path
  // and VSIXInstaller silently wrote nothing. Log into the product logs directory (created
  // by [Dirs]) so VSIX install/uninstall output is actually available for diagnosis.
  Result := ExpandConstant(VsixInstallerLogArgumentPrefix + '"{commonappdata}\VsIdeBridge\logs\' + LogFileName + '"');
end;

function IsVisualStudioRunning(): Boolean;
var
  ExitCode: Integer;
  TempFile: string;
  Content: AnsiString;
begin
  Result := False;
  TempFile := ExpandConstant('{tmp}\devenv-check.txt');
  // tasklist filter is exact-name; redirect to a temp file and scan for the process.
  if Exec(ExpandConstant('{cmd}'), '/C tasklist /FI "IMAGENAME eq devenv.exe" /NH > "' + TempFile + '" 2>&1',
          '', SW_HIDE, ewWaitUntilTerminated, ExitCode) then
  begin
    if LoadStringFromFile(TempFile, Content) then
      Result := Pos('devenv.exe', LowerCase(string(Content))) > 0;
  end;
  DeleteFile(TempFile);
end;

// True only when a devenv.exe with a REAL window is running (a user is actually using VS).
// A leftover zombie devenv.exe has no window, so tasklist /V reports its title as "N/A";
// a real instance's title always contains "Microsoft Visual Studio". This lets the installer
// silently clear zombies (no user data) while still warning about an in-use Visual Studio.
function RealVisualStudioRunning(): Boolean;
var
  ExitCode: Integer;
  TempFile: string;
  Content: AnsiString;
begin
  Result := False;
  TempFile := ExpandConstant('{tmp}\devenv-v-check.txt');
  if Exec(ExpandConstant('{cmd}'), '/C tasklist /V /FI "IMAGENAME eq devenv.exe" /FO LIST > "' + TempFile + '" 2>&1',
          '', SW_HIDE, ewWaitUntilTerminated, ExitCode) then
  begin
    if LoadStringFromFile(TempFile, Content) then
      Result := Pos('visual studio', LowerCase(string(Content))) > 0;
  end;
  DeleteFile(TempFile);
end;

// Force-close every devenv.exe and poll until they are actually gone (a lingering
// zombie devenv.exe can outlive the visible windows and still hold the extension).
// Returns True when no devenv.exe remains.
function ForceCloseVisualStudio(): Boolean;
var
  ExitCode, Attempt: Integer;
begin
  Log('ForceCloseVisualStudio: taskkill /F /IM devenv.exe');
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM devenv.exe', '', SW_HIDE, ewWaitUntilTerminated, ExitCode);
  for Attempt := 1 to 20 do  { poll up to ~5s for the process (and its handles) to release }
  begin
    if not IsVisualStudioRunning() then
    begin
      Result := True;
      Sleep(500);  { small grace so the OS releases the extension file handles }
      Exit;
    end;
    Sleep(250);
  end;
  Result := not IsVisualStudioRunning();
end;

// Clear the *background* processes that also block VSIXInstaller's silent install but
// hold no user data: MSBuild.exe (node-reuse build workers), VBCSCompiler.exe (the Roslyn
// compiler server) and the ServiceHub.* hosts. Repeated builds can leave dozens of zombie
// MSBuild.exe alive; VSIXInstaller's BlockingProcessBlocker then fails every /install with
// exit 2004 - this was the real reason the extension appeared to "not update". These are
// safe to kill unconditionally (unlike devenv, which may hold unsaved work).
procedure KillVsBackgroundBlockers();
var
  ExitCode: Integer;
begin
  Log('KillVsBackgroundBlockers: clearing MSBuild/VBCSCompiler/ServiceHub before VSIXInstaller.');
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM MSBuild.exe', '', SW_HIDE, ewWaitUntilTerminated, ExitCode);
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM VBCSCompiler.exe', '', SW_HIDE, ewWaitUntilTerminated, ExitCode);
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /FI "IMAGENAME eq ServiceHub.*"', '', SW_HIDE, ewWaitUntilTerminated, ExitCode);
  Sleep(500);  { let the OS release handles before VSIXInstaller enumerates blockers }
end;

// The VSIX install path no longer passes /shutdownprocesses, so the VSIXInstaller blockers
// must be cleared here. Visual Studio (devenv.exe) may hold unsaved work, so it is only
// warned about - force-close is an explicit button; Cancel simply skips the extension (it
// can be installed later from {app}\vsix). The background build blockers, which have no
// user data, are killed automatically once devenv is gone. Returns True to proceed with
// the extension install, False to skip it.
function EnsureVisualStudioClosedForVsix(): Boolean;
var
  Response: Integer;
begin
  Result := True;
  while IsVisualStudioRunning() do
  begin
    { A leftover zombie devenv.exe (no window) has no unsaved work, so kill it without
      prompting - this is the case that kept silently failing when the user "closed VS"
      but a background devenv lingered. Only a REAL, in-use Visual Studio (with a window)
      gets the warning. Silent installs always force-close. }
    if WizardSilent() or (not RealVisualStudioRunning()) then
    begin
      Log('EnsureVisualStudioClosedForVsix: no in-use Visual Studio window; clearing leftover devenv.exe.');
      if not ForceCloseVisualStudio() then
      begin
        Log('EnsureVisualStudioClosedForVsix: devenv.exe would not terminate; skipping extension install.');
        Result := False;
        Exit;
      end;
      Break;  { devenv cleared; proceed to background-blocker cleanup below }
    end;

    { Real Visual Studio is open. Plain MsgBox with known return codes (IDYES/IDNO/IDCANCEL);
      YES is the default so an auto-dismissed dialog still closes VS and proceeds rather than
      silently skipping. }
    Response := MsgBox(
      'Visual Studio is open, and the VS IDE Bridge extension cannot be updated while it is'
      + ' running.'#13#10#13#10 +
      'YES  -  close Visual Studio now and continue installing the extension.'#13#10 +
      'NO  -  I will save my work and close it myself; check again.'#13#10 +
      'CANCEL  -  skip the extension for now (you can install it later from '
      + ExpandConstant('{app}') + '\vsix).',
      mbConfirmation, MB_YESNOCANCEL);

    if Response = IDYES then
    begin
      Log('EnsureVisualStudioClosedForVsix: force-closing Visual Studio (YES).');
      if not ForceCloseVisualStudio() then
        Log('EnsureVisualStudioClosedForVsix: devenv.exe still present after force-close; will re-prompt.');
      { loop re-checks; if it is gone the while exits and we proceed }
    end
    else if Response = IDNO then
    begin
      { Retry: fall through and re-check. }
    end
    else
    begin
      { CANCEL: leave VS alone and skip the VSIX step. }
      Log('EnsureVisualStudioClosedForVsix: user cancelled; skipping the extension install.');
      Result := False;
      Exit;
    end;
  end;

  { devenv is gone (or was never running / was force-closed); clear the background build
    processes that also block VSIXInstaller before the caller runs /install. }
  KillVsBackgroundBlockers();
end;

function GetLegacyVsixUninstallParameters(): string;
begin
  Result := GetVsixLogFileArgument('vsix-uninstall-legacy.log') + ' /uninstall:{#LegacyVsixId}';
end;

function GetCurrentVsixUninstallParameters(): string;
begin
  Result := GetVsixLogFileArgument('vsix-uninstall.log') + ' /uninstall:{#VsixId}';
end;

function GetVsixInstallParameters(): string;
begin
  Result := GetVsixLogFileArgument('vsix-install.log') + ' "' + ExpandConstant('{app}\vsix\VsIdeBridge.vsix') + '"';
end;

function RemoveQuotes(const Value: string): string;
begin
  Result := Value;
  if (Length(Result) >= 2) and (Result[1] = '"') and (Result[Length(Result)] = '"') then
  begin
    Delete(Result, Length(Result), 1);
    Delete(Result, 1, 1);
  end;
end;

function GetInstalledUninstallString(): string;
var
  KeyPath: string;
  UninstallString: string;
begin
  KeyPath :=
    'Software\Microsoft\Windows\CurrentVersion\Uninstall\{#emit SetupSetting("AppId")}_is1';
  UninstallString := '';

  if not RegQueryStringValue(HKLM, KeyPath, 'UninstallString', UninstallString) then
    RegQueryStringValue(HKCU, KeyPath, 'UninstallString', UninstallString);

  Result := UninstallString;
end;

function UninstallOldVersionIfPresent(): Boolean;
var
  UninstallCommand: string;
  ExitCode: Integer;
begin
  Result := True;
  UninstallCommand := GetInstalledUninstallString();
  if UninstallCommand = '' then
    Exit;

  UninstallCommand := RemoveQuotes(UninstallCommand);
  Log('Previous version detected. Running uninstall command: ' + UninstallCommand);

  if not Exec(
    UninstallCommand,
    '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART',
    '',
    SW_HIDE,
    ewWaitUntilTerminated,
    ExitCode) then
  begin
    Log('Failed to launch previous uninstaller.');
    Result := False;
    Exit;
  end;

  if ExitCode <> 0 then
  begin
    Log(Format('Previous uninstaller exited with code %d.', [ExitCode]));
    Result := False;
    Exit;
  end;

  Log('Previous version uninstall completed successfully.');
end;

function InitializeSetup(): Boolean;
begin
  Result := UninstallOldVersionIfPresent();
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  ExitCode: Integer;
begin
  Result := '';
  // Stop the service gracefully via SCM first so it cannot be auto-restarted
  // while Inno Setup is copying the service binary.  Taskkill alone bypasses
  // the SCM and can lose the race against an immediate SCM restart.
  Log('PrepareToInstall: stopping VsIdeBridgeService via sc.exe...');
  Exec(ExpandConstant('{sys}\sc.exe'), 'stop VsIdeBridgeService', '', SW_HIDE, ewWaitUntilTerminated, ExitCode);
  Log(Format('sc stop VsIdeBridgeService exited with code %d.', [ExitCode]));
  Sleep(2000);  { Give the service time to drain and release its file handles. }

  Log('PrepareToInstall: killing residual VsIdeBridgeService.exe processes...');
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM VsIdeBridgeService.exe', '', SW_HIDE, ewWaitUntilTerminated, ExitCode);
  Log(Format('taskkill VsIdeBridgeService.exe exited with code %d.', [ExitCode]));
  Sleep(500);  { Brief pause before Inno Setup proceeds with file copies. }
end;


procedure ShowVsixInstallerMissingMessage();
begin
  Log('VSIXInstaller.exe not found; VSIX install skipped.');
  MsgBox(
    'Visual Studio VSIXInstaller.exe was not found by vswhere or the fallback Program Files scan. The VSIX step was skipped.' + InstallerLineBreak +
    'You can install {app}\vsix\VsIdeBridge.vsix manually later.',
    mbInformation,
    MB_OK);
end;

function GetPostInstallStepCount(const HasVsixInstallerPath: Boolean): Integer;
begin
  Result := 2;

  if WizardIsTaskSelected('service') then
    Result := Result + 6;  { stop, delete, create, description, failure, failureflag }

  if WizardIsTaskSelected('startservice') then
    Result := Result + 1;

  if HasVsixInstallerPath then
    Result := Result + 1;  { single VSIXInstaller /install (the old version's uninstaller already removed the previous VSIX) }
end;

procedure UpdatePostInstallProgress(const Position, Max: Integer; const Status, SubStatus: string);
begin
  PostInstallProgressPage.SetText(Status, SubStatus);

  if Max > 0 then
    PostInstallProgressPage.SetProgress(Position, Max);
end;

function RunInstallerCommand(
  const Status,
  SubStatus,
  Filename,
  Parameters: string;
  const Required: Boolean): Boolean;
var
  ExitCode: Integer;
begin
  Result := True;
  Log(Status + ' [' + SubStatus + ']: ' + Filename + ' ' + Parameters);

  if not Exec(Filename, Parameters, '', SW_HIDE, ewWaitUntilTerminated, ExitCode) then
  begin
    Log(Format('%s failed to launch. Result code %d.', [Status, ExitCode]));
    Result := False;
    if Required then
      RaiseException(Format('%s failed to launch. Result code %d.', [Status, ExitCode]));
    Exit;
  end;

  if ExitCode <> 0 then
  begin
    Log(Format('%s exited with code %d.', [Status, ExitCode]));
    Result := False;
    if Required then
      RaiseException(Format('%s failed with exit code %d.', [Status, ExitCode]));
  end;
end;

procedure RunPostInstallStep(
  var StepIndex: Integer;
  const TotalSteps: Integer;
  const Status,
  SubStatus,
  Filename,
  Parameters: string;
  const Required: Boolean);
begin
  UpdatePostInstallProgress(StepIndex, TotalSteps, Status, SubStatus);
  RunInstallerCommand(Status, SubStatus, Filename, Parameters, Required);
  StepIndex := StepIndex + 1;
  UpdatePostInstallProgress(StepIndex, TotalSteps, Status, SubStatus);
end;

procedure RunPythonRuntimeSetupStep(var StepIndex: Integer; const TotalSteps: Integer);
begin
  UpdatePostInstallProgress(
    StepIndex,
    TotalSteps,
    'Configuring bridge Python runtime...',
    GetSelectedPythonProvisioningMode());
  WritePythonRuntimeConfig();
  StepIndex := StepIndex + 1;
  UpdatePostInstallProgress(
    StepIndex,
    TotalSteps,
    'Configuring bridge Python runtime...',
    GetSelectedPythonProvisioningMode());
end;

procedure RunHttpServerSetupStep(var StepIndex: Integer; const TotalSteps: Integer);
var
  HttpMode: string;
begin
  if WizardIsTaskSelected('enablehttp') then
    HttpMode := 'enabled'
  else
    HttpMode := 'disabled';

  UpdatePostInstallProgress(
    StepIndex,
    TotalSteps,
    'Configuring MCP HTTP endpoints...',
    HttpMode);
  WriteHttpServerConfig();
  StepIndex := StepIndex + 1;
  UpdatePostInstallProgress(
    StepIndex,
    TotalSteps,
    'Configuring MCP HTTP endpoints...',
    HttpMode);
end;

procedure RunPostInstallActions();
var
  ScPath: string;
  StepIndex: Integer;
  TotalSteps: Integer;
  VsixInstallerPath: string;
  VsixAllowed: Boolean;
begin
  ScPath := ExpandConstant('{sys}\sc.exe');
  VsixInstallerPath := ResolveVsixInstallerPath();

  { Resolve whether the VS extension can be installed BEFORE the progress page is shown -
    a MsgBox displayed on top of the TOutputProgressWizardPage can lose modality and be
    auto-dismissed. This also clears the background build blockers up front. }
  VsixAllowed := True;
  if VsixInstallerPath <> '' then
    VsixAllowed := EnsureVisualStudioClosedForVsix();

  TotalSteps := GetPostInstallStepCount(VsixInstallerPath <> '');
  StepIndex := 0;

  if TotalSteps > 0 then
  begin
    UpdatePostInstallProgress(0, TotalSteps, PostInstallPageTitle, PostInstallPageDescription);
    PostInstallProgressPage.Show;
  end;

  try
    RunPythonRuntimeSetupStep(StepIndex, TotalSteps);
    RunHttpServerSetupStep(StepIndex, TotalSteps);

    if WizardIsTaskSelected('service') then
    begin
      RunPostInstallStep(
        StepIndex,
        TotalSteps,
        'Stopping previous VS IDE Bridge service...',
        'sc stop',
        ScPath,
        GetServiceCommandParameters('stop'),
        False);
      RunPostInstallStep(
        StepIndex,
        TotalSteps,
        'Removing previous VS IDE Bridge service registration...',
        'sc delete',
        ScPath,
        GetServiceCommandParameters('delete'),
        False);
      RunPostInstallStep(
        StepIndex,
        TotalSteps,
        'Registering VS IDE Bridge service...',
        'sc create',
        ScPath,
        GetServiceCreateParameters(),
        True);
      RunPostInstallStep(
        StepIndex,
        TotalSteps,
        'Configuring VS IDE Bridge service...',
        'sc description',
        ScPath,
        GetServiceDescriptionParameters(),
        True);
      RunPostInstallStep(
        StepIndex,
        TotalSteps,
        'Configuring VS IDE Bridge service recovery...',
        'sc failure',
        ScPath,
        GetServiceFailureParameters(),
        False);
      RunPostInstallStep(
        StepIndex,
        TotalSteps,
        'Configuring VS IDE Bridge service recovery flag...',
        'sc failureflag',
        ScPath,
        GetServiceFailureFlagParameters(),
        False);
    end;

    if WizardIsTaskSelected('startservice') then
    begin
      RunPostInstallStep(
        StepIndex,
        TotalSteps,
        'Starting VS IDE Bridge service...',
        'sc start',
        ScPath,
        GetServiceCommandParameters('start'),
        True);
    end;

    if VsixInstallerPath <> '' then
    begin
      // InitializeSetup already ran the previous version's uninstaller (which removes the
      // old VSIX via its own [UninstallRun]), so a single /install is all that is needed
      // here - the old uninstall-current / uninstall-legacy steps were redundant and were
      // the bulk of the perceived install "hang" (three VSIXInstaller cold-starts).
      { VsixAllowed was resolved above, before the progress page was shown. }
      if VsixAllowed then
        RunPostInstallStep(
          StepIndex,
          TotalSteps,
          'Installing VS IDE Bridge extension into Visual Studio...',
          'VSIXInstaller /install',
          VsixInstallerPath,
          GetVsixInstallParameters(),
          True)
      else
      begin
        VsixInstallSkipped := True;
        Log('RunPostInstallActions: VSIX install skipped at user request (Visual Studio left open).');
      end;
    end
    else
    begin
      ShowVsixInstallerMissingMessage();
    end;

    if TotalSteps > 0 then
      UpdatePostInstallProgress(TotalSteps, TotalSteps, 'VS IDE Bridge setup completed.', '');
  finally
    if TotalSteps > 0 then
      PostInstallProgressPage.Hide;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if (CurStep = ssPostInstall) and (not PostInstallCompleted) then
  begin
    RunPostInstallActions();
    PostInstallCompleted := True;
  end;
end;

procedure CurPageChanged(CurPageID: Integer);
begin
  if CurPageID = wpFinished then
  begin
    if VsixInstallSkipped then
      WizardForm.FinishedLabel.Caption := WizardForm.FinishedLabel.Caption + InstallerDoubleLineBreak +
        'Visual Studio extension install was skipped because Visual Studio was left open. '
        + 'Install it later by closing Visual Studio and running {app}\vsix\VsIdeBridge.vsix.'
    else if HasVsixInstaller() then
      WizardForm.FinishedLabel.Caption := WizardForm.FinishedLabel.Caption + InstallerDoubleLineBreak +
        'Visual Studio extension installed: {#VsixId}.'
    else
      WizardForm.FinishedLabel.Caption := WizardForm.FinishedLabel.Caption + InstallerDoubleLineBreak +
        'Visual Studio extension install was skipped (VSIXInstaller.exe not found).';
  end;
end;












