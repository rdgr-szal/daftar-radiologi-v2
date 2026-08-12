; Inno Setup Script for Daftar Radiologi
; Compatible with Inno Setup 6.x

#define MyAppName "Daftar Radiologi"
#ifndef MyAppVersion
  #define MyAppVersion "2.0.0"
#endif
#define MyAppPublisher "Jabatan Radiologi"
#define MyAppExeName "DaftarRadiologi.exe"
#define MyAppAppId "C6B82C94-8120-4F25-B470-3E6B47C590E2"

[Setup]
AppId={{C6B82C94-8120-4F25-B470-3E6B47C590E2}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=dist
OutputBaseFilename=DaftarRadiologi_v{#MyAppVersion}_Setup
SetupIconFile=Daftar_Radiologi\static\img\favicon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Copy all build files from PyInstaller output folder
Source: "Daftar_Radiologi\dist\DaftarRadiologi\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; Optional: Embed WebView2 Bootstrapper installer if available in project root
Source: "MicrosoftEdgeWebview2Setup.exe"; DestDir: "{app}"; Flags: ignoreversion skipifnotsilent; Check: FileExistsInSrc('MicrosoftEdgeWebview2Setup.exe')

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; Auto-run embedded WebView2 Setup silently if included
Filename: "{app}\MicrosoftEdgeWebview2Setup.exe"; Parameters: "/silent /install"; StatusMsg: "Installing Microsoft Edge WebView2 Runtime..."; Flags: waituntilterminated; Check: FileExists(ExpandConstant('{app}\MicrosoftEdgeWebview2Setup.exe'))
; Option to launch app after installation
Filename: "{app}\{#MyAppName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Dirs]
; Ensure Pendaftaran data folder exists and is strictly protected on uninstall
Name: "{app}\Pendaftaran"; Flags: uninsneveruninstall

[UninstallDelete]
; Protect Pendaftaran folder and Excel files from deletion during uninstall
Type: filesandordirs; Name: "{app}\_internal"
Type: filesandordirs; Name: "{app}\templates"
Type: filesandordirs; Name: "{app}\static"

[Code]
function FileExistsInSrc(const FileName: String): Boolean;
begin
  Result := FileExists(ExpandConstant('{src}\' + FileName)) or FileExists(FileName);
end;

var
  OptionPage: TWizardPage;
  RadioUpgrade: TRadioButton;
  RadioReinstall: TRadioButton;
  RadioUninstall: TRadioButton;
  IsExistingInstallation: Boolean;

function IsAppInstalled(): Boolean;
var
  RegKey: String;
begin
  RegKey := 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{' + '{#MyAppAppId}' + '}_is1';
  Result := RegValueExists(HKEY_LOCAL_MACHINE, RegKey, 'UninstallString') or
            RegValueExists(HKEY_CURRENT_USER, RegKey, 'UninstallString');
end;

procedure InitializeWizard();
var
  Lbl: TLabel;
begin
  IsExistingInstallation := IsAppInstalled();

  if IsExistingInstallation then
  begin
    OptionPage := CreateCustomPage(wpWelcome, 'Existing Installation Detected', 
      'An existing version of Daftar Radiologi was found on this system. Please choose an action:');

    RadioUpgrade := TRadioButton.Create(OptionPage);
    RadioUpgrade.Parent := OptionPage.Surface;
    RadioUpgrade.Caption := 'Upgrade / Update (Recommended)';
    RadioUpgrade.Font.Style := [fsBold];
    RadioUpgrade.Left := ScaleX(15);
    RadioUpgrade.Top := ScaleY(15);
    RadioUpgrade.Width := ScaleX(400);
    RadioUpgrade.Checked := True;

    Lbl := TLabel.Create(OptionPage);
    Lbl.Parent := OptionPage.Surface;
    Lbl.Caption := 'Updates application files while strictly preserving patient records (Pendaftaran folder) and clinic settings.';
    Lbl.Left := ScaleX(35);
    Lbl.Top := ScaleY(35);
    Lbl.Width := ScaleX(410);
    Lbl.WordWrap := True;

    RadioReinstall := TRadioButton.Create(OptionPage);
    RadioReinstall.Parent := OptionPage.Surface;
    RadioReinstall.Caption := 'Fresh Install / Repair';
    RadioReinstall.Font.Style := [fsBold];
    RadioReinstall.Left := ScaleX(15);
    RadioReinstall.Top := ScaleY(75);
    RadioReinstall.Width := ScaleX(400);

    Lbl := TLabel.Create(OptionPage);
    Lbl.Parent := OptionPage.Surface;
    Lbl.Caption := 'Overwrites and repairs all core application files if you are experiencing launch issues.';
    Lbl.Left := ScaleX(35);
    Lbl.Top := ScaleY(95);
    Lbl.Width := ScaleX(410);
    Lbl.WordWrap := True;

    RadioUninstall := TRadioButton.Create(OptionPage);
    RadioUninstall.Parent := OptionPage.Surface;
    RadioUninstall.Caption := 'Uninstall Existing Installation';
    RadioUninstall.Font.Style := [fsBold];
    RadioUninstall.Left := ScaleX(15);
    RadioUninstall.Top := ScaleY(135);
    RadioUninstall.Width := ScaleX(400);

    Lbl := TLabel.Create(OptionPage);
    Lbl.Parent := OptionPage.Surface;
    Lbl.Caption := 'Launches the uninstaller to remove Daftar Radiologi from your computer.';
    Lbl.Left := ScaleX(35);
    Lbl.Top := ScaleY(155);
    Lbl.Width := ScaleX(410);
    Lbl.WordWrap := True;
  end;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var
  UninstallerPath: String;
  ResultCode: Integer;
begin
  Result := True;
  if IsExistingInstallation and (OptionPage <> nil) and (CurPageID = OptionPage.ID) then
  begin
    if RadioUninstall.Checked then
    begin
      if MsgBox('Are you sure you want to uninstall Daftar Radiologi? Your patient records in Pendaftaran folder will be kept safe.', 
        mbConfirmation, MB_YESNO) = IDYES then
      begin
        RegQueryStringValue(HKEY_LOCAL_MACHINE, 
          'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{' + '{#MyAppAppId}' + '}_is1', 
          'UninstallString', UninstallerPath);
        if UninstallerPath = '' then
          RegQueryStringValue(HKEY_CURRENT_USER, 
            'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{' + '{#MyAppAppId}' + '}_is1', 
            'UninstallString', UninstallerPath);

        if UninstallerPath <> '' then
        begin
          Exec(RemoveQuotes(UninstallerPath), '', '', SW_SHOWNORMAL, ewWaitUntilTerminated, ResultCode);
        end;
      end;
      Result := False; // Abort remaining wizard setup steps
    end;
  end;
end;
