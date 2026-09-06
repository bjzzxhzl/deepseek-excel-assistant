; ExcelAI - Inno Setup 安装脚本
; 编译: ISCC.exe setup.iss
#define MyAppName "ExcelAI"
#define MyAppVersion "0.41"
#define MyAppPublisher "DSH"
#define MyAppId "7c1e9a24-5d3f-4b8e-9c2a-0f6d1b4e8a5c"
#define LocalAddinId "1537f254-10aa-41d5-aed2-0a00b89da104"

[Setup]
AppId={{7c1e9a24-5d3f-4b8e-9c2a-0f6d1b4e8a5c}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
; 保留旧安装目录与注册键，确保现有 DeepSeek Excel Assistant 安装可原位升级。
DefaultDirName={localappdata}\DeepSeekExcelAssistant
DefaultGroupName={#MyAppName}
PrivilegesRequired=lowest
OutputDir=dist
OutputBaseFilename=ExcelAI-Standalone-Setup-v0.41
Compression=lzma2
SolidCompression=yes
SetupIconFile=app.ico
UninstallDisplayIcon={app}\app.ico
WizardStyle=modern
DisableProgramGroupPage=yes
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "taskpane.html"; DestDir: "{app}"; Flags: ignoreversion
Source: "app.js"; DestDir: "{app}"; Flags: ignoreversion
Source: "app.css"; DestDir: "{app}"; Flags: ignoreversion
Source: "manifest.xml"; DestDir: "{app}"; Flags: ignoreversion
Source: "manifest-standalone.xml"; DestDir: "{app}"; Flags: ignoreversion
Source: "server.mjs"; DestDir: "{app}"; Flags: ignoreversion
Source: "icon16.png"; DestDir: "{app}"; Flags: ignoreversion
Source: "icon32.png"; DestDir: "{app}"; Flags: ignoreversion
Source: "icon64.png"; DestDir: "{app}"; Flags: ignoreversion
Source: "icon80.png"; DestDir: "{app}"; Flags: ignoreversion
Source: "app.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "icon-master-ai.png"; DestDir: "{app}"; Flags: ignoreversion
Source: "DeepSeekExcelAssistant-Standalone-Autoload.xlsm"; DestDir: "{app}\carrier"; Flags: ignoreversion
Source: "configure-autoload.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "start-server.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "start-server.vbs"; DestDir: "{app}"; Flags: ignoreversion
Source: "start-all.vbs"; DestDir: "{app}"; Flags: ignoreversion
Source: "uninstall.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "CHANGELOG.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "STORE-SUBMISSION.md"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEF\Developer"; ValueType: string; ValueName: "{#LocalAddinId}"; ValueData: "{app}\manifest-standalone.xml"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\DeepSeekExcelAssistant"; ValueType: string; ValueName: "Path"; ValueData: "{app}\carrier\"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\DeepSeekExcelAssistant"; ValueType: dword; ValueName: "AllowSubfolders"; ValueData: "0"
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\DeepSeekExcelAssistant"; ValueType: string; ValueName: "Description"; ValueData: "ExcelAI startup carrier"

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{sys}\wscript.exe"; Parameters: """{app}\start-all.vbs"""; WorkingDir: "{app}"; IconFilename: "{app}\app.ico"; Comment: "启动 ExcelAI"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{sys}\wscript.exe"; Parameters: """{app}\start-all.vbs"""; WorkingDir: "{app}"; IconFilename: "{app}\app.ico"; Comment: "启动 ExcelAI"

[UninstallDelete]
Type: files; Name: "{userappdata}\Microsoft\Excel\XLSTART\DeepSeekExcelAssistant-Autoload.xlsm"
Type: files; Name: "{userappdata}\Microsoft\Excel\XLSTART\DeepSeekExcelAssistant-Autoload.xlsx"
Type: filesandordirs; Name: "{app}"

[InstallDelete]
Type: files; Name: "{userdesktop}\DeepSeek Excel 助手.lnk"
Type: files; Name: "{userdesktop}\DeepSeek Excel Assistant.lnk"
Type: files; Name: "{userprograms}\DeepSeek Excel 助手.lnk"
Type: files; Name: "{userprograms}\DeepSeek Excel Assistant.lnk"
Type: files; Name: "{userappdata}\Microsoft\Excel\XLSTART\DeepSeekExcelAssistant-Autoload.xlsx"
Type: files; Name: "{userappdata}\Microsoft\Excel\XLSTART\DeepSeekExcelAssistant-Autoload.xlsm"
Type: files; Name: "{app}\carrier\DeepSeekExcelAssistant-Autoload.xlsm"

[Code]
const
  AddinId = '{#LocalAddinId}';
  LegacyDeveloperAddinId = '{#MyAppId}';
  DeveloperKey = 'Software\Microsoft\Office\16.0\WEF\Developer';
  ValidationCacheKey = 'Software\Microsoft\Office\16.0\Common\CustomUIValidationCache';
  ExcelOptionsKey = 'Software\Microsoft\Office\16.0\Excel\Options';
  InstallerStateKey = 'Software\DeepSeekExcelAssistant';

function CarrierOpenArgument: String;
begin
  Result := '/r "' + ExpandConstant('{app}\carrier\DeepSeekExcelAssistant-Standalone-Autoload.xlsm') + '"';
end;

procedure RemoveCarrierStartupRegistration;
var
  ValueName, StoredArgument, CurrentArgument: String;
begin
  if RegQueryStringValue(HKCU, InstallerStateKey, 'ExcelStartupValue', ValueName) and
     RegQueryStringValue(HKCU, InstallerStateKey, 'ExcelStartupArgument', StoredArgument) then
  begin
    if RegQueryStringValue(HKCU, ExcelOptionsKey, ValueName, CurrentArgument) and
       (CompareText(CurrentArgument, StoredArgument) = 0) then
      RegDeleteValue(HKCU, ExcelOptionsKey, ValueName);
  end;
  RegDeleteValue(HKCU, InstallerStateKey, 'ExcelStartupValue');
  RegDeleteValue(HKCU, InstallerStateKey, 'ExcelStartupArgument');
end;

procedure RegisterCarrierStartup;
var
  I: Integer;
  ValueName, CurrentArgument, ArgumentValue: String;
  Registered: Boolean;
begin
  ArgumentValue := CarrierOpenArgument;
  Registered := False;
  for I := 0 to 999 do
  begin
    if I = 0 then
      ValueName := 'OPEN'
    else
      ValueName := 'OPEN' + IntToStr(I);

    if not RegValueExists(HKCU, ExcelOptionsKey, ValueName) then
    begin
      if RegWriteStringValue(HKCU, ExcelOptionsKey, ValueName, ArgumentValue) then
        Registered := True;
    end
    else if RegQueryStringValue(HKCU, ExcelOptionsKey, ValueName, CurrentArgument) and
            (CompareText(CurrentArgument, ArgumentValue) = 0) then
      Registered := True;

    if Registered then
      Break;
  end;

  if not Registered then
    RaiseException('无法为 Excel 分配启动载体注册项。');

  RegWriteStringValue(HKCU, InstallerStateKey, 'ExcelStartupValue', ValueName);
  RegWriteStringValue(HKCU, InstallerStateKey, 'ExcelStartupArgument', ArgumentValue);
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Result := '';
  if FindWindowByClassName('XLMAIN') <> 0 then
    Result := '请先关闭所有 Excel 窗口，然后重试安装。';
end;

procedure ClearAddinValidationCache(IncludeLegacy: Boolean);
var
  ValueNames: TArrayOfString;
  I: Integer;
begin
  if RegGetValueNames(HKCU, ValidationCacheKey, ValueNames) then
  begin
    for I := 0 to GetArrayLength(ValueNames) - 1 do
      if (Pos(LowerCase(AddinId), LowerCase(ValueNames[I])) = 1) or
         (IncludeLegacy and
          (Pos(LowerCase(LegacyDeveloperAddinId), LowerCase(ValueNames[I])) = 1)) then
        RegDeleteValue(HKCU, ValidationCacheKey, ValueNames[I]);
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    RemoveCarrierStartupRegistration;
    { Migrate old EXE sideload registration. Microsoft 365 centralized
      deployment is not stored under WEF\Developer and is not removed. }
    RegDeleteValue(HKCU, DeveloperKey, LegacyDeveloperAddinId);
    RegisterCarrierStartup;
    ClearAddinValidationCache(True);
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
  begin
    RemoveCarrierStartupRegistration;
    RegDeleteKeyIncludingSubkeys(HKCU, InstallerStateKey);
    RegDeleteValue(HKCU, DeveloperKey, AddinId);
    ClearAddinValidationCache(False);
  end;
end;
