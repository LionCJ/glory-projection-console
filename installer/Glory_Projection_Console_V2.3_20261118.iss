; Glory 榮耀 · 投影控制台 V2.3 安裝包（2026-11-18 實機測試）
; 使用 Inno Setup 6.7 編譯
; 安裝來源：..\dist\Glory_榮耀_執行版_V2.3_20261118（自包含 publish 輸出）
; 請勿把舊 dist 或 obj_corrupt_* 掃進安裝包

#define MyAppName "Glory 榮耀 · 投影控制台"
#define MyAppVersion "2.3.0"
#define MyAppPublisher "Glory 榮耀"
#define MyAppExeName "ProjectionManager.exe"
#define MyAppDirName "Glory 投影控制台"

[Setup]
AppId={{0DF6D9B6-3EE7-4E4E-BEE8-814DADF1E284}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} V{#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppCopyright=Copyright (C) 2026 Glory 榮耀
DefaultDirName={autopf}\{#MyAppDirName}
DefaultGroupName=Glory 榮耀
DisableProgramGroupPage=yes
OutputDir=..\dist\installer
OutputBaseFilename=Glory_榮耀_投影控制台_V2.3_安裝包_20261118
SetupIconFile=..\Assets\glory.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0.17763
CloseApplications=yes
RestartApplications=no
UsePreviousAppDir=yes
AllowNoIcons=yes
InfoBeforeFile=..\docs\快速開始.txt
VersionInfoVersion=2.3.0.0
VersionInfoCompany=Glory 榮耀
VersionInfoDescription=Glory 榮耀投影控制台 V2.3 安裝包（2026-11-18）
VersionInfoProductName=Glory 榮耀 · 投影控制台
VersionInfoProductVersion=2.3.0
VersionInfoCopyright=Copyright (C) 2026 Glory 榮耀
SetupLogging=yes

[Languages]
Name: "chinesetraditional"; MessagesFile: "Languages\ChineseTraditional.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; 主程式與自包含 .NET 8 執行環境
Source: "..\dist\Glory_榮耀_執行版_V2.3_20261118\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; 操作手冊與快速開始
Source: "..\docs\Glory_榮耀_投影控制台_操作手冊_V2.3.docx"; DestDir: "{app}\docs"; Flags: ignoreversion
Source: "..\docs\Glory_榮耀_投影控制台_操作手冊_V2.3.pdf"; DestDir: "{app}\docs"; Flags: ignoreversion
Source: "..\docs\快速開始.txt"; DestDir: "{app}"; Flags: ignoreversion
; WebView2 完整離線安裝程式（不複製到安裝目錄，僅在需要時解壓執行）
Source: "prereq\MicrosoftEdgeWebView2RuntimeInstallerX64.exe"; DestDir: "{tmp}"; Flags: dontcopy

[Icons]
Name: "{group}\Glory 榮耀 · 投影控制台"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\操作手冊（PDF）"; Filename: "{app}\docs\Glory_榮耀_投影控制台_操作手冊_V2.3.pdf"
Name: "{autodesktop}\Glory 榮耀 · 投影控制台"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,Glory 榮耀 · 投影控制台}"; Flags: nowait postinstall skipifsilent

[Code]
const
  WebView2RegKey = 'SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}';
  WebView2RegKey64 = 'SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}';

function WebView2Installed: Boolean;
var
  Version: String;
begin
  Result := RegQueryStringValue(HKLM64, WebView2RegKey, 'pv', Version) and (Version <> '');
  if not Result then
    Result := RegQueryStringValue(HKLM64, WebView2RegKey64, 'pv', Version) and (Version <> '');
  if not Result then
    Result := RegQueryStringValue(HKCU, WebView2RegKey, 'pv', Version) and (Version <> '');
end;

function PowerPointInstalled: Boolean;
var
  AppPath: String;
begin
  Result := RegQueryStringValue(HKLM64, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE', '', AppPath) and FileExists(AppPath);
  if not Result then
    Result := RegQueryStringValue(HKLM, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE', '', AppPath) and FileExists(AppPath);
end;

function InitializeSetup: Boolean;
begin
  Result := True;
  if not PowerPointInstalled then
    MsgBox('提醒：未偵測到 Microsoft PowerPoint。' + #13#10 +
           '主 PPT 放映需要 PowerPoint 2016 或更新版本；素材、經文與網頁投影不受影響。' + #13#10 +
           '如只在實機測試素材功能，可繼續安裝。', mbInformation, MB_OK);
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
begin
  if CurStep = ssInstall then
  begin
    if not WebView2Installed then
    begin
      ExtractTemporaryFile('MicrosoftEdgeWebView2RuntimeInstallerX64.exe');
      if not Exec(ExpandConstant('{tmp}\MicrosoftEdgeWebView2RuntimeInstallerX64.exe'),
                  '/silent /install', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
      begin
        MsgBox('WebView2 執行環境安裝失敗（無法啟動安裝程式）。' + #13#10 +
               '網頁投影可能無法使用；其他功能不受影響。', mbError, MB_OK);
      end
      else if (ResultCode <> 0) and (ResultCode <> 3010) then
      begin
        MsgBox('WebView2 執行環境安裝回傳代碼 ' + IntToStr(ResultCode) + '。' + #13#10 +
               '如網頁投影空白，請手動執行安裝包內的 WebView2 安裝程式。', mbInformation, MB_OK);
      end;
    end;
  end;
end;
