; ─────────────────────────────────────────────────────────────
; 山羊联合育种管理系统 — Inno Setup 安装包脚本
; 构建顺序: pyinstaller build.spec  ->  make_cert_and_sign.ps1  ->  iscc installer.iss
; 说明:
;   - dist\SheepBreeding.exe 由 PyInstaller 生成（单文件）
;   - resources\SheepBreeding.cer 由签名脚本导出（自签名证书公钥）
;   - resources\MicrosoftEdgeWebView2Setup.exe 需手动下载放置（微软 Evergreen 引导）
;   - cert_thumb.iss 由签名脚本自动生成（写入证书指纹，供卸载时清理）
; ─────────────────────────────────────────────────────────────

#include "cert_thumb.iss"
#ifdef CertThumbprint
; 由 cert_thumb.iss 提供
#endif

#define MyAppName "山羊联合育种管理系统"
#define MyAppVersion "3.0.0"
#define MyAppPublisher "区品改站shouzhangsun"
#define MyAppURL "https://localhost"
#define MyAppExeName "SheepBreeding.exe"
; ── 以下为发布者/版本号配置区 ──────────────────────────────
; MyAppPublisher : 出现在「程序和功能」、UAC/SmartScreen 的发布者名称，
;                  需与自签名证书 CN 保持一致（见 make_cert_and_sign.ps1 的 $Publisher）。
; MyAppVersion   : 语义化版本号（主.次.修订），安装包与 EXE 文件属性均使用它。
; 修改发布者或版本号时，请同步修改 make_cert_and_sign.ps1 顶部同名变量。
; ───────────────────────────────────────────────────────────

[Setup]
; 固定 GUID，保证升级/卸载一致
AppId={{C9A1E2B3-7F4D-4C8A-9B6E-2D5F8A1C3B7E}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=installer
OutputBaseFilename=Setup_{#MyAppName}
Compression=lzma2
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64os
PrivilegesRequired=admin
UninstallDisplayIcon={app}\{#MyAppExeName}
WizardStyle=modern
LanguageDetectionMethod=uilanguage
; 文件属性中的版本信息（右键 EXE → 属性 → 详细信息）
VersionInfoVersion={#MyAppVersion}
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription={#MyAppName}
VersionInfoCopyright=Copyright (C) 2026 {#MyAppPublisher}
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion={#MyAppVersion}
; 图标（若存在则使用，否则用默认）
#ifexist "resources\app.ico"
SetupIconFile=resources\app.ico
#endif

[Languages]
Name: "chinesesimplified"; MessagesFile: "Languages\ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "快捷方式设置:"; Flags: unchecked

[Files]
; 主程序（PyInstaller 单文件产物）
Source: "dist\SheepBreeding.exe"; DestDir: "{app}"; Flags: ignoreversion
; 自签名证书公钥（若存在则打包并导入受信任根）
#ifexist "resources\SheepBreeding.cer"
Source: "resources\SheepBreeding.cer"; DestDir: "{app}"; Flags: ignoreversion
#endif
; WebView2 运行环境引导（若缺失则安装；需手动下载放入 resources）
#ifexist "resources\MicrosoftEdgeWebView2Setup.exe"
Source: "resources\MicrosoftEdgeWebView2Setup.exe"; DestDir: "{app}"; Flags: ignoreversion
#endif

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; 导入自签名证书到受信任根，消除 SmartScreen 警告（仅当 cer 存在）
#ifexist "resources\SheepBreeding.cer"
Filename: "certutil.exe"; Parameters: "-addstore -f ""Root"" ""{app}\SheepBreeding.cer"""; StatusMsg: "正在信任发布者证书…"; Flags: runhidden; Check: CertBundled
#endif
; 若系统缺少 x86 WebView2，但已存在其他架构：先卸载（清除「已安装」拦截），再装 x86
#ifexist "resources\MicrosoftEdgeWebView2Setup.exe"
Filename: "{app}\MicrosoftEdgeWebView2Setup.exe"; Parameters: "/uninstall /purge"; StatusMsg: "正在重置 WebView2 运行环境…"; Flags: runhidden; Check: NeedWebView2Reset
#endif
; 若系统缺少 x86 WebView2，则静默安装
#ifexist "resources\MicrosoftEdgeWebView2Setup.exe"
Filename: "{app}\MicrosoftEdgeWebView2Setup.exe"; Parameters: "/silent /install"; StatusMsg: "正在安装 WebView2 运行环境…"; Flags: runhidden; Check: NotWebView2Installed
#endif

[UninstallDelete]
Type: files; Name: "{app}\SheepBreeding.cer"
Type: files; Name: "{app}\MicrosoftEdgeWebView2Setup.exe"

[Code]
function CertBundled(): Boolean;
begin
  Result := FileExists(ExpandConstant('{app}\SheepBreeding.cer'));
end;

function WebView2Installed(): Boolean;
begin
  // 本程序为 32 位 exe，依赖 x86 WebView2 运行时（注册在 32 位视图 WOW6432Node）。
  // 用 HKLM32 显式查 32 位注册表视图，避免被 x64 WebView2 误判为已安装。
  Result := RegKeyExists(HKLM32, 'SOFTWARE\Microsoft\EdgeWebView\Applications')
         or RegKeyExists(HKCU, 'SOFTWARE\Microsoft\EdgeWebView\Applications');
end;

function NotWebView2Installed(): Boolean;
begin
  Result := not WebView2Installed();
end;

function AnyWebView2Installed(): Boolean;
begin
  // 任意架构的 WebView2 均已视为「已装」（用于判断是否需要先卸载再重装 x86）
  Result := RegKeyExists(HKLM64, 'SOFTWARE\Microsoft\EdgeWebView\Applications')
         or RegKeyExists(HKLM32, 'SOFTWARE\Microsoft\EdgeWebView\Applications')
         or RegKeyExists(HKCU, 'SOFTWARE\Microsoft\EdgeWebView\Applications');
end;

function NeedWebView2Reset(): Boolean;
begin
  // x86 缺失但系统已有其他架构 WebView2：需先卸载（清除「已安装」拦截）再装 x86
  Result := (not WebView2Installed()) and AnyWebView2Installed();
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  Res: Integer;
begin
  if CurUninstallStep = usPostUninstall then
  begin
    // 卸载时按指纹移除导入的根证书（指纹为空时不动作）
    #ifdef CertThumbprint
    if '{#CertThumbprint}' <> '' then
      Exec('certutil.exe', '-delstore -f Root {#CertThumbprint}', '', SW_HIDE, ewWaitUntilTerminated, Res);
    #endif
  end;
end;
