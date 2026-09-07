; Inno Setup 安装包脚本：把 dist/FileConverter/（onedir 双 exe）打成 setup.exe。
;
; 构建（项目根目录执行）：
;   pyinstaller packaging/FileConverter.spec --noconfirm --clean
;   iscc /DAppVersion=1.0.0 packaging/installer.iss
; 输出：dist/FileConverter-<版本>-windows-x64-setup.exe
;
; 免管理员安装（PrivilegesRequired=lowest），装到 %LOCALAPPDATA%\Programs\FileConverter，
; 开始菜单快捷方式 + 可选桌面快捷方式，自带卸载程序。

#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif
; AppTag 用于输出文件名（如 v1.0.0），与 zip 产物命名保持一致
#ifndef AppTag
  #define AppTag AppVersion
#endif

[Setup]
AppId={{7C4E9A2B-5D81-4F3A-A6C2-9E0B1D8F3A57}
AppName=FileConverter
AppVersion={#AppVersion}
AppVerName=FileConverter {#AppVersion}
DefaultDirName={autopf}\FileConverter
DefaultGroupName=FileConverter
OutputDir=..\dist
OutputBaseFilename=FileConverter-{#AppTag}-windows-x64-setup
SetupIconFile=..\file_converter\assets\icon.ico
UninstallDisplayIcon={app}\FileConverter.exe
Compression=lzma2/ultra64
SolidCompression=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
WizardStyle=modern

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式 (&D)"; GroupDescription: "附加任务:"

[Files]
Source: "..\dist\FileConverter\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\FileConverter"; Filename: "{app}\FileConverter.exe"
Name: "{autodesktop}\FileConverter"; Filename: "{app}\FileConverter.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\FileConverter.exe"; Description: "启动 FileConverter (&L)"; Flags: nowait postinstall skipifsilent
