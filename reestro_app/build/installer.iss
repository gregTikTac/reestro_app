; Inno Setup — установщик ЕГРН-Парсера.
;
; ПЕРЕД компиляцией обязательно:
;   1) build\build_exe.bat  (создаст dist\ЕГРН-Парсер\ с _internal\python311.dll)
;   2) python build\verify_dist.py  (проверка — должно быть OK)
;   3) Только потом Compile этот .iss
;
; Если dist пустой — установщик соберётся битым!

#define AppName "ЕГРН-Парсер"
#define AppVersion "0.1.3"
#define AppExe "ЕГРН-Парсер.exe"

[Setup]
AppName={#AppName}
AppVersion={#AppVersion}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
OutputDir=..\dist_installer
OutputBaseFilename=ЕГРН-Парсер_Setup_{#AppVersion}
Compression=lzma2
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=lowest
WizardStyle=modern

[Languages]
Name: "ru"; MessagesFile: "compiler:Languages\Russian.isl"

[Files]
; Вся папка сборки PyInstaller (one-folder)
Source: "..\dist\{#AppName}\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Создать ярлык на рабочем столе"; GroupDescription: "Дополнительно:"

[Run]
Filename: "{app}\{#AppExe}"; Description: "Запустить {#AppName}"; Flags: nowait postinstall skipifsilent
