; Inno Setup Script for АЛЬВАРЕС AI

[Setup]
AppName=АЛЬВАРЕС AI
AppVersion=1.0
AppPublisher=12 штурмова рота
DefaultDirName={autopf}\AlvaresAI
DefaultGroupName=АЛЬВАРЕС AI
OutputBaseFilename=AlvaresAI_Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest

[Languages]
Name: "ukrainian"; MessagesFile: "compiler:Languages\Ukrainian.isl"

[Files]
; Вся папка dist\Alvares
Source: "dist\Alvares\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; Робочі файли (onlyifdoesntexist — щоб не затерти дані при оновленні)
Source: "Табель_Багатомісячний.xlsx"; DestDir: "{app}"; Flags: ignoreversion onlyifdoesntexist
Source: "BR_4ShB.xlsx"; DestDir: "{app}"; Flags: ignoreversion onlyifdoesntexist
Source: "app.db"; DestDir: "{app}"; Flags: ignoreversion

[Dirs]
Name: "{app}\output"

[Icons]
Name: "{group}\АЛЬВАРЕС AI"; Filename: "{app}\Alvares.exe"
Name: "{autodesktop}\АЛЬВАРЕС AI"; Filename: "{app}\Alvares.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Створити ярлик на робочому столі"; GroupDescription: "Додаткові ярлики:"

[Run]
Filename: "{app}\Alvares.exe"; Description: "Запустити АЛЬВАРЕС AI"; Flags: nowait postinstall skipifsilent
