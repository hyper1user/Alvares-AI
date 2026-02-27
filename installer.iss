; Inno Setup Script for АЛЬВАРЕС AI

[Setup]
AppId={{alvares-ai-app}
AppName=АЛЬВАРЕС AI
AppVersion=1.7.1
AppPublisher=12 штурмова рота
DefaultDirName={autopf}\AlvaresAI
DefaultGroupName=АЛЬВАРЕС AI
SetupIconFile=alvares.ico
OutputBaseFilename=AlvaresAI_Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
CloseApplications=force

[Languages]
Name: "ukrainian"; MessagesFile: "compiler:Languages\Ukrainian.isl"

[Files]
; Вся папка dist\Alvares (крім шаблонів — вони окремо)
Source: "dist\Alvares\*"; DestDir: "{app}"; Excludes: "_internal\templates\rozp_template.docx,_internal\templates\pozition_template.docx,_internal\templates\rozp_Variant_A.docx,_internal\templates\rozp_Variant_B.docx"; Flags: ignoreversion recursesubdirs createallsubdirs

; Шаблони БР — не перезаписувати якщо користувач редагував
Source: "dist\Alvares\_internal\templates\rozp_template.docx"; DestDir: "{app}\_internal\templates"; Flags: onlyifdoesntexist
Source: "dist\Alvares\_internal\templates\pozition_template.docx"; DestDir: "{app}\_internal\templates"; Flags: onlyifdoesntexist
Source: "dist\Alvares\_internal\templates\rozp_Variant_A.docx"; DestDir: "{app}\_internal\templates"; Flags: onlyifdoesntexist
Source: "dist\Alvares\_internal\templates\rozp_Variant_B.docx"; DestDir: "{app}\_internal\templates"; Flags: onlyifdoesntexist

; Робочі файли (onlyifdoesntexist — щоб не затерти дані при оновленні)
Source: "Табель_Багатомісячний.xlsx"; DestDir: "{app}"; Flags: ignoreversion onlyifdoesntexist
Source: "BR_4ShB.xlsx"; DestDir: "{app}"; Flags: ignoreversion onlyifdoesntexist
Source: "ROP.txt"; DestDir: "{app}"; Flags: ignoreversion onlyifdoesntexist
Source: "app.db"; DestDir: "{app}"; Flags: ignoreversion

[Dirs]
Name: "{app}\output"

[Icons]
Name: "{group}\АЛЬВАРЕС AI"; Filename: "{app}\Alvares.exe"; IconFilename: "{app}\_internal\alvares.ico"
Name: "{autodesktop}\АЛЬВАРЕС AI"; Filename: "{app}\Alvares.exe"; IconFilename: "{app}\_internal\alvares.ico"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Створити ярлик на робочому столі"; GroupDescription: "Додаткові ярлики:"

[Run]
Filename: "{app}\Alvares.exe"; Description: "Запустити АЛЬВАРЕС AI"; Flags: nowait postinstall skipifsilent
