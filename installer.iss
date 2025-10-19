[Setup]
AppName=DNS Tester
AppVersion=1.0.0
AppPublisher=pcoof
AppPublisherURL=https://github.com/pcoof/dns
AppSupportURL=https://github.com/pcoof/dns/issues
AppUpdatesURL=https://github.com/pcoof/dns/releases
DefaultDirName={autopf}\DNS Tester
DefaultGroupName=DNS Tester
AllowNoIcons=yes
LicenseFile=LICENSE
OutputDir=Output
OutputBaseFilename=DNS-Tester-Setup
SetupIconFile=icon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "chinesesimp"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 6.1

[Files]
Source: "dist\DNS-Tester.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "icon.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "dns_servers.ini"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "LICENSE"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\DNS Tester"; Filename: "{app}\DNS-Tester.exe"; IconFilename: "{app}\icon.ico"
Name: "{group}\{cm:UninstallProgram,DNS Tester}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\DNS Tester"; Filename: "{app}\DNS-Tester.exe"; IconFilename: "{app}\icon.ico"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\DNS Tester"; Filename: "{app}\DNS-Tester.exe"; IconFilename: "{app}\icon.ico"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\DNS-Tester.exe"; Description: "{cm:LaunchProgram,DNS Tester}"; Flags: nowait postinstall skipifsilent

[Registry]
Root: HKLM; Subkey: "Software\DNS Tester"; ValueType: string; ValueName: "InstallPath"; ValueData: "{app}"

[UninstallDelete]
Type: filesandordirs; Name: "{app}"