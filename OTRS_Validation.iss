; -- Example1.iss --
; Demonstrates copying 3 files and creating an icon.

; SEE THE DOCUMENTATION FOR DETAILS ON CREATING .ISS SCRIPT FILES!

[Setup]
AppName=OTRS Validation
AppVersion=1.1.000
WizardStyle=modern
DefaultDirName={autopf}\OTRS Validation
DefaultGroupName=OTRS Validation
; UninstallDisplayIcon=
SetupIconFile=C:\Users\S4sh\Documents\GitHub\OTRS_Validation\rocket.ico
Compression=lzma2
SolidCompression=yes
OutputDir=.

[Files]
Source: "C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python36_64\Scripts\dist\OTRS_Validation.exe"; DestDir: "{app}"; CopyMode: alwaysoverwrite;
Source: "C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python36_64\Scripts\chromedriver.exe"; DestDir: "{app}"; CopyMode: alwaysoverwrite;

[Icons]
Name: "{group}\OTRS Validation"; Filename: "{app}\OTRS_Validation.exe"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}";

[Run]
Filename: "{app}\OTRS_Validation.exe"; Flags: nowait postinstall unchecked