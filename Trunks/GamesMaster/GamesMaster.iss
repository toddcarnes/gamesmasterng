#define AppDate "31 December 2007"
#define Source SourcePath
#define HomePage "http://www.mykoala.net"
#define vbFiles Source + "\VB60Files"

#define FileVersion GetFileVersion(Source+"\GamesMaster.exe")
#define AppVersion Copy(FileVersion,1,rpos(".",Copy(FileVersion,1,rpos(".",FileVersion)-1))-1)+Copy(FileVersion,rpos(".",FileVersion))

[Setup]
AppID=GalaxyNGGamesMaster
AppName=Games Master
AppVersion={#AppVersion}
AppVerName=Games Master {#AppVersion}
AppPublisher=Ian Evans
AppPublisherURL={#HomePage}
AppSupportURL={#HomePage}
AppUpdatesURL={#HomePage}
DefaultDirName={pf}\GalaxyNG
DefaultGroupName=GalaxyNG
LicenseFile={#Source}\package\License.rtf
InfoBeforeFile={#Source}\package\Before.rtf
InfoAfterFile={#Source}\package\After.rtf
OutputDir={#Source}\Distributions
SourceDir={#Source}
OutputBaseFilename=GamesMasterSetupV{#AppVersion}
MinVersion=4.1.1998,4.0.1381sp6
AppCopyright=Copyright © Ian Llewelyn Evans 2007-2008
UserInfoPage=false
ChangesAssociations=true
VersionInfoVersion={#FileVersion}
VersionInfoTextVersion=Games Master V{#AppVersion}
VersionInfoCompany=Ian Evans
VersionInfoDescription=Games Master Installation V{#AppVersion}
ShowLanguageDialog=no
UninstallDisplayIcon={app}\GamesMaster.exe
UninstallDisplayName=Games Master Version {#AppVersion}
AppReadmeFile=
UsePreviousUserInfo=false
DisableDirPage=false
DisableProgramGroupPage=false
AllowRootDirectory=true
AllowUNCPath=false
VersionInfoCopyright=Copyright © Ian Llewelyn Evans 2007-2008
AlwaysShowDirOnReadyPage=true
AlwaysShowGroupOnReadyPage=true

[Icons]
Name: {group}\Uninstall Games Master; Filename: {uninstallexe}
Name: {group}\Games Master; Filename: {app}\GamesMaster.exe
Name: {commondesktop}\Games Master; Filename: {app}\GamesMaster.exe

[Files]
Source: {#Source}\GamesMaster.exe; DestDir: {app}; Flags: comparetimestamp
Source: {#Source}\GamesMaster.txt; DestDir: {app}; Flags: comparetimestamp
Source: {#Source}\Package\GamesMaster.ini; DestDir: {app}; DestName: GamesMaster.ini; Flags: onlyifdoesntexist
Source: {#Source}\Package\License.rtf; DestDir: {app}; Flags: comparetimestamp
;GalaxyNG
Source: {#Source}\package\galaxyng.exe; DestDir: {app}; Flags: comparetimestamp
Source: {#Source}\Package\gpl-3.0.txt; DestDir: {app}; Flags: comparetimestamp
;Info-Zip
Source: {#Source}\package\zip32.dll; DestDir: {app}; Flags: comparetimestamp
Source: {#Source}\Package\Info-Zip License.txt; DestDir: {app}; Flags: comparetimestamp
;VB6
Source: {#vbFiles}\stdole2.tlb; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regtypelib
Source: {#vbFiles}\msvbvm60.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: {#vbFiles}\oleaut32.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: {#vbFiles}\olepro32.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: {#vbFiles}\asycfilt.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile
;VB6 Controls
Source: {#vbFiles}\mshflxgd.OCX; DestDir: {sys}; Flags: restartreplace sharedfile uninsneveruninstall regserver
Source: {#vbFiles}\tabctl32.OCX; DestDir: {sys}; Flags: restartreplace sharedfile uninsneveruninstall regserver
Source: {#vbFiles}\mscomctl.OCX; DestDir: {sys}; Flags: restartreplace sharedfile uninsneveruninstall regserver
Source: {#vbFiles}\ws2_32.dll; DestDir: {sys}; Flags: restartreplace sharedfile uninsneveruninstall

[INI]
Filename: {app}\gamesmaster.ini; Section: Folders; Key: GalaxyNGHome; String: {app}\
Filename: {app}\gamesmaster.ini; Section: FileNames; Key: Executable; String: {app}\galaxyng.exe
Filename: {app}\gamesmaster.ini; Section: EMail; Key: ServerName; String: {computername}

[Dirs]
Name: {app}\data
Name: {app}\inbox
Name: {app}\log
Name: {app}\notices
Name: {app}\orders
Name: {app}\outbox
Name: {app}\reports
Name: {app}\statistics

[Run]
Filename: {app}\GamesMaster.exe; Parameters: -showoptions; WorkingDir: {app}; Tasks: " SetOptions"

[Tasks]
Name: SetOptions; Description: Set Program Options Now; Flags: unchecked checkablealone
