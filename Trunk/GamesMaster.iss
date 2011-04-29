#define AppDate "16 May 2010"
#define Source SourcePath
#define HomePage "http://sourceforge.net/projects/gamesmasterng/"
#define vbFiles Source + "\VB60Files"

#define FileVersion GetFileVersion(Source+"\GamesMaster.exe")
#define AppVersion Copy(FileVersion,1,rpos(".",Copy(FileVersion,1,rpos(".",FileVersion)-1))-1)+Copy(FileVersion,rpos(".",FileVersion))

[Setup]
AppID=GamesMasterNG
AppName=Games Master NG
AppVersion={#AppVersion}
AppVerName=Games Master NG {#AppVersion}
AppPublisher=Ian Evans
AppPublisherURL={#HomePage}
AppSupportURL={#HomePage}
AppUpdatesURL={#HomePage}
DefaultDirName={pf}\GamesMasterNG
DefaultGroupName=GamesMasterNG
LicenseFile={#Source}\Package\COPYING.txt
InfoBeforeFile={#Source}\package\Before.rtf
InfoAfterFile={#Source}\package\After.rtf
OutputDir={#Source}\Distributions
SourceDir={#Source}
OutputBaseFilename=GamesMasterNG_SetupV{#AppVersion}
MinVersion=4.1.1998,4.0.1381sp6
AppCopyright=Copyright 2007 -2011 Ian Evans
UserInfoPage=false
ChangesAssociations=true
VersionInfoVersion={#FileVersion}
VersionInfoTextVersion=Games Master NG V{#AppVersion}
VersionInfoCompany=Ian Evans
VersionInfoDescription=Games Master NG Installation V{#AppVersion}
ShowLanguageDialog=no
UninstallDisplayIcon={app}\GamesMaster.exe
UninstallDisplayName=Games Master NG Version {#AppVersion}
UsePreviousUserInfo=false
AllowRootDirectory=true
AllowUNCPath=false
VersionInfoCopyright=Copyright 2007 - 2011 Ian Evans
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
;GalaxyNG
Source: {#Source}\Package\galaxyng.exe; DestDir: {app}; Flags: comparetimestamp
Source: {#Source}\Package\COPYING.txt; DestDir: {app}; Flags: comparetimestamp
;Info-Zip
Source: {#Source}\Package\zip32.dll; DestDir: {app}; Flags: comparetimestamp
Source: {#Source}\Package\Info-Zip License.txt; DestDir: {app}; Flags: comparetimestamp
Source: {#Source}\Package\Changes.txt; DestDir: {app}; Flags: comparetimestamp
;VB6
Source: {#vbFiles}\stdole2.tlb; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regtypelib
Source: {#vbFiles}\msvbvm60.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: {#vbFiles}\oleaut32.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: {#vbFiles}\olepro32.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: {#vbFiles}\comcat.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: {#vbFiles}\asycfilt.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile
;VB6 Controls
Source: {#vbFiles}\mshflxgd.OCX; DestDir: {sys}; Flags: restartreplace sharedfile uninsneveruninstall regserver
Source: {#vbFiles}\tabctl32.OCX; DestDir: {sys}; Flags: restartreplace sharedfile uninsneveruninstall regserver
Source: {#vbFiles}\mscomctl.OCX; DestDir: {sys}; Flags: restartreplace sharedfile uninsneveruninstall regserver
Source: {#vbFiles}\MSWinSck.ocx; DestDir: {sys}; Flags: restartreplace sharedfile uninsneveruninstall regserver
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
Filename: {app}\GamesMaster.exe; Parameters: -showoptions; WorkingDir: {app}; Flags: postinstall unchecked
Filename: {app}\Changes.txt; Description: View Changes included in this version; Flags: shellexec postinstall; WorkingDir: {app}

[InnoIDE_PreCompile]
Name: {#Source}\Package\signcode.exe; Parameters: " -cn ""Ian Evans"" -s ""TrustedPeople"" -n ""Games Master NG Version {#AppVersion} ({#AppDate})"" -sp chain -t http://timestamp.verisign.com/scripts/timstamp.dll ""{#Source}\GamesMaster.exe"" -i {#HomePage}"; Flags: AbortOnError; 

[InnoIDE_PostCompile]
Name: {#Source}\Package\signcode.exe; Parameters: " -cn ""Ian Evans"" -s ""TrustedPeople"" -n ""Games Master NG Setup Version {#AppVersion} ({#AppDate})"" -sp chain -t http://timestamp.verisign.com/scripts/timstamp.dll ""{#Source}\Distributions\GamesMasterNG_SetupV{#AppVersion}.exe"" -i {#HomePage}"; Flags: abortonerror; 
