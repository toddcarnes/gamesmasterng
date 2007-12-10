#define AppDate "19 April 2006"
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
DefaultDirName={pf}\Games Master
DefaultGroupName=Games Master
LicenseFile=
InfoBeforeFile=
InfoAfterFile=
OutputDir={#Source}\Distributions
SourceDir={#Source}
OutputBaseFilename=GamesMasterSetup {#AppVersion}
MinVersion=4.01.1998,4.00.1381sp6
AppCopyright=Copyright © Ian Llewelyn Evans 2007
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
DisableDirPage=true
DisableProgramGroupPage=true

[Icons]
Name: {group}\Uninstall Games Master; Filename: {uninstallexe}
Name: {group}\Games Master; Filename: {app}\GamesMaster.exe
Name: {commondesktop}\Games Master; Filename: {app}\GamesMaster.exe

[Files]
Source: {#Source}\GamesMaster.exe; DestDir: {app}; Flags: comparetimestamp
;VB6
Source: {#vbFiles}\stdole2.tlb; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regtypelib
Source: {#vbFiles}\msvbvm60.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: {#vbFiles}\oleaut32.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: {#vbFiles}\olepro32.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: {#vbFiles}\asycfilt.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile
;Source: {#vbFiles}\comcat.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: {#vbFiles}\COMDLG32.OCX; DestDir: {sys}; Flags: regserver sharedfile uninsneveruninstall restartreplace
Source: {#vbFiles}\MSFLXGRD.OCX; DestDir: {sys}; Flags: restartreplace sharedfile uninsneveruninstall regserver
