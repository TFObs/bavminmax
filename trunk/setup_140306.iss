; 'D:\Dokumente und Einstellungen\Jörg\Eigene Dateien\Visual Basic Projekte\BAVMinMax2\code\Paket\SETUP.LST' imported by ISTool version 5.1.5

[Setup]
AppName=BAVMinMax
AppVerName=BAVMinMax
PrivilegesRequired=admin
DefaultDirName={pf}\BAVMinMax
DefaultGroupName=BAVMinMax

[Languages]
Name: german; MessagesFile: compiler:Languages\German.isl

[Tasks]
Name: desktopicon; Description: {cm:CreateDesktopIcon}; GroupDescription: {cm:AdditionalIcons}; Flags: unchecked

[Files]

Source: weitergabe\MSCAL.OCX; DestDir: {app}; Flags: regserver sharedfile promptifolder
Source: weitergabe\COMDLG32.OCX; DestDir: {app}; Flags: regserver sharedfile promptifolder
Source: weitergabe\MSHFLXGD.OCX; DestDir: {app}; Flags: regserver sharedfile promptifolder
Source: weitergabe\BAVMinMax.exe; DestDir: {app}; Flags: promptifolder
Source: weitergabe\BAV_Sterne.mdb; DestDir: {app}
Source: weitergabe\BAVMinMax.chm; DestDir: {app}
Source: weitergabe\sternkoord.txt; DestDir: {app}

[Icons]
Name: {group}\BAVMinMax; Filename: {app}\BAVMinMax.exe; WorkingDir: {app}
