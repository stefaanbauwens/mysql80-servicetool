[Setup]
OutputDir=.
OutputBaseFilename=setup-MySQL80-ServiceTool
AppName=MySQL80 Service Tool
AppVerName=1.0.0
AppId={{77B2C974-87E2-4925-BE67-D7AC343C67D0}
DefaultDirName={pf}\MySQL80-ServiceTool
DisableDirPage=true
ShowLanguageDialog=no
UninstallDisplayName=MySQL80 Service Tool
UninstallDisplayIcon={app}\MySQL80-ServiceTool.exe
AppPublisher=SteBaDev
AppVersion=1.0.0
DisableProgramGroupPage=yes

[Files]
Source: MySQL80-ServiceTool.exe; DestDir: {app}
Source: setup-task.xml; DestDir: {tmp}

[Icons]
Name: "{userdesktop}\MySQL80 Service Tool"; Filename: "{sys}\schtasks.exe"; WorkingDir: "{app}"; Flags: runminimized; IconFilename: "{app}\MySQL80-ServiceTool.exe"; Parameters: "/RUN /TN MySQL80-ServiceTool"; MinVersion: 0,6.0

[Run]
Filename: "{sys}\schtasks.exe"; Parameters: "/CREATE /XML {tmp}\setup-task.xml /TN MySQL80-ServiceTool"; WorkingDir: "{app}"; Flags: runascurrentuser skipifdoesntexist runhidden; MinVersion: 0,6.0

[ThirdPartySettings]
CompileLogMethod=append
