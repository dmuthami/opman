[Setup]
AppName=BlueTrax 
AppVerName=BlueTrax AVL v2.2
AppPublisher=RiverCross Technologies
AppPublisherURL=http://www.rivercrosstech.com
AppSupportURL=http://www.rivercrosstech.com
AppUpdatesURL=http://www.rivercrosstech.com
DefaultDirName={pf}\BlueTrax 
DefaultGroupName=BlueTrax 
AllowNoIcons=yes
OutputDir=C:\Programming\Inno scripts\BlueTrax\Binaries
OutputManifestFile=BlueTrax AVL v2.2 - Manifest.txt
OutputBaseFilename=BlueTrax AVL v2.2
SetupIconFile=C:\Programming\Inno scripts\BlueTrax\Dependencies\rcticon.ico
Compression=lzma
SolidCompression=yes
AlwaysRestart=yes

[Languages]
Name: "eng"; MessagesFile: "compiler:Default.isl"
Name: "bra"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"
Name: "cat"; MessagesFile: "compiler:Languages\Catalan.isl"
Name: "cze"; MessagesFile: "compiler:Languages\Czech.isl"
Name: "dan"; MessagesFile: "compiler:Languages\Danish.isl"
Name: "dut"; MessagesFile: "compiler:Languages\Dutch.isl"
Name: "fre"; MessagesFile: "compiler:Languages\French.isl"
Name: "ger"; MessagesFile: "compiler:Languages\German.isl"
Name: "hun"; MessagesFile: "compiler:Languages\Hungarian.isl"
Name: "ita"; MessagesFile: "compiler:Languages\Italian.isl"
Name: "nor"; MessagesFile: "compiler:Languages\Norwegian.isl"
Name: "pol"; MessagesFile: "compiler:Languages\Polish.isl"
Name: "por"; MessagesFile: "compiler:Languages\Portuguese.isl"
Name: "rus"; MessagesFile: "compiler:Languages\Russian.isl"
Name: "slo"; MessagesFile: "compiler:Languages\Slovenian.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]

;dlls and configuration files
Source: "C:\Programming\Inno scripts\BlueTrax\Bluetrax v2.2 Release\config.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax v2.2 Release\BlueTrax AVL v2.2.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax v2.2 Release\BlueTrax AVL v2.2.exe.config"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax v2.2 Release\AE32.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax v2.2 Release\Interop.MO2MO.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax v2.2 Release\MO2MO.dll"; DestDir: "{win}\System32"; Flags: ignoreversion
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax v2.2 Release\MO2MO.dll"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\BlueTrax "; Filename: "{app}\BlueTrax AVL v2.2.exe"
Name: "{userdesktop}\BlueTrax "; Filename: "{app}\BlueTrax AVL v2.2.exe"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\BlueTrax "; Filename: "{app}\BlueTrax AVL v2.2.exe"; Tasks: quicklaunchicon

[Run]
Filename: "{sys}\regsvr32"; Parameters: "/s {win}\System32\MO2MO.dll";


;Write Necessary registry values
[code]
 var
  FinishedInstal: Boolean;
function InitializeSetup(): Boolean;
begin
  Result := True;
  if Result = False then
    MsgBox('InitializeSetup:' #13#13 'Ok, bye bye.', mbInformation, MB_OK);
end;
// Database Values...
procedure Something();
begin

end;

procedure DeinitializeSetup();
 var
  Path: String;
begin
  if FinishedInstal then begin                      //Postgres Values

  if RegQueryStringValue(HKEY_LOCAL_MACHINE, 'System\ControlSet001\Control\Session Manager\Environment','Path', Path) then
  begin
    //MsgBox('List of values:'#13#10#13#10 + Path, mbInformation, MB_OK);
  end else
  begin
    // add any code to handle failure here
  end;
    


  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    FinishedInstal := True;
end;

 function yes(): Boolean;
 begin
 end;
 
procedure CurPageChanged(CurPageID: Integer);
begin
  case CurPageID of
    wpWelcome :
       yes();
      //MsgBox('CurPageChanged:' #13#13 'Welcome to the [Code] scripting demo. This demo will show you some possibilities of the scripting support.' #13#13 'The scripting engine used is RemObjects Pascal Script by Carlo Kok. See http://www.remobjects.com/?ps for more information.', mbInformation, MB_OK);
    wpFinished :
       yes();
      //MsgBox('CurPageChanged:' #13#13 'Welcome to final page of this demo. Click Finish to exit.', mbInformation, MB_OK);
  end;
end;

[UninstallDelete]
Type: files; Name: "{app}\BlueTrax.url"

