[Setup]
AppName=OPman
AppVerName=OPman
AppPublisher=RiverCross Technologies
AppPublisherURL=http://www.rivercrosstech.com
AppSupportURL=http://www.rivercrosstech.com
AppUpdatesURL=http://www.rivercrosstech.com
DefaultDirName={pf}\OPman
DefaultGroupName=OPman
AllowNoIcons=yes
OutputDir=C:\Programming\vbnet_2003\Version23apr11
;OutputManifestFile=BlueTrax AVL 2.1 - Manifest.txt
OutputBaseFilename=OPman
 ;SetupIconFile=C:\Programming\vbnet_2003\Version23apr11\Ramani logo.ico
SetupIconFile=C:\Programming\vbnet_2003\Version23apr11\opman.ico
;SetupIconFile=C:\Programming\vbnet_2003\Version23apr11\georiginea.ico
Password=123
Compression=lzma
SolidCompression=yes
AlwaysRestart=no

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
;Source: "C:\Programming\C#\sperhec\hec2\bin\Release\Interop.Excel.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\vbnet_2003\Version23apr11\ExpTree_Demo\bin\OPman.exe"; DestDir: "{app}"; Flags: ignoreversion

;database and other dependencies
;Msi installer. latest version for dot net
Source: "C:\Programming\C#\dependencies\msiexec.exe"; DestDir: "{win}\System32"; Flags: ignoreversion deleteafterinstall
;Source: "C:\Programming\C#\dependencies\postgresql-8.0.msi"; DestDir: "{win}\System32"; Flags: ignoreversion  deleteafterinstall
;Source: "C:\Programming\C#\dependencies\postgresql-8.0-int.msi"; DestDir: "{win}\System32"; Flags: ignoreversion  deleteafterinstall
Source: "C:\Programming\C#\dependencies\dotNetFramework1.1\dotnetfx v1.1.exe"; DestDir: "{win}\System32"; Flags: ignoreversion  deleteafterinstall
Source: "C:\Programming\C#\dependencies\psqlodbc-08_01_0200\psqlodbc.msi"; DestDir: "{win}\System32"; Flags: ignoreversion  deleteafterinstall
;Source: "C:\Programming\C#\dependencies\RCL_DB.bat"; DestDir: "{win}\System32"; Flags: ignoreversion  deleteafterinstall
Source: "C:\Programming\C#\dependencies\psqlodbc-08_01_0200\upgrade.bat"; DestDir: "{win}\System32"; Flags: ignoreversion  deleteafterinstall
;  -----------copying dlls   and  configuration files

Source: "C:\Programming\vbnet_2003\Version23apr11\ExpTree_Demo\bin\VSEssentials.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\vbnet_2003\Version23apr11\ExpTree_Demo\bin\WinWordControl.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\vbnet_2003\Version23apr11\ExpTree_Demo\bin\adodb.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\vbnet_2003\Version23apr11\ExpTree_Demo\bin\AMS.TextBox.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\vbnet_2003\Version23apr11\ExpTree_Demo\bin\AxInterop.MSMask.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\vbnet_2003\Version23apr11\ExpTree_Demo\bin\ExpTreeLib.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\vbnet_2003\Version23apr11\ExpTree_Demo\bin\Interop.MSMask.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\vbnet_2003\Version23apr11\ExpTree_Demo\bin\Interop.Office.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\vbnet_2003\Version23apr11\ExpTree_Demo\bin\Interop.VBIDE.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\vbnet_2003\Version23apr11\ExpTree_Demo\bin\Interop.Word.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\vbnet_2003\Version23apr11\ExpTree_Demo\bin\MSMASK32.OCX"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace sharedfile regserver
Source: "C:\Programming\vbnet_2003\Version23apr11\ExpTree_Demo\bin\exporttoexcel.dll"; DestDir: "{app}"; Flags: ignoreversion
;Source: "C:\Programming\vbnet_2003\Version23apr11\ExpTree_Demo\bin\OPman_RTR.exe.config"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\vbnet_2003\Version23apr11\ExpTree_Demo\bin\stdole.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\vbnet_2003\Version23apr11\ExpTree_Demo\bin\Stimulsoft.Controls.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\vbnet_2003\Version23apr11\ExpTree_Demo\bin\AMS.Profile.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\vbnet_2003\Version23apr11\ExpTree_Demo\bin\output.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\vbnet_2003\Version23apr11\ExpTree_Demo\bin\output2.txt"; DestDir: "{app}"; Flags: ignoreversion
;Source: "C:\Programming\vbnet_2003\Version23apr11\ExpTree_Demo\bin\Config.dll"; DestDir: "{app}"; Flags: ignoreversion

;Folders
;Source: "C:\Programming\vbnet_2003\Version23apr11\ExpTree_Demo\bin\ar\*"; DestDir: "{app}\ar"; Flags: ignoreversion recursesubdirs createallsubdirs
;Source: "C:\Programming\vbnet_2003\Version23apr11\ExpTree_Demo\bin\en\*"; DestDir: "{app}\en"; Flags: ignoreversion recursesubdirs createallsubdirs
;Source: "C:\Programming\vbnet_2003\Version23apr11\ExpTree_Demo\bin\en-GB\*"; DestDir: "{app}\en-GB"; Flags: ignoreversion recursesubdirs createallsubdirs
;Source: "C:\Programming\vbnet_2003\Version23apr11\ExpTree_Demo\bin\vi\*"; DestDir: "{app}\vi"; Flags: ignoreversion recursesubdirs createallsubdirs


  ;register mo2mo
;Filename: "{sys}\regsvr32"; Parameters: "/s {win}\System32\MO2MO.dll";

;Source: "{src}\MapClick.exe"; DestDir: "{app}"; Flags: external
;Source: "{src}\MapWinGISOCXOnly.exe"; DestDir: "{app}"; Flags: external
;Source: "{src}\Data\*"; DestDir: "{app}"; Flags: external
;Source: "{src}\countylabel.csv"; DestDir: "{app}"; Flags: external

;you can use {src} to indicate that the files are located in the same location as the setup.exe.
; and the external flag prevents the files from being compiled into the setup.exe.

[Icons]
Name: "{group}\OPman"; Filename: "{app}\OPman.exe"
Name: "{userdesktop}\OPman"; Filename: "{app}\OPman.exe"; Tasks: desktopicon
Name: "{group}\{cm:UninstallProgram,OPman}"; Filename: "{uninstallexe}"
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\OPman"; Filename: "{app}\OPman.exe"; Tasks: quicklaunchicon




[Run]
;Filename: "{win}\System32\msiexec.exe"; parameters: /i {win}\System32\postgresql-8.0.msi;
Filename: "{win}\System32\dotnetfx v1.1.exe"; Parameters: "/Q";
Filename: {win}\System32\msiexec.exe; parameters: /i {win}\System32\psqlodbc.msi /qb!;
Filename: "{win}\System32\upgrade.bat"; Parameters: "/Q";



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
        RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
  'BI', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'BoolsAsChar', '1');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'ByteaAsLongVarBinary', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'CancelAsFreeStmt', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'CommLog', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'ConnSettings', '');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'Database', 'RCL_DB');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'Debug', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'Description', '');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'DisallowPremature', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'Driver', ExpandConstant('{win}\system32\psqlodbca.dll'));
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'ExtraSysTablePrefixes', 'dd_;');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'FakeOidIndex', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'Fetch', '100');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'Ksqo', '1');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'LFConversion', '1');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'LowerCaseIdentifier', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'MaxLongVarcharSize', '8190');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'MaxVarcharSize', '254');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'Optimizer', '1');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'Parse', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'Password', 'postgres');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'Port', '');
   RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'ReadOnly', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'RowVersioning', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'Servername', '');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'ShowOidColumn', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'ShowSystemTables', '0');
   RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'ShowSystemTables', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'SSLmode', 'prefer');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'TextAsLongVarchar', '1');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'TrueIsMinus1', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'UniqueIndex', '1');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'UnknownAsLongVarChar', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'UnkownSizes', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'UpdatableCursors', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'UseDeclareFetch', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'Username', '');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\RCL_DB',
    'UserServerSidePrepare', '0');


    //Database Environment Path Variables
  //RegWriteStringValue(HKEY_LOCAL_MACHINE, 'System\ControlSet001\Control\Session Manager\Environment',
   // 'PGHOME', ExpandConstant('{pf}\PostgreSQL\8.0'));
 // RegWriteStringValue(HKEY_LOCAL_MACHINE, 'System\ControlSet001\Control\Session Manager\Environment',
   // 'PGDATA', ExpandConstant('{pf}\PostgreSQL\8.0\Data'));
 // RegWriteStringValue(HKEY_LOCAL_MACHINE, 'System\ControlSet001\Control\Session Manager\Environment',
 //   'PGLIB', ExpandConstant('{pf}\PostgreSQL\8.0\Lib'));
 // RegWriteStringValue(HKEY_LOCAL_MACHINE, 'System\ControlSet001\Control\Session Manager\Environment',
   // 'PGHOST', ExpandConstant('localhost'));
 // RegWriteStringValue(HKEY_LOCAL_MACHINE, 'System\ControlSet001\Control\Session Manager\Environment',
   // 'Path', ExpandConstant('{pf}\PostgreSQL\8.0\Bin;' + Path));


    // Database ODBC Control Panel Extension
      RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\ODBC Data Sources',
        'RCL_DB', 'PostgreSQL ANSI');
      //RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\ODBC Data Sources',
       // 'FMS', 'Driver do Microsoft Access (*.mdb)');

      // Register DLLs and OXCs copied to the {win}\System32 directory!

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



