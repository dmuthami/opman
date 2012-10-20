[Setup]
AppName=BlueTrax 
AppVerName=BlueTrax AVL 2.2
AppPublisher=RiverCross Technologies
AppPublisherURL=http://www.rivercrosstech.com
AppSupportURL=http://www.rivercrosstech.com
AppUpdatesURL=http://www.rivercrosstech.com
DefaultDirName={pf}\BlueTrax 
DefaultGroupName=BlueTrax 
AllowNoIcons=yes
OutputDir=C:\Programming\Inno scripts\BlueTrax\2.2binaries
OutputManifestFile=BlueTrax v2.2-Full-Manifest.txt
OutputBaseFilename=BlueTrax
SetupIconFile=C:\Programming\Inno scripts\BlueTrax\Dependencies\rcticon.ico
Compression=lzma
SolidCompression=yes
AlwaysRestart=yes
DiskSpanning=yes
SlicesPerDisk=1
DiskSliceSize=636000000

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
Source: "C:\WINDOWS\Fonts\esri_1.ttf"; DestDir: "{win}\Fonts"; Flags: ignoreversion
Source: "C:\WINDOWS\Fonts\esri_2.ttf"; DestDir: "{win}\Fonts"; Flags: ignoreversion
Source: "C:\WINDOWS\Fonts\esri_3.ttf"; DestDir: "{win}\Fonts"; Flags: ignoreversion
Source: "C:\WINDOWS\Fonts\esri_4.ttf"; DestDir: "{win}\Fonts"; Flags: ignoreversion
Source: "C:\WINDOWS\Fonts\esri_5.ttf"; DestDir: "{win}\Fonts"; Flags: ignoreversion
Source: "C:\WINDOWS\Fonts\esri_6.ttf"; DestDir: "{win}\Fonts"; Flags: ignoreversion
Source: "C:\WINDOWS\Fonts\esri_7.ttf"; DestDir: "{win}\Fonts"; Flags: ignoreversion
Source: "C:\WINDOWS\Fonts\esri_8.ttf"; DestDir: "{win}\Fonts"; Flags: ignoreversion
Source: "C:\WINDOWS\Fonts\esri_9.ttf"; DestDir: "{win}\Fonts"; Flags: ignoreversion
Source: "C:\Programming\Inno scripts\BlueTrax\Dependencies\TTMIOS__.TTF"; DestDir: "{win}\Fonts"; Flags: ignoreversion

;dlls and configuration files

Source: "C:\Programming\Inno scripts\BlueTrax\Bluetrax v2.2 Release\config.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax v2.2 Release\BlueTrax AVL v2.2.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax v2.2 Release\BlueTrax AVL v2.2.exe.config"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax v2.2 Release\AE32.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax v2.2 Release\Interop.MO2MO.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax v2.2 Release\MO2MO.dll"; DestDir: "{win}\System32"; Flags: ignoreversion
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax v2.2 Release\MO2MO.dll"; DestDir: "{app}"; Flags: ignoreversion

Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax v2.2 Release\Interop.VBIDE.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax v2.2 Release\Interop.Microsoft.Office.Core.dll"; DestDir: "{app}"; Flags: ignoreversion

Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax Release\AxInterop.MapObjects2.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax Release\bState.ini"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax Release\Config.ini"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax Release\CustomSymbol.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax Release\Initialize.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax Release\Interop.AFCustom.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax Release\Interop.Excel.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax Release\Interop.MapObjects2.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax Release\Interop.Shell.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax Release\postgis.sql"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax Release\rctgraph.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax Release\RctMo.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax Release\stdole.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax Release\AFCust20.tlb"; DestDir: "{app}"; Flags: ignoreversion


;Folders
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax Release\Asters\*"; DestDir: "{app}\Asters"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax Release\MS-Access\*"; DestDir: "{app}\MS-Access"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax Release\init\*"; DestDir: "{app}\Init"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax Release\Images\*"; DestDir: "{app}\Images"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Programming\Inno scripts\BlueTrax\BlueTrax Release\Vectors\*"; DestDir: "{app}\Vectors"; Flags: ignoreversion recursesubdirs createallsubdirs

;DataBase And dot net framework and its dependencies
Source: "C:\Programming\Inno scripts\BlueTrax\dotnetframework\dotnetfx.exe"; DestDir: "{win}\System32"; Flags: ignoreversion  deleteafterinstall
Source: "C:\Programming\Inno scripts\BlueTrax\Dependencies\odbc_net.msi"; DestDir: "{win}\System32"; Flags: ignoreversion  deleteafterinstall
Source: "C:\Programming\Inno scripts\BlueTrax\Dependencies\psqlodbc.msi"; DestDir: "{win}\System32"; Flags: ignoreversion  deleteafterinstall
Source: "C:\Programming\Inno scripts\BlueTrax\DataBase\postgresql-8.0.msi"; DestDir: "{win}\System32"; Flags: ignoreversion  deleteafterinstall
Source: "C:\Programming\Inno scripts\BlueTrax\DataBase\postgresql-8.0-int.msi"; DestDir: "{win}\System32"; Flags: ignoreversion  deleteafterinstall
;Source: "C:\Programming\Inno scripts\BlueTrax\DataBase\postgresql-8.0-int.msi"; DestDir: "{app}"; Flags: ignoreversion  deleteafterinstall
Source: "C:\Programming\Inno scripts\BlueTrax\Dependencies\psqlshell.exe"; DestDir: "{app}"; Flags: ignoreversion  deleteafterinstall
;this is the .net odbc support driver it has to go to the application folder
Source: "C:\Programming\Inno scripts\BlueTrax\Dependencies\adodb.dll"; DestDir: "{app}"; Flags: ignoreversion


;MapObjcets
Source: "C:\Programming\Inno scripts\BlueTrax\MapObjects\MO21rt.EXE"; DestDir: "{app}"; Flags: ignoreversion deleteafterinstall
Source: "C:\Programming\Inno scripts\BlueTrax\MapObjects\MO21sp3.EXE"; DestDir: "{app}"; Flags: ignoreversion deleteafterinstall

;Msi installer. latest version for dot net
Source: "C:\Programming\Inno scripts\BlueTrax\Dependencies\msiexec.exe"; DestDir: "{win}\System32"; Flags: ignoreversion deleteafterinstall

;sql Cnverter
Source: "C:\Programming\Inno scripts\BlueTrax\Converter\*"; DestDir: "{app}\Converter"; Flags: ignoreversion recursesubdirs createallsubdirs

; NOTE: Don't use "Flags: ignoreversion" on any shared system files strange things might happen

[INI]
Filename: "{app}\BlueTrax.url"; Section: "InternetShortcut"; Key: "URL"; String: "http://www.rivercrosstech.com"

[Icons]
Name: "{group}\BlueTrax "; Filename: "{app}\BlueTrax AVL v2.2.exe"
Name: "{group}\Setup Database "; Filename: "{app}\Converter\DB Converter.exe"
Name: "{group}\{cm:ProgramOnTheWeb,BlueTrax }"; Filename: "{app}\BlueTrax.url"
Name: "{group}\{cm:UninstallProgram,BlueTrax }"; Filename: "{uninstallexe}"
Name: "{userdesktop}\BlueTrax "; Filename: "{app}\BlueTrax AVL v2.2.exe"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\BlueTrax "; Filename: "{app}\BlueTrax AVL v2.2.exe"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\BlueTrax.exe"; Description: "{cm:LaunchProgram,BlueTrax }"; Flags: nowait postinstall skipifsilent
Filename: "{win}\System32\dotnetfx.EXE"; Parameters: "/Q";
Filename: {win}\System32\msiexec.exe; parameters: /i {win}\System32\odbc_net.msi /qb!;
;Filename: "{app}\psqlshell.EXE";
Filename: "{app}\MO21rt.EXE"; Parameters: "/ABCDEFGHIJKLMNOPQX";
Filename: "{app}\MO21sp3.EXE";
Filename: "{win}\System32\msiexec.exe"; parameters: /i {win}\System32\postgresql-8.0.msi;

;register mo2mo
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
        RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
  'BI', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'BoolsAsChar', '1');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'ByteaAsLongVarBinary', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'CancelAsFreeStmt', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'CommLog', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'ConnSettings', '');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'Database', 'BlueBase');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'Debug', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'Description', '');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'DisallowPremature', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'Driver', ExpandConstant('{win}\system32\psqlodbc.dll'));
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'ExtraSysTablePrefixes', 'dd_;');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'FakeOidIndex', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'Fetch', '100');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'Ksqo', '1');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'LFConversion', '1');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'LowerCaseIdentifier', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'MaxLongVarcharSize', '8190');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'MaxVarcharSize', '254');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'Optimizer', '1');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'Parse', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'Password', 'postgres');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'Port', '5432');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'Protocol', '6.4');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'ReadOnly', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'RowVersioning', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'Servername', 'localhost');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'ShowOidColumn', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'ShowSystemTables', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'TextAsLongVarChar', '1');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'TrueIsMinus1', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'UniqueIndex', '1');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'UnknownAsLongVarChar', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'UnkownSizes', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'UpdatableCursors', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'UseDeclareFetch', '0');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'Username', 'postgres');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseV2',
    'UserServerSidePrepare', '0');
    
    //Rams Database Values

  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseBackup',
  'DBQ', ExpandConstant('{app}\Ms-Access\BlueTrax.mdb'));
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseBackup',
  'Driver', ExpandConstant('{win}\System32\odbcjt32.dll'));
  RegWriteDWordValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseBackup',
  'DriverId', 25);
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseBackup',
  'FIL', 'MS Access;');
  RegWriteDWordValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseBackup',
  'SafeTransactions', 0);
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseBackup',
  'UID', '');
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseBackup\Engines\Jet',
  'ImplicitCommitSync', '');

  RegWriteDWordValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseBackup\Engines\Jet',
  'MaxBufferSize', 2048);
  RegWriteDWordValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseBackup\Engines\Jet',
  'PageTimeout', 5);
  RegWriteDWordValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseBackup\Engines\Jet',
  'Threads', 3);
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\BlueBaseBackup\Engines\Jet',
  'UserCommitSync', 'Yes');


  // Database ODBC Control Panel Extension
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\ODBC Data Sources',
  'BlueBaseBackup', 'Microsoft Access Driver (*.mdb)');
    
  if RegQueryStringValue(HKEY_LOCAL_MACHINE, 'System\ControlSet001\Control\Session Manager\Environment','Path', Path) then
  begin
    //MsgBox('List of values:'#13#10#13#10 + Path, mbInformation, MB_OK);
  end else
  begin
    // add any code to handle failure here
  end;
    
  //Database Environment Path Variables
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'System\ControlSet001\Control\Session Manager\Environment',
    'PGHOME', ExpandConstant('{pf}\PostgreSQL\8.0'));
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'System\ControlSet001\Control\Session Manager\Environment',
    'PGDATA', ExpandConstant('{pf}\PostgreSQL\8.0\Data'));
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'System\ControlSet001\Control\Session Manager\Environment',
    'PGLIB', ExpandConstant('{pf}\PostgreSQL\8.0\Lib'));
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'System\ControlSet001\Control\Session Manager\Environment',
    'PGHOST', ExpandConstant('localhost'));
  RegWriteStringValue(HKEY_LOCAL_MACHINE, 'System\ControlSet001\Control\Session Manager\Environment',
    'Path', ExpandConstant('{pf}\PostgreSQL\8.0\Bin;' + Path));
    
    // Database ODBC Control Panel Extension
      RegWriteStringValue(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\ODBC Data Sources',
        'BlueBaseV2', 'PostgreSQL');
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

