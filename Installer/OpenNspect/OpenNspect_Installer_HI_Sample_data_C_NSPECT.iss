; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
; DLE 12/23/2011: ESRI data version AppId
;AppId={{DB15D326-3441-4107-8444-35E75A2ABD7F}
; TIF Data AppID
AppId={{6D382664-1488-4A1E-900C-1F5DE3BC4AA5}
AppName=OpenNSPECT
AppVerName=OpenNSPECT
AppPublisher=NOAA Coastal Services Center
VersionInfoCompany=NOAA Coastal Services Center
VersionInfoDescription=NOAA Coastal Services Center
VersionInfoTextVersion=1.2.0
VersionInfoVersion=1.2.0
AppPublisherURL=http://www.csc.noaa.gov/
AppSupportURL=http://www.csc.noaa.gov/
AppUpdatesURL=http://www.csc.noaa.gov/

;The database references files at C:\NSPECT, so we will need to use that location.
; DLE 12/23/2011: Testing letting user define their own directory so changed the below lines
;DefaultDirName={userdocs}\OpenNSPECT
;DisableDirPage=no
DefaultDirName=C:\NSPECT
DisableDirPage=yes

DefaultGroupName=OpenNSPECT
EnableDirDoesntExistWarning=yes
OutputBaseFilename=OpenNSPECT_Setup
AllowNoIcons=yes
DisableProgramGroupPage=yes
;LicenseFile=License.rtf
Compression=lzma/ultra
SolidCompression=yes
InternalCompressLevel=ultra

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Components]
Name: main; Description: "MapWindow Plugin"; ExtraDiskSpaceRequired: 0; Flags: fixed; Types: Standard_Installation
Name: sample_data; Description: "Sample database and project files"; Types: Standard_Installation;
Name: help; Description: "Help Files"; Types: Standard_Installation

[Dirs]
Name: {app}; Permissions: everyone-modify
Name: {app}\projects; Permissions: everyone-modify
Name: {app}\workspace; Permissions: everyone-modify
;Name: {app}\metadata; Permissions: everyone-modify

[Files]
;Source: "License.rtf"; DestDir: "{app}"; Flags: ignoreversion
; DLE 12/23/2011: Changing Base directory to use new tiff data instead of ESRI grids as in ESRI NSPECT
;Source: "Base Files\*"; Excludes: "*..svnbridge*,\help\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs uninsneveruninstall; Components: sample_data
;Source: "Base Files\help\*"; Excludes: "*..svnbridge*"; DestDir: "{app}\help"; Flags: ignoreversion; Components: help
;Source: "Base Files\bin\OpenNspect.MapWindow4Plugin.dll";  Check: GetMWPluginDestination; DestDir: "{code:PluginDestination}"; Flags: ignoreversion overwritereadonly; Components: main
Source: "HI_Sample_Data\*"; Excludes: "*..svnbridge*,\help\*,\bin,coefficients\ccap2001 - Copy for import"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs uninsneveruninstall; Components: sample_data
Source: "HI_Sample_Data\help\*"; Excludes: "*..svnbridge*,NSPECT - v1.1OLD.chm,NSPECT_old_v1_1_.chm"; DestDir: "{app}\help"; Flags: ignoreversion; Components: help
Source: "HI_Sample_Data\bin\OpenNspect.MapWindow4Plugin.dll";  Check: GetMWPluginDestination; DestDir: "{code:PluginDestination}"; Flags: ignoreversion overwritereadonly; Components: main

[Registry]
;Installation Path
Root: HKLM; Subkey: "Software\OpenNSPECT"; ValueType: string; ValueName: InstallPath; ValueData: "{app}"; Flags: uninsdeletekey deletekey deletevalue createvalueifdoesntexist; Components: sample_data

[Types]
Name: Standard_Installation; Description: Standard OpenNSPECT installation; Flags: iscustom

[Messages]
SelectDirDesc=Where should [name] data files and uninstaller be placed?

[Run]
; Make sure the windows installer gets installed before you try
; to use it below (msiexec)


[Code]

var sPluginDest : String;

//
// Search for the path where MW was installed.  Return true if path found.
// Set variable to plugin folder
//

function GetMWPluginDestination(): Boolean;
var
  i:      Integer;
  len:    Integer;

begin
  sPluginDest := '';

  RegQueryStringValue( HKLM, 'SOFTWARE\Classes\MapWindowConfig\DefaultIcon', '', sPluginDest );
  len := Length(sPluginDest);
  if len > 0 then
  begin
    i := len;
    while sPluginDest[i] <> '\' do
      begin
        i := i-1;
      end;

    i := i+1;
    Delete(sPluginDest, i, Len-i+1);
    Insert('plugins\OpenNSPECT\', sPluginDest, i);
  end;
  len := Length(sPluginDest);
  Result := len > 0;
end;

//
//  Use this function to return path to install plugin
//
function PluginDestination(Param: String) : String;
begin
   Result := sPluginDest;
end;

