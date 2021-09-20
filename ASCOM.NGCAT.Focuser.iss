[Setup]
AppID={{835503b2-4cbd-4675-b5ad-40f2af06180a}}
AppName=ASCOM NGCAT Focuser Driver for Robofocus
AppVerName=ASCOM NGCAT Focuser Driver 1.0.5
AppVersion=1.0.5
AppPublisher=Cedric Raguenaud <cedric@raguenaud.earth>
AppPublisherURL=mailto:cedric@raguenaud.earth
AppSupportURL=https://github.com/cedric-r/ASCOM.NGCAT.Focuser
AppUpdatesURL=https://github.com/cedric-r/ASCOM.NGCAT.Focuser
VersionInfoVersion=1.0.5
MinVersion=0,6.0
DefaultDirName="{cf}\ASCOM\Focuser"
DisableDirPage=yes
DisableProgramGroupPage=yes
OutputDir="."
OutputBaseFilename="NGCAT Focuser Setup"
Compression=lzma
SolidCompression=yes
; Put there by Platform if Driver Installer Support selected
WizardImageFile="C:\Program Files (x86)\ASCOM\Platform 6 Developer Components\Installer Generator\Resources\WizardImage.bmp"
LicenseFile="k:\astro\ngcat\ASCOM.NGCAT.Focuser\License"
; {cf}\ASCOM\Uninstall\Focuser folder created by Platform, always
UninstallFilesDir="{cf}\ASCOM\Uninstall\Focuser\NGCAT Focuser"

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Dirs]
Name: "{cf}\ASCOM\Uninstall\Focuser\NGCAT Focuser"
; TODO: Add subfolders below {app} as needed (e.g. Name: "{app}\MyFolder")

[Files]
Source: "k:\astro\ngcat\ASCOM.NGCAT.Focuser\ASCOM.NGCAT.Focuser\bin\Debug\ASCOM.NGCAT.Focuser.dll"; DestDir: "{app}"
; Require a read-me HTML to appear after installation, maybe driver's Help doc
Source: "k:\astro\ngcat\ASCOM.NGCAT.Focuser\readme.txt"; DestDir: "{app}"; Flags: isreadme
; TODO: Add other files needed by your driver here (add subfolders above)
Source: "k:\astro\ngcat\ASCOM.NGCAT.Focuser\ASCOM.NGCAT.Focuser\bin\Debug\ASCOM.NGCAT.Focuser.dll.config"; DestDir: "{app}"
Source: "k:\astro\ngcat\ASCOM.NGCAT.Focuser\ASCOM.NGCAT.Focuser\bin\Debug\ASCOM.DeviceInterfaces.dll"; DestDir: "{app}"
Source: "k:\astro\ngcat\ASCOM.NGCAT.Focuser\ASCOM.NGCAT.Focuser\bin\Debug\ASCOM.DeviceInterfaces.xml"; DestDir: "{app}"
Source: "k:\astro\ngcat\ASCOM.NGCAT.Focuser\ASCOM.NGCAT.Focuser\bin\Debug\ASCOM.Utilities.dll"; DestDir: "{app}"
Source: "k:\astro\ngcat\ASCOM.NGCAT.Focuser\ASCOM.NGCAT.Focuser\bin\Debug\ASCOM.Utilities.xml"; DestDir: "{app}"
Source: "k:\astro\ngcat\ASCOM.NGCAT.Focuser\ASCOM.NGCAT.Focuser\bin\Debug\ASCOM.Exceptions.dll"; DestDir: "{app}"
Source: "k:\astro\ngcat\ASCOM.NGCAT.Focuser\ASCOM.NGCAT.Focuser\bin\Debug\ASCOM.Exceptions.xml"; DestDir: "{app}"
Source: "k:\astro\ngcat\ASCOM.NGCAT.Focuser\ASCOM.NGCAT.Focuser\bin\Debug\ASCOM.Utilities.Video.dll"; DestDir: "{app}"
Source: "k:\astro\ngcat\ASCOM.NGCAT.Focuser\ASCOM.NGCAT.Focuser\bin\Debug\ASCOM.Astrometry.dll"; DestDir: "{app}"
Source: "k:\astro\ngcat\ASCOM.NGCAT.Focuser\ASCOM.NGCAT.Focuser\bin\Debug\ASCOM.Astrometry.xml"; DestDir: "{app}"
Source: "k:\astro\ngcat\ASCOM.NGCAT.Focuser\ASCOM.NGCAT.Focuser\bin\Debug\ASCOM.SettingsProvider.dll"; DestDir: "{app}"
Source: "k:\astro\ngcat\ASCOM.NGCAT.Focuser\ASCOM.NGCAT.Focuser\bin\Debug\ASCOM.SettingsProvider.xml"; DestDir: "{app}"
Source: "k:\astro\ngcat\ASCOM.NGCAT.Focuser\ASCOM.NGCAT.Focuser\bin\Debug\ASCOM.Attributes.dll"; DestDir: "{app}"
Source: "k:\astro\ngcat\ASCOM.NGCAT.Focuser\ASCOM.NGCAT.Focuser\bin\Debug\ASCOM.Attributes.xml"; DestDir: "{app}"
Source: "k:\astro\ngcat\ASCOM.NGCAT.Focuser\ASCOM.NGCAT.Focuser\bin\Debug\ASCOM.Controls.dll"; DestDir: "{app}"
Source: "k:\astro\ngcat\ASCOM.NGCAT.Focuser\ASCOM.NGCAT.Focuser\bin\Debug\ASCOM.Controls.xml"; DestDir: "{app}"
Source: "k:\astro\ngcat\ASCOM.NGCAT.Focuser\ASCOM.NGCAT.Focuser\bin\Debug\ASCOM.Internal.Extensions.dll"; DestDir: "{app}"
Source: "k:\astro\ngcat\ASCOM.NGCAT.Focuser\ASCOM.NGCAT.Focuser\bin\Debug\ASCOM.Internal.Extensions.xml"; DestDir: "{app}"

; Only if driver is .NET
[Run]
; Only for .NET assembly/in-proc drivers
Filename: "{dotnet4032}\regasm.exe"; Parameters: "/codebase ""{app}\ASCOM.NGCAT.Focuser.dll"""; Flags: runhidden 32bit
Filename: "{dotnet4064}\regasm.exe"; Parameters: "/codebase ""{app}\ASCOM.NGCAT.Focuser.dll"""; Flags: runhidden 64bit; Check: IsWin64




; Only if driver is .NET
[UninstallRun]
; Only for .NET assembly/in-proc drivers
Filename: "{dotnet4032}\regasm.exe"; Parameters: "-u ""{app}\ASCOM.NGCAT.Focuser.dll"""; Flags: runhidden 32bit
; This helps to give a clean uninstall
Filename: "{dotnet4064}\regasm.exe"; Parameters: "/codebase ""{app}\ASCOM.NGCAT.Focuser.dll"""; Flags: runhidden 64bit; Check: IsWin64
Filename: "{dotnet4064}\regasm.exe"; Parameters: "-u ""{app}\ASCOM.NGCAT.Focuser.dll"""; Flags: runhidden 64bit; Check: IsWin64




[CODE]
//
// Before the installer UI appears, verify that the (prerequisite)
// ASCOM Platform 6.2 or greater is installed, including both Helper
// components. Utility is required for all types (COM and .NET)!
//
function InitializeSetup(): Boolean;
var
   U : Variant;
   H : Variant;
begin
   Result := TRUE;  // Assume failure
   // check that the DriverHelper and Utilities objects exist, report errors if they don't
   // try
   //   H := CreateOLEObject('DriverHelper.Util');
   // except
   //   MsgBox('The ASCOM DriverHelper object has failed to load, this indicates a serious problem with the ASCOM installation', mbInformation, MB_OK);
   // end;
   // try
   //   U := CreateOLEObject('ASCOM.Utilities.Util');
   // except
   //   MsgBox('The ASCOM Utilities object has failed to load, this indicates that the ASCOM Platform has not been installed correctly', mbInformation, MB_OK);
   // end;
   // try
   //   if (U.IsMinimumRequiredVersion(6,2)) then	// this will work in all locales
   //      Result := TRUE;
   // except
   // end;
   // if(not Result) then
   //   MsgBox('The ASCOM Platform 6.2 or greater is required for this driver.', mbInformation, MB_OK);
end;

// Code to enable the installer to uninstall previous versions of itself when a new version is installed
procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
  UninstallExe: String;
  UninstallRegistry: String;
begin
  if (CurStep = ssInstall) then // Install step has started
	begin
      // Create the correct registry location name, which is based on the AppId
      UninstallRegistry := ExpandConstant('Software\Microsoft\Windows\CurrentVersion\Uninstall\{#SetupSetting("AppId")}' + '_is1');
      // Check whether an extry exists
      if RegQueryStringValue(HKLM, UninstallRegistry, 'UninstallString', UninstallExe) then
        begin // Entry exists and previous version is installed so run its uninstaller quietly after informing the user
          MsgBox('Setup will now remove the previous version.', mbInformation, MB_OK);
          Exec(RemoveQuotes(UninstallExe), ' /SILENT', '', SW_SHOWNORMAL, ewWaitUntilTerminated, ResultCode);
          sleep(1000);    //Give enough time for the install screen to be repainted before continuing
        end
  end;
end;

