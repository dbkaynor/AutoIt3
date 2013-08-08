#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseUpx=n
#AutoIt3Wrapper_UseAnsi=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#comments-start PROGRAM HEADER
	;******************************************************************************************
	;** Intel Corporation, MPG MPAD
	;** Title			:  		incMisc.au3
	;** Description	:
	;**		Misc function include file.
	;**
	;** Revision: 	Rev 2.0.64
	;******************************************************************************************
	;******************************************************************************************
	;** Revision History:
	;**
	;** Update for Rev 2.0.0		- Dick Lin 03/02/2006
	;**		- Initial release
	;**
	;** Update for Rev 2.0.1		- Dick Lin 08/02/2006
	;**		- Added GetHDAudioDriverSourceDir() function.
	;**
	;** Update for Rev 2.0.2		- Dick Lin 09/22/2006
	;**		- ScriptStarting() use @WorkingDir for GetFileTimeStamp() function.
	;**
	;** Update for Rev 2.0.3		- Dick Lin 11/16/2006
	;** 	- Added GetInstallScriptExe() function.
	;**		- Modified RunVManager() function to support redesing test script.
	;**
	;** Update for Rev 2.0.4		- Dick Lin 11/27/2006
	;**		- Added GetAppletExe() function.
	;**
	;** Update for Rev 2.0.5		- Dick Lin 12/15/2006
	;**		- Added ($CmdLine[1] <> "-InstallCHECK") check for RunVManager.
	;**
	;** Update for Rev 2.0.6		- Dick Lin 	01/08/2007
	;**		- Added PCMark2005 for Vista
	;**		- Added PCMark2005
	;**
	;** Update for Rev 2.0.7		- Dick Lin	01/23/2007
	;**		- Added Power Management Test Tool
	;**
	;** Update for Rev 2.0.8		- Dick Lin 	01/24/2007
	;**		- Added PwrTestGUI() from TmPTSleep.au3
	;**
	;** Update for Rev 2.0.9		- Dick Lin 02/01/2007
	;**		- Added ScriptEndingRebootVManagerUI() function.
	;**
	;** Update for Rev 2.0.10		- Dick Lin 02/06/2007
	;**		- Added GetISTTps() function.
	;**
	;** Update for Rev 2.0.11		- Dick Lin 02/07/2007
	;**		- Added IsAppletInstalled() function.
	;**
	;** Update for Rev 2.0.12		- Dick Lin 02/13/2007
	;**		- GetInstallScriptExe() funciton added IBASES 2.0.
	;**
	;** Update for Rev 2.0.13		- Dick Lin 02/15/2007
	;**		- GetInstallScriptExe() funciton added "Microsoft WHQL DTM(Driver Test Manager) Log Viewer"
	;**
	;** Update for Rev 2.0.14		- Dick Lin 02/21/2007
	;**		- Added _IsPressed() to check Enter key.
	;**		- Replace Select with Switch.
	;**
	;** Update for Rev 2.0.15		- Dick Lin 03/06/2007
	;**		- Added GetIGDSourceDir() function.
	;**
	;** Update for Rev 2.0.16		- Dick Lin 03/17/2007
	;**		- Added BusyIdleGUI() function.
	;**		- Added KillTask() function.
	;**
	;** Update for Rev 2.0.17		- Dick Lin 05/15/2007
	;**		- Added PCMk05BusyIdleGUI() function.
	;**
	;** Update for Rev 2.0.18		- Dick Lin 06/11/2007
	;**		- Add GetTATSourceDir() Cantiga support.
	;**		- Added "SysMark 2007" for IsAppletInstalled() function.
	;**		- Added "Microsoft .NET Framework V2.0" for GetInstallScriptExe() function.
	;**
	;** Update for Rev 2.0.l9		- Dick Lin 06/18/2007
	;**		- Added WriteBIdleLog() function.
	;**
	;** Update for Rev 2.0.20		- Dick Lin 06/22/2007
	;**		- Added InstallSP2XDHotfix() function.
	;**
	;** Update for Rev 2.0.21		- Dick Lin 06/26/2007
	;**		- Use $BS_DEFPUSHBUTTON for Enter key press.
	;**		- Fixed GetIGDSourceDir() didn't return SourceDir issue.
	;**
	;** Update for Rev 2.0.22		- Michelle Tran 07/10/2007
	;**		- Added "CAMARILLO SECTION" to IsAppletInstalled() function
	;**        	- Added Intel Extended Thermal Model to Switch Case
	;**			- Added Intel TPT to Switch Case
	;**			- Added Intel ETMDemo to Switch Case
	;**			- Added TestApp2 to Switch Case
	;**			- Added Graphics Frequency Display to Switch Case
	;**			- Added Intel GfX_DX9_Stress to Switch Case
	;**
	;**	Update for Rev 2.0.23		- Dick Lin 07/10/2007
	;**		- Update IsAppletInstalled($strScriptName) for Glaze.
	;**
	;** Update for Rev 2.0.24		- Dick Lin 07/25/2007
	;**		- Added "WMI Code Creator"
	;**
	;** Update for Rev 2.0.25		- Dick Lin 08/09/2007
	;**		- Fixed GetIGDSourceDir() undefined $strResult for Ohlone.
	;**
	;** Update for Rev 2.0.26		- Dick Lin 08/14/2007
	;**		- Removed LICENSED-CV SECTION from IsAppletInstalled() function.
	;**
	;** Update for Rev 2.0.27		- Dick Lin 08/15/2007
	;**		- Modified GetIGDSourceDir() to support Cantiga IGD.
	;**		- Added "Intel Graphics Media Accelerator Driver for Cantiga" to IsAppletInstalled() function.
	;**
	;** Update for Rev 2.0.28		- Dick Lin 08/21/2007
	;**		- Added "Windows Vista Update (KB938194)" for IsAppletInstalled() function.
	;**		- Added "Windows Vista Update (KB938979)" for IsAppletInstalled() function.
	;**		- Added InstallVistaKB938194() function.
	;**		- Added InstallVistaKB938979() function.
	;**
	;** Update for Rev 2.0.29		- Dick Lin 08/30/2007
	;**		- Added couple UAC related functions.
	;**
	;** Update for Rev 2.0.30		- Dick Lin 10/15/2007
	;**		- Added MobileMark 2007 for IsAppletInstalled().
	;**
	;** Update for Rev 2.0.31		- Dick Lin 12/13/2007
	;**		- Added InstallXPKB940566() fuction.
	;**
	;** Update for Rev 2.0.32		- Dick Lin 02/22/2008
	;**		- Added IsWMPlayer10Installed()
	;**
	;** Update for Rev 2.0.33		- André Nadeau 6/3/2008
	;**		- Added Unreal Tournament 3 Demo Test support
	;**
	;** Update for Rev 2.0.34		- Jarek Szymanski 7/17/2008
	;**		- Added 3DMark Vantage for IsAppletInstalled()
	;**
	;** Update for Rev 2.0.35		- Jarek Szymanski 7/24/2008
	;**		- Added PCMark Vantage for IsAppletInstalled()
	;**		- Changed InterVideoDVD WinDVD IsAppletInstalled() case; Corel WinDVD case - moved from version 8 to 9
	;**
	;** Update for Rev 2.0.36		- Jarek Szymanski 7/28/2008
	;**		- Changed CyberLink PowerDVD IsAppletInstalled() case; Corel WinDVD case - moved to version 8
	;**
	;**	Update for Rev 2.0.37		- Jarek Szymanski 8/6/2008
	;**		- Updated RebootSystem() entry to ShowMessage("Rebooting... ", $strMsg)
	;**		- it makes more sense displayed on the screen
	;**
	;**	Update for Rev 2.0.38		- Jarek Szymanski 8/22/2008
	;**		- added entry to GetInstallScriptExe() for: DebugView, TXT Driver, SINIT ACM driver, TXT MPG App test script and 56. MPG TXT TEST
	;**		- added entry to IsAppletInstalled() for: DebugView, TXT Driver, SINIT ACM driver and TXT MPG App test script
	;**
	;**	Update for Rev 2.0.39		- Jarek Szymanski 8/26/2008
	;**		- fixed some spelling errors
	;**
	;**	Update for Rev 2.0.40		- Jarek Szymanski 8/27/2008
	;**		- added HeavyLoad section in IsAppletInstalled() function
	;**		- added entries for Cinebench R10
	;**		- removed entry for Cinebench 2003 from IsAppletInstalled() function
	;**
	;**	Update for Rev 2.0.41		- Jarek Szymanski 9/02/2008
	;**		- added entry to GetInstallScriptExe() for: HeavyLoad install script and 57. HeavyLoad Stress Test
	;**		- added entry to IsAppletInstalled() for: HeavyLoad install script and 57. HeavyLoad Stress Test script
	;**
	;**	Update for Rev 2.0.42		- Jarek Szymanski 9/04/2008
	;**		- added entry to GetInstallScriptExe() for: 3DMark Vantage install script and 57. HeavyLoad Stress Test
	;**
	;**	Update for Rev 2.0.43		- Jarek Szymanski 9/11/2008
	;**		- added entry to GetInstallScriptExe() for PCMark Vantage install script
	;**
	;**	Update for Rev 2.0.44		- André Nadeau 6/3/2008
	;**		- Added X3 Terran Conflict rolling Demo Test support
	;**
	;**	Update for Rev 2.0.45		- Jarek Szymanski 11/12/2008
	;**		- changed entry in IsApplet Installed() for Frequency Display Tool
	;**
	;**	Update for Rev 2.0.46		- Jarek Szymanski 11/18/2008
	;**		- added support in IsApplet Installed() for AutoMate6
	;**
	;** Update for Rev 2.0.47		- Andre Nadeau 1/26/09
	;**		- updated the path for install ProcLoad 1.1
	;**		- GVCycle and CSwitch
	;**
	;** Update for Rev 2.0.48		- Jarek Szymanski 2/3/09
	;**		- updated ISAppletInstalled() for .NET 3.0 case
	;**		- added ImdotNET30.exe to GetInstallScriptExe()
	;**
	;** Update for Rev 2.0.49		- Jarek Szymanski 2/17/09
	;**		- updated GetIGDSourceDir(): added Win 7 support: IsWin7_32 and IsWin7_64
	;**		- updated GetIGDSourceDir(): changed folder paths according to locations on Chakotay
	;**		- updated IsAppletInstalled() for Auburndale graphics install script
	;**
	;** Update for Rev 2.0.50		- Jarek Szymanski 2/25/09
	;**		- added ImdotNET30_64.exe to GetInstallScriptExe() and install check for IsAppletInstalled()
	;**
	;** Update for Rev 2.0.51		- Jarek Szymanski 3/5/09
	;**		- added new file name for Intel C State Residency Monitor Utility for IsAppletInstalled()
	;**
	;** Update for Rev 2.0.52		- Jarek Szymanski 3/10/09
	;**		- updated path for .NET 2.0 install check in IsAppletInstalled() function
	;**
	;** Update for Rev 2.00.53		- Jarek Szymanski 3/13/2009
	;**	- updated 'X3 Terran Conflict Demo' name to ALL capital letters to capital letters
	;**
	;** Update for Rev 2.00.54		- Jarek Szymanski 3/20/2009
	;**	- added WinFlash to IsAppletInstalled()
	;**
	;** Update for Rev 2.00.55		- Jarek Szymanski 3/27/09
	;**		- updated GetIGDSourceDir(): source files for Cantiga
	;**
	;** Update for Rev 2.00.56		- Jarek Szymanski 3/31/09
	;**		- updated GetIGDSourceDir(): source files for Crestline and Calistoga
	;**
	;** Update for Rev 2.0.57		- Jarek Szymanski 4/2/2009
	;**		- moved on to AutoIt 3.3.00
	;**		- added #include: <EditConstants.au3>; <WindowsConstants.au3>
	;**
	;** Update for Rev 2.0.58 		- Jarek Szymanski 4/9/2009
	;**		- added _VersionNumber($strDestDir): function checks for *.vmgr file in app install folder and gets information about tool name and it's version
	;**		  and writes this information to Manager.txt log file
	;**
	;** Update for Rev 2.0.59		- Jarek Szymanski 05/7/2009
	;**	- adjusted _Version() format
	;**	- added MEI SOL Driver support in IsAppletInstalled()
	;**
	;** Update for Rev 2.0.60		- Jarek Szymanski 08/04/2009
	;**	- added Real-Time HDR Image-Based Lighting for GetInstallScriptExe() function
	;**
	;** Update for Rev 2.00.61		- Jarek Szymanski 8/13/09
	;**		- updated GetIGDSourceDir(): source files for SandyBridge
	;**		- updated GetIGDSourceDir(): source files for Pineview changed for Win 7 from Vista source folder to Win 7 source folder
	;**
	;** Update for Rev 2.00.62		- Michelle Tran		12/23/2009
	;**	- added Intel Turbo Boost Technology Driver for GetInstallScriptExe() function
	;**
	;** Update for Rev 2.00.63		- Jarek Szymanski 1/14/10
	;**		- updated GetIGDSourceDir(): source files for SandyBridge-CPT (Cougar Point)
	;**
	;** Update for Rev 2.00.64		- Jarek Szymanski 1/28/10
	;**		- updated GetIGDSourceDir(): source files for SandyBridge-CPT (Cougar Point) - changed Win 7 dirs temoprary for Vista
	;**
	;******************************************************************************************
#comments-end HEADER

#include-once

#include <GUIConstants.au3>
#include <EditConstants.au3> ;added for AutoIt 3.3.0.0
#include <WindowsConstants.au3> ;added for AutoIt 3.3.0.0
#include <Date.au3>
#include <incOS.au3>
#include <Misc.au3>
#include <Array.au3>
#include <File.au3>

;***************************************************************************
;** Function: 		WriteLog(strMsg)
;** Parameters:
;**		$strmsg - string to write into log file.
;** Description:
;**		This function is called to write string to log file - c:\logs\manager.txt.
;** Return:
;**		None
;** Usage:
;**		WriteLog("message you want write")
;**
;***************************************************************************
Func WriteLog($strMsg)

	$strLogFile = @HomeDrive & "\Logs\Manager.txt"

	; Open file for append. If file not exists, will create
	$strFileHandle = FileOpen($strLogFile, 1)

	$strMsg = _Now() & @TAB & @TAB & $strMsg
	FileWriteLine($strFileHandle, $strMsg)

	FileClose($strFileHandle)

EndFunc   ;==>WriteLog

Func _WriteLogVersion($strMsg)

	$strLogFile = @HomeDrive & "\Logs\Version.txt"

	; Open file for append. If file not exists, will create
	$strFileHandle = FileOpen($strLogFile, 1)

	;$strMsg = _Now() & @TAB & @TAB & $strMsg
	FileWriteLine($strFileHandle, $strMsg)

	FileClose($strFileHandle)

EndFunc   ;==>_WriteLogVersion

;***************************************************************************
;** Function: 		GetFileTimeStamp($fileName)
;** Parameters:
;**		$strFileName - the file name you want get the timestamp.
;** Description:
;**		This function is called to get file's modification timestamp.
;** Return:
;**		A string contains modification timestamp of the file.
;** Usage:
;**		$timeStamp = GetFileTimeStamp("c:\dir\filename.ext")
;**
;***************************************************************************
Func GetFileTimeStamp($strFileName)
	$strValue = ""
	$strTime = FileGetTime($strFileName, 0)
	If Not @error Then
		$strValue = $strTime[1] & "/" & $strTime[2] & "/" & $strTime[0] & " " & $strTime[3] & ":" & $strTime[4] & ":" & $strTime[5]
	Else
		WriteLog("GetFileTimeStamp ERROR")
	EndIf

	Return $strValue

EndFunc   ;==>GetFileTimeStamp

;***************************************************************************
;** Function: 		InstallAlready($strFileName)
;** Parameters:
;**		$strFileName - file name used to check
;** Description:
;**		This function is called to check if application already installed.
;** Return:
;**		1 for success, otherwise 0.
;** Usage:
;**		; If app already installed, abort this install script. Write error to log and registry.
;**		$strFileName = @ProgramFilesDir & "\Intel Corporation\Frequency Display\FreqDsp.exe"
;**		If InstallAlready($strFileName) then
;**			...
;**
;***************************************************************************
Func InstallAlready($strFileName)
	Return FileExists($strFileName)
EndFunc   ;==>InstallAlready

;***************************************************************************
;** Function: 		SetWin2KAutoLogon()
;** Parameters:
;**		None
;** Description:
;**		This function is called to setup auto logon for Win2K
;** Return:
;**		None
;** Usage:
;**		SetWin2KAutoLogon()
;**
;***************************************************************************
Func SetWin2KAutoLogon()
	If IsWin2KFamily() Then
		$strKey = "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon"
		$strSubKey = "AutoAdminLogon"
		RegWrite($strKey, $strSubKey, "REG_SZ", "1")

		$strKey = "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon"
		$strSubKey = "AutoLogonCount"
		RegWrite($strKey, $strSubKey, "REG_DWORD", 0x1000)
	EndIf
EndFunc   ;==>SetWin2KAutoLogon

;***************************************************************************
;** Function: 		SetupRestartApp($strApp)
;** Parameters:
;**		strApp - full path application name to restart after reboot.
;** Description:
;**		This function is called to setup restart apps after system reboot.
;** Return:
;**		None
;** Usage:
;**		SetupRestartApp(strApp)
;**
;***************************************************************************
Func SetupRestartApp($strApp)

	;$strKey	= "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
	$strKey = "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
	$strSubKey = "VManager"
	$intResult = RegWrite($strKey, $strSubKey, "REG_SZ", $strApp)

	; Write log
	WriteLog("Set restart applet : " & $strApp & " Return Value: " & $intResult)
EndFunc   ;==>SetupRestartApp

;***************************************************************************
;** Function: 		SetRestartVManager()
;** Parameters:
;**		None
;** Description:
;**		This function is called to start VManager without UI display.
;** Return:
;**		None
;** Usage:
;**		SetRestartVManager()
;**
;***************************************************************************
Func SetRestartVManager()

	$strApp = @HomeDrive & "\Scripts\VManager.exe -NOUI"
	SetupRestartApp($strApp)
EndFunc   ;==>SetRestartVManager

;***************************************************************************
;** Function: 		SetRestartVManagerUI()
;** Parameters:
;**		None
;** Description:
;**		This function is called to start VManager with UI display.
;** Return:
;**		None
;** Usage:
;**		SetRestartVManagerUI()
;**
;***************************************************************************
Func SetRestartVManagerUI()

	$strApp = @HomeDrive & "\Scripts\VManager.exe"
	SetupRestartApp($strApp)
EndFunc   ;==>SetRestartVManagerUI

;***************************************************************************
;** Function: 		RebootSystem()
;** Parameters:
;**		None
;** Description:
;**		This function is called to reboot system.
;** Return:
;**		None
;** Usage:
;**		RebootSystem()
;**
;***************************************************************************
Func RebootSystem()
	; Display
	$strMsg = "System will reboot in few seconds."
	ShowMessage("Rebooting... ", $strMsg)

	Shutdown(6) ;Force a reboot
	Exit
EndFunc   ;==>RebootSystem

;***************************************************************************
;** Function: 		CleanupCallRunVManager()
;** Parameters:
;**		None
;** Description:
;**		This function is used by cleanup routine to run VManager without command line switch.
;** Return:
;**		None
;** Usage:
;**		CleanupCallRunVManager()
;**
;***************************************************************************
Func CleanupCallRunVManager()

	$strVManager = @HomeDrive & "\Scripts\VManager.exe"
	$strDir = @HomeDrive & "\Scripts"

	If Not FileExists($strVManager) Then
		WriteLog("CleanupCallRunVManager: VManager.exe does not exist.")
		Return
	EndIf

	; Run VManager if file exists
	;Run($strVManager, $strDir)
	Run(@ComSpec & " /c " & $strVManager, "", @SW_HIDE)
EndFunc   ;==>CleanupCallRunVManager

;***************************************************************************
;** Function: 		RunVManager()
;** Parameters:
;**		None
;** Description:
;**		This function is used by script to run VManager.
;** Return:
;**		None
;** Usage:
;**		RunVManager()
;**
;***************************************************************************
Func RunVManager()
	WriteLog("RunVManager() function.")

	$strVMgr = @HomeDrive & "\scripts\VManager.exe"
	$strInstallTemp = @HomeDrive & "\scripts\InstTemp.csv"
	$strTestTemp = @HomeDrive & "\scripts\TestTemp.csv"
	$strInstallTemp2 = @HomeDrive & "\scripts\InstallTemp.xml"
	$strTestTemp2 = @HomeDrive & "\scripts\TestTemp.xml"

	;MsgBox(0, "RunVManager", $CmdLine[0] & " " & $CmdLine[1])
	; TO DO: TEMP SOLUTION
	If (($CmdLine[0] > 0) And ($CmdLine[1] <> "-InstallCHECK")) Then
		$strExeName = ""
		If GetInstallScriptExe($CmdLine[1], $strExeName) Then
			$strCmd = @ComSpec & " /c " & $strExeName & ' "' & $CmdLine[2] & '"'
			WriteLog($strCmd)
			Run($strCmd, "", @SW_HIDE)
		Else
			MsgBox(0, "ERROR", $CmdLine[0] & " " & $CmdLine[1])
		EndIf
	Else
		; Only call VManager with 2 temp csv files exist and VManager.exe exists.
		If Not FileExists($strVMgr) Then
			Return
		ElseIf Not ((FileExists($strInstallTemp) And FileExists($strTestTemp)) Or (FileExists($strInstallTemp2) And FileExists($strTestTemp2))) Then
			Exit
		EndIf

		; Only run VManager.exe if the script is called by VManager.
		$strVManager = @HomeDrive & "\Scripts\VManager.exe"
		Run(@ComSpec & " /c " & $strVManager & " -NOUI", "", @SW_HIDE)
	EndIf
EndFunc   ;==>RunVManager

;***************************************************************************
;** Function: 		CloseISTWindow()
;** Parameters:
;**		None
;** Description:
;**		This function is called to close the IST window popup after Win2K system boot.
;** Return:
;**		None
;** Usage:
;**		CloseISTWindow()
;**
;***************************************************************************
Func CloseISTWindow()
	; Set WinTitleMatchMode to 2 - 1=start, 2=subStr, 3=exact, 4=advanced
	Opt("WinTitleMatchMode", 2)

	; Close IST Winodow
	$strTitle = "Intel SpeedStep(R) technology"
	;WinClose($strTitle)
	$strText = "Intel(R) SpeedStep(TM) technology Applet"
	If WinExists($strTitle, $strText) Then
		ControlClick($strTitle, "OK", "OK")
	EndIf
EndFunc   ;==>CloseISTWindow

;***************************************************************************
;** Function: 		ExitWithErrorMessage($strScriptName, $strError)
;** Parameters:
;**		$strScriptName - script name call this function
;**		$strError - error message from the caller
;** Description:
;**		This function is called to delete registry string value.
;** Return:
;**		None
;** Usage:
;** 	ExitWithErrorMessage(SCRIPT_NAME, "ERROR: can't copy from server")
;**
;***************************************************************************
Func ExitWithErrorMessage($strScriptName, $strError)
	RecordErrorMessage($strScriptName, $strError)

	; Run VManager
	RunVManagerError()
EndFunc   ;==>ExitWithErrorMessage

;***************************************************************************
;** Function: 		ShowMessageWithoutHeader($strScriptName, $strValue)
;** Parameters:
;**		$strScriptName - script name called this function
;**		$strValue - string to display/write to log.
;** Description:
;**		This function is called to display message to user and write to log file
;**		without script name in front.
;** Return:
;**		None
;** Usage:
;**		ShowMessage($strScriptName, $strValue)
;**
;***************************************************************************
Func ShowMessageWithoutHeader($strScriptName, $strValue)
	;strValue = strCat(strScriptName, strValue)
	WriteLog($strValue)
	ShowMessage($strScriptName, $strValue)
EndFunc   ;==>ShowMessageWithoutHeader

;***************************************************************************
;** Function: 		ShowMessage($strScriptName, $strValue)
;** Parameters:
;**		None
;** Description:
;**		This function is called to display message to user and write to log file.
;** Return:
;**		None
;** Usage:
;**		ShowMessage($strScriptName, $strValue)
;**
;***************************************************************************
Func ShowMessage($strScriptName, $strValue)
	$strMsg = $strScriptName & " " & $strValue
	WriteLog($strMsg)

	; Splash text
	;SplashTextOn($strScriptName, $strMsg, 500, 75)
	SplashTextOn($strScriptName, $strMsg, 450, 50, -1, -1, 2, "", 10)
	Sleep(4000)
	SplashOff()
EndFunc   ;==>ShowMessage

;***************************************************************************
;** Function: 		RecordErrorMessage($strScriptName, $strError)
;** Parameters:
;**		$strScriptName - script calling this function
;**		$strError - error message to write/show.
;** Description:
;**		This function is called to write error message to register and display to user.
;** Return:
;**		None
;** Usage:
;**
;***************************************************************************
Func RecordErrorMessage($strScriptName, $strError)
	$strKey = "HKLM\Software\Intel\MPG\TESTAPP"
	$strSubKey = $strScriptName
	$strPreloadServer = RegWrite($strKey, $strSubKey, "REG_SZ", $strError)

	If @error Then
		WriteLog("UpdateRegistryStatus Function Failed.")
	EndIf

	; Show message
	ShowMessageWithoutHeader($strScriptName, $strError)

	; Write ending message to log file. Display to user.
	ShowMessage($strScriptName, ": Ending Script Execution")

EndFunc   ;==>RecordErrorMessage

;***************************************************************************
;** Function: 		AbortProgram($strScriptName)
;** Parameters:
;**		$strScriptName - script name call this routine.
;** Description:
;**		This function is called to allow use to quite the script running.
;** Return:
;**		None
;** Usage:
;**		AbortProgram(SCRIPT_NAME)
;**
;***************************************************************************
Global $Paused
Global $Escape
Global $terminateNow
Global $scriptNameforTerminate
Func AbortProgram($scriptName)
	; For the Terminate() function.
	$scriptNameforTerminate = $scriptName

	$msg = "Press Esc to terminate script, Pause/Break to pause within 3 seconds."
	SplashTextOn($scriptName, $msg, 450, 50, -1, -1, 2, "", 10)
	$begin = TimerInit()

	; Set PAUSE/ESC HotKeys
	HotKeySet("{PAUSE}", "TogglePause")
	HotKeySet("{ESC}", "TerminateScript")

	; 5 seconds to press ESC/PAUSE key
	While True
		$dif = TimerDiff($begin)

		If Not $Escape And Not $terminateNow And Not $Paused Then
			If $dif > 3000 Then
				ExitLoop
			EndIf
		EndIf
	WEnd

	SplashOff()

EndFunc   ;==>AbortProgram

;***************************************************************************
;** Function: 		TogglePause()
;** Parameters:
;**		None
;** Description:
;**		This function is called when user press PAUSE key to pause script running.
;** Return:
;**		None
;** Usage:
;**		This function is set during the HotKeySet() function.
;**
;***************************************************************************
Func TogglePause()
	$Paused = Not $Paused
	While $Paused
		Sleep(100)
		ToolTip('Script is "Paused"', 250, 250)
	WEnd
	ToolTip("")
EndFunc   ;==>TogglePause

;***************************************************************************
;** Function: 		TerminateScript()
;** Parameters:
;**		None
;** Description:
;**		This function is called when user press ESC key to quit script running
;** Return:
;**		None
;** Usage:
;**		This function is set during the HotKeySet() function.
;**
;***************************************************************************
Func TerminateScript()
	$Escape = 1
	$buttonClick = MsgBox(4, "Aborting Program Execution !!!", "Are you sure you want to terminate program ?")
	If $buttonClick == 6 Then
		$terminateNow = 1
		$strError = "TERMINATED: User selected to terminate program."
		ExitWithErrorMessage($scriptNameforTerminate, $strError)
		Exit
	EndIf
	$Escape = 0
EndFunc   ;==>TerminateScript

;***************************************************************************
;** Function: 		DirExists($strPath)
;** Parameters:
;**		None
;** Description:
;**		This function is called to determine a directory exists or not.
;** Return:
;**		1 for directory exists, otherwise 0.
;** Usage:
;**		This function is set during the HotKeySet() function.
;**
;***************************************************************************
Func DirExists($strPath)
	$intResult = DirGetSize($strPath)

	If $intResult == -1 Then
		Return False
	Else
		Return True
	EndIf

EndFunc   ;==>DirExists

;***************************************************************************
;** Function: 		CloseFoundNewHardwareWindow()
;** Parameters:
;**	None
;** Description:
;**	This function is called to close the "Found New Hardware" window.
;** Return:
;**	None
;** Usage:
;**	CloseFoundNewHardwareWindow()
;**
;***************************************************************************
Func CloseFoundNewHardwareWindow()
	$strName = "Found New Hardware Wizard"

	If WinExists($strName) Then
		ControlClick($strName, "Cancel", 2)
	EndIf

EndFunc   ;==>CloseFoundNewHardwareWindow

;***************************************************************************
;** Function: 		PrepareStarting($strScriptName, $strScriptVersion)
;** Parameters:
;**		$strScriptName - script name calling this function
;**		$strArgCount - command line arg count
;** Description:
;**		This function is called prepare running script. Add extender, unmap drives...
;** Return:
;**	None
;** Usage:
;**		PrepareForScriptRunning(SCRIPT_NAME, $strScriptVersion)
;**
;***************************************************************************
Func ScriptStarting($strScriptName, $strScriptExe, $strScriptVersion)

	; Write beginning message to log file. Display to user.
	WriteLog(" ")
;~ 	If ($strScriptExe == "VMLauncher.exe") Then
;~ 		$strFileTimestamp = GetFileTimeStamp(@HomeDrive & "\VMLauncher\" & $strScriptExe)
;~ 	Else
;~ 		$strFileTimestamp = GetFileTimeStamp(@HomeDrive & "\scripts\" & $strScriptExe)
;~ 	EndIf

	; Close 2 windows for Vista x64.
	CloseVistaWindow()

	If FileExists(@HomeDrive & "\scripts\" & $strScriptExe) Then
		$strFile = @HomeDrive & "\scripts\" & $strScriptExe
	ElseIf FileExists(@HomeDrive & "\VMLauncher\" & $strScriptExe) Then
		$strFile = @HomeDrive & "\VMLauncher\" & $strScriptExe
	ElseIf FileExists(@WorkingDir & "\" & $strScriptExe) Then
		$strFile = @WorkingDir & "\" & $strScriptExe
	Else
		$strFile = @WorkingDir & "\bin\" & $strScriptExe
	EndIf

	$strFileTimestamp = GetFileTimeStamp($strFile)
	WriteLog($strScriptName & " " & $strScriptVersion & " " & $strScriptExe & " " & $strFileTimestamp)

	ShowMessage($strScriptName, ": Beginning Script Execution")

	; Unmap all network drives
	UnmapNetworkDrives()

	; Close IST Windows
	CloseISTWindow()
EndFunc   ;==>ScriptStarting

;***************************************************************************
;** Function: 		ScriptEnding($strScriptName, $strMsg)
;** Parameters:
;**		$strScriptName - script name calling this function
;**		$strMsg - string for registry
;** Description:
;**		This function is called to display ending message then call VManager.
;** Return:
;**		None
;** Usage:
;**		ScriptEndingReboot(SCRIPT_NAME, $strMsg)
;**
;***************************************************************************
Func ScriptEnding($strScriptName, $strMsg)

	; Set registry key indicate script status
	UpdateRegistryStatus($strScriptName, $strMsg)

	; Write ending message to log file. Display to use.
	ShowMessage($strScriptName, ": Ending Script Execution")

	; Unmap drive
	UnmapNetworkDrives()

	; Run VManager
	RunVManager()

EndFunc   ;==>ScriptEnding

;***************************************************************************
;** Function: 		ScriptEndingNoVManager($strScriptName, $strMsg)
;** Parameters:
;**		$strScriptName - script name calling this function
;**		$strMsg - string for registry
;** Description:
;**		This function is called to display ending message without running VManager.
;** Return:
;**		None
;** Usage:
;**		ScriptEndingReboot(SCRIPT_NAME, $strMsg)
;**
;***************************************************************************
Func ScriptEndingNoVManager($strScriptName, $strMsg)

	; Set registry key indicate script status
	UpdateRegistryStatus($strScriptName, $strMsg)

	; Write ending message to log file. Display to use.
	ShowMessage($strScriptName, ": Ending Script Execution")

	; Unmap drive
	UnmapNetworkDrives()

EndFunc   ;==>ScriptEndingNoVManager


;***************************************************************************
;** Function: 		ScriptEndingReboot($strScriptName, $strApp, $strMsg)
;** Parameters:
;**		$strScriptName - script name calling this function
;**		$strApp - after reboot, run this apps
;**		$strMsg - string for registry
;** Description:
;**		This function is called to display ending message then reboot.
;** Return:
;**		None
;** Usage:
;**		ScriptEndingReboot($SCRIPT_NAME, $strApp, $strMsg)
;**
;***************************************************************************
Func ScriptEndingReboot($strScriptName, $strApp, $strMsg)

	; Set Win2K auto logon/Restart VManager registry
	SetWin2KAutoLogon()

	; Reboot system with $strApp
	$strApp = StringUpper($strApp)
	If $strApp == "VMANAGER.EXE" Then
		SetRestartVManager()
	Else
		$strApp = @HomeDrive & "\Scripts\" & $strApp
		SetupRestartApp($strApp)
	EndIf

	; Set registry key indicate script status
	UpdateRegistryStatus($strScriptName, $strMsg)

	; Write ending message to log file. Display to use.
	ShowMessage($strScriptName, ": Ending Script Execution")

	; Unmap before starting mapping.
	UnmapNetworkDrives()

	; RebootSystem
	RebootSystem()

EndFunc   ;==>ScriptEndingReboot

;***************************************************************************
;** Function: 		ScriptEndingRebootVManager($strScriptName, $strMsg)
;** Parameters:
;**		$strScriptName - script name calling this function
;**		$strMsg - string for registry
;** Description:
;**		This function is called to display ending message then reboot.
;** Return:
;**		None
;** Usage:
;**		ScriptEndingReboot(SCRIPT_NAME, $strMsg)
;**
;***************************************************************************
Func ScriptEndingRebootVManager($strScriptName, $strMsg)

	; Set Win2K auto logon/Restart VManager registry
	SetWin2KAutoLogon()

	; Set Restart VManager
	SetRestartVManager()

	; Set registry key indicate script status
	UpdateRegistryStatus($strScriptName, $strMsg)

	; Write ending message to log file. Display to use.
	ShowMessage($strScriptName, ": Ending Script Execution")

	; Unmap before starting mapping.
	UnmapNetworkDrives()

	; RebootSystem
	RebootSystem()

EndFunc   ;==>ScriptEndingRebootVManager

;***************************************************************************
;** Function: 		ScriptEndingRebootVManagerUI($strScriptName, $strMsg)
;** Parameters:
;**		$strScriptName - script name calling this function
;**		$strMsg - string for registry
;** Description:
;**		This function is called to display ending message then reboot.
;** Return:
;**		None
;** Usage:
;**		ScriptEndingReboot(SCRIPT_NAME, $strMsg)
;**
;***************************************************************************
Func ScriptEndingRebootVManagerUI($strScriptName, $strMsg)

	; Set Win2K auto logon/Restart VManager registry
	SetWin2KAutoLogon()

	; Set Restart VManager
	SetRestartVManagerUI()

	; Set registry key indicate script status
	UpdateRegistryStatus($strScriptName, $strMsg)

	; Write ending message to log file. Display to use.
	ShowMessage($strScriptName, ": Ending Script Execution")

	; Unmap before starting mapping.
	UnmapNetworkDrives()

	; RebootSystem
	RebootSystem()

EndFunc   ;==>ScriptEndingRebootVManagerUI

;***************************************************************************
;** Function: 		ScriptEndingRebootOthers($strScriptName, $strApp, $strMsg)
;** Parameters:
;**		$strScriptName - script name calling this function
;**		$strApp - after reboot, run this apps
;**		$strMsg - ending message to show/write
;** Description:
;**		This function is called to display ending message then reboot.
;** Return:
;**		None
;** Usage:
;**		ScriptEndingReboot(SCRIPT_NAME, $strApp, $strMsg)
;**
;***************************************************************************
Func ScriptEndingRebootOthers($strScriptName, $strApp, $strMsg)

	; Set Win2K auto logon/Restart VManager registry
	SetWin2KAutoLogon()

	; Reboot system with $strApp
	; For PLDWrapper.exe, it passed $strApp with full path.
	$result = StringInStr($strApp, "PLDWrapper.exe")
	If (Not $result) Then
		$strApp = @HomeDrive & "\Scripts\" & $strApp
	EndIf
	SetupRestartApp($strApp)

	; Set registry key indicate script status
	UpdateRegistryStatus($strScriptName, $strMsg)


	; Write ending message to log file. Display to use.
	ShowMessage($strScriptName, ": Ending Script Execution")

	; Unmap before starting mapping.
	UnmapNetworkDrives()

	; RebootSystem
	RebootSystem()

EndFunc   ;==>ScriptEndingRebootOthers

;***************************************************************************
;** Function: 		ScriptEndingNoRebootOthers($strScriptName, $strApp, $strMsg)
;** Parameters:
;**		$strScriptName - script name calling this function
;**		$strApp - after reboot, run this apps
;**		$strMsg - message for registry
;** Description:
;**		This function is called to display ending message then reboot.
;** Return:
;**		None
;** Usage:
;**		ScriptEndingReboot(SCRIPT_NAME, $strApp, "INSTALLED")
;**
;***************************************************************************
Func ScriptEndingNoRebootOthers($strScriptName, $strApp, $strMsg)

	; Set Win2K auto logon/Restart VManager registry
	SetWin2KAutoLogon()

	; Reboot system with $strApp
	$strApp = @HomeDrive & "\Scripts\" & $strApp
	SetupRestartApp($strApp)

	; Set registry key indicate script status
	UpdateRegistryStatus($strScriptName, $strMsg)

	; Write ending message to log file. Display to use.
	ShowMessage($strScriptName, ": Ending Script Execution")

	; Unmap before starting mapping.
	UnmapNetworkDrives()

EndFunc   ;==>ScriptEndingNoRebootOthers

;***************************************************************************
;** Function: 		ActivateWinDVD()
;** Parameters:
;**		None
;** Description:
;**		This function is called to active WinDVD
;** Return:
;**		None
;** Usage:
;**		ActivateWinDVD()
;**
;***************************************************************************
Func ActivateWinDVD()
	; Write some msg.
	WriteLog("Running ActivateWinDVD() function.")

	$strFile = @ProgramFilesDir & "\InterVideo\DVD7\WinDVD.exe"
	If Not FileExists($strFile) Then
		Return False
	EndIf
	Run(@ComSpec & " /c " & '"' & $strFile & '"', "", @SW_HIDE)
	Sleep(2000)

	; Close "Security Alert" window
	$strTitle = "Windows Security Alert"
	$strText = "Do you want to keep blocking this program?"
	If WinExists($strTitle) Then
		WinClose($strTitle)
	EndIf

	; Set "Do not display this dialog again.
	$strTitle = "InterVideo Product Registration"
	$strText = "Please register this product to be eligible for Intervideo product support"
	If Not WinWait($strTitle, $strText, 100) Then
		WriteLog("Step 2 of 4")
		Return False
	EndIf
	ControlCommand($strTitle, "Do not display this dialog again.", 213, "Check", "")
	ControlClick($strTitle, "&Continue", "&Continue")

	; Close WinDVD
	$strTitle = "InterVideo WinDVD 7"
	If WinExists($strTitle) Then
		WinClose($strTitle)
	EndIf

	Return True

EndFunc   ;==>ActivateWinDVD

;***************************************************************************
;** Function: 		SetupRunNext(strScriptName, strRunNext)
;** Parameters:
;**	strScriptName - script calling this function
;**	strRunNext - full path of executable filename.
;** Description:
;**	This function is called to write scriptname and RunNext full path to registry.
;** Return:
;**	None
;** Usage:
;**
;***************************************************************************
Func SetupRunNext($strScriptName, $strRunNext)
	; ScriptName registry
	$strKey = "HKLM\Software\Intel\MPG\TESTAPP"
	$strSubKey = $strScriptName
	RegWrite($strKey, $strSubKey, "REG_SZ", "")

	; RunNext Registry
	$strKey = "HKLM\Software\Intel\MPG\TESTAPP"
	$strSubKey = "RunNext"
	RegWrite($strKey, $strSubKey, "REG_SZ", $strRunNext)

EndFunc   ;==>SetupRunNext

;***************************************************************************
;** Function: 		CheckInstallStatus(strScriptName, strRunNext)
;** Parameters:
;**		strScriptName - script name call this function
;**		strRunNext - run next executable
;** Description:
;**		This function is called to check the applet installation status.
;**		It's called by test script to make sure the applet already installed before running test.
;** Return:
;**		1
;** Usage:
;**		strRunNext = "C:\Scripts\TmIST.exe"
;**		strScriptName = "Intel CPU Frequency Display"
;**		if !CheckInstallStatus(strScriptName, strRunNext) then
;**			WriteLog("%strScriptName% CheckRequiredApplication @FALSE")
;**			return @FALSE
;**		endif
;**
;***************************************************************************
Func CheckInstallStatus($strScriptName, $strRunNext)

	$strKey = "HKLM\Software\Intel\MPG\TESTAPP"
	$strSubKey = $strScriptName
	$strStatus = RegRead($strKey, $strSubKey)

	; For Testing
	WriteLog("Registry Status: " & $strStatus)

	If $strStatus <> "" Then
		$strStatus = StringUpper($strStatus)
		If ($strStatus == "INSTALLED") Or ($strStatus == "EXECUTED") Then
			WriteLog($strScriptName & " Function CheckInstallStatus() " & " INSTALLED")
			Return True
		Else
			WriteLog($strScriptName & " Function CheckInstallStatus() " & " FAILED")
			Return False
		EndIf
	Else
		; Setup RunNext registry
		SetupRunNext($strScriptName, $strRunNext)
		WriteLog($strScriptName & ": Function CheckInstallStatus() Call TmInstCh.exe")

		; Call TmInstCk.exe
		$strFile = @HomeDrive & "\Scripts\TmInstCkAU3.exe"
		Run(@ComSpec & " /c " & '"' & $strFile & '"', "", @SW_HIDE)
		Exit
	EndIf

EndFunc   ;==>CheckInstallStatus

;***************************************************************************
;** Function: 		CheckRunStatus(strScriptName, strRunNext)
;** Parameters:
;**		strScriptName - script name call this function
;**		strRunNext - run next executable
;** Description:
;**		This function is called to check the applet running status.
;**		It's called by test script to make sure the applet is running.
;** Return:
;**		None
;** Usage:
;**		strScriptName = "Intel CPU Frequency Display"
;**		if !CheckRunStatus(strScriptName, strRunNext) then
;**			WriteLog("%strScriptName% CheckRunStatus @FALSE")
;**			return @FALSE
;**		endif
;**
;***************************************************************************
Func CheckRunStatus($strScriptName, $strRunNext)

	; RunNowStatus
	$strKey = "HKLM\Software\Intel\MPG\TESTAPP\RunNowStatus"
	$strSubKey = $strScriptName
	$strStatus = RegRead($strKey, $strSubKey)

	If $strStatus <> "" Then
		$strStatus = StringUpper($strStatus)
		If $strStatus == "RUNNING" Then
			WriteLog($strScriptName & " Function CheckRunStatus() RUNNING")
			Return True
		Else
			Return False
		EndIf
	Else
		; Update RunNow registry
		$strKey = "HKLM\Software\Intel\MPG\TESTAPP"
		$strSubKey = "RunNow"
		RegWrite($strKey, $strSubKey, "REG_SZ", $strScriptName)

		; Setup RunNext Registry
		SetupRunNext($strScriptName, $strRunNext)
		WriteLog($strScriptName & " Function CheckRunStatus() Call TmRunApp.exe")

		; Call TmRunApp.exe
		$strFile = @HomeDrive & "\Scripts\TmRunAppAU3.exe"
		Run(@ComSpec & " /c " & '"' & $strFile & '"', "", @SW_HIDE)
		Exit
	EndIf

EndFunc   ;==>CheckRunStatus

;***************************************************************************
;** Function: 		CheckRunNext()
;** Parameters:
;**		None
;** Description:
;**		This function is called to get the full name of executable to run.
;** Return:
;**		None
;** Usage:
;**		CheckRunNext()
;**
;***************************************************************************
Func CheckRunNext()
	; RunNowStatus
	$strKey = "HKLM\Software\Intel\MPG\TESTAPP"
	$strSubKey = "RunNext"
	$strRunNext = RegRead($strKey, $strSubKey)

	If $strRunNext <> "" Then
		; RunNext file
		Run(@ComSpec & " /c " & '"' & $strRunNext & '"', "", @SW_HIDE)
	EndIf

EndFunc   ;==>CheckRunNext

;***************************************************************************
;** Function: 		GetHDAudioDriverSourceDir(ByRef $strExeDir, ByRef $strExeFile)
;** Parameters:
;**		None
;** Description:
;**		This function is called to run setup.
;** Return:
;**		None
;***************************************************************************
Func GetHDAudioDriverSourceDir(ByRef $strExeDir, ByRef $strExeFile)
	If (IsWinXP64Family()) Then
		$strExeFile = "KB901105.exe"
		$strExeDir = "XP64"
	ElseIf (IsWinXPSP1()) Then
		$strExeFile = "kb888111xpsp1.exe"
		$strExeDir = "XP"
	ElseIf (IsWinXPSP2()) Then
		$strExeFile = "kb888111xpsp2.exe"
		$strExeDir = "XP"
	ElseIf (IsWin2KSP4()) Then
		$strExeFile = "kb888111w2ksp4.exe"
		$strExeDir = "Win2K"
	EndIf
EndFunc   ;==>GetHDAudioDriverSourceDir

;***************************************************************************
;** Function: 		CloseVistaWindow()
;** Parameters:
;**		None
;** Description:
;**		This function is called to close 2 of Vista x64 windows popup after system boot.
;** Return:
;**		None
;***************************************************************************
Func CloseVistaWindow()
	$strTitle = "Set Network Location"
	If WinExists($strTitle) Then
		WinClose($strTitle)
	EndIf

	$strTitle = "New Display Detected"
	If WinExists($strTitle) Then
		WinClose($strTitle)
	EndIf

EndFunc   ;==>CloseVistaWindow

;***************************************************************************
;** Function: 		GetInstallScriptExe($strScriptName)
;** Parameters:
;**		None
;** Description:
;**		This function is called to check get the executable file name by checking the
;**		script name.
;** Return:
;**		None
;***************************************************************************
Func GetInstallScriptExe($strScriptName, ByRef $strExeName)

	WriteLog("Running GetInstallScriptExe() function: " & $strScriptName)

	$blnReturn = True

	Switch $strScriptName

		; Intel CPU Frequency Display
		Case "Intel CPU Frequency Display"
			$strExeName = @HomeDrive & "\Scripts\ImCPUFrq.exe"

			; Intel CSwitch Utility
		Case "Intel CSwitch Utility"
			$strExeName = @HomeDrive & "\Scripts\ImCSwt.exe"

			; Intel GVCycle Utility
		Case "Intel GVCycle Utility"
			$strExeName = @HomeDrive & "\Scripts\ImGVCyl.exe"

			; WinBez
		Case "WinBez"
			$strExeName = @HomeDrive & "\Scripts\ImWinBez.exe"

			; NX Hammer
		Case "NX Hammer"
			$strExeName = @HomeDrive & "\Scripts\ImNXHammer.exe"

			; Intel SpeedStep 3 Applet
		Case "Intel SpeedStep 3 Applet"
			$strExeName = @HomeDrive & "\Scripts\ImIST.exe"

			; PCMark2005 for Vista
		Case "PCMark2005 for Vista"
			$strExeName = @HomeDrive & "\Scripts\ImPCMk05Vista.exe"

			; PCMark2005
		Case "PCMark2005"
			$strExeName = @HomeDrive & "\Scripts\ImPCMk05.exe"

		Case "Unreal Tournament 3 Demo"
			$strExeName = @HomeDrive & "\Scripts\ImUT3Demo.exe"

			; X3 Terran Conflict Demo
		Case "X3 Terran Conflict Demo"
			$strExeName = @HomeDrive & "\Scripts\ImX3TC.exe"

			; Intel C State Residency Monitor Utility
		Case "Intel C State Residency Monitor Utility"
			$strExeName = @HomeDrive & "\Scripts\ImCStResMon.exe"

			; Power Management Test Tool
		Case "Power Management Test Tool"
			$strExeName = @HomeDrive & "\Scripts\ImPwrTest.exe"

			; IBASES 2.0
		Case "IBASES 2.0"
			$strExeName = @HomeDrive & "\Scripts\ImIBASES.exe"

			; 3DMark 01
		Case "3DMark 2001"
			$strExeName = @HomeDrive & "\Scripts\Im3DMk01.exe"

			; Serious Sam 2 Demo
		Case "Serious Sam 2 Demo"
			$strExeName = @HomeDrive & "\Scripts\ImSerSam2.exe"

			; PCMark2002
		Case "PCMark2002"
			$strExeName = @HomeDrive & "\Scripts\ImPCMk02.exe"

			; SiSoftware Sandra Pro
		Case "SiSoftware Sandra Pro"
			$strExeName = @HomeDrive & "\Scripts\ImSandraPro.exe"

			; SiSoftware Sandra Pro
		Case "SiSoftware Sandra Pro"
			$strExeName = @HomeDrive & "\Scripts\ImSandraPro.exe"

			; Intel ProcLoad Utility
		Case "Intel ProcLoad Utility"
			$strExeName = @HomeDrive & "\Scripts\ImProcLo.exe"

			; 3DMark 2006
		Case "3DMark 2006"
			$strExeName = @HomeDrive & "\Scripts\Im3DMk06.exe"

			; Glaze V3.1 Setup
		Case "Glaze V3.1 Setup"
			$strExeName = @HomeDrive & "\Scripts\ImGlaze.exe"

			; Final Reality Benchmark 1.01
		Case "Final Reality Benchmark 1.01"
			$strExeName = @HomeDrive & "\Scripts\ImFinalR.exe"

			; 3DMark 2003 Professional Version
		Case "3DMark 2003 Professional Version"
			$strExeName = @HomeDrive & "\Scripts\Im3DMk03P.exe"

			; Microsoft .NET Framework V2.0
		Case "Microsoft .NET Framework V2.0"
			$strExeName = @HomeDrive & "\Scripts\ImDotnet20.exe"

			; MobileMark 2007
		Case "MobileMark 2007"
			$strExeName = @HomeDrive & "\Scripts\ImMMk07.exe"


			; Windows Media Encoder 9
		Case "Windows Media Encoder 9"
			$strExeName = @HomeDrive & "\Scripts\ImMEncoder.exe"

			; Windows MediaPlayer 10
		Case "Windows MediaPlayer 10"
			$strExeName = @HomeDrive & "\Scripts\ImMPlayer.exe"

			; Debug View
		Case "DebugView"
			$strExeName = @HomeDrive & "\Scripts\ImDbgview.exe"

			; TXT Driver
		Case "MPGTXT Driver"
			$strExeName = @HomeDrive & "\Scripts\ImMPGTXT.exe"

			; ACM Driver
		Case "SINIT ACM"
			$strExeName = @HomeDrive & "\Scripts\ImACM.exe"

			; TXT MPG GUI install script
		Case "TXT MPG App"
			$strExeName = @HomeDrive & "\Scripts\ImTXTMPGGUI.exe" ; file necessary to run test 56. MPG TXT TEST

			; CineBench R10 install script
		Case "CineBench R10"
			$strExeName = @HomeDrive & "\Scripts\ImCineBench.exe"

			; HeavyLoad
		Case "HeavyLoad"
			$strExeName = @HomeDrive & "\Scripts\ImHvyLd.exe"

			; 3DMark Vantage
		Case "3DMark Vantage"
			$strExeName = @HomeDrive & "\Scripts\Im3DMkVN.exe"

			; PCMark Vantage
		Case "PCMark Vantage"
			$strExeName = @HomeDrive & "\Scripts\ImPCMkVN.exe"

			; ImTAT.au3
		Case "Intel Thermal Analysis Tool"
			$strExeName = @HomeDrive & "\Scripts\ImTAT.exe"

			; ImdotNET30.au3
		Case ".NET 3.0"
			$strExeName = @HomeDrive & "\Scripts\ImdotNET30.exe"

			; ImdotNET30.au3
		Case ".NET 3.0 x64"
			$strExeName = @HomeDrive & "\Scripts\ImdotNET30_64.exe"
			
			; Imrthdribl.au3
		Case "Real-Time HDR Image-Based Lighting"
			$strExeName = @HomeDrive & "\Scripts\Imrthdribl.exe"

			; Intel Turbo Boost Technology driver
		Case "Intel Turbo Boost Technology Driver"
			$strExeName = @HomeDrive & "\Scripts\ImIPS.exe"
			; ***** SPECIAL CASE *****

			; Application Install Check TEST VERSION
		Case "Application Install Check"
			$strExeName = @HomeDrive & "\Scripts\TmInstallCheck.exe"

			; 08. WinBez C4 Stress Test
		Case "08. WinBez C4 Stress Test"
			$strExeName = @HomeDrive & "\Scripts\TmWinBez.exe"

			; 09. C4 Stress 24 Hour Test
		Case "09. C4 Stress 24 Hour Test"
			$strExeName = @HomeDrive & "\Scripts\TmC4Strs.exe"

			; 34. CineBench R10
		Case "34. CineBench R10"
			$strExeName = @HomeDrive & "\Scripts\TmCineBench.exe"


			; "48. 3DMark 2006 Benchmark Test"
		Case "48. 3DMark 2006 Benchmark Test"
			$strExeName = @HomeDrive & "\Scripts\Tm3DMk06.exe"

			; 49. Busy Idle Test for Vista"
;~ 	Case "49. Busy Idle Test for Vista"
;~ 		$strExeName   = @HomeDrive & "\Scripts\TmBIdleVista.exe"

			; "AutoIt3 Debugger"
		Case "AutoIt3 Debugger"
			$strExeName = @HomeDrive & "\Scripts\ImAutoIt3Debugger.exe"

			; 52. MobileMark2007 Productivity 2007 Test
		Case "52. MobileMark2007 Productivity 2007 Test"
			$strExeName = @HomeDrive & "\Scripts\TmMMk2007Product.exe"

			; "53. MobileMark2007 DVD 2007 Test"
		Case "53. MobileMark2007 DVD 2007 Test"
			$strExeName = @HomeDrive & "\Scripts\TmMMk2007DVD.exe"

			; "54. MobileMark2007 Reader 2007 Test"
		Case "54. MobileMark2007 Reader 2007 Test"
			$strExeName = @HomeDrive & "\Scripts\TmMMk2007Reader.exe"

			; "55. Unreal Tournament 3 Demo Test"
		Case "55. Unreal Tournament 3 Demo Test"
			$strExeName = @HomeDrive & "\Scripts\TmUT3Demo.exe"

		Case "56. MPG TXT Test"
			$strExeName = @HomeDrive & "\Scripts\TmMPGTXT.exe" ; added for "56. MPG TXT Test" purposes

		Case "57. HeavyLoad Stress Test"
			$strExeName = @HomeDrive & "\Scripts\TmHvyLd.exe"

		Case "58. 3DMark Vantage Benchmark Test"
			$strExeName = @HomeDrive & "\Scripts\Tm3DMkVn.exe"

		Case "59. PCMark Vantage Benchmark Test"
			$strExeName = @HomeDrive & "\Scripts\TmPCMkVn.exe"

		Case "60. X3 Terran Conflict Rolling Demo"
			$strExeName = @HomeDrive & "\Scripts\TmX3TC.exe"

		Case Else
			; Something else ?
			$blnReturn = False
	EndSwitch

	Return $blnReturn

EndFunc   ;==>GetInstallScriptExe

;***************************************************************************
;** Function: 		RunVManager()
;** Parameters:
;**		None
;** Description:
;**		This function is used by script to run VManager.
;** Return:
;**		None
;** Usage:
;**		RunVManager()
;**
;***************************************************************************
Func RunVManagerError()
	WriteLog("RunVManager() function.")

	$strVMgr = @HomeDrive & "\scripts\VManager.exe"
	$strInstallTemp = @HomeDrive & "\scripts\InstTemp.csv"
	$strTestTemp = @HomeDrive & "\scripts\TestTemp.csv"
	$strInstallTemp2 = @HomeDrive & "\scripts\InstallTemp.xml"
	$strTestTemp2 = @HomeDrive & "\scripts\TestTemp.xml"


	; Only call VManager with 2 temp csv files exist and VManager.exe exists.
	If Not FileExists($strVMgr) Then
		Return
	ElseIf Not ((FileExists($strInstallTemp) And FileExists($strTestTemp)) Or (FileExists($strInstallTemp2) And FileExists($strTestTemp2))) Then
		Exit
	EndIf

	; Only run VManager.exe if the script is called by VManager.
	$strVManager = @HomeDrive & "\Scripts\VManager.exe"
	Run(@ComSpec & " /c " & $strVManager & " -NOUI", "", @SW_HIDE)

EndFunc   ;==>RunVManagerError

;***************************************************************************
;** Function: 		TileWindowsHorizontally()
;** Parameters:
;**		None
;** Description:
;**		This function is called to tile windows horizontally
;** Return:
;**		None
;** Usage:
;**		TileWindowsVertically()
;**
;***************************************************************************
Func TileWindowsHorizontally()
	$objShell = ObjCreate("Shell.Application") ; Get the Windows Shell Object
	$objShell.TileHorizontally ; Get the collection of open shell
EndFunc   ;==>TileWindowsHorizontally

#CS
	;***************************************************************************
	;** Function: 		TileWindowsHorizontally-TAT()
	;** Parameters:
	;**		None
	;** Description:
	;**		This function is called to tile windows horizontally excluding TAT.exe
	;** Return:
	;**		None
	;** Usage:
	;**		TileWindowsVertically()-TAT()
	;**
	;***************************************************************************
	Func TileWindowsHorizontally-TAT()
	
	$objShell = ObjCreate("TAT.Application")    ; Get the Windows Shell Object
	$objShell.TileHorizontally 	          ; Get the collection of open shell
	$objShell = ObjCreate("Shell.Application")    ; Get the Windows Shell Object
	$objShell.TileHorizontally 	          ; Get the collection of open shell
	EndFunc   ;==>TileWindowsHorizontally
#CE
;***************************************************************************
;** Function: 		TileWindowsVertically()
;** Parameters:
;**		None
;** Description:
;**		This function is called to tile windows vertically
;** Return:
;**		None
;** Usage:
;**		TileWindowVertically()
;**
;***************************************************************************
Func TileWindowsVertically()
	$objShell = ObjCreate("Shell.Application") ; Get the Windows Shell Object
	$objShell.TileVertically ; Get the collection of open shell
EndFunc   ;==>TileWindowsVertically

;***************************************************************************
;** Function: 		PwrTestGUI()
;** Parameters:
;**	None
;** Description:
;**	This function is called to run test.
;** Return:
;**	None
;***************************************************************************
Func PwrTestGUI(ByRef $intCycles, ByRef $intDelay, ByRef $intSleep, ByRef $strSState, ByRef $strCyclingThrough)
	#include <GUIConstants.au3>

	; == GUI generated with Koda ==
	$Form1 = GUICreate("PwrTest SLEEP Scenario", 416, 475, 192, 125)
	$Group1 = GUICtrlCreateGroup("Number of Cycles", 16, 16, 385, 129)
	$rad10Cycles = GUICtrlCreateRadio("10", 32, 40, 113, 17)
	GUICtrlSetState(-1, $GUI_CHECKED)
	$rad20Cycles = GUICtrlCreateRadio("20", 32, 64, 113, 17)
	$txtCycles = GUICtrlCreateInput("3000", 272, 112, 121, 21)
	$rad30Cycles = GUICtrlCreateRadio("30", 32, 88, 113, 17)
	$rad100Cycles = GUICtrlCreateRadio("100", 152, 40, 81, 17)
	$rad200Cycles = GUICtrlCreateRadio("200", 152, 64, 73, 17)
	$rad300Cycles = GUICtrlCreateRadio("300", 152, 88, 73, 17)
	$rad1000Cycles = GUICtrlCreateRadio("1000", 248, 40, 129, 17)
	$rad2000Cycles = GUICtrlCreateRadio("2000", 248, 64, 113, 17)
	$radEnterCycles = GUICtrlCreateRadio("Enter number of cycles", 248, 88, 137, 17)
	$rad50Cycles = GUICtrlCreateRadio("50", 32, 112, 113, 17)
	$rad500Cycles = GUICtrlCreateRadio("500", 152, 112, 113, 17)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	$Group2 = GUICtrlCreateGroup("Sleep Time in Seconds", 216, 160, 185, 105)
	$rad60Sleep = GUICtrlCreateRadio("60 seconds", 232, 184, 113, 17)
	GUICtrlSetState(-1, $GUI_CHECKED)
	$radEnterSleep = GUICtrlCreateRadio("Enter number in seconds", 232, 208, 145, 17)
	$txtSleep = GUICtrlCreateInput("30", 264, 232, 121, 21)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	$Group3 = GUICtrlCreateGroup("Delay Time in Seconds", 16, 160, 185, 105)
	$rad90Delay = GUICtrlCreateRadio("90 seconds", 32, 184, 113, 17)
	GUICtrlSetState(-1, $GUI_CHECKED)
	$radEnterDelay = GUICtrlCreateRadio("Enter number in seconds", 32, 208, 145, 17)
	$txtDelay = GUICtrlCreateInput("30", 64, 232, 121, 21)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	$Group4 = GUICtrlCreateGroup("Cycling Throught All Power States", 216, 280, 185, 105)
	$radInOrder = GUICtrlCreateRadio("In Order", 240, 304, 113, 17)
	GUICtrlSetState(-1, $GUI_CHECKED)
	GUICtrlSetState(-1, $GUI_DISABLE)
	$radRandomly = GUICtrlCreateRadio("Randomly", 240, 328, 113, 17)
	GUICtrlSetState(-1, $GUI_DISABLE)
	$Group5 = GUICtrlCreateGroup("Power State", 16, 280, 185, 105)
	$radS3 = GUICtrlCreateRadio("S3", 32, 304, 113, 17)
	GUICtrlSetState(-1, $GUI_CHECKED)
	$radS4 = GUICtrlCreateRadio("S4", 32, 328, 113, 17)
	$radALL = GUICtrlCreateRadio("All", 32, 352, 113, 17)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	$btnOK = GUICtrlCreateButton("OK", 312, 408, 75, 25, $BS_DEFPUSHBUTTON)
	$btnCancel = GUICtrlCreateButton("Cancel", 312, 440, 75, 25, 0)
	GUISetState(@SW_SHOW)

	; Return value
	$retVal = True
	$blnInputCycle = False
	$blnInputDelay = False
	$blnInputSleep = False

	While 1
		$msg = GUIGetMsg()

		Switch $msg
			Case $btnOK
				$msg = $GUI_EVENT_CLOSE
				ExitLoop
			Case $btnCancel
				$msg = $GUI_EVENT_CLOSE
				$retVal = False
				ExitLoop
			Case $GUI_EVENT_CLOSE

				If $blnInputCycle Then
					$intCycles = GUICtrlRead($txtCycles)
				EndIf
				If $blnInputDelay Then
					$intDelay = GUICtrlRead($txtDelay)
				EndIf
				If $blnInputSleep Then
					$intSleep = GUICtrlRead($txtSleep)
				EndIf

				ExitLoop

				; Get Cycles
			Case $rad10Cycles
				$intCycles = 10
				$blnInputCycle = False
			Case $rad20Cycles
				$intCycles = 20
				$blnInputCycle = False
			Case $rad30Cycles
				$intCycles = 30
				$blnInputCycle = False
			Case $rad50Cycles
				$intCycles = 50
				$blnInputCycle = False
			Case $rad100Cycles
				$intCycles = 100
				$blnInputCycle = False
			Case $rad200Cycles
				$intCycles = 200
				$blnInputCycle = False
			Case $rad300Cycles
				$intCycles = 300
				$blnInputCycle = False
			Case $rad500Cycles
				$intCycles = 500
				$blnInputCycle = False
			Case $rad1000Cycles
				$intCycles = 1000
				$blnInputCycle = False
			Case $rad2000Cycles
				$intCycles = 2000
				$blnInputCycle = False
			Case $radEnterCycles
				$blnInputCycle = True

				; Get Delay time
			Case $rad90Delay
				$intDelay = 90
				$blnInputDelay = False
			Case $radEnterDelay
				$blnInputDelay = True

				; Get Sleep time
			Case $rad60Sleep
				$intSleep = 90
				$blnInputSleep = False
			Case $radEnterSleep
				$blnInputSleep = True

				; Get power state
			Case $radS3
				$strSState = "3"
				GUICtrlSetState($radInOrder, $GUI_DISABLE)
				GUICtrlSetState($radRandomly, $GUI_DISABLE)
			Case $radS4
				$strSState = "4"
				GUICtrlSetState($radInOrder, $GUI_DISABLE)
				GUICtrlSetState($radRandomly, $GUI_DISABLE)
			Case $radALL
				$strSState = "all"
				GUICtrlSetState($radInOrder, $GUI_ENABLE)
				GUICtrlSetState($radRandomly, $GUI_ENABLE)

				; Get
			Case $radInOrder
				$strCyclingThrough = "all"
			Case $radRandomly
				$strCyclingThrough = "rnd"

		EndSwitch
	WEnd

	; Close DLL
	;DllClose($dll)

	; Get textbox input
	If $retVal Then
		If $blnInputCycle Then
			$intCycles = GUICtrlRead($txtCycles)
		EndIf
		If $blnInputDelay Then
			$intDelay = GUICtrlRead($txtDelay)
		EndIf
		If $blnInputSleep Then
			$intSleep = GUICtrlRead($txtSleep)
		EndIf

		; Write PwrTest GUI file
		$strPwrTestGUIFile = @HomeDrive & "\scripts\PwrTestGUI.txt"
		$strFileHandle = FileOpen($strPwrTestGUIFile, 1)

		$strLog = @HomeDrive & "\logs\PwrTestLog.wtl"
		$strExe = @HomeDrive & "\APPS\PwrTest\PwrTest.exe /sleep "
		If ($strSState <> "3") And ($strSState <> "4") Then
			$strCmd = $strExe & "/c:" & $intCycles & " /d:" & $intDelay & " /p:" & $intSleep & " /s:" & $strCyclingThrough & " /l:" & $strLog
		Else
			$strCmd = $strExe & "/c:" & $intCycles & " /d:" & $intDelay & " /p:" & $intSleep & " /s:" & $strSState & " /l:" & $strLog
		EndIf
		FileWriteLine($strFileHandle, $strCmd)

		FileClose($strFileHandle)
	EndIf


	Return $retVal

EndFunc   ;==>PwrTestGUI

;***************************************************************************
;** Function: 		GetISTTps(ByRef $strSwitchingApplet, ByRef $intTPS)
;** Parameters:
;**		None
;** Description:
;**		This function is called to display GUI for user to select input.
;** Return:
;**		1 for success, otherwise 0.
;**
;***************************************************************************
Func GetISTTps(ByRef $intTPS)

	$strKey = "HKLM\SOFTWARE\Intel\MPG\TESTSCRIPT"
	$intTPS = RegRead($strKey, "ISTCFG")
	If ($intTPS) Then
		Return True
	EndIf

	$retVal = True
	$intTPS = 100

	#include <GUIConstants.au3>
	; == GUI generated with Koda ==
	$Form1 = GUICreate("IST Transitioning GUI", 324, 287, 192, 125)
	$Group1 = GUICtrlCreateGroup("IST Transitions/Second", 16, 16, 289, 193)
	$rad1TPS = GUICtrlCreateRadio("1 Transition Per Second", 40, 48, 177, 17)
	$rad5TPS = GUICtrlCreateRadio("5 Transitions Per Second", 40, 80, 169, 17)
	$rad10TPS = GUICtrlCreateRadio("10 Transitions Per Second", 40, 112, 177, 17)
	$rad100TPS = GUICtrlCreateRadio("100 Transition Per Second", 40, 144, 177, 17)
	$radNoneTPS = GUICtrlCreateRadio("None: CPU doesn't support IST", 40, 176, 177, 17)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	$btnOK = GUICtrlCreateButton("OK", 224, 224, 75, 25, $BS_DEFPUSHBUTTON)
	$btnCancel = GUICtrlCreateButton("Cancel", 224, 256, 75, 25, 0)
	GUISetState(@SW_SHOW)


	; Set 10 tps as default, GVCycle as default
	GUICtrlSetState($rad100TPS, $GUI_CHECKED)

	GUISetState(@SW_SHOW)
	While 1
		$msg = GUIGetMsg()

		Switch $msg
			Case $btnOK
				$msg = $GUI_EVENT_CLOSE
				ExitLoop
			Case $btnCancel
				$msg = $GUI_EVENT_CLOSE
				$retVal = False
				ExitLoop
			Case $GUI_EVENT_CLOSE
				ExitLoop
			Case $rad1TPS
				$intTPS = 1
			Case $rad5TPS
				$intTPS = 5
			Case $rad10TPS
				$intTPS = 10
			Case $rad100TPS
				$intTPS = 100
			Case $radNoneTPS
				$intTPS = 0

		EndSwitch

	WEnd
	GUISetState(@SW_HIDE)

	; Close DLL
	;DllClose($dll)

	; Write registry
	If ($retVal) Then
		$strKey = "HKLM\SOFTWARE\Intel\MPG\TESTSCRIPT"
		RegWrite($strKey, "ISTCFG", "REG_SZ", $intTPS)
	EndIf

	Return $retVal

EndFunc   ;==>GetISTTps

;***************************************************************************
;** Function: 		IsAppletInstalled($strScriptName)
;** Parameters:
;**		None
;** Description:
;**		This function is called to check if applet is already installed.
;** Return:
;**		None
;***************************************************************************
Func IsAppletInstalled($strScriptName)
	$strFileName = ""
	$blnCheckFile = True
	$blnResult = False

	Switch $strScriptName
		; *** APPLICATION SECTION ***
		; ImFoobar2K.wbt
		Case "Foobar2000 Audio Player"
			$strFileName = @ProgramFilesDir & "\foobar2000\foobar2000.exe"

			; ImMEncoder.wbt
		Case "Windows Media Encoder 9"
			$strFileName = @ProgramFilesDir & "\Windows Media Components\Encoder\wmenc.exe"

			; ImQTime.wbt
		Case "QuickTime Media Player"
			$strFileName = @ProgramFilesDir & "\QuickTime\QuickTimePlayer.exe"

			; ImHCT121.wbt
		Case "HCT 12.1"
			$blnCheckFile = False
			$blnResult = InstallHCTAlready()

			; ImMPlayer.wbt
		Case "Windows MediaPlayer 10"
			$blnCheckFile = False
			$blnResult = IsWMPlayer10Installed()

			; ImIST.wbt
		Case "Intel SpeedStep 3 Applet"
			$blnCheckFile = False
			$blnResult = IsISTInstalled()

			; ImNXHammer.au3
		Case "NX Hammer"
			$strFileName = @HomeDrive & "\APPS\NX_Hammer\NxHammer.exe"

			; ImMConnect.wbt
		Case "Windows Media Connect"
			$strFileName = @ProgramFilesDir & "\Windows Media Connect\mswmc.exe"

			; ImFirefox.wbt
		Case "Mozilla Firefox"
			$strFileName = @ProgramFilesDir & "\Mozilla Firefox\firefox.exe"

			; ImAIM.wbt
		Case "AOL Instant Messager Installation"
			$strFileName = @ProgramFilesDir & "\AOL\AIM.exe"

			; ImMSN.wbt
		Case "MSN Messager"
			$strFileName = @ProgramFilesDir & "\MSN Messenger\msnmsgr.exe"

			; ImiTunes.wbt
		Case "iTunes+QuickTime"
			$strFileName = @ProgramFilesDir & "\iTunes\iTunes.exe"

			; ImIE.wbt
		Case "Internet Explorer"
			$blnCheckFile = False

			; ImPMTE.au3
		Case "PMTE"
			$strFileName = @HomeDrive & "\APPS\PMTE\pmte.exe"

			; ImDotnetFW.wbt
		Case "Microsoft .NET Framework V1.1"
			$strFileName = @WindowsDir & "\Microsoft.NET\Framework\v1.1.4322\ConfigWizards.exe"

			; ImPwrTest.au3
		Case "Power Management Test Tool"
			$strFileName = @HomeDrive & "\APPS\PwrTest\PwrTest.exe"

			; ImDTMLogViewer.au3
		Case "Microsoft WHQL DTM(Driver Test Manager) Log Viewer"
			$strFileName = @ProgramFilesDir & "\Microsoft\WHQL DTM Log Viewer\WHQL DTM Log Viewer.exe"

			; *** BENCHMARKSOFTWARE SECTION ***
			; Im3DMk01.wbt
		Case "3DMark 2001"
			$strFileName = @ProgramFilesDir & "\MadOnion.com\3DMark2001 SE\3DMark2001SE.exe"

			;ImFinalR.wbt
		Case "Final Reality Benchmark 1.01"
			$strFileName = @ProgramFilesDir & "\Final Reality\FR.exe"

			; ImGlaze.wbt
		Case "Glaze V3.1 Setup"
			;$strFileName   = @ProgramFilesDir & "\Evans & Sutherland\Glaze\Glaze.exe"
			$strFileName = @HomeDrive & "\Apps\Glaze\Glaze.exe"

			; ImIBASES.au3
		Case "IBASES 2.0"
			$strFileName = @ProgramFilesDir & "\Intel\IPEAK\IBASES\AGPTest\AGPTest.exe"

			; ImKSink.wbt - RoboCopy can't handle copy to c:\.
		Case "Intel KNI Kitchen Sink"
			$strFileName = @HomeDrive & "\KNIStress\KNISetup.BAT"

			; ImSerSam.wbt
		Case "Serious Sam SE Demo"
			$strFileName = @ProgramFilesDir & "\Croteam\Serious Sam - The Second Encounter Demo\Bin\SeriousSam.exe"

			; ImUT2004.wbt
		Case "Unreal Tournament 2004 Demo"
			$strFileName = @HomeDrive & "\UT2004Demo\System\UT2004.exe"

			; ImAquaMk3.wbt
		Case "AquaMark3 Benchmark"
			$strFileName = @ProgramFilesDir & "\AquaMark3\aquamark.exe"

			; ImCineBench.au3
		Case "CineBench R10"
			If IsWinXP64Family() Or IsWinVista64() Then
				$strFileName = @HomeDrive & "\apps\CineBench_R10\CINEBENCH R10 64Bit.exe"
			Else
				$strFileName = @HomeDrive & "\apps\CineBench_R10\CINEBENCH R10.exe"
			EndIf

			; ImPCMk02.wbt
		Case "PCMark2002"
			$strFileName = @ProgramFilesDir & "\MadOnion.com\PCMark2002\PCMark2002.exe"

			; ImPCMk05.wbt
		Case "PCMark2005"
			$strFileName = @ProgramFilesDir & "\Futuremark\PCMark05\PCMark05.exe"

			; ImPCMk05Vista.au3
		Case "PCMark2005 for Vista"
			$strFileName = @ProgramFilesDir & "\Futuremark\PCMark05\PCMark05.exe"

			; Im3DMk03P.wbt
		Case "3DMark 2003 Professional Version"
			$strFileName = @ProgramFilesDir & "\Futuremark\3DMark03\3DMark03.exe"

			; Im3DMk05.wbt
		Case "3DMark 2005 Business Edition"
			$strFileName = @ProgramFilesDir & "\Futuremark\3DMark05\3DMark05.exe"

			; ImSerSam2.au3
		Case "Serious Sam 2 Demo"
			$strFileName = @HomeDrive & "\APPS\Serious_Sam_II_Demo\Bin\Sam2.exe"

			; ImDotnet20.au3
		Case "Microsoft .NET Framework V2.0"
			$strFileName = @WindowsDir & "\Microsoft.NET\Framework\v2.0.50727\CONFIG"
			;$blnCheckFile = False
			;$blnResult = IsInstallDotnetFramework20Installed()

			; WinBez
		Case "WinBez"
			$strFileName = @HomeDrive & "\APPS\WinBez\winbez.exe"

			; 3DMark 2006
		Case "3DMark 2006"
			$strFileName = @ProgramFilesDir & "\Futuremark\3DMark06\3DMark06.exe"

			; Glaze V3.1 Setup
		Case "Glaze V3.1 Setup"
			;$strFileName = @ProgramFilesDir & "\Evans & Sutherland\Glaze\Glaze.exe"
			$strFileName = @HomeDrive & "\Apps\Glaze\Glaze.exe"

			; ImSMk07.au3
		Case "SYSmark 2007"
			$strFileName = @ProgramFilesDir & "\BAPCo\SYSmark 2007 Preview\bin\Sysmark2007.exe"

			;ImMMk07.au3
		Case "MobileMark 2007"
			$strFileName = @ProgramFilesDir & "\BAPCo\MobileMark 2007\bin\MobileMark2007.exe"

			;ImUT3Demo.au3
		Case "Unreal Tournament 3 Demo"
			$strFileName = @ProgramFilesDir & "\Unreal Tournament 3 Demo\Binaries\UT3Demo.exe"

			;ImX3TC.au3
		Case "X3 Terran Conflict Demo"
			$strFileName = @ProgramFilesDir & "\EGOSOFT\X3 Terran Conflict Rolling Demo\X3TC_Demo.exe"

			;Im3DMkVN.au3
		Case "3DMark Vantage"
			$strFileName = @ProgramFilesDir & "\Futuremark\3DMark Vantage\3DMarkVantage.exe"

			;ImPCMkVN.au3
		Case "PCMark Vantage"
			$strFileName = @ProgramFilesDir & "\Futuremark\PCMark Vantage\bin\PCMark.exe"

			;ImHvyLd.au3
		Case "HeavyLoad"
			$strFileName = @ProgramFilesDir & "\JAM Software\HeavyLoad\HeavyLoad.exe"


			; *** CAMARILLO SECTION ***
			;ImEtm.au3
		Case "Intel Extended Thermal Model"
			$blnCheckFile = False
			$blnResult = IsCamarilloInstalled()

			; ImTPT.au3
		Case "Intel TPT"
			$strFileName = @ProgramFilesDir & "\Intel Corporation\TPT\TPT.exe"

			; ImGfxFreqDisp.au3
		Case "Graphics Frequency Display"
			$strFileName = @ProgramFilesDir & "\Intel Corporation\Graphics Frequency Display\GfxFreqDisp.exe"

			; ImEtmDemo.au3
		Case "Intel Extended Thermal Model Demo"
			If IsWinXPSP2() Or IsWinVista32() Then
				$strFileName = @HomeDrive & "\APPS\ETMDemo\EtmDemo32.exe"
			Else
				$strFileName = @HomeDrive & "\APPS\ETMDemo\EtmDemo64.exe"
			EndIf

			; ImTestApp2.au3
		Case "TestApp2"
			$strFileName = @HomeDrive & "\APPS\TestApp2\TestApp2.exe"

			;ImGfxDX9Stress.au3
		Case "GFX_DX9_Stress"
			$strFileName = @HomeDrive & "\APPS\GFX_DX9_Stress\powerprofile.exe"

			; *** OSUPDATE SECTION ***
			; ImSP2QFE.au3
		Case "Windows XP SP2 GV3 QFE (KB896256)"
			$blnCheckFile = False
			$blnResult = InstallSP2QFEAlready()

			; ImXP64SP1GV3QFE.au3
		Case "Windows XP x64 SP1 GV3 QFE (KB896256)"
			$blnCheckFile = False
			$blnResult = IsInstallGV3QFEAlready()

			; ImEmerald.wbt
		Case "MCE Emerald Update"
			$blnCheckFile = False
			$strMCEDir = @WindowsDir & "\$NtUninstallKB900325$"
			$blnResult = DirExists($strMCEDir)

			; ImMCEQFE.au3
		Case "MCE Emerald Update QFE"
			$blnCheckFile = False
			$strMceQFEDir = @WindowsDir & "\$NtUninstallKB914548$"
			$blnResult = DirExists($strMceQFEDir)

			; ImUSB2Hotfix.au3
		Case "Windows XP SP2 USB2 Hotfix (KB918005)"
			$blnCheckFile = False
			$blnResult = IsInstallUSBHotfixAlready()

			; ImUSB2HotfixXP64.au3
		Case "Windows XP x64 SP2 USB2 Hotfix (KB918005)"
			$blnCheckFile = False
			$blnResult = IsInstallXP64USBHotfixAlready()

			; ImXP64USBYBHotfix.au3
		Case "Windows XP x64 SP1 USB Yellow Bang Hotfix (KB921411)"
			$blnCheckFile = False
			$blnResult = IsInstallUSBYBHotfixAlready()

			; ImHDAudioHotFix.au3
		Case "High Definition Audio Class Driver Hotfix (KB901105/KB888111)"
			$blnCheckFile = False
			$blnResult = IsInstallHDHotfixAlready()

			; ImXPSP2USBYBHotfix.au3
		Case "Windows XP SP2 USB Yellow Bang Hotfix (KB921411)"
			$blnCheckFile = False
			$blnResult = IsInstallUSBYBHotfixAlready()

			; ImDSTPatch.au3
		Case "Daylight Saving Time(DST) Patch"
			$blnCheckFile = False
			$blnResult = IsInstallDSTPatchAlready()

			; ImSP2XDHotfix.au3
		Case "Windows XP SP2 XD-bit S3 Hotfix (KB889673)"
			$blnCheckFile = False
			$blnResult = InstallSP2XDHotfix()

			; ImVistaKB938194.au3
		Case "Windows Vista Update (KB938194)"
			$blnCheckFile = False
			$blnResult = InstallVistaKB938194()

			; ImVistaKB938979.au3
		Case "Windows Vista Update (KB938979)"
			$blnCheckFile = False
			$blnResult = InstallVistaKB938979()

			; ImVistaKB939008.au3
		Case "Windows Vista ReadyBoot QFE (KB939008)"
			$blnCheckFile = False
			$blnResult = InstallVistaKB939008()

			; ImPenrynC6QFE.au3
		Case "Windows XP SP2 Penryn C6 QFE (KB940566)"
			$blnCheckFile = False
			$blnResult = InstallXPKB940566()

		Case ".NET 3.0"
			$strFileName = @HomeDrive & "\WINNT\Microsoft.NET\Framework\v3.0\WPF"

		Case ".NET 3.0 x64"
			$strFileName = @HomeDrive & "\WINNT\Microsoft.NET\Framework64\v3.0\WPF"

			; *** DRIVER SECTION ***
			; ImInf.au3
		Case "Intel Chipset Software Installation Utility"
			$blnCheckFile = False
			$blnResult = InstallInfAlready()

			; ImDX90.au3
		Case "Direct X 9.0 Drivers"
			$strFileName = @SystemDir & "\d3d9.dll"

			; ImMCodec.wbt
		Case "Windows Media Video 9 VCM Codec"
			$strFileName = @ProgramFilesDir & "\WMV9_VCM\license.txt"

			; ImKLCodec.wbt
		Case "K-Lite Codec"
			$strFileName = @ProgramFilesDir & "\K-Lite Codec Pack\unins000.exe"

			; ImKLDivX.wbt
		Case "K-Lite DivX Codec"
			$strFileName = @ProgramFilesDir & "\K-Lite Codec Pack\unins000.exe"

			; ImKLXvid.wbt
		Case "K-Lite Xvid Codec"
			$strFileName = @ProgramFilesDir & "\K-Lite Codec Pack\unins000.exe"

			; ImIMSM.au3
		Case "Intel Matrix Storage Manager"
			$strFileName = @ProgramFilesDir & "\Intel\Intel Matrix Storage Manager\IAANTmon.exe"
			
			; ImMEI_SOL.au3
		Case "MEI SOL Driver"
			$strFileName = @ProgramFilesDir & "\Intel\Intel(R) Active Management Technology\uninstall\ar-SA\license.txt"
			
;~ 	; ImMuroc.wbt
;~ 	Case "Muroc 10 - Intel Wireless Driver Installer"
;~
;~
;~ 	; ImLoudon.wbt
;~ 	Case "Loudon - Intel Wireless Driver Installer"

			; *** Utility Section ***
			; ImCSwt.au3
		Case "Intel CSwitch Utility"
			$strFileName = @ProgramFilesDir & "\Intel Corporation\Intel(R) CSwitch Utility\CSwitch.exe"

			; ImProcLo.wbt
		Case "Intel ProcLoad Utility"
			$strFileName = @ProgramFilesDir & "\Intel Corporation\Intel(R) ProcLoad Utility\ProcLoad.exe"

			; ImWinPM.au3
		Case "Intel WinPM"
			$strFileName = @ProgramFilesDir & "\Intel Corporation\WinPM\WinPM.exe"

			; ImCPUFrq.au3
		Case "Intel CPU Frequency Display"
			$strFileName = @ProgramFilesDir & "\Intel Corporation\Intel(R) Frequency Display Tool\FreqDsp.exe"

			; ImTAT.au3
		Case "Intel Thermal Analysis Tool"
			$strFileName = @ProgramFilesDir & "\Intel Corporation\Intel(R) Thermal Analysis Tool\TAT.exe"
			;$strFileName = @ProgramFilesDir & "\Intel Corporation\Thermal Analysis Tool\TAT.exe"

			; ImVerInf.wbt
		Case "Intel WinVerInfo"
			$strFileName = @ProgramFilesDir & "\Intel Corp\WinVerInfo\WinInfo.sys"

			; ImGVCyl.au3
		Case "Intel GVCycle Utility"
			$strFileName = @ProgramFilesDir & "\Intel Corporation\Intel(R) GVCycle Tool\GVCycle.exe"

			; ImCStResMon.au3
		Case "Intel C State Residency Monitor Utility"
			$strFileName = @ProgramFilesDir & "\Intel Corporation\CStResMon\ReadMe.html"

			; DiskLoad
		Case "DiskLoad"
			$strFileName = @HomeDrive & "\bin\DiskLoad.exe"

			; WinFlash
		Case "Intel WinFlash Utility"
			$strFileName = @ProgramFilesDir & "\Intel Corporation\Intel(R) WinFlash Tool\WinFlash.exe"

			; *** VIDEO SECTION ***
			; ImIGD.au3
		Case "Intel Extreme Graphics"
			$blnCheckFile = False
			$blnResult = InstallSRIGDAlready()

			; ImGraphics.au3
		Case "Intel Graphics Media Accelerator Driver"
			$blnCheckFile = False
			$blnResult = InstallSRIGDAlready()

			; ImGraphicsCT.au3
		Case "Intel Graphics Media Accelerator Driver for Cantiga"
			$blnCheckFile = False
			$blnResult = InstallSRIGDAlready()

			; ImGraphicsAB.au3
		Case "Intel Graphics Media Accelerator Driver for Calpella / PineView"
			$blnCheckFile = False
			$blnResult = InstallSRIGDAlready()

		Case "Intel Graphics Media Accelerator Driver for SandyBridge / Cougar Point"
			$blnCheckFile = False
			$blnResult = InstallSRIGDAlready()

			; ImATIpcie.wbt
		Case "ATI PCI-Express Video Adapter"
			$blnCheckFile = False

			; ImATIagp.wbt
		Case "ATI AGP Video Adapter"
			$blnCheckFile = False

			; ImNVpcie.wbt
		Case "Nvidia PCI Express Video Adapter"
			$blnCheckFile = False

			; *** LICENSED-General SECTION ***
			; ImPwrDVD.au3
		Case "CyberLink PowerDVD 8"
			$strFileName = @ProgramFilesDir & "\CyberLink\PowerDVD8\PowerDVD8.exe"

			; ImWinDVD.au3
		Case "Corel WinDVD 9"
			$strFileName = @ProgramFilesDir & "\Corel\DVD9\WinDVD.exe"

			; ImSandraPro.au3
		Case "SiSoftware Sandra Pro"
			$strFileName = @HomeDrive & "\Program Files\SiSoftware\SiSoftware Sandra Professional Business XI.SP3\sandra.exe"

			; ImOffice.wbt
		Case "Microsoft Office 2003"
			$strFileName = @ProgramFilesDir & "\Microsoft Office\OFFICE11\OUTLOOK.EXE"

			; ImWDBench5.wbt
		Case "WorldBench 5"
			$strFileName = @HomeDrive & "\WorldBench\wb.exe"

			; *** LICENSED-CV SECTION ***

			; *** Development ***

		Case "DebugView"
			$strFileName = @HomeDrive & "\Apps\DebugView\Dbgview.exe"

		Case "HeavyLoad"
			$strFileName = @ProgramFilesDir & "\JAM Software\HeavyLoad\HeavyLoad.exe"


			; *** TXT SECTION ***

			; MPGTXT Driver
			;Case "MPGTXT Driver"													; grayed out because currently not needed (added when addind support for test script 56. MPG TXT Test)
			;$strFileName = @HomeDrive & "\Apps\MPGTXT\MpgTxt.inf"

			;SINIT ACM
			;Case "SINIT ACM"														; grayed out because currently not needed (added when addind support for test script 56. MPG TXT Test)
			;$strFileName = @WindowsDir & "\System32\drivers\ACM.BIN"

			; TXT MPG App
		Case "TXT MPG App"
			$strFileName = @HomeDrive & "\Apps\TXT Tests\TXT MPG App.exe" ; file necessary to run test 56. MPG TXT Test

			; *** DEVELOPMENT SECTION ***
			; ImAutoIt3.au3
		Case "AutoIt3"
			$strFileName = @ProgramFilesDir & "\AutoIt3\AutoIt3.exe"

			; ImAutoItDebugger.au3
		Case "AutoIt3 Debugger"
			$strFileName = @ProgramFilesDir & "\AutoIt3\AutoIt Debugger\AutoIt Debugger.exe"

			; ImSciTE4.au3
		Case "SciTE for AutoIt3"
			$strFileName = @ProgramFilesDir & "\AutoIt3\SciTE\SciTE.exe"

			; ImWMICodeCreator.au3
		Case "WMI Code Creator"
			$strFileName = @HomeDrive & "\Apps\WMICodeCreator\WMICodeCreator.exe"

			; *** Custom SECTION ***

			; ImWMICodeCreator.au3
		Case "Automate V6"
			$strFileName = @ProgramFilesDir & "\AutoMate 6\AMTA.exe"

		Case Else
			; Something else ?

	EndSwitch

	If ($blnCheckFile) Then
		Return FileExists($strFileName)
	Else
		Return $blnResult
	EndIf

EndFunc   ;==>IsAppletInstalled

;***************************************************************************
;** Function: 		GetTATSourceDir()
;** Parameters:
;**		None
;** Description:
;**		This function is called to get the TAT soruce dir.
;** Return:
;**		string source dir name.
;***************************************************************************
Func GetTATSourceDir()
	$strPlatfromFamily = StringUpper(GetPlatformFamilyName())
	;if IsWinXP64Family() Or ($strPlatfromFamily == "CALISTOGA") Or ($strPlatfromFamily == "CRESTLINE") _
	;Or ($strPlatfromFamily == "CANTIGA") Then
	If $strPlatfromFamily <> "ALVISO" Then
		$strSubDir = "Thermal_Analysis_Tool_for_multi-core_CPUs"
	Else
		$strSubDir = "Thermal_Analysis_Tool"
	EndIf

	Return $strSubDir
EndFunc   ;==>GetTATSourceDir

;***************************************************************************
;** Function: 		GetIGDSourceDir()
;** Parameters:
;**		None
;** Description:
;**		This function is called to check if get the IGD source dir string.
;**
;**		Win2k & WinXP (32 bit) Intel video drivers for the Alviso platform are now available from:
;**			\\lestat1\drivers\video\Intel\Extreme_Graphics\Win2K-XP.Alviso
;**		The Alviso platform does not support 64 bit operating systems and the 32 bit Vista graphics drivers are
;**		built in to the Vista OS for this old platform.
;**
;**		Win2k & WinXP (32 bit) Intel video drivers for the Calistoga & Crestline platforms are now available from:
;**			\\lestat1\drivers\video\Intel\Extreme_Graphics\Win2K-XP
;**
;**		WinXP x64 Intel video drivers for the Calistoga & Crestline platforms are now available from:
;**			\\lestat1\drivers\video\Intel\Extreme_Graphics\XP64
;**
;**		Vista 32 bit Intel video drivers for the Calistoga & Crestline platforms are now available from:
;**			\\lestat1\drivers\video\Intel\Extreme_Graphics\Vista\32bit
;**
;**		Vista 64 bit Intel video drivers for the Calistoga & Crestline platforms are now available from:
;**			\\lestat1\drivers\video\Intel\Extreme_Graphics\Vista\64bit
;**
;** Return:
;**		IGD source dir string.
;**
;***************************************************************************
Func GetIGDSourceDir()

	$strResult = False

	$strPlatfromFamily = StringUpper(GetPlatformFamilyName())

	; Windows XP / 2K - Calistoga, Crestline, Cantiga, IbexPeak
	If IsWin2KFamily() Or IsWinXPFamily() Then
		If ($strPlatfromFamily == "CALISTOGA") Then
			$strResult = "Win2K-XP.Calistoga"
		ElseIf ($strPlatfromFamily == "CRESTLINE") Then
			$strResult = "Win2K-XP.Crestline"
		ElseIf ($strPlatfromFamily == "CANTIGA") Then
			$strResult = "Win2k-XP.Cantiga"
		ElseIf ($strPlatfromFamily == "IBEXPEAK") Then
			$strResult = "Win2K-XP.IbexPeak"
		ElseIf ($strPlatfromFamily == "PINEVIEW") Then
			$strResult = "Win2K-XP.Pineview"
		ElseIf ($strPlatfromFamily == "SANDYBRIDGE") Then
			$strResult = "Win2K-XP.SandyBridge"
		ElseIf ($strPlatfromFamily == "SANDYBRIDGE-CPT") Then
			$strResult = "Win2K-XP.SandyBridge-CPT"
			;Else
			;$strResult = "Win2K-XP.Alviso"
			;	WriteLog("Unsupported platform - IGD error")
		EndIf

		; Windows XP x64 - Calistoga, Crestline, Cantiga, IbexPeak
	ElseIf IsWinXP64Family() Then
		If ($strPlatfromFamily == "CANTIGA") Then
			$strResult = "XP64.Cantiga"
		ElseIf ($strPlatfromFamily == "IBEXPEAK") Then
			$strResult = "XP64.IbexPeak"
		ElseIf ($strPlatfromFamily == "CALISTOGA") Then
			$strResult = "XP64.Calistoga"
		ElseIf ($strPlatfromFamily == "CRESTLINE") Then
			$strResult = "XP64.Crestline"
		ElseIf ($strPlatfromFamily == "PINEVIEW") Then
			$strResult = "XP64.Pineview"
		ElseIf ($strPlatfromFamily == "SANDYBRIDGE") Then
			$strResult = "XP64.SandyBridge"
		ElseIf ($strPlatfromFamily == "SANDYBRIDGE-CPT") Then
			$strResult = "XP64.SandyBridge-CPT"
			;Else
			;$strResult = "XP64"
			;WriteLog("Unsupported platform - IGD error")
		EndIf

		; Vista x32 - Calistoga, Crestline, Cantiga, IbexPeak
	ElseIf IsWinVista32() Then
		If ($strPlatfromFamily == "CALISTOGA") Then
			$strResult = "Vista.Calistoga\32bit"
		ElseIf ($strPlatfromFamily == "CRESTLINE") Then
			$strResult = "Vista.Crestline\32bit"
		ElseIf ($strPlatfromFamily == "CANTIGA") Then
			$strResult = "Vista.Cantiga\32bit"
		ElseIf ($strPlatfromFamily == "IBEXPEAK") Then
			$strResult = "Vista.IbexPeak\32bit"
		ElseIf ($strPlatfromFamily == "PINEVIEW") Then
			$strResult = "Vista.Pineview\32bit"
		ElseIf ($strPlatfromFamily == "SANDYBRIDGE") Then
			$strResult = "Vista.SandyBridge\32bit"
		ElseIf ($strPlatfromFamily == "SANDYBRIDGE-CPT") Then
			$strResult = "Vista.SandyBridge-CPT\32bit"
			;Else
			;WriteLog("Unsupported platform - IGD error")
		EndIf

		; Vista x64 - Calistoga, Crestline, Cantiga, IbexPeak
	ElseIf IsWinVista64() Then
		If ($strPlatfromFamily == "CALISTOGA") Then
			$strResult = "Vista.Calistoga\64bit"
		ElseIf ($strPlatfromFamily == "CRESTLINE") Then
			$strResult = "Vista.Crestline\64bit"
		ElseIf ($strPlatfromFamily == "CANTIGA") Then
			$strResult = "Vista.Cantiga\64bit"
		ElseIf ($strPlatfromFamily == "IBEXPEAK") Then
			$strResult = "Vista.IbexPeak\64bit"
		ElseIf ($strPlatfromFamily == "PINEVIEW") Then
			$strResult = "Vista.Pineview\64bit"
		ElseIf ($strPlatfromFamily == "SANDYBRIDGE") Then
			$strResult = "Vista.SandyBridge\64bit"
		ElseIf ($strPlatfromFamily == "SANDYBRIDGE-CPT") Then
			$strResult = "Vista.SandyBridge-CPT\64bit"
			;Else
			;WriteLog("Unsupported platform - IGD error")
		EndIf

		; Windows 7 x32 - Calistoga, Crestline, Cantiga, IbexPeak
	ElseIf IsWin7_32() Then
		If ($strPlatfromFamily == "CALISTOGA") Then
			$strResult = "Win7.Calistoga\32bit"
		ElseIf ($strPlatfromFamily == "CRESTLINE") Then
			$strResult = "Win7.Crestline\32bit"
		ElseIf ($strPlatfromFamily == "CANTIGA") Then
			$strResult = "Win7.Cantiga\32bit"
		ElseIf ($strPlatfromFamily == "IBEXPEAK") Then
			$strResult = "Vista.IbexPeak\32bit"
		ElseIf ($strPlatfromFamily == "PINEVIEW") Then
			$strResult = "Win7.Pineview\32bit"
		ElseIf ($strPlatfromFamily == "SANDYBRIDGE") Then
			$strResult = "Win7.SandyBridge\32bit"
		ElseIf ($strPlatfromFamily == "SANDYBRIDGE-CPT") Then
			;$strResult = "Win7.SandyBridge-CPT\32bit"
			$strResult = "Vista.SandyBridge-CPT\32bit"
			;Else
			;WriteLog("Unsupported platform - IGD error")
		EndIf

		; Windows 7 - Calistoga, Crestline, Cantiga, IbexPeak
	ElseIf IsWin7_64() Then
		If ($strPlatfromFamily == "CALISTOGA") Then
			$strResult = "Win7.Calistoga\64bit"
		ElseIf ($strPlatfromFamily == "CRESTLINE") Then
			$strResult = "Win7.Crestline\64bit"
		ElseIf ($strPlatfromFamily == "CANTIGA") Then
			$strResult = "Win7.Cantiga\64bit"
		ElseIf ($strPlatfromFamily == "IBEXPEAK") Then
			$strResult = "Vista.IbexPeak\64bit"
		ElseIf ($strPlatfromFamily == "PINEVIEW") Then
			$strResult = "Win7.Pineview\64bit"
		ElseIf ($strPlatfromFamily == "SANDYBRIDGE") Then
			$strResult = "Win7.SandyBridge\64bit"
		ElseIf ($strPlatfromFamily == "SANDYBRIDGE-CPT") Then
			;$strResult = "Win7.SandyBridge-CPT\64bit"
			$strResult = "Vista.SandyBridge-CPT\64bit"
		EndIf

	Else
		; Vista x32/x64 for older platform
		;	The Alviso platform does not support 64 bit operating systems and the 32 bit Vista graphics drivers are
		;	built in to the Vista OS for this old platform.
		WriteLog("Unsupported platform - IGD error")
		$strResult = False
	EndIf

	Return $strResult
EndFunc   ;==>GetIGDSourceDir

;***************************************************************************
;** Function: 		DiskloadGUI($strDisk, $strName)
;** Parameters:
;**		None
;** Description:
;**		This function is called to get the user's input for Disklaod logical disk/thread.
;**
;** Return:
;**		True/False.
;**
;***************************************************************************
Func DiskloadGUI($strDisk, $strName)

	; Check registry value
	$strKey = "HKLM\SOFTWARE\Intel\MPG\TESTSCRIPT"
	$strSubKey = "DiskLoadCmd"
	$strResult = RegRead($strKey, $strSubKey)
	If $strResult Then
		Return True
	EndIf

	Dim $chkDisk[15]
	$retVal = True

	#Region ### START Koda GUI section ### Form=C:\MY PROJECT\AutoIt3\SCRIPTS\frmDiskload.kxf
	$Form1_1 = GUICreate("DiskLoad Tool", 346, 279, 193, 115)
	$Group1 = GUICtrlCreateGroup("Select Logical Disk(s)", 16, 16, 313, 169)

	For $intIndex = 1 To $strDisk[0]

		$chkDisk[$intIndex - 1] = GUICtrlCreateCheckbox($strDisk[$intIndex] & $strName[$intIndex], 40, 25 + 20 * $intIndex, 145, 17)
		If StringInStr($strDisk[$intIndex], "C:") Then
			GUICtrlSetState($chkDisk[$intIndex - 1], $GUI_CHECKED)
		EndIf
		;$Radio2 = GUICtrlCreateRadio("ARadio2", 56, 60, 113, 17)
		;$Radio3 = GUICtrlCreateRadio("ARadio3", 56, 80, 113, 17)
	Next

	GUICtrlCreateGroup("", -99, -99, 1, 1)
	$btnOK = GUICtrlCreateButton("OK", 264, 208, 75, 25, $BS_DEFPUSHBUTTON)
	$btnCancel = GUICtrlCreateButton("Cancel", 264, 240, 75, 25, 0)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				ExitLoop
			Case $btnOK
				$msg = $GUI_EVENT_CLOSE
				ExitLoop
			Case $btnCancel
				$msg = $GUI_EVENT_CLOSE
				$retVal = False
				ExitLoop
		EndSwitch
	WEnd

	; Close DLL
	;DllClose($dll)

	; Get checkbox state
	If $retVal Then
		$strCmd = ""
		For $intIndex = 1 To $strDisk[0]
			$chkState = GUICtrlRead($chkDisk[$intIndex - 1])

			If $chkState == $GUI_CHECKED Then
				$intThread = InputBox("DiskLoad Tool", "Enter desired number of threads (1 - 10) for " & $strDisk[$intIndex], "10")

				; The Cancel button was pushed.
				If @error == 1 Then
					Return False
				Else
					$strCmd = $strCmd & $strDisk[$intIndex] & $intThread & " "
				EndIf

			EndIf

		Next

		; Update registry
		RegWrite("HKLM\SOFTWARE\Intel\MPG\TESTSCRIPT", "DiskLoadCmd", "REG_SZ", $strCmd)

	EndIf

	Return $retVal

EndFunc   ;==>DiskloadGUI

;***************************************************************************
;** Function: 		BusyIdleGUI(ByRef $intTPS, ByRef $strVideo)
;** Parameters:
;**		None
;** Description:
;**		This function is called to get the user's input for Busy/Idle test.
;**
;** Return:
;**		True/False.
;**
;***************************************************************************
Func BusyIdleGUI(ByRef $intTPS, ByRef $strVideo)

	$strKey = "HKLM\SOFTWARE\Intel\MPG\TESTSCRIPT"
	$intTPS = RegRead($strKey, "ISTCFG")
	If ($intTPS) Then
		Return True
	EndIf

	$retVal = True
	$intTPS = 100
	If (IsWinVista64() Or IsWin7_64()) Then
		$strVideo = "3DMark 2001"
	Else
		$strVideo = "Glaze V3.1 Setup"
	EndIf

	#Region ### START Koda GUI section ### Form=c:\my project\autoit3\scripts\frmbusyidle.kxf
	$Form1 = GUICreate("IST Transitioning GUI", 324, 370, 193, 126)
	$Group1 = GUICtrlCreateGroup("IST Transitions/Second", 16, 16, 289, 145)
	$rad1TPS = GUICtrlCreateRadio("1 Transition Per Second", 40, 40, 177, 17)
	$rad5TPS = GUICtrlCreateRadio("5 Transitions Per Second", 40, 64, 169, 17)
	$rad10TPS = GUICtrlCreateRadio("10 Transitions Per Second", 40, 88, 177, 17)
	$radNoneTPS = GUICtrlCreateRadio("None: CPU doesn't support IST", 40, 136, 177, 17)
	$rad100TPS = GUICtrlCreateRadio("100 Transition Per Second", 40, 112, 177, 17)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	$btnOK = GUICtrlCreateButton("OK", 232, 288, 75, 25, $BS_DEFPUSHBUTTON)
	$btnCancel = GUICtrlCreateButton("Cancel", 232, 320, 75, 25, 0)
	$Group2 = GUICtrlCreateGroup("Select Video Application", 16, 176, 185, 177)
	$radGlaze = GUICtrlCreateRadio("Glaze", 40, 200, 113, 17)
	$radFinalR = GUICtrlCreateRadio("Final Reality", 40, 224, 113, 17)
	$rad3DMk2001 = GUICtrlCreateRadio("3DMark 2001", 40, 248, 113, 17)
	$rad3DMk2003 = GUICtrlCreateRadio("3DMark 2003", 40, 272, 113, 17)
	$rad3DMk2005 = GUICtrlCreateRadio("3DMark 2005", 40, 296, 113, 17)
	$rad3DMk2006 = GUICtrlCreateRadio("3DMark 2006", 40, 320, 113, 17)


	GUICtrlCreateGroup("", -99, -99, 1, 1)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

	; Set 10 tps as default, Glaze as default for Vista32, 3DMark2001 for Vista64.
	GUICtrlSetState($rad100TPS, $GUI_CHECKED)
	If (IsWinVista64() Or IsWin7_64()) Then
		GUICtrlSetState($rad3DMk2001, $GUI_CHECKED)
	Else
		GUICtrlSetState($radGlaze, $GUI_CHECKED)
	EndIf

	; Grey out
	If (IsWinVista64() Or IsWin7_64()) Then
		GUICtrlSetState($radGlaze, $GUI_DISABLE)
		GUICtrlSetState($radFinalR, $GUI_DISABLE)
	EndIf

	; TO DO: TEMPORART GREYOUT
	GUICtrlSetState($rad3DMk2003, $GUI_DISABLE)
	GUICtrlSetState($rad3DMk2005, $GUI_DISABLE)
	GUICtrlSetState($rad3DMk2006, $GUI_DISABLE)


	While 1
		$nMsg = GUIGetMsg()
		Select
			Case $nMsg == $btnOK
				$msg = $GUI_EVENT_CLOSE
				ExitLoop
			Case $nMsg == $btnCancel
				$msg = $GUI_EVENT_CLOSE
				$retVal = False
				ExitLoop
			Case $nMsg == $GUI_EVENT_CLOSE
				ExitLoop

				; IST TPS
			Case $nMsg == $rad1TPS
				$intTPS = 1
			Case $nMsg == $rad5TPS
				$intTPS = 5
			Case $nMsg == $rad10TPS
				$intTPS = 10
			Case $nMsg == $rad100TPS
				$intTPS = 100
			Case $nMsg == $radNoneTPS
				$intTPS = 0

				; Video applet
			Case $nMsg == $radGlaze
				$strVideo = "Glaze V3.1 Setup"
			Case $nMsg == $radFinalR
				$strVideo = "Final Reality Benchmark 1.01"
			Case $nMsg == $rad3DMk2001
				$strVideo = "3DMark 2001"
			Case $nMsg == $rad3DMk2003
				$strVideo = "3DMark 2003 Professional Version"
			Case $nMsg == $rad3DMk2005
				$strVideo = "3DMark 2005 Business Edition"
			Case $nMsg == $rad3DMk2006
				$strVideo = "3DMark 2006"

		EndSelect
	WEnd

	GUISetState(@SW_HIDE)

	; Write registry
	If ($retVal) Then
		$strKey = "HKLM\SOFTWARE\Intel\MPG\TESTSCRIPT"
		RegWrite($strKey, "ISTCFG", "REG_SZ", $intTPS)
	EndIf

	Return $retVal

EndFunc   ;==>BusyIdleGUI

;***************************************************************************
;** Function: 		KillTask(strTask)
;** Parameters:
;**	None
;** Description:
;**	This function is called to kill the specific task passed to this function.
;** Return:
;**	None
;***************************************************************************
Func KillTask($strTask)


	If IsWin2KFamily() Then
		Run(@HomeDrive & "\bin\pv.exe -k -f " & $strTask)
	Else
		Run(@SystemDir & "\taskkill.exe /im " & $strTask & " /F /T")
	EndIf

EndFunc   ;==>KillTask

;***************************************************************************
;** Function: 		RunScheduler(strScriptName)
;** Parameters:
;**	None
;** Description:
;**	This function is called to run scheduler.
;** Return:
;**	Script name of the video application.
;***************************************************************************
Func RunScheduler($strScriptName)

	; Set BusyIdleScript registry value to
	$strKey = "HKLM\Software\Intel\MPG\TESTSCRIPT\BusyIdle"
	$strSubKey = "TestScript"

	RegWrite($strKey, $strSubKey, "REG_SZ", $strScriptName)
	WriteLog("Set MPG\TESTAPP\BusyIdle[TestScript] registry to " & $strScriptName)

	; Run TmSchedule.exe
	$strExe = @HomeDrive & "\scripts\TmScheduler.exe " & $strScriptName
	RunWait($strExe)

	; Check status
	$strKey = "HKLM\Software\Intel\MPG\TESTAPP"
	$strSubKey = "Busy/Idle Scheduler Script"

	$strStatus = RegRead($strKey, $strSubKey)
	$blnResult = StringInStr($strStatus, "ERROR")
	If $blnResult Then
		Return False
	EndIf

	Return True

EndFunc   ;==>RunScheduler

;***************************************************************************
;** Function: 		GetNextScheduleTime()
;** Parameters:
;**	None
;** Description:
;**	This function is called to next scheduled time.
;** Return:
;**	string value of the registry
;***************************************************************************
Func GetNextScheduleTime()

	$strKey = "HKLM\Software\Intel\MPG\TESTSCRIPT\BusyIdle"
	$strSubKey = "NextScheduleTime"

	$strNextScheduleTime = RegRead($strKey, $strSubKey)

	$strDate = GetDate($strNextScheduleTime)
	$strTime = GetTime($strNextScheduleTime)

	$strNextScheduleTime = $strDate & " at " & $strTime

	Return $strNextScheduleTime

EndFunc   ;==>GetNextScheduleTime

;***************************************************************************
;** Function: 		GetTime(strNext)
;** Parameters:
;**	strNext - a YYYY:MM:DD:HH:MM:SS format time
;** Description:
;**	This function is called to get the extract time from a YYYY:MM:DD:HH:MM:SS format time.
;** Return:
;**	None
;***************************************************************************
Func GetTime($strNext)
	$strTime = StringRight($strNext, 8)

	Return $strTime

EndFunc   ;==>GetTime

;***************************************************************************
;** Function: 		GetDate(strNext)
;** Parameters:
;**		strNext - a YYYY:MM:DD:HH:MM:SS format time
;** Description:
;**		This function is called to get the extract time from a YYYY:MM:DD:HH:MM:SS format time.
;** Return:
;**	None
;***************************************************************************
Func GetDate($strNext)
	; Vista Schtasks.exe need /SD mm/dd/yyyy format
	$strDate = StringLeft($strNext, 10)

	$strArray = StringSplit($strDate, ":")
	;$strArray[1] contains "YYYY"
	;$strArray[2] contains "MM"
	;$strArray[3] contains "DD"
	$strDate = $strArray[2] & "/" & $strArray[3] & "/" & $strArray[1]

	Return $strDate

EndFunc   ;==>GetDate

;***************************************************************************
;** Function: 		WriteIdleLog(strMsg)
;** Parameters:
;**	None
;** Description:
;**	This function is called to write string to idle log file - c:\logs\Idle.txt.
;** Return:
;**	None
;** Usage:
;**	WriteLog("message you want write")
;**
;***************************************************************************
Func WriteIdleLog($strMsg)

	$strIdleFile = @HomeDrive & "\logs\Idle.txt"

	; Open file for append. If file not exists, will create
	$strFileHandle = FileOpen($strIdleFile, 1)

	FileWriteLine($strFileHandle, $strMsg)

	FileClose($strFileHandle)

EndFunc   ;==>WriteIdleLog


;***************************************************************************
;** Function: 		PwrTestGUI()
;** Parameters:
;**	None
;** Description:
;**	This function is called to run test.
;** Return:
;**	None
;***************************************************************************
Func BusyIdlePwrTestGUI(ByRef $intCycles, ByRef $intDelay, ByRef $intSleep, ByRef $strSState, ByRef $strCyclingThrough, $strPwrTestGUIFile)
	;#include <GUIConstants.au3>

	; == GUI generated with Koda ==
	$Form1 = GUICreate("PwrTest SLEEP Scenario", 416, 475, 192, 125)
	$Group1 = GUICtrlCreateGroup("Number of Cycles", 16, 16, 385, 129)
	$rad10Cycles = GUICtrlCreateRadio("10", 32, 40, 113, 17)
	GUICtrlSetState(-1, $GUI_CHECKED)
	$rad20Cycles = GUICtrlCreateRadio("20", 32, 64, 113, 17)
	$txtCycles = GUICtrlCreateInput("3000", 272, 112, 121, 21)
	$rad30Cycles = GUICtrlCreateRadio("30", 32, 88, 113, 17)
	$rad100Cycles = GUICtrlCreateRadio("100", 152, 40, 81, 17)
	$rad200Cycles = GUICtrlCreateRadio("200", 152, 64, 73, 17)
	$rad300Cycles = GUICtrlCreateRadio("300", 152, 88, 73, 17)
	$rad1000Cycles = GUICtrlCreateRadio("1000", 248, 40, 129, 17)
	$rad2000Cycles = GUICtrlCreateRadio("2000", 248, 64, 113, 17)
	$radEnterCycles = GUICtrlCreateRadio("Enter number of cycles", 248, 88, 137, 17)
	$rad50Cycles = GUICtrlCreateRadio("50", 32, 112, 113, 17)
	$rad500Cycles = GUICtrlCreateRadio("500", 152, 112, 113, 17)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	$Group2 = GUICtrlCreateGroup("Sleep Time in Seconds", 216, 160, 185, 105)
	$rad60Sleep = GUICtrlCreateRadio("60 seconds", 232, 184, 113, 17)
	GUICtrlSetState(-1, $GUI_CHECKED)
	$radEnterSleep = GUICtrlCreateRadio("Enter number in seconds", 232, 208, 145, 17)
	$txtSleep = GUICtrlCreateInput("30", 264, 232, 121, 21)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	$Group3 = GUICtrlCreateGroup("Delay Time in Seconds", 16, 160, 185, 105)
	$rad90Delay = GUICtrlCreateRadio("90 seconds", 32, 184, 113, 17)
	GUICtrlSetState(-1, $GUI_CHECKED)
	$radEnterDelay = GUICtrlCreateRadio("Enter number in seconds", 32, 208, 145, 17)
	$txtDelay = GUICtrlCreateInput("30", 64, 232, 121, 21)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	$Group4 = GUICtrlCreateGroup("Cycling Throught All Power States", 216, 280, 185, 105)
	$radInOrder = GUICtrlCreateRadio("In Order", 240, 304, 113, 17)
	GUICtrlSetState(-1, $GUI_CHECKED)
	GUICtrlSetState(-1, $GUI_DISABLE)
	$radRandomly = GUICtrlCreateRadio("Randomly", 240, 328, 113, 17)
	GUICtrlSetState(-1, $GUI_DISABLE)
	$Group5 = GUICtrlCreateGroup("Power State", 16, 280, 185, 105)
	$radS3 = GUICtrlCreateRadio("S3", 32, 304, 113, 17)
	GUICtrlSetState(-1, $GUI_CHECKED)
	$radS4 = GUICtrlCreateRadio("S4", 32, 328, 113, 17)
	$radALL = GUICtrlCreateRadio("All", 32, 352, 113, 17)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	$btnOK = GUICtrlCreateButton("OK", 312, 408, 75, 25, $BS_DEFPUSHBUTTON)
	$btnCancel = GUICtrlCreateButton("Cancel", 312, 440, 75, 25, 0)
	GUISetState(@SW_SHOW)

	; Return value
	$retVal = True
	$blnInputCycle = False
	$blnInputDelay = False
	$blnInputSleep = False

	While True
		$msg = GUIGetMsg()

		Switch $msg
			Case $btnOK
				$msg = $GUI_EVENT_CLOSE
				ExitLoop
			Case $btnCancel
				$msg = $GUI_EVENT_CLOSE
				$retVal = False
				ExitLoop
			Case $GUI_EVENT_CLOSE

				If $blnInputCycle Then
					$intCycles = GUICtrlRead($txtCycles)
				EndIf
				If $blnInputDelay Then
					$intDelay = GUICtrlRead($txtDelay)
				EndIf
				If $blnInputSleep Then
					$intSleep = GUICtrlRead($txtSleep)
				EndIf

				ExitLoop

				; Get Cycles
			Case $rad10Cycles
				$intCycles = 10
				$blnInputCycle = False
			Case $rad20Cycles
				$intCycles = 20
				$blnInputCycle = False
			Case $rad30Cycles
				$intCycles = 30
				$blnInputCycle = False
			Case $rad50Cycles
				$intCycles = 50
				$blnInputCycle = False
			Case $rad100Cycles
				$intCycles = 100
				$blnInputCycle = False
			Case $rad200Cycles
				$intCycles = 200
				$blnInputCycle = False
			Case $rad300Cycles
				$intCycles = 300
				$blnInputCycle = False
			Case $rad500Cycles
				$intCycles = 500
				$blnInputCycle = False
			Case $rad1000Cycles
				$intCycles = 1000
				$blnInputCycle = False
			Case $rad2000Cycles
				$intCycles = 2000
				$blnInputCycle = False
			Case $radEnterCycles
				$blnInputCycle = True

				; Get Delay time
			Case $rad90Delay
				$intDelay = 90
				$blnInputDelay = False
			Case $radEnterDelay
				$blnInputDelay = True

				; Get Sleep time
			Case $rad60Sleep
				$intSleep = 90
				$blnInputSleep = False
			Case $radEnterSleep
				$blnInputSleep = True

				; Get power state
			Case $radS3
				$strSState = "3"
				GUICtrlSetState($radInOrder, $GUI_DISABLE)
				GUICtrlSetState($radRandomly, $GUI_DISABLE)
			Case $radS4
				$strSState = "4"
				GUICtrlSetState($radInOrder, $GUI_DISABLE)
				GUICtrlSetState($radRandomly, $GUI_DISABLE)
			Case $radALL
				$strSState = "all"
				GUICtrlSetState($radInOrder, $GUI_ENABLE)
				GUICtrlSetState($radRandomly, $GUI_ENABLE)

				; Get
			Case $radInOrder
				$strCyclingThrough = "all"
			Case $radRandomly
				$strCyclingThrough = "rnd"

		EndSwitch
	WEnd

	; Get textbox input
	If $retVal Then
		If $blnInputCycle Then
			$intCycles = GUICtrlRead($txtCycles)
		EndIf
		If $blnInputDelay Then
			$intDelay = GUICtrlRead($txtDelay)
		EndIf
		If $blnInputSleep Then
			$intSleep = GUICtrlRead($txtSleep)
		EndIf

		; Write PwrTest GUI file
		;;;$strPwrTestGUIFile = @HomeDrive & "\scripts\PwrTestGUI.txt"
		$strFileHandle = FileOpen($strPwrTestGUIFile, 1)

		$strLog = @HomeDrive & "\logs\PwrTestLog.wtl"
		$strExe = @HomeDrive & "\APPS\PwrTest\PwrTest.exe /sleep "
		If ($strSState <> "3") And ($strSState <> "4") Then
			$strCmd = $strExe & "/c:" & $intCycles & " /d:" & $intDelay & " /p:" & $intSleep & " /s:" & $strCyclingThrough & " /l:" & $strLog
		Else
			$strCmd = $strExe & "/c:" & $intCycles & " /d:" & $intDelay & " /p:" & $intSleep & " /s:" & $strSState & " /l:" & $strLog
		EndIf
		FileWriteLine($strFileHandle, $strCmd)

		; Call
		$strCmd = @HomeDrive & "\scripts\TmBIdleVista.exe -SCHEDULER"
		FileWriteLine($strFileHandle, $strCmd)
		FileWriteLine($strFileHandle, "Exit")

		FileClose($strFileHandle)
	EndIf


	Return $retVal

EndFunc   ;==>BusyIdlePwrTestGUI

;***************************************************************************
;** Function: 		PCMk05BusyIdleGUI()
;** Parameters:
;**	None
;** Description:
;**	This function is called to run test.
;** Return:
;**	None
;***************************************************************************
Func PCMk05BusyIdleGUI()

	; No need to run this if we already have it done.
	$strKey = "HKLM\Software\Intel\MPG\TESTSCRIPT\BusyIdle"
	$strSubKey = "PCMk05BusyIdleCycle"

	$intValue = RegRead($strKey, $strSubKey)
	If $intValue Then
		Return True
	EndIf

	#Region ### START Koda GUI section ### Form=c:\my project\autoit3\scripts\frmpcmk05bidle.kxf
	$Form1_1 = GUICreate("PCMk05 Busy Idle Test", 419, 208, 193, 126)
	$Group1 = GUICtrlCreateGroup("Number of Cycles", 16, 16, 385, 129)
	$rad10Cycles = GUICtrlCreateRadio("10", 32, 40, 113, 17)
	GUICtrlSetState(-1, $GUI_CHECKED)
	$rad20Cycles = GUICtrlCreateRadio("20", 32, 64, 113, 17)
	$txtCycles = GUICtrlCreateInput("3000", 272, 112, 121, 21)
	$rad30Cycles = GUICtrlCreateRadio("30", 32, 88, 113, 17)
	$rad100Cycles = GUICtrlCreateRadio("100", 152, 40, 81, 17)
	$rad200Cycles = GUICtrlCreateRadio("200", 152, 64, 73, 17)
	$rad300Cycles = GUICtrlCreateRadio("300", 152, 88, 73, 17)
	$rad1000Cycles = GUICtrlCreateRadio("1000", 248, 40, 129, 17)
	$rad2000Cycles = GUICtrlCreateRadio("2000", 248, 64, 113, 17)
	$radEnterCycles = GUICtrlCreateRadio("Enter number of cycles", 248, 88, 137, 17)
	$rad500Cycles = GUICtrlCreateRadio("500", 152, 112, 113, 17)
	$rad50Cycles = GUICtrlCreateRadio("50", 32, 112, 113, 17)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	$btnOK = GUICtrlCreateButton("OK", 240, 168, 75, 25, $BS_DEFPUSHBUTTON)
	$btnCancel = GUICtrlCreateButton("Cancel", 320, 168, 75, 25, 0)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

	$retVal = True
	$blnInputCycle = False
	$intCycles = 10

	While True
		$msg = GUIGetMsg()

		Switch $msg
			Case $btnOK
				$msg = $GUI_EVENT_CLOSE
				ExitLoop
			Case $btnCancel
				$msg = $GUI_EVENT_CLOSE
				$retVal = False
				ExitLoop
			Case $GUI_EVENT_CLOSE

				If $blnInputCycle Then
					$intCycles = GUICtrlRead($txtCycles)
				EndIf

				ExitLoop

				; Get Cycles
			Case $rad10Cycles
				$intCycles = 10
				$blnInputCycle = False
			Case $rad20Cycles
				$intCycles = 20
				$blnInputCycle = False
			Case $rad30Cycles
				$intCycles = 30
				$blnInputCycle = False
			Case $rad50Cycles
				$intCycles = 50
				$blnInputCycle = False
			Case $rad100Cycles
				$intCycles = 100
				$blnInputCycle = False
			Case $rad200Cycles
				$intCycles = 200
				$blnInputCycle = False
			Case $rad300Cycles
				$intCycles = 300
				$blnInputCycle = False
			Case $rad500Cycles
				$intCycles = 500
				$blnInputCycle = False
			Case $rad1000Cycles
				$intCycles = 1000
				$blnInputCycle = False
			Case $rad2000Cycles
				$intCycles = 2000
				$blnInputCycle = False
			Case $radEnterCycles
				$blnInputCycle = True

		EndSwitch
	WEnd

	; Close DLL
	;DllClose($dll)

	; Get textbox input
	If $retVal Then
		If $blnInputCycle Then
			$intCycles = GUICtrlRead($txtCycles)

		EndIf

		; Write registry
		$strKey = "HKLM\Software\Intel\MPG\TESTSCRIPT\BusyIdle"
		$strSubKey = "PCMk05BusyIdleCycle"
		RegWrite($strKey, $strSubKey, "REG_SZ", $intCycles)
		WriteBIdleLog("***********************")
		WriteBIdleLog("*****BUSY/IDLE Total Cycle: " & $intCycles)
		WriteBIdleLog("***********************")
	EndIf

	Return $retVal
EndFunc   ;==>PCMk05BusyIdleGUI

;***************************************************************************
;** Function: 		WriteBIdleLog(strMsg)
;** Parameters:
;**		$strmsg - string to write into log file.
;** Description:
;**		This function is called to write string to log file - c:\logs\manager.txt.
;** Return:
;**		None
;** Usage:
;**		WriteLog("message you want write")
;**
;***************************************************************************
Func WriteBIdleLog($strMsg)

	$strLogFile = @HomeDrive & "\Logs\BIdle.txt"

	; Open file for append. If file not exists, will create
	$strFileHandle = FileOpen($strLogFile, 1)

	$strMsg = _Now() & @TAB & @TAB & $strMsg
	FileWriteLine($strFileHandle, $strMsg)

	FileClose($strFileHandle)

EndFunc   ;==>WriteBIdleLog

;***************************************************************************
;** Function: 		InstallSP2XDHotfix()
;** Parameters:
;**		None
;** Description:
;**		This function is called to check registry for XD hotfix installation.
;** Return:
;**		True/False
;** Usage:
;**		$blnInstall = InstallSP2XDHotfix()
;**
;***************************************************************************
Func InstallSP2XDHotfix()
	$strKey = "HKLM\SOFTWARE\Microsoft\Updates\Windows XP\SP3\KB889673"
	$strSubKey = "Description"
	$var = RegRead($strKey, $strSubKey)
	If $var Then
		Return True
	Else
		Return False
	EndIf

EndFunc   ;==>InstallSP2XDHotfix

;***************************************************************************
;** Function: 		InstallVistaKB938194()
;** Parameters:
;**		None
;** Description:
;**		This function is called to check registry for KB938192 hotfix installation.
;** Return:
;**		True/False
;** Usage:
;**		$blnInstall = InstallVistaKB938194()
;**
;***************************************************************************
Func InstallVistaKB938194()
	If IsWinVista64() Then
		$strKey = "HKLM64\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages\Package_for_KB938194~31bf3856ad364e35~amd64~~6.0.1.2"
	Else
		$strKey = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages\Package_for_KB938194~31bf3856ad364e35~x86~~6.0.1.2"
	EndIf
	$strSubKey = "CurrentState"
	$var = RegRead($strKey, $strSubKey)
	If $var == 7 Then
		Return True
	Else
		Return False
	EndIf
EndFunc   ;==>InstallVistaKB938194

;***************************************************************************
;** Function: 		InstallVistaKB938979()
;** Parameters:
;**		None
;** Description:
;**		This function is called to check registry for KB938979 hotfix installation.
;** Return:
;**		True/False
;** Usage:
;**		$blnInstall = InstallVistaKB938194()
;**
;***************************************************************************
Func InstallVistaKB938979()

	If IsWinVista64() Then
		$strKey = "HKLM64\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages\Package_for_KB938979~31bf3856ad364e35~amd64~~6.0.1.2"
	Else
		$strKey = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages\Package_for_KB938979~31bf3856ad364e35~x86~~6.0.1.2"
	EndIf

	$strSubKey = "CurrentState"
	$var = RegRead($strKey, $strSubKey)
	If $var == 7 Then
		Return True
	Else
		Return False
	EndIf
EndFunc   ;==>InstallVistaKB938979

;***************************************************************************
;** Function: 		InstallVistaKB939008()
;** Parameters:
;**		None
;** Description:
;**		This function is called to check registry for KB939008 hotfix installation.
;** Return:
;**		True/False
;** Usage:
;**		$blnInstall = InstallVistaKB938194()
;**
;***************************************************************************
Func InstallVistaKB939008()

	If IsWinVista64() Then
		$strKey = "HKLM64\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages\Package_1_for_KB939008~31bf3856ad364e35~amd64~~6.0.1.0"
	Else
		$strKey = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages\Package_1_for_KB939008~31bf3856ad364e35~amd64~~6.0.1.0"
	EndIf

	$strSubKey = "CurrentState"
	$var = RegRead($strKey, $strSubKey)
	If $var == 7 Then
		Return True
	Else
		Return False
	EndIf
EndFunc   ;==>InstallVistaKB939008

;***************************************************************************
;** Function: DisableUAC()
;** Parameters:
;**	None
;** Description:
;**	  This function is called to disable UAC.
;** Return:
;**	  None.
;**
;***************************************************************************
Func DisableUAC()
	WriteLog("RunSetup() function - Disabling UAC")
	$strCmd = " /k %windir%\System32\reg.exe ADD HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System /v EnableLUA /t REG_DWORD /d 0 /f"
	Run(@ComSpec & $strCmd, "", "")

EndFunc   ;==>DisableUAC

;***************************************************************************
;** Function: EnableUAC()
;** Parameters:
;**	None
;** Description:
;**	  This function is called to enable/disable Vista UAC.
;** Return:
;**	  NONE.
;***************************************************************************
Func EnableUAC()
	WriteLog(" RunSetup() function - Enabling UAC")
	$strCmd = " /k %windir%\System32\reg.exe ADD HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System /v EnableLUA /t REG_DWORD /d 1 /f"
	Run(@ComSpec & $strCmd, "", "")

EndFunc   ;==>EnableUAC

;***************************************************************************
;** Function: IsUACEnabled()
;** Parameters:
;**	None
;** Description:
;**	  This function is called to check UAC status.
;** Return:
;**	  NONE.
;***************************************************************************
Func IsUACEnabled()
	If IsWinVista64() Then
		$strKey = "HKLM64\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
	Else
		$strKey = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
	EndIf
	$strSubKey = "EnableLUA"
	$var = RegRead($strKey, $strSubKey)

	If $var == 1 Then
		Return True
	Else
		Return False
	EndIf
EndFunc   ;==>IsUACEnabled

;***************************************************************************
;** Function: 		CheckCustomizationSupport()
;** Parameters:
;**		 None
;** Description:
;**		This function is called to check if we need do customization support.
;** Return:
;**		None
;** Usage:
;**		CheckCustomizationSupport()
;**
;***************************************************************************
Func CheckCustomizationSupport()
	; TESTING ONLY
	;CreateVManagerCustomization()

	; Do Customization support?
	If IsCustomizationSupport() Then
		; Unmap drive
		UnmapNetworkDrives()

		CustomizationSupport()

		; Remove the registry value
		DeleteVManagerCustomization()
	Else
		CheckXMLBackupFile()
	EndIf
EndFunc   ;==>CheckCustomizationSupport

;***************************************************************************
;** Function: 		MappingGUI(ByRef $strLocation, ByRef $strUserName, ByRef $strPassword)
;** Parameters:
;**		ByRef $strLocation
;**		ByRef $strUserName
;**		ByRef $strPassword
;** Description:
;**		This function is called to show GUI for user to enter drive mapping info.
;** Return:
;**		None
;***************************************************************************
Func MappingGUI(ByRef $strLocation, ByRef $strUserName, ByRef $strPassword)

	#Region ### START Koda GUI section ### Form=c:\my project\autoit3\scripts\frmcustom.kxf
	$Form1 = GUICreate("User Info for Mapping", 354, 208, -1, -1)
	$txtLocation = GUICtrlCreateInput("", 80, 40, 249, 21, BitOR($ES_AUTOHSCROLL, $WS_CLIPSIBLINGS))
	$txtUserName = GUICtrlCreateInput("", 80, 80, 249, 21)
	$strPassword = GUICtrlCreateInput("", 80, 120, 249, 21, BitOR($ES_PASSWORD, $ES_AUTOHSCROLL))
	$btnOK = GUICtrlCreateButton("OK", 176, 168, 75, 25, $BS_DEFPUSHBUTTON)
	$btnCancel = GUICtrlCreateButton("Cancel", 264, 168, 75, 25, 0)
	$Label1 = GUICtrlCreateLabel("Location", 24, 48, 45, 17)
	$Label2 = GUICtrlCreateLabel("User Name", 16, 88, 57, 17)
	$Label3 = GUICtrlCreateLabel("Password", 24, 128, 50, 17)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

	$blnReturn = True
	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				ExitLoop
			Case $btnOK
				$msg = $GUI_EVENT_CLOSE
				ExitLoop
			Case $btnCancel
				$msg = $GUI_EVENT_CLOSE
				$blnReturn = False
				ExitLoop
		EndSwitch
	WEnd

	If $blnReturn Then
		$strLocation = GUICtrlRead($txtLocation)
		$strUserName = GUICtrlRead($txtUserName)
		$strPassword = GUICtrlRead($strPassword)
	EndIf
	GUIDelete()

	Return $blnReturn
EndFunc   ;==>MappingGUI

;***************************************************************************
;** Function: 		CustomizationGUI(ByRef $intSelection)
;** Parameters:
;**		ByRef $intSelection
;** Description:
;**		This function is called to show customization GUI for customer to select.
;** Return:
;**		None
;***************************************************************************
Func CustomizationGUI(ByRef $intSelection)

	#Region ### START Koda GUI section ### Form=c:\my project\autoit3\scripts\frmcustselect.kxf
	$Form1 = GUICreate("Customization GUI", 303, 176, 193, 115)
	$Group1 = GUICtrlCreateGroup("Select your choice", 16, 16, 265, 97)
	$radBrowse = GUICtrlCreateRadio("Browse to select folder", 56, 48, 193, 17)
	$radMapping = GUICtrlCreateRadio("Mapping network drive", 56, 80, 193, 17)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	$btnOK = GUICtrlCreateButton("OK", 120, 136, 75, 25, $BS_DEFPUSHBUTTON)
	$btnCancel = GUICtrlCreateButton("Cancel", 208, 136, 75, 25, 0)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

	; Code needed for this GUI
	GUICtrlSetState($radBrowse, $GUI_CHECKED)
	$blnReturn = True


	While True
		$nMsg = GUIGetMsg()

		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				ExitLoop
			Case $btnOK
				$msg = $GUI_EVENT_CLOSE
				ExitLoop
			Case $btnCancel
				$msg = $GUI_EVENT_CLOSE
				$blnReturn = False
				ExitLoop

			Case $radBrowse
				$intSelection = 1
			Case $radMapping
				$intSelection = 2

		EndSwitch
	WEnd

	; Hide GUI
	;GUISetState(@SW_HIDE)
	GUIDelete()

	Return $blnReturn

EndFunc   ;==>CustomizationGUI

;***************************************************************************
;** Function: 		AppendXmlFile()
;** Parameters:
;**	None
;** Description:
;**		This function is called to implementation customization support.
;** Return:
;**		None
;***************************************************************************
Func AppendXmlFile($strInstall, $strInstallApp)
	Dim $aRecords

	If FileExists($strInstallApp) Then
		$retVal = _ReplaceStringInFile($strInstall, "</CRB>", "")

		; Read
		If Not _FileReadToArray($strInstallApp, $aRecords) Then
			MsgBox(4096, "Error", " Error reading log to Array error:" & @error)
			Exit
		EndIf

		; Write
		$file = FileOpen($strInstall, 1)
		For $x = 1 To $aRecords[0]
			FileWrite($file, $aRecords[$x] & @CRLF)
			;Msgbox(0,'Record:' & $x, $aRecords[$x])
		Next

		FileWrite($file, "</CRB>")
		FileClose($file)

	EndIf

	FileDelete($strInstallApp)

EndFunc   ;==>AppendXmlFile

;***************************************************************************
;** Function: 		DoCustomizationCopy()
;** Parameters:
;**	None
;** Description:
;**		This function is called to implementation customization support.
;** Return:
;**		None
;***************************************************************************
Func DoCustomizationCopy($strFolder)
	; XML/EXE files
	FileCopy($strFolder & "\*.xml", @HomeDrive & "\scripts", 1)
	FileCopy($strFolder & "\*.exe", @HomeDrive & "\scripts", 1)

	; Dir copy
	DirCopy($strFolder & "\CUST", @HomeDrive & "\CUST", 1)

	; REPLACE Dir copy
	If DirExists($strFolder & "\REPLACE") Then
		DirCopy($strFolder & "\REPLACE", @HomeDrive & "\", 1)
	EndIf

	; Append xml file
	$strInstall = @HomeDrive & "\scripts\Install.xml"
	$strInstallApp = @HomeDrive & "\scripts\InstallApp.xml"
	$strTest = @HomeDrive & "\scripts\Test.xml"
	$strTestApp = @HomeDrive & "\scripts\TestApp.xml"
	If FileExists($strInstallApp) Or FileExists($strTestApp) Then
		AppendXmlFile($strInstall, $strInstallApp)
		AppendXmlFile($strTest, $strTestApp)
	EndIf

	; Backup files
	$strInstallBackup = @HomeDrive & "\scripts\InstallBackup.xml"
	$strTestBackup = @HomeDrive & "\scripts\TestBackup.xml"
	FileCopy($strInstall, $strInstallBackup, 1)
	FileCopy($strTest, $strTestBackup, 1)
EndFunc   ;==>DoCustomizationCopy

;***************************************************************************
;** Function: 		CustomizationSupport()
;** Parameters:
;**	None
;** Description:
;**		This function is called to implementation customization support.
;** Return:
;**		None
;***************************************************************************
Func CustomizationSupport()
	$intSelection = 1
	If CustomizationGUI($intSelection) Then
		Switch $intSelection
			Case 1
				$strPreloadServer = StringUpper(GetPreLoadServerName())
				If (($strPreloadServer == "CHAKOTAY") Or ($strPreloadServer == "LESTAT1")) Then
					$strID = "MTN\script"
					$strPassword = "in84tel2"

					$blnResult = DriveMapAdd("*", "\\Chakotay\SoftVal", 0, $strID, $strPassword)
					$blnResult = DriveMapAdd("*", "\\Chakotay\initVal", 0, $strID, $strPassword)
				EndIf

			Case 2
				$strLocation = ""
				$strUserName = ""
				$strPassword = ""
				While True
					If MappingGUI($strLocation, $strUserName, $strPassword) Then
						$blnResult = DriveMapAdd("*", $strLocation, 0, $strUserName, $strPassword)
						If Not $blnResult Then
							MsgBox(0, "Drive Mapping Failed !!!", "Please try again.")
						Else
							ExitLoop
						EndIf
					Else
						ExitLoop
					EndIf
				WEnd
		EndSwitch

		$strFolder = FileSelectFolder("Choose a folder.", "", 2)

		; Do File Copy
		If $strFolder Then
			DoCustomizationCopy($strFolder)
		EndIf
	Else
		CheckXMLBackupFile()
	EndIf

EndFunc   ;==>CustomizationSupport

;***************************************************************************
;** Function: 		CheckXMLBackupFile()
;** Parameters:
;**		None
;** Description:
;**		This function is called to check if the xml backup files exist or not.
;** Return:
;**		None
;** Usage:
;**		CheckXMLBackupFile()
;**
;***************************************************************************
Func CheckXMLBackupFile()
	; Use backfile if exists
	$strInstall = @HomeDrive & "\scripts\Install.xml"
	$strTest = @HomeDrive & "\scripts\Test.xml"
	$strInstallBackup = @HomeDrive & "\scripts\InstallBackup.xml"
	$strTestBackup = @HomeDrive & "\scripts\TestBackup.xml"
	If (FileExists($strInstallBackup) Or FileExists($strTestBackup)) Then
		FileCopy($strInstallBackup, $strInstall, 1)
		FileCopy($strTestBackup, $strTest, 1)
	EndIf
EndFunc   ;==>CheckXMLBackupFile

;***************************************************************************
;** Function: 		InstallXPKB940566()
;** Parameters:
;**		None
;** Description:
;**		This function is called to check registry for XD hotfix installation.
;** Return:
;**		True/False
;** Usage:
;**		$blnInstall = InstallSP2XDHotfix()
;**
;***************************************************************************
Func InstallXPKB940566()
	$strKey = "HKLM\SOFTWARE\Microsoft\Updates\Windows XP\SP3\KB940566-v2"
	$strSubKey = "Description"
	$var = RegRead($strKey, $strSubKey)
	If $var == "Hotfix for Windows XP (KB940566-v2)" Then
		Return True
	Else
		Return False
	EndIf

EndFunc   ;==>InstallXPKB940566

;***************************************************************************
;** Function: 		IsInstallDotnetFramework20Installed()
;** Parameters:
;**		None
;** Description:
;**		This function is called to check registry for XD hotfix installation.
;** Return:
;**		True/False
;** Usage:
;**		$blnInstall = IsInstallDotnetFramework20Installed()
;**
;***************************************************************************
Func IsInstallDotnetFramework20Installed()
	$strKey = "HKLM\SOFTWARE\Microsoft\.NETFramework\v2.0.50727"
	$strSubKey = "Description"
	$var = RegRead($strKey, $strSubKey)
	If $var == "Hotfix for Windows XP (KB940566-v2)" Then
		Return True
	Else
		Return False
	EndIf

EndFunc   ;==>IsInstallDotnetFramework20Installed

;***************************************************************************
;** Function: 		IsWMPlayer10Installed()
;** Parameters:
;**	None
;** Description:
;**		This function is called to check if Windows Media Player 10 installation.
;** Return:
;**		True or False
;***************************************************************************
Func IsWMPlayer10Installed()
	; Get system drive
	If IsWinXP64Family() Then
		$strKey = "HKLM\Software\Intel\VMgr"
		$strSubKey = "Windows MediaPlayer 10"
		$strResult = RegRead($strKey, $strSubKey)

		$intPos = StringInStr($strResult, "INSTALLED")
		If $intPos Then
			Return True
		Else
			Return False
		EndIf

	Else
		$strFileName = @ProgramFilesDir & "\Windows Media Player\wmplayer.exe"
		$strFileTS = FileGetTime($strFileName)
		$strTemp = $strFileTS[1] & "/" & $strFileTS[2] & "/" & $strFileTS[0]
		$intPosition = StringInStr($strTemp, "9/22/2004")

		If $intPosition Then
			Return True
		Else
			Return False
		EndIf
	EndIf

EndFunc   ;==>IsWMPlayer10Installed

;***************************************************************************
;** Function: 		_VersionNumber($strDestDir)
;** Parameters:
;**	None
;** Description:
;**		This function is called to write to log: software version number with time and date when particular script is executed
;** Return:
;**		none
;***************************************************************************

Func _VersionNumber($strDestDir)
	; Shows the filenames of all files in the current directory.
	$search = FileFindFirstFile($strDestDir & "\" & "*.vmgr")

	; tells us application name based on install folder if *.vmgr file is not present
	$app = StringTrimLeft($strDestDir, 6)



	; Check if the search was successful
	If $search = -1 Then

		; since we have different folders for install files we need to determine which part to display if there is an error
		If StringInStr($strDestDir, "\APPS\") Then
			$app = StringTrimLeft($strDestDir, 6)

		ElseIf StringInStr($strDestDir, "\Drivers\") Then
			$app = StringTrimLeft($strDestDir, 9)

		ElseIf StringInStr($strDestDir, "\Software_Updates\") Then
			$app = StringTrimLeft($strDestDir, 18)

		ElseIf StringInStr($strDestDir, "\Drivers\Video\") Then
			$app1 = StringTrimLeft($strDestDir, 9)
			$app = StringTrimRight($strDestDir, 6) & " Driver" ; just to get Video to work 

		EndIf

;;		_WriteLogVersion("")
;;		_WriteLogVersion("!!!.........Version Number NOT FOUND for " & $app & " .........!!!")
;;		_WriteLogVersion("")
		_WriteLogVersion("INSTALL TIME " & _Now())
		_WriteLogVersion("!...Version Number NOT FOUND for " & $app & " ...!")
		_WriteLogVersion("")
		_WriteLogVersion("")
;;		_WriteLogVersion("!!!.........Version Number NOT FOUND for " & $app & " .........!!!")
;;		_WriteLogVersion("")


		;EndIf

	ElseIf $search = 1 Then

		$file = FileFindNextFile($search)
		;$keptFile = $file ;puts file name into $keptFile variable

		; Close the search handle
		FileClose($search)

		$version = StringTrimRight($file, 5)

	;;	_WriteLogVersion("")
	;;	_WriteLogVersion("-=........................................=-")
	;;	_WriteLogVersion("-=......VERSION NUMBER FIELD START........=-")
	;;	_WriteLogVersion("-=........................................=-")
	;;	_WriteLogVersion("")
		_WriteLogVersion("NAME & VERSION: " & $version)
		_WriteLogVersion("INSTALL TIME " & _Now())
		;_WriteLogVersion("  INSTALL " & $date)
		;_WriteLogVersion("  INSTALL " & $time )
		_WriteLogVersion("")
		_WriteLogVersion("")
;;		_WriteLogVersion("-=........................................=-")
;;		_WriteLogVersion("-=......VERSION NUMBER FIELD END..........=-")
;;		_WriteLogVersion("-=........................................=-")
;;		_WriteLogVersion("")

	EndIf

	WriteLog("Version feature executed")

EndFunc   ;==>_VersionNumber
