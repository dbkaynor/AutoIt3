#comments-start HEADER
;******************************************************************************************
;** Intel Corporation, MPG MPAD
;** Title			:  		incReg.wbt
;** Description	: 	
;**		Registry related include file.
;**
;** Revision: 	Rev 2.0.10
;******************************************************************************************
;******************************************************************************************
;** Revision History:
;** 
;** Update for Rev 2.0.0		- Dick Lin 03/02/2006
;**		- Initial release
;**
;** Update for Rev 2.0.1		- Dick Lin 12/07/2006
;**		- Added IsISTInstalled() function.
;**
;** Update for Rev 2.0.2		- Dick Lin 01/31/2007
;**		- Added GetCleanupRebootFlag() function.
;**
;** Update for Rev 2.0.2		- Dick Lin 05/15/2007
;**		- Added SetBusyIdleSuspendState() function.
;** 	- Added GetBusyIdleSuspendState() function.
;**
;** Update for Rev 2.0.3		- Dick Lin 06/05/2007
;**		- GetServerNameSubnet() support Portable Preload Server
;**
;** Update for Rev 2.0.3		- Dick Lin 06/08/2007
;**		- Added GetPlatformFamilyRSDT() function.
;**		- Added GetPlatformRSDT() function.
;**
;** Update for Rev 2.0.4		- Michelle Tran 07/10/2007
;**		- Added IsCamarilloInstalled() function.
;**
;** Update for Rev 2.0.5		- Michelle Tran 07/11/2007
;**	- Updated IsCamarilloInstalled() function to read a registry value to the 64-bit Windows Vista
;**
;** Update for Rev 2.0.6		- Dick Lin 08/22/2007
;**		- Added GetRebootFlag() function.
;**
;** Update for Rev 2.0.7		- Dick Lin 09/10/2007
;**		- Added DeleteTestRegistry() function.
;**
;** Update for Rev 2.0.8		- Dick Lin 09/17/2007
;**		- Fixed InstallInfAlready() registry incorrect issue.
;**		- Fixed IsInstallUSBYBHotfixAlready() registry incorrect issue.
;** 
;** Update for Rev 2.0.9		- Dick Lin 10/17/2007
;**		- Added IsCustomizationSupport() function.
;**		- Added DeleteVManagerCustomization() function.
;**		- Added DeleteVManagerCustomization() function.
;**
;**	Update for Rev 2.1.0		- Jarek Szymanski 10/14/2008
;**		- Added Calpella / (IbexPeak) platform to GetRSDT function
;**		- Added 2nd 10.10.248 address for Akasha in GetServerNameSubnet()
;**
;** Update for Rev 2.1.0		- Andre Nadeau 1/9/2009
;**		- Added Support IBEXPEAK
;**
;******************************************************************************************
#comments-end HEADER

#include-once

; AutoIt3 inlcude files
#include <Misc.au3>
#include <inet.au3>

;***************************************************************************
;** Function: 		GetPreLoadServerName()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to get preload server name from registry.
;** Return:
;**		the Preload server name read from registry.
;** Usage:
;**		$strPreloadServer = GetPreLoadServerName()
;**	
;***************************************************************************
Func GetPreLoadServerName()
	$strKey = "HKLM\SOFTWARE\Intel\MPG"
	$strSubKey = "PreloadServer"
	$preloadServer = RegRead($strKey, $strSubKey)
		
 	; TO DO: get from HKEY_LOCAL_MACHINE\HARDWARE\ACPI\RSDT\PTLTD_\Ohlone00
 	If $preloadServer == "" Then
 		; Get from RSDT
		$preloadServer = GetServerNameSubnet()
 	EndIf

	Return $preloadServer
EndFunc

;***************************************************************************
;** Function: 		GetPreLoadType()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to get preload server type from registry.
;** Return:
;**		the Preload type read from registry.
;** Usage:
;**		$strPreloadType = GetPreLoadType()
;**	
;***************************************************************************
Func GetPreLoadType()
	$strKey = "HKLM\SOFTWARE\Intel\MPG"
	$strSubKey = "PreloadType"
	$var = RegRead($strKey, $strSubKey)
		
	Return $var
EndFunc

;***************************************************************************
;** Function: 		GetPlatformName()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to get platform name from registry.
;** Return:
;**		Return the value read from registry.
;** Usage:
;**		$strPreloadServer = GetPlatformName()
;**	
;***************************************************************************
Func GetPlatformName()
	$strKey = "HKLM\SOFTWARE\Intel\MPG"
	$strSubKey = "PlatformName"
	$strPlatform = RegRead($strKey, $strSubKey)
		
	; Get from HKEY_LOCAL_MACHINE\HARDWARE\ACPI\RSDT\PTLTD_\Ohlone00
	If $strPlatform = "" Then
		; Get from RSDT
		$strPlatform = GetPlatformRSDT()
	EndIf

	Return $strPlatform
EndFunc

;***************************************************************************
;** Function: 		GetPlatformFamilyName()
;** Parameters:
;**	None
;** Description: 				 
;**		This function is called to get platfomr family name from registry.
;** Return:
;**		Return the value read from registry.
;** Usage:
;**		$strPreloadServer = GetPlatformFamilyName()
;**	
;***************************************************************************
Func GetPlatformFamilyName()
	$strKey = "HKLM\SOFTWARE\Intel\MPG"
	$strSubKey = "PlatformFamily"
	$strPlatformFamily = RegRead($strKey, $strSubKey)
		
	; Get from HKEY_LOCAL_MACHINE\HARDWARE\ACPI\RSDT\PTLTD_\Ohlone00
	If $strPlatformFamily = "" Then
		; Get from RSDT
		$strPlatformFamily = GetPlatformFamilyRSDT()
	EndIf

	Return $strPlatformFamily
EndFunc

;***************************************************************************
;** Function: 		GetServerNameSubnet()
;** Parameters:
;**	None
;** Description: 				 
;**		This function is called to get preload server name by determining subnet.
;** Return:
;**		Server string name got from subnet, or "" for can't find it.
;** Usage:
;**
;***************************************************************************
Func GetServerNameSubnet()

	; Get local machine IP address
	$ipAddress = @IPAddress1
	$dot = "."
	
	; Find the 3rd occurance of "."
	$location = StringInStr($ipAddress, $dot, 0, 3)	
	;MsgBox(0, "Location", $location)
	
	$ipSubNet = StringTrimRight($ipAddress, StringLen($ipAddress) - $location + 1)	
	;MsgBox(0, $ipAddress, $ipSubNet)
	; Chakotay IP address
	If (($ipSubNet == "10.10.242") Or ($ipSubNet == "10.10.243") Or ($ipSubNet == "10.10.244") _
		Or ($ipSubNet == "10.10.245") Or ($ipSubNet == "192.168.94")) Then
		Return "CHAKOTAY"
	; Akasha IP address
	ElseIf ($ipSubNet == "10.10.247") Or ($ipSubNet == "10.10.248") Then
		Return "AKASHA"
	ElseIf Ping("192.1.168.10", 250) Then
		; Portable Preload Server w/ fixed IP
		Return "P1151-NEW-PARIS" ; js 10-13-08 changed from "192.1.168.10" due to wrapper changes
	ElseIf Ping("P1151-NEW-PARIS", 250) Then
		; Portable Preload Server w/ DHCP IP
		Return "P1151-NEW-PARIS"
	Else 
		Return "DVD"
	EndIf
	
EndFunc	

;***************************************************************************
;** Function: 		UpdateRegistryStatus($strScriptName, $strStatus)
;** Parameters:
;**		$strScriptName - script calling this function
;**		$strStatus - status string to write
;** Description: 				 
;**		This function is called to write status to registry.
;** Return:
;**		None
;** Usage:
;**		UpdateRegistryStatus($strScriptName, $strMsg)	
;**
;***************************************************************************
Func UpdateRegistryStatus($strScriptName, $strStatus)
	; Set registry key indicate script has ran
	$strKey = "HKLM\Software\Intel\MPG\TESTAPP"
	$strSubKey = $strScriptName
	$var = RegWrite($strKey, $strSubKey, "REG_SZ", $strStatus)
	
	If @error Then
		WriteLog("UpdateRegistryStatus Function Failed.")
	EndIf
EndFunc	

;***************************************************************************
;** Function: 		today()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to Return the current date in mm/dd/yyyy form.
;** Return:
;**		Return the current date in mm/dd/yyyy form.
;** Usage:
;**		$strToday = today()	
;**
;***************************************************************************
Func today()  ;Return the current date in mm/dd/yyyy form
    Return (@MON & "/" & @MDAY & "/" & @YEAR)
EndFunc

;***************************************************************************
;** Function: 		InstallOnly()
;** Parameters:
;**	None
;** Description: 				 
;**	This function is called to get check if VManager is in InstallOnly mode
;** Return:
;**	@TRUE for InstallOnly, else @FALSE
;** Usage:
;**	CleanupRegistry(SCRIPT_NAME)
;** 
;** Usage:
;**  	; Install Only ?
;**	if InstallOnly() then
;**		...
;**	endif
;***************************************************************************
Func InstallOnly()

	$strKey    = "HKLM\Software\Intel\MPG\TESTAPP"
	$strSubKey = "InstallOnly"
	$strStatus = RegRead($strKey, $strSubKey)

	If $strStatus <> "" Then
		If StringUpper($strStatus) == "YES" Then
			Return True
		Else
			Return False
		EndIf
	EndIf
	
	Return False

EndFunc

;***************************************************************************
;** Function: 		SetupCleanupScript(strScript, strApp, blnReboot)
;** Parameters:
;**	strScript - test script name need cleanup
;**	strApp - Clean up routine to call
;**	blnReboot - reboot required flag after running cleanup
;** Description: 				 
;**	This function is called to prepare registry for CleanUp script to run
;** Return:
;**	None
;** Usage:
;**		SetupCleanupScript(strScript, strApp, blnReboot)
;**
;***************************************************************************
Func SetupCleanupScript($strScript, $strApp, $blnReboot)
	WriteLog($strScript & " Setting  up cleanup script.")

	$strKey	   = "HKLM\Software\Intel\MPG\TESTAPP"
	$strSubKey = "CleanUp"
	RegWrite($strKey, $strSubKey, "REG_SZ", $strScript)


	$strSubKey = "CleanUpApp"
	RegWrite($strKey, $strSubKey, "REG_SZ", $strApp)

	$strSubKey = "Reboot"
	RegWrite($strKey, $strSubKey, "REG_SZ", $blnReboot)

EndFunc

;***************************************************************************
;** Function: 		IsISTInstalled()
;** Parameters:
;**		None
;** Description:
;**		This function is called to check if IST Applet is installed.
;** Return:
;**		True or False.
;***************************************************************************
Func IsISTInstalled()
	
	$strKey = "HKLM\Software\Intel\PRPC"
	$strSubKey = "Command"
	
	$val = RegRead($strKey, $strSubKey)
	If $val Then
		Return True
	Else
		Return False
	EndIf
	

	Return False
EndFunc 

;***************************************************************************
;** Function: 		IsCamarilloInstalled()
;** Parameters:
;**		None
;** Description:
;**		This function is called to check if ETM driver is installed.
;** Return:
;**		True or False.
;***************************************************************************
Func IsCamarilloInstalled()
	
	If IsWinVista64() Then
		$strKey = "HKLM64\Software\Intel\IIF2"
	Else
		$strKey = "HKLM\Software\Intel\IIF2"
	EndIf
	
	$strSubKey = "Install"
	
	$val = RegRead($strKey, $strSubKey)
	If $val Then
		Return True
	Else
		Return False
	EndIf
	
	Return False
EndFunc 

;***************************************************************************
;** Function: 		IsISTEnabled()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to IST enable or disable status.
;** Return:
;**		1 - IST enabled, 0 - IST disabled.
;***************************************************************************
Func IsISTEnabled()
	
	$strKey = "HKLM\Software\Intel\PRPC"
	$strSubKey = "flags"
	
	$intVal = RegRead($strKey, $strSubKey)

	; Fliter (Decimal 268435456) Hex 00010000 00000000 00000000 00000000 - use this to get the first 4 bits of flags registry.
	$intFilter = 0x10000000	
		
	Return BitAND($intVal, $intFilter)
	
EndFunc

;***************************************************************************
;** Function: 		GetISTExitCode()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to get the IST exit code from registry.
;** Return:
;**		Value of IST ExitCode.
;***************************************************************************
Func GetISTExitCode()
	$strKey = "HKLM\Software\Intel\PRPC"
	$strSubKey = "ExitCode"
	
	Return RegRead($strKey, $strSubKey)
	
EndFunc

;***************************************************************************
;** Function: 		CheckISTExitCode()
;** Parameters:
;**	None
;** Description: 				 
;**		This function is called to check if IST ExitCode all right.
;** Return:
;**		String name of the processor.
;***************************************************************************
Func CheckISTExitCode()

	; Get IST exit code from registry.
	$intExitCode = GetISTExitCode()

	Select
		; Initialization was successful
		Case $intExitCode = 0
			$strMsg = "ExitCode: " & $intExitCode
			SetupISTRegistry()
						
		; 0x102 - The enable flag bit of the Flags value is 0
		Case $intExitCode = 0x102
			$strMsg = "ExitCode: " & $intExitCode & " - The enable flag bit of the Flags value is 0."
		
		; 0x106 - BIOS has disabled IST
		Case $intExitCode = 0x106
			$strMsg = "ExitCode: " & $intExitCode & " - BIOS has disabled IST."
			
		; IST ExitCode does not equal to 0x106. ExitCode is assumed to be 0x105 (which means that a CPU is not IST capable.
		; The IST applet will uninstall itself. The install manager will run the next script on the list.
		; 0x105
		Case $intExitCode = 0x105			
			; TODO - Some OK clik (DON'T KNOW HOW TO GENERATE THIS SITUATION YET.
			$strMsg = "ExitCode: " & $intExitCode & " - CPU is not IST capable."
			
		Case Else 
			$strMsg = "ExitCode: " & $intExitCode			
		
	EndSelect	
	WriteLog($strMsg)

	Return $intExitCode
	
EndFunc

;***************************************************************************
;** Function: 		SetupISTRegistry()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to setup Win2K IST registry after install IST.
;** Return:
;**		None
;** Usage:
;**		SetupISTRegistry()
;**
;***************************************************************************
Func SetupISTRegistry()

	; IST PRPC registry key.
;~ 	$strKey = "HKLM\Software\Intel\PRPC"
;~ 	RegRead($strKey)
;~ 	
;~ 	strKey = "Software\Intel\PRPC"
;~ 	intFlag = 1
;~ 	if !RegistryExistKey(@REGMachine, strKey, intFlag) then
;~ 		strMsg = "IST: Could not find IST registry key PRPC"
;~ 		WriteLog(strMsg)
;~ 		return
;~ 	endif

	; Set reserved1 value to 0
	$strKey = "HKLM\Software\Intel\PRPC"
	$strSubKey = "Reserved1"
	RegWrite($strKey, $strSubKey, "REG_DWORD", 0)
	
	; Set value to override OS transition restrictions
	$strKey = "HKLM\Software\Intel\PRPC"
	$strSubKey = "flags"
	$intFlags = RegRead($strKey, $strSubKey)
	
	Select 
		Case $intFlags =  0x10000202 		; 0x10000202
			$intNewValue = 0x12000202		; 0x12000202
			RegWrite($strKey, $strSubKey, "REG_DWORD", $intNewValue)
							
		Case $intFlags =  0x10000232 		; 0x10000232
			$intNewValue = 0x12000232		; 0x12000232
			RegWrite($strKey, $strSubKey, "REG_DWORD", $intNewValue)	
			
		; Do nothing. It's already changed.
		Case $intFlags = 0x12000202		;301990402	
		Case $intFlags = 0x12000202 		;301990450				
			
		Case Else		; Default
			$strMsg = "IST - flag not reset to override OS transition restrictions"
			WriteLog($strMsg)
				
	EndSelect

EndFunc

;***************************************************************************
;** Function: 		GetCleanupRebootFlag()
;** Parameters:
;**	None
;** Description: 				 
;**	This function is called to get the boolean flag for CleanUp routine.
;** Return:
;**	@TRUE for reboot required, otherwise @FALSE
;** Usage:
;**		blnReboot = GetCleanupRebootFlag()
;**
;***************************************************************************
Func GetCleanupRebootFlag()
	$strKey = "HKLM\SOFTWARE\Intel\MPG\TESTAPP"
	$strSubKey = "Reboot"
	$blnReboot = RegRead($strKey, $strSubKey)
	
	If $blnReboot == 1 Then
		Return True
	Else
		Return False
	EndIf
	
EndFunc

;***************************************************************************
;** Function: 		InstallInfAlready()
;** Parameters:
;**	None
;** Description:
;**	This function is called to check if INF driver already installed.
;** Return:
;**	None
;***************************************************************************
Func InstallInfAlready()
	
	If IsWinXP64Family() Or IsWinVista64() Then
		$strKey = "HKLM64\Software\Intel\InfInst"
	Else
		$strKey = "HKLM\Software\Intel\InfInst"
	EndIf
	$strSubKey = "Install"
	
	$val = RegRead($strKey, $strSubKey)
	If StringUpper($val) == "SUCCESS" Then
		Return True
	Else
		Return False
	EndIf
	

	Return False
EndFunc   ;==>InstallInfAlready

;***************************************************************************
;** Function: 		InstallSP2QFEAlready()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to check if QFE is installed or not.
;** Return:
;**		1 for true, otherwise 0.
;**
;***************************************************************************
Func InstallSP2QFEAlready()
	
	$strKey = "HKLM\SOFTWARE\Microsoft\Updates\Windows XP\SP3\KB896256"
	$strSubKey = "Description"
		
	$var = RegRead($strKey, $strSubKey)

	If $var <> "" Then
		Return True
	Else
		Return False
	EndIf
	
EndFunc

;***************************************************************************
;** Function: 		IsInstallGV3QFEAlready()
;** Parameters:
;**	None
;** Description: 				 
;**	This function is called to check if QFE is installed or not.
;** Return:
;**	None
;***************************************************************************
Func IsInstallGV3QFEAlready()
	; TO DO: NEED HKLM64 - wait for next release or beta version.
	$strKey = "HKLM64\SOFTWARE\Microsoft\Updates\Windows XP Version 2003\SP2\KB896256"
	$strSubKey = "Description"
	$val = RegRead($strKey, $strSubKey)

	WriteLog($val)
	
	If $val == "" Then
		Return False
	Else
		Return True
	EndIf

EndFunc

;***************************************************************************
;** Function: 		IsInstallUSBHotfixAlready()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to check if QFE is installed or not.
;** Return:
;**		1 for true, otherwise 0.
;**
;***************************************************************************
Func IsInstallUSBHotfixAlready()
	
	$strKey = "HKLM\SOFTWARE\Microsoft\Updates\Windows XP\SP3\KB918005"
	$strSubKey = "Description"
		
	$var = RegRead($strKey, $strSubKey)

	If $var <> "" Then
		Return True
	Else
		Return False
	EndIf
	
EndFunc

;***************************************************************************
;** Function: 		IsInstallXP64USBHotfixAlready()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to check if QFE is installed or not.
;** Return:
;**		1 for true, otherwise 0.
;**
;***************************************************************************
Func IsInstallXP64USBHotfixAlready()
	
	$strKey = "HKLM64\SOFTWARE\Microsoft\Updates\Windows XP Version 2003\SP2\KB918005"
	$strSubKey = "Description"
		
	$var = RegRead($strKey, $strSubKey)

	If $var <> "" Then
		Return True
	Else
		Return False
	EndIf
	
EndFunc

;***************************************************************************
;** Function: 		IsInstallUSBYBHotfixAlready()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to check if QFE is installed or not.
;** Return:
;**		1 for true, otherwise 0.
;**
;***************************************************************************
Func IsInstallUSBYBHotfixAlready()
	
	$strKey = "HKLM\SOFTWARE\Microsoft\Updates\Windows XP\SP3\KB921411"
	$strSubKey = "Description"
		
	$var = RegRead($strKey, $strSubKey)

	If $var <> "" Then
		Return True
	Else
		Return False
	EndIf
	
EndFunc

;***************************************************************************
;** Function: 		IsInstallHDHotfixAlready()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to check if QFE is installed or not.
;** Return:
;**		1 for true, otherwise 0.
;**
;***************************************************************************
Func IsInstallHDHotfixAlready()
	
	If (IsWinXP64Family()) Then
		$strKey = "HKLM64\SOFTWARE\Microsoft\Updates\Windows XP Version 2003\SP2\KB901105"
		$strSubKey = "InstalledBy"
	ElseIf (IsWinXPSP1()) Then
		$strKey = "HKLM\SOFTWARE\Microsoft\Updates\Windows XP Version 2003\SP3\KB888111WXP"	
		$strSubKey = "InstalledBy"
	ElseIf (IsWinXPSP2()) Then
		$strKey = "HKLM\SOFTWARE\Microsoft\Updates\Windows XP Version 2003\SP3\KB888111WXPSP2"	
		$strSubKey = "InstalledBy"
	ElseIf (IsWin2KSP4()) Then
		$strKey = "HKLM\SOFTWARE\Microsoft\Updates\KB888111"	
		$strSubKey = "Installed"
	EndIf	
			
	$var = RegRead($strKey, $strSubKey)

	If $var <> "" Then
		Return True
	Else
		Return False
	Endif
EndFunc

;***************************************************************************
;** Function: 		InstallSRIGDAlready()
;** Parameters:
;**	None
;** Description: 				 
;**	This function is called to check if IGD alaredy installed.
;** Return:
;**	None
;***************************************************************************
Func InstallSRIGDAlready()

	If IsWinXP64Family() Or IsWinVista64() Then
		$strKey = "HKLM64\Software\Intel\IGDI"
	Else
		$strKey = "HKLM\Software\Intel\IGDI"
	EndIf
	$strSubKey = "Install"
		
	$var = RegRead($strKey, $strSubKey)  
	$var = StringUpper($var) 
	
	If $var == "SUCCESS" Then
		Return True
	Else
		Return False
	EndIf
	
EndFunc

;***************************************************************************
;** Function: 		InstallNapaIGDAlready()
;** Parameters:
;**	None
;** Description: 				 
;**	This function is called to check if IGD alaredy installed.
;** Return:
;**	None
;***************************************************************************
Func InstallNapaIGDAlready()

	If IsWinVista64() Then
		$strKey = "HKLM64\Software\Intel\IGDI"
	Else
		$strKey = "HKLM\Software\Intel\IGDI"
	EndIf
	$strSubKey = "Install"
		
	$var = RegRead($strKey, $strSubKey)  
	$var = StringUpper($var) 
	
	If $var == "SUCCESS" Then
		Return True
	Else
		Return False
	EndIf
	
EndFunc

;***************************************************************************
;** Function: 		InstallHCTAlready()
;** Parameters:
;**	None
;** Description: 				 
;**	This function is called to check if HCT 121.1 SP1 already installed.
;** Return:
;**	None
;***************************************************************************
Func InstallHCTAlready()
	
	$strKey = "HKLM\SOFTWARE\Microsoft\HCT"
	$strSubKey = "InstallPath"
		
	$var = RegRead($strKey, $strSubKey)

	If $var <> "" Then
		Return True
	Else
		Return False
	EndIf	
														  
EndFunc

;***************************************************************************
;** Function: 		IsInstallHDHotfixAlready()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to check if QFE is installed or not.
;** Return:
;**		1 for true, otherwise 0.
;**
;***************************************************************************
Func IsInstallDSTPatchAlready()
	If IsWinXP64Family() Then
		$strKey = "HKLM64\SOFTWARE\Microsoft\Updates\Windows XP Version 2003\SP3\KB931836"
	Else
		$strKey = "HKLM\SOFTWARE\Microsoft\Updates\Windows XP\SP3\KB931836"
	EndIf
	$strSubKey = "Description"
		
	$var = RegRead($strKey, $strSubKey)

	If $var <> "" Then
		Return True
	Else
		Return False
	EndIf	
EndFunc	

;***************************************************************************
;** Function: 		UpdateCycleCount()
;** Parameters:
;**	None
;** Description: 				 
;**	This function is called to get Busy/Idle cycle count.
;** Return:
;**	string value of the registry
;***************************************************************************
Func UpdateCycleCount()
	$strKey = "HKLM\Software\Intel\MPG\TESTSCRIPT\BusyIdle"
	$strSubKey = "CycleCount"

	$intCount = RegRead($strKey, $strSubKey)
	If $intCount <> "" Then
		RegWrite($strKey, $strSubKey, "REG_SZ", $intCount + 1)
	Else
		RegWrite($strKey, $strSubKey, "REG_SZ", 1)
	EndIf
	
EndFunc

;***************************************************************************
;** Function: 		GetCycleCount()
;** Parameters:
;**	None
;** Description: 				 
;**	This function is called to get Busy/Idle cycle count.
;** Return:
;**	string value of the registry
;***************************************************************************
Func GetCycleCount()
	$strKey = "HKLM\Software\Intel\MPG\TESTSCRIPT\BusyIdle"
	$strSubKey = "CycleCount"
		
	$intCount = RegRead($strKey, $strSubKey)
	
	If $intCount <> "" Then
		return $intCount
	Else
		Return 0
	EndIf

EndFunc

;***************************************************************************
;** Function: 		GetBusyIdleCurrState()
;** Parameters:
;**	None
;** Description: 				 
;**	This function is called to get current busy/idle state. 
;** Return:
;**	string value of the registry
;***************************************************************************
Func GetBusyIdleCurrState()
		
	$strKey = "HKLM\Software\Intel\MPG\TESTSCRIPT\BusyIdle"
	$strSubKey = "CurrentState"
		
	$strCurrentState = RegRead($strKey, $strSubKey)

	Return $strCurrentState

EndFunc

;***************************************************************************
;** Function: 		GetBusyIdleNewState()
;** Parameters:
;**	None
;** Description: 				 
;**	This function is called to get current busy/idle state. 
;** Return:
;**	string value of the registry
;***************************************************************************
Func GetBusyIdleNewState()
	

	$strKey = "HKLM\Software\Intel\MPG\TESTSCRIPT\BusyIdle"
	$strSubKey = "NewState"
		
	$strNewState = RegRead($strKey, $strSubKey)

	Return $strNewState

EndFunc

;***************************************************************************
;** Function: 		GetBusyIdleNewState()
;** Parameters:
;**	None
;** Description: 				 
;**	This function is called to get current busy/idle state. 
;** Return:
;**	string value of the registry
;***************************************************************************
Func SetBusyIdleState($strCurrState)
	
	$strKey = "HKLM\Software\Intel\MPG\TESTSCRIPT\BusyIdle"
	$strSubKey = "CurrentState"		
	RegWrite($strKey, $strSubKey, "REG_SZ", $strCurrState)

	$strSubKey = "NewState"
	Switch $strCurrState
		Case "IDLE"
			$strNewState = "BUSY"
		Case "BUSY"
			$strNewState = "SUSPEND"
		Case "SUSPEND"
			$strNewState = "IDLE"
		
	EndSwitch
	RegWrite($strKey, $strSubKey, "REG_SZ", $strNewState)

EndFunc

;***************************************************************************
;** Function: 		CheckBusyIdleConfig()
;** Parameters:
;**	None
;** Description: 				 
;**	This function is called to check if the registry value exists. 
;** Return:
;**	1 - "OK", 2 - "RUNNING", or 0 - "ERROR"
;***************************************************************************
Func CheckBusyIdleConfig()
	

	$strKey = "HKLM\Software\Intel\MPG\TESTSCRIPT\BusyIdle"
	$strSubKey = "CONFIG"		
	$strState = RegRead($strKey, $strSubKey)
	
	If Not $strState Then
		$strState = "GUI"
		RegWrite($strKey, $strSubKey, "REG_SZ", $strState)
	EndIf

	Return $strState

EndFunc

;***************************************************************************
;** Function: 		UpdateBusyIdleConfig()
;** Parameters:
;**	None
;** Description: 				 
;**	This function is called to check if the registry value exists. 
;** Return:
;**	1 - "OK", 2 - "RUNNING", or 0 - "ERROR"
;***************************************************************************
Func UpdateBusyIdleConfig($strState)
	

	$strKey = "HKLM\Software\Intel\MPG\TESTSCRIPT\BusyIdle"
	$strSubKey = "CONFIG"		

	RegWrite($strKey, $strSubKey, "REG_SZ", $strState)

EndFunc

;***************************************************************************
;** Function: 		GetBusyIdleSuspendState()
;** Parameters:
;**	None
;** Description: 				 
;**	This function is called to get current busy/idle state. 
;** Return:
;**	string value of the registry
;***************************************************************************
Func GetBusyIdleSuspendState()
	

	$strKey = "HKLM\Software\Intel\MPG\TESTSCRIPT\BusyIdle"
	$strSubKey = "SuspendState"
		
	$strNewState = RegRead($strKey, $strSubKey)

	Return $strNewState

EndFunc

;***************************************************************************
;** Function: 		GetBusyIdleNewState()
;** Parameters:
;**	None
;** Description: 				 
;**	This function is called to get current busy/idle state. 
;** Return:
;**	string value of the registry
;***************************************************************************
Func SetBusyIdleSuspendState($strCurrState)
	
	$strKey = "HKLM\Software\Intel\MPG\TESTSCRIPT\BusyIdle"
	$strSubKey = "SuspendState"		
	RegWrite($strKey, $strSubKey, "REG_SZ", $strCurrState)

;~ 	$strSubKey = "NewState"
;~ 	Switch $strCurrState
;~ 		Case "IDLE"
;~ 			$strNewState = "BUSY"
;~ 		Case "BUSY"
;~ 			$strNewState = "SUSPEND"
;~ 		Case "SUSPEND"
;~ 			$strNewState = "IDLE"
;~ 		
;~ 	EndSwitch
	RegWrite($strKey, $strSubKey, "REG_SZ", $strCurrState)

EndFunc

;***************************************************************************
;** Function: 		GetRSDT()
;** Parameters:
;**		ByRef $strFamily
;**		ByRef $strPlatform
;**	None
;** Description: 				 
;**		This function is called to get Platform Family Name and Platform Name from RSDT.
;** Return:
;**		ByRef $strFamily, ByRef $strPlatform
;***************************************************************************
Func GetRSDT(ByRef $strFamily, ByRef $strPlatform)
	$strBiosKey = RegEnumKey("HKEY_LOCAL_MACHINE\HARDWARE\ACPI\RSDT\", 1)
	;$strKey = RegEnumKey("HKEY_LOCAL_MACHINE\HARDWARE\ACPI\RSDT\PTLTD_\", 1)
	$strKey = RegEnumKey("HKEY_LOCAL_MACHINE\HARDWARE\ACPI\RSDT\" & $strBiosKey & "\", 1)
	$strPlatform = ""
	$strFamily = ""
	Switch StringUpper($strKey)
		; Alviso Family - Ohlone, Guadalupe, South Bay
		Case "OHLONE00"
			$strFamily = "Alviso"
			$strPlatform = "Sonoma (Alviso / ICH6-M)"
		Case "GUADALUP"
			$strFamily = "Alviso"
			$strPlatform = "Sonoma (Alviso / ICH6-M)"
		; Calistoga Family - Capell Valley, North Bay, Jamison Canyon, Dardon Canyon LV, Dardon Canyon ULV
		Case "CAPELL00"
			$strFamily = "Calistoga"
			$strPlatform = "Napa (Calistoga / ICH7-M)"
		Case "JAMISONC"
			$strFamily = "Calistoga"
			$strPlatform = "Napa (Calistoga / ICH7-M)"
		Case "DARDONCL"
			$strFamily = "Calistoga"
			$strPlatform = "Napa (Calistoga / ICH7-M)"
		Case "DARDONCU"
			$strFamily = "Calistoga"
			$strPlatform = "Napa (Calistoga / ICH7-M)"
		; Crestline Family - Fountain Grove, Matanzas, Mayacama, Oakmont
		Case "FTNGROVE"
			$strFamily = "Crestline"
			$strPlatform = "Santa Rosa (Crestline / ICH8-M)"
		Case "MATANZAS"
			$strFamily = "Crestline"
			$strPlatform = "Santa Rosa (Crestline / ICH8-M)"
		Case "OAKMONT0"
			$strFamily = "Crestline"
			$strPlatform = "Santa Rosa (Crestline / ICH8-M)"
		; Cantiga Family - Pillar Rock
		Case "PILLARRK"
			$strFamily = "Cantiga"
			$strPlatform = "Montevina (Cantiga / ICH9-M)"
		;This is the Calpella ERB which uses the Cantiga / ICH9-M silicon
		Case "CALPELAE"
			$strFamily = "Cantiga"
			$strPlatform = "Calpella (Montello / GreenPeak)"
		;This is the Calpella ERB which uses the IBEXPEAK silicon
		Case "CALPELAC"
			$strFamily = "IbexPeak"
			$strPlatform = "Calpella / (IbexPeak)"
		; GENERIC
		Case Else
			$strFamily = "GENERIC"
			$strPlatform = "GENERIC"
	EndSwitch	
EndFunc	

;***************************************************************************
;** Function: 		GetPlatformFamilyRSDT()
;** Parameters:
;**	None
;** Description: 				 
;**		This function is called to get Platform Family Name from RSDT
;** Return:
;**		String of Platform Family
;***************************************************************************
Func GetPlatformFamilyRSDT()

	$strFamily = ""
	$strPlatform = ""
	GetRSDT($strFamily, $strPlatform)
	
	Return $strFamily
EndFunc

;***************************************************************************
;** Function: 		GetPlatformRSDT()
;** Parameters:
;**	None
;** Description: 				 
;**		This function is called to get Platform Name from RSDT
;** Return:
;**		String of Platform name
;***************************************************************************
Func GetPlatformRSDT()
	$strFamily = ""
	$strPlatform = ""
	GetRSDT($strFamily, $strPlatform)
	
	Return $strPlatform
EndFunc

;***************************************************************************
;** Function: 		GetRebootFlag(strScriptName)
;** Parameters:
;**	strScriptName - script name to check
;** Description: 				 
;**	This function is called to get the script reboot flag. It's used for optional reboot.
;** Return:
;**	None
;** Usage:
;**
;**	blnRebootFlag = GetRebootFlag(SCRIPT_NAME)
;**	if blnRebootFlag then
;**		; Restart now
;**		SetWin2KAutoLogon()
;**		SetRestartVManager()	
;**		...
;**	endif
;***************************************************************************
Func GetRebootFlag($strScriptName)
	$strKey    = "HKLM\Software\Intel\MPG\TESTAPP\VMgrRebootFlag"
	$strSubKey = $strScriptName
	$var = RegRead($strKey, $strSubKey)

	;MsgBox(0, $strSubKey, $var)
	
	Return $var
EndFunc

;***************************************************************************
;** Function: 		RegistryDelete($strKey)
;** Parameters:
;**		$strKey
;** Description: 				 
;**		This function is called to delete registry key.
;** Return:
;**		None
;***************************************************************************
Func RegistryDelete($strKey)
	For $i= 1 to 50
		$var = RegEnumKey($strKey, $i)
		If @error <> 0 then ExitLoop
		;MsgBox(4096, "SubKey #" & $i & " under HKLM\Software: ", $var)
		RegDelete($strKey & "\" & $var)	
	Next
	RegDelete($strKey)	
EndFunc


;***************************************************************************
;** Function: 		DeleteTestRegistry($strHKLM)
;** Parameters:
;**		$strHKLM
;** Description: 				 
;**	This function is called to delete registry keys used for VManager/Script.
;** Return:
;**		None
;***************************************************************************
Func DeleteTestRegistry()
	;MsgBox(0, "BEFORE", "DeleteTestRegistry")
		; Deletes HKLM\Software\Intel\MPG\TESTAPP key 
		$strKey =  "HKLM\Software\Intel\MPG\TESTSCRIPT"
		RegistryDelete($strKey)	
			
		; Deletes HKLM\Software\Intel\MPG\TESTSCRIPT key 
		$strKey = "HKLM\Software\Intel\MPG\TESTAPP\BusyIdle"
		RegistryDelete($strKey)		
		
		; Deletes HKLM\Software\Intel\MPG\TESTSCRIPT key 
		$strKey = "HKLM\Software\Intel\MPG\TESTAPP\BusyIdle\DiskLoad"
		RegistryDelete($strKey)		
						
		; Deletes HKLM\Software\Intel\MPG\TESTSCRIPT key 
		$strKey = "HKLM\Software\Intel\MPG\TESTAPP"
		RegistryDelete($strKey)	
		
		; Deletes HKLM\Software\Intel\MPG\TESTSTATUS key
		$strKey = "HKLM\Software\Intel\MPG\TESTSTATUS"
		RegistryDelete($strKey)	
		
		; Deletes HKLM\Software\Intel\MPG\TmSState key
		$strKey = "HKLM\Software\Intel\MPG\TmSState"
		RegistryDelete($strKey)	
			
		; Deletes HKLM\Software\Intel\MPG\VMLauncher value		
		$strKey = "HKLM\Software\Intel\MPG"
		RegDelete($strKey, "VMLauncher")			
	;MsgBox(0, "AFTER", "DeleteTestRegistry")		
EndFunc

;***************************************************************************
;** Function: 		IsCustomizationSupport
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to check if we need do customization support.
;** Return:
;**		True/False.
;** Usage:
;**
;**		$blnCustomization = IsCustomizationSupport()
;**
;***************************************************************************
Func IsCustomizationSupport()
	$strKey    = "HKLM\Software\Intel\MPG"
	$strSubKey = "VManagerCustomization"
	$strStatus = RegRead($strKey, $strSubKey)

	If $strStatus <> "" Then
		If StringUpper($strStatus) == "YES" Then
			Return True
		Else
			Return False
		EndIf
	EndIf
	
	Return False

EndFunc

;***************************************************************************
;** Function: 		CreateVManagerCustomization
;** Parameters:
;**	None
;** Description: 				 
;**		This function is called to create registry value for Customization support.
;** Return:
;**	None
;***************************************************************************
Func CreateVManagerCustomization()
	$strKey = "HKLM\Software\Intel\MPG"
	$strSubKey = "VManagerCustomization"		

	RegWrite($strKey, $strSubKey, "REG_SZ", "YES")	
EndFunc

;***************************************************************************
;** Function: 		CreateVManagerCustomization
;** Parameters:
;**	None
;** Description: 				 
;**		This function is called to create registry value for Customization support.
;** Return:
;**	None
;***************************************************************************
Func DeleteVManagerCustomization()
	$strKey = "HKLM\Software\Intel\MPG"
	$strSubKey = "VManagerCustomization"		

	RegDelete ($strKey, $strSubKey)	
EndFunc
