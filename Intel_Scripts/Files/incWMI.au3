#comments-start HEADER
;******************************************************************************************
;** Intel Corporation, MPG MPAD
;** Title			:  		incWIM.wbt
;** Description	: 	
;**		Registry related include file.
;**
;** Revision: 	Rev 2.0.0
;******************************************************************************************
;******************************************************************************************
;** Revision History:
;** 
;** Update for Rev 2.0.0		- Dick Lin 11/21/2006
;**		- Initial release
;**
;** Update for Rev 2.0.1		- Dick Lin 11/27/2006
;**		- Added WMIIsNXCapable() function.
;**		- Added WMIGetNXSettingsValue() function.
;**
;** Update for Rev 2.0.2		- Dick Lin 01/12/2007
;**		- Added WMIIsDCPowerSource() function.
;**
;******************************************************************************************
#comments-end HEADER

#include-once

#include <Misc.au3>
#include <inet.au3>

;***************************************************************************
;** Function: 		WMIGetOperatingSystemName()
;** Parameters:
;**	None
;** Description: 				 
;**		This function is called to get OS name.
;** Return:
;**		String of OS Name.
;***************************************************************************
Func WMIGetOperatingSystemName()
    $strOSName = "N/A"
	$strComputer = "." 
    $objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
    If @error Then Return $strOSName
    
    $colItems = $objWMIService.ExecQuery ("SELECT * FROM Win32_OperatingSystem")

    For $objItem In $colItems
		If $objItem = "" Then ExitLoop
		$strOSName = $objItem.Caption
    next
	
	Return $strOSName

EndFunc

;***************************************************************************
;** Function: 		GetServicePackVersion()
;** Parameters:
;**	None
;** Description: 				 
;**	This function is called to get Windows service pack version.
;** Return:
;**	string name of OS.
;***************************************************************************
Func WMIGetServicePackVersion()
    $strCSDVersion = "N/A"
	$strComputer = "."
    $objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
    If @error Then Return False
    
    $colItems = $objWMIService.ExecQuery ("SELECT * FROM Win32_OperatingSystem")

    For $objItem In $colItems
		If $objItem = "" Then ExitLoop		
		$strCSDVersion = $objItem.CSDVersion
    next
	
	Return $strCSDVersion

EndFunc

;***************************************************************************
;** Function: 		GetBIOSVersion()
;** Parameters:
;**	None
;** Description: 				 
;**	This function is called to get BIOS version.
;** Return:
;**	string of BIOS version.
;***************************************************************************
Func WMIGetBIOSVersion()
	
    $strBIOSVersion = "N/A"
	$strComputer = "."
    $objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
    If @error Then Return $strBIOSVersion
    
    $colItems = $objWMIService.ExecQuery ("SELECT * FROM Win32_BIOS")

    For $objItem In $colItems
		If $objItem = "" Then ExitLoop
		$strBIOSVersion = $objItem.SMBIOSBIOSVersion  
    next

	Return $strBIOSVersion
EndFunc

;***************************************************************************
;** Function: 		WMIGetCPUName()
;** Parameters:
;**	None
;** Description: 				 
;**		This function is called to get processor name.
;** Return:
;**		String of CPU name.
;***************************************************************************
Func WMIGetCPUName()
    $strCPUName = "N/A"
	$strComputer = "."
    $objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
    If @error Then Return $strCPUName
    
    $colItems = $objWMIService.ExecQuery ("SELECT * FROM Win32_Processor")

    For $objItem In $colItems
		If $objItem = "" Then ExitLoop
		$strCPUName = $objItem.Name
    next
	
	Return $strCPUName
	
EndFunc

;***************************************************************************
;** Function: 		WMIGetCPUID()
;** Parameters:
;**	None
;** Description: 				 
;**		This function is called to get CPU stepping info.
;** Return:
;**		String of CPU stepping information.
;***************************************************************************
Func WMIGetCPUID()
    $strCPUID = "N/A"
	$strComputer = "."
    $objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
    If @error Then Return $strCPUID
    
    $colItems = $objWMIService.ExecQuery ("SELECT * FROM Win32_Processor")

    For $objItem In $colItems
		If $objItem = "" Then ExitLoop
		$strCPUID = $objItem.Caption
    Next
	
	Return $strCPUID	
	
EndFunc

;***************************************************************************
;** Function: 		WMIGetChipsetName()
;** Parameters:
;**	None
;** Description: 				 
;**	This function is called to get chipset info.
;** Return:
;**	string of BIOS version.
;***************************************************************************
Func WMIGetChipsetName()
    $strChipset = "N/A"
	$strComputer = "."
    $objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
    If @error Then Return $strChipset 
    
    $colItems = $objWMIService.ExecQuery ("SELECT * FROM Win32_ComputerSystem")

    For $objItem In $colItems
		If $objItem = "" Then ExitLoop
		$strChipset = $objItem.Model
    Next
	
	Return $strChipset	
	
EndFunc

;***************************************************************************
;** Function: 		IsNXCapable()
;** Parameters:
;**	None
;** Description: 				 
;**	This function is called to check if system is NX(no-execute or 
;**	Data Execution Prevention) capable.
;** Return:
;**	@TRUE or @FALSE
;***************************************************************************
Func WMIIsNXCapable()
	; Not for 2K.
	If IsWin2KFamily() Then
		Return False
	EndIf
	
    $blnNX = False
	$strComputer = "."
    $objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
    If @error Then Return $blnNX 
		    
    $colItems = $objWMIService.ExecQuery ("SELECT * FROM Win32_OperatingSystem")

    For $objItem In $colItems
		If $objItem = "" Then ExitLoop
		$blnNX = $objItem.DataExecutionPrevention_Available
    Next
	
	If $blnNX == -1 Then
		Return True
	Else
		Return False
	endif
EndFunc

;***************************************************************************
;** Function: 		GetNXSettingsValue()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to get NX setting value.
;** Return:	
;**		DEP Data Execution Prevention Support Policy
;**		0 - Always Off
;**			DEP is turned off for all 32-bit applications on the computer with no exceptions. 
;**			This setting is not available for the user interface. 
;**
;**		1 - Always On
;**			DEP is enabled for all 32-bit applications on the computer. 
;**			This setting is not available for the user interface. 
;**
;**		2 - Opt In
;**			DEP is enabled for a limited number of binaries, the kernel, and all Windows services. 
;**			However, it is off by default for all 32-bit applications. 
;**			A user or administrator must explicitly choose either the AlwaysOn 
;**			or the OptOut setting before DEP can be applied to 32-bit applications.
;**
;**		3 - Opt Out
;**			DEP is enabled by default for all 32-bit applications. 
;**			A user or administrator can explicitly remove support for a 32-bit application by adding the application 
;**			to an exceptions list. 
;**
;***************************************************************************
Func WMIGetNXSettingsValue()
	$intNXSetting = 0
	$strComputer = "."
    $objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
	If @error Then Return $intNXSetting 
    
    $colItems = $objWMIService.ExecQuery ("SELECT * FROM Win32_OperatingSystem")

    For $objItem In $colItems
		If $objItem = "" Then ExitLoop
		$intNXSetting = $objItem.DataExecutionPrevention_SupportPolicy
    Next
	
	Return $intNXSetting	
EndFunc	

;***************************************************************************
;** Function: 		WMIIsDCPowerSource()
;** Parameters:
;**	None
;** Description: 				 
;**		This function is called to get the power source.
;**
;**		BatteryStatus 
;**		Data type: uint16
;**		Access type: Read-only
;**
;**		Status of the battery. The value 10 (Undefined) is not valid in the CIM schema because in DMI it represents 
;**		that no battery is installed. In this case, the object should not be instantiated. 
;**		This property is inherited from CIM_Battery.
;**
;**		Value Meaning 
;**		1 The battery is discharging 
;**		2 The system has access to AC so no battery is being discharged. However, the battery is not necessarily charging. 
;**		3 Fully Charged 
;**		4 Low 
;**		5 Critical 
;**		6 Charging 
;**		7 Charging and High 
;**		8 Charging and Low 
;**		9 Charging and Critical 
;**		10 Undefined 
;**		11 Partially Charged 
;**
;** Return:
;**		@TRUE for DC power source, @FALSE for AC power source.
;** Usage:
;**		blnACPower = WMIIsDCPowerSource()
;**
;***************************************************************************
Func WMIIsDCPowerSource()
	$intBatteryStatus = 0
	$strComputer = "."
    $objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
	If @error Then Return $intBatteryStatus 
    
    $colItems = $objWMIService.ExecQuery ("SELECT * FROM Win32_Battery")

    For $objItem In $colItems
		If $objItem = "" Then ExitLoop
		$intBatteryStatus = $objItem.BatteryStatus
    Next
	
	If (($intBatteryStatus == 2) Or ($intBatteryStatus == 10)) Then
		Return False
	Else
		Return True
	EndIf
EndFunc

;***************************************************************************
;** Function: 		WMIGetOperatingSystemVersion()
;** Parameters:
;**	None
;** Description: 				 
;**		This function is called to get OS version.
;** Return:
;**		String of OS Name.
;***************************************************************************
Func WMIGetOperatingSystemVersion()
    $strOSVersion = "N/A"
	$strComputer = "." 
    $objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
    If @error Then Return $strOSVersion
    
    $colItems = $objWMIService.ExecQuery ("SELECT * FROM Win32_OperatingSystem")

    For $objItem In $colItems
		If $objItem = "" Then ExitLoop
		$strOSVersion = $objItem.Version
    Next
	
	Return $strOSVersion

EndFunc

;***************************************************************************
;** Function: 		WMIGetOperatingSystemBuildNumber()
;** Parameters:
;**	None
;** Description: 				 
;**		This function is called to get OS version.
;** Return:
;**		String of OS Name.
;***************************************************************************
Func WMIGetOperatingSystemBuildNumber()
    $strOSBuildNumber = "N/A"
	$strComputer = "." 
    $objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
    If @error Then Return $strOSBuildNumber
    
    $colItems = $objWMIService.ExecQuery ("SELECT * FROM Win32_OperatingSystem")

    For $objItem In $colItems
		If $objItem = "" Then ExitLoop
		$strOSBuildNumber = $objItem.BuildNumber
    Next
	
	Return $strOSBuildNumber

EndFunc

;***************************************************************************
;** Function: 		WMIGetOperatingSystemInstallDate()
;** Parameters:
;**	None
;** Description: 				 
;**		This function is called to get OS install date.
;** Return:
;**		String of OS Name.
;***************************************************************************
Func WMIGetOperatingSystemInstallDate()
    $strOSBInstallDate = "N/A"
	$strComputer = "." 
    $objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
    If @error Then Return $strOSBInstallDate
    
    $colItems = $objWMIService.ExecQuery ("SELECT * FROM Win32_OperatingSystem")

    For $objItem In $colItems
		If $objItem = "" Then ExitLoop
		$strOSBInstallDate = $objItem.InstallDate
    Next
	
	Return $strOSBInstallDate

EndFunc

;***************************************************************************
;** Function: 		WMIGetUserName()
;** Parameters:
;**	None
;** Description: 				 
;**		This function is called to get OS install date.
;** Return:
;**		String of OS Name.
;***************************************************************************
Func WMIGetUserName()
;~     $strUserName = "N/A"
;~ 	$strComputer = "." 
;~     $objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
;~     If @error Then Return $strUserName
;~     
;~     $colItems = $objWMIService.ExecQuery ("SELECT * FROM Win32_ComputerSystem")

;~     For $objItem In $colItems
;~ 		If $objItem = "" Then ExitLoop
;~ 		$strUserName = $objItem.UserName
;~     Next
	
	Return EnvGet("USERNAME")

EndFunc

;***************************************************************************
;** Function: 		WMIGetNumberOfProcessor()
;** Parameters:
;**	None
;** Description: 				 
;**		This function is called to get the number of processor.
;** Return:
;**		integer number of processor.
;***************************************************************************
Func WMIGetNumberOfProcessor()
    $strCSNoOfProcessors = "N/A"
	$strComputer = "." 
    $objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
    If @error Then Return $strCSNoOfProcessors
    
    $colItems = $objWMIService.ExecQuery ("SELECT * FROM Win32_ComputerSystem")

    For $objItem In $colItems
		If $objItem = "" Then ExitLoop
		$strCSNoOfProcessors = $objItem.NumberOfProcessors
    Next
	
	Return $strCSNoOfProcessors

EndFunc

;***************************************************************************
;** Function: 		WMIGetTotalPhysicalMemory()
;** Parameters:
;**	None
;** Description: 				 
;**	This function is called to get total physical memory in MB.
;** Return:
;**	Total physical memory.
;***************************************************************************
Func WMIGetTotalPhysicalMemory()
    $intTotalPhysicalMemory = "N/A"
	$strComputer = "." 
    $objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
    If @error Then Return $intTotalPhysicalMemory
    
    $colItems = $objWMIService.ExecQuery ("SELECT * FROM Win32_ComputerSystem")

    For $objItem In $colItems
		If $objItem = "" Then ExitLoop
		$intTotalPhysicalMemory = $objItem.TotalPhysicalMemory
    Next
	
	$intTotalPhysicalMemory = $intTotalPhysicalMemory/1000000
	
	Return $intTotalPhysicalMemory

EndFunc

;***************************************************************************
;** Function: 		WMIGetLogicalDiskName()
;** Parameters:
;**	None
;** Description: 				 
;**	WMI call to get the logical disk drive of the system
;** Return:
;**		String of logical drive delimeter with ",".
;***************************************************************************
Func WMIGetLogicalDiskName(ByRef $strLogicalDisk, ByRef $strDiskName)
	$strLogicalDisk = ""
	$strDiskName = ""
	
    $strName = "N/A"
	$strVolumeName = "N/A"
	$strComputer = "." 
    $objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
    If @error Then Return $strName
    
    $colItems = $objWMIService.ExecQuery ("SELECT * FROM Win32_LogicalDisk")

    For $objItem In $colItems
		If $objItem = "" Then ExitLoop
		$strName = $objItem.Name
		$strVolumeName = $objItem.VolumeName
		;$strLogicalDisk = $strLogicalDisk & $strName & $strVolumeName & "," 		
		$strLogicalDisk = $strLogicalDisk & $strName & "," 		
		$strDiskName = $strDiskName & $strVolumeName & "," 
	Next
	

	$strLogicalDisk = StringTrimRight ( $strLogicalDisk, 1 )
	$strDiskName = StringTrimRight ( $strDiskName, 1 )
	
EndFunc

;***************************************************************************
;** Function: 		CheckFloppyDriveStatus()
;** Parameters:
;**	None
;** Description: 				 
;**	This function is called to make sure floppy is in A: drive and not write protected.
;** Return:
;**	None
;***************************************************************************
Func CheckFloppyDriveStatus()

	; Make sure A: Drive has Disk in?
	$var = DriveStatus( "A:" )
	While $var <> "READY"
		MsgBox(0, "ERROR", "Drive A: is not ready.")
		$var = DriveStatus( "A:" )
	WEnd
	
	; Make sure a:\drive is read/write	
	$handle = Fileopen("A:\abcxyz.zyx", 1) ;Try and create a file
	While $handle == -1
		MsgBox(0, "ERROR", "A:\Drive is Write Protected!")
		$handle= Fileopen("A:\abcxyz.zyx", 1) ;Try and create a file	
	WEnd

	Fileclose($handle)
	Filedelete("A:\abcxyz.zyx") ;Delete the file
	
EndFunc