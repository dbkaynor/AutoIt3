#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_icon=..\..\ICONS\ImDriver.ico
#AutoIt3Wrapper_outfile=ImStdIR.exe
#AutoIt3Wrapper_Res_Description=Win7 Standard Infrared Driver
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_Res_LegalCopyright=Copyright © 2010 Intel Corporation. All rights reserved.
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Field=Compiler|AutoIt Version: %AutoItVer%
#AutoIt3Wrapper_Res_Field=Build Date|03/19/2010
#AutoIt3Wrapper_Res_Field=Author|Chris Shorey
#AutoIt3Wrapper_Res_Field=Updated By|Chris Shorey
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Allow_Decompile=n

#comments-start PROGRAM HEADER
	;******************************************************************************************
	;** Intel Corporation, MPG MPAD
	;** Title			:  		ImStdIR.au3
	;** Description	:
	;**		(Standard Infrared Port) driver update for Built-in Infrared Device for Win7
	;**
	;** Revision: 	Rev 1.0.0
	;******************************************************************************************
	;******************************************************************************************
	;** Revision History:
	;**
	;**
	;**	- Initial release - Chris Shorey 03/19/2010
	;**
	;******************************************************************************************
#comments-end PROGRAM HEADER

; Script/File name
Dim Const $SCRIPT_NAME = "Win7 Standard Infrared Driver"
Dim Const $SCRIPT_FILENAME = "ImStdIR.au3"
Dim Const $SCRIPT_VERSION = "V1.0.0"

; Include files
#include <Constants.au3>
#include <incAll.au3>


;***************************************************************************
;** Function: 		RunSetup()
;** Parameters:
;**	None
;** Description:
;**	This function is called to run setup.
;** Return:
;**	None
;***************************************************************************


Func RunSetup($strINFPath)

	WriteLog("Running " & $SCRIPT_NAME & " RunSetup() function.")

	; Run setup

	; Load driver
	$pid = Run(@ComSpec & " /c devcon " & $strINFPath & " ACPI\PNP0510", "", "", $STDOUT_CHILD)
	Sleep(500)
	Local $line
	While 1
		$line = StdoutRead($pid)
		If @error Then ExitLoop
		MsgBox(0, "STDOUT read:", $line)

		$intInstalled = StringInStr($line, "Successfully")
		MsgBox(0, "Search result:", "$intInstalled = " & $intInstalled)

		If $intInstalled <> 0 Then
			ExitLoop
		EndIf
	WEnd

	If $intInstalled <> 0 Then
		MsgBox(0, "Search result:", $intInstalled & " Driver Installed")
		Return True
	Else
		MsgBox(0, "Search result:", $intInstalled & " Driver did NOT Install")
		Return False
	EndIf

	Return True
EndFunc   ;==>RunSetup


;***************************************************************************
;** Main Program
;***************************************************************************

; Prepre for running
ScriptStarting($SCRIPT_NAME, $SCRIPT_FILENAME, $SCRIPT_VERSION)

If Not IsSupportedOS() Then
	$strError = "ERROR: Unsupported OS"
	ExitWithErrorMessage($SCRIPT_NAME, $strError)
	Exit
EndIf

;check devcon if ID exists and if does, is driver already installed
$pid = Run(@ComSpec & " /c devcon status ACPI\PNP0510", "", "", $STDOUT_CHILD)
Sleep(500)
Local $line
While 1
	$line = StdoutRead($pid)
	If @error Then ExitLoop
	MsgBox(0, "STDOUT read:", $line)

	$intNoDevice = StringInStr($line, "No Matching Devices found.")
	$intDriverLoaded = StringInStr($line, "Driver is running")
	$intNoDriver = StringInStr($line, "Device has a problem")
	MsgBox(0, "Search result:", "$intNoDevice = " & $intNoDevice & " | $intDriverLoaded = " & $intDriverLoaded & " | $intNoDriver = " & $intNoDriver)

	If $intNoDevice <> 0 Then
		$strError = "ABORTED: Hardware Device not found"
		ExitWithErrorMessage($SCRIPT_NAME, $strError)
		Exit
	ElseIf $intDriverLoaded <> 0 Then
		$strError = "ABORTED: Driver is already loaded"
		ExitWithErrorMessage($SCRIPT_NAME, $strError)
		Exit
	ElseIf $intNoDriver <> 0 Then
		ExitLoop
	EndIf
WEnd

; Allow user press ESC to abort the program
AbortProgram($SCRIPT_NAME)

If @OSArch = "X86" Then
	$strDriverPath = "\Drivers\Infrared_Port\Win7_32"
ElseIf @OSArch = "X64" Then
	$strDriverPath = "\Drivers\Infrared_Port\Win7_64"
Else
	$strError = "ERROR: Unsupported Operating System Architecture."
	ExitWithErrorMessage($SCRIPT_NAME, $strError)
EndIf


; Local available or need to get from server
$strFile = @HomeDrive & $strDriverPath & "\netirsir.inf"

If Not FileExists($strFile) Then
	; Assign source/destination dir for copy.
	$strSourceDir = $strDriverPath
	$strDestDir = $strDriverPath

	; Copy from server.
	If Not CopyFromServer($SCRIPT_NAME, $strSourceDir, $strDestDir) Then
		$strError = "ERROR: Could not copy from network."
		ExitWithErrorMessage($SCRIPT_NAME, $strError)
		Exit
	EndIf
EndIf

; Call RunSetup
If Not RunSetup($strFile) Then
	$strError = "ERROR: RunSetup() function failed."
	ExitWithErrorMessage($SCRIPT_NAME, $strError)
	Exit
EndIf

; Calls for *.VMGR file to get name and version name based on name of this file
$strDestDir = $strDriverPath ; app's destination folder to make it valid in any conditions
_VersionNumber($strDestDir)


; Ending
ScriptEnding($SCRIPT_NAME, "INSTALLED")

Exit

;***************************************************************************
;** End Of Program
;***************************************************************************
