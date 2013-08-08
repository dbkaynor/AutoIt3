#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_icon=..\ICONS\10.ico
#AutoIt3Wrapper_outfile=t:\temp\ImRealTek.exe
#AutoIt3Wrapper_Res_Description=Win7 Realtek Driver
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_Res_LegalCopyright=Copyright © 2010 Intel Corporation. All rights reserved.
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Field=Compiler|AutoIt Version: %AutoItVer%
#AutoIt3Wrapper_Res_Field=Build Date|04/06/2010
#AutoIt3Wrapper_Res_Field=Author|Doug Kaynor
#AutoIt3Wrapper_Res_Field=Updated By|Doug Kaynor
#AutoIt3Wrapper_UseX64=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#comments-start PROGRAM HEADER
	;******************************************************************************************
	;** Intel Corporation, MPG MPAD
	;** Title			:  		ImRealTek.au3
	;** Description	:
	;**		RealTek driver update for the RealTek audio Device for Win2k,XP,XP-64 and Vista
	;**
	;** Revision: 	Rev 2.0.0
	;******************************************************************************************
	;******************************************************************************************
	;** Revision History:
	;**
	;** Update for Rev 2.0.0		- Doug Kaynor 04/06/2006
	;**	- Initial release
	;**
	;******************************************************************************************
#comments-end PROGRAM HEADER

; Script/File name
Dim Const $SCRIPT_NAME = "RealTek audio driver"
Dim Const $SCRIPT_FILENAME = "ImRealTek.exe"
Dim Const $SCRIPT_VERSION = "V2.0.0"

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
;***************************************************************************
;** Function: 		RunSetup()
;** Parameters:
;**	None
;** Description:
;**	This function is called to run setup.
;** Return:
;**	None
;**************************************************************************

Func RunSetup($strINFPath)
	; Load driver
	WriteLog($SCRIPT_FILENAME & " DebugMSG: " & $strINFPath)
	$pid = RunWait(@ComSpec & " /c " & $strINFPath & " /s")
	FileCopy(@HomeDrive & "\RHDSetup.log", @HomeDrive & "\logs\.", 1)
	;test for success
	$pid = Run(@ComSpec & " /c devcon status HDAUDIO*", "", "", $STDOUT_CHILD)
	Sleep(500)
	Local $line
	While 1
		$line = StdoutRead($pid)
		If @error Then ExitLoop

		If StringInStr($line, "Driver is running") > 0 Then
			WriteLog($SCRIPT_FILENAME & " RealTek device is running")
			Return True
		EndIf
	WEnd

	Return False
EndFunc   ;==>RunSetup


;***************************************************************************
;** Main Program
;***************************************************************************

; Prepare for running
ScriptStarting($SCRIPT_NAME, $SCRIPT_FILENAME, $SCRIPT_VERSION)

If Not IsSupportedOS() Then
	$strError = "ERROR: Unsupported OS"
	ExitWithErrorMessage($SCRIPT_NAME, $strError)
	Exit
EndIf

If @OSArch = "X86" Then
	$Devcon = @HomeDrive & "\bin\devcon.exe"
ElseIf @OSArch = "X64" Then
	$Devcon = @HomeDrive & "\bin\devcon_64.exe"
Else
	$strError = "ERROR: Unsupported Operating System Architecture."
	ExitWithErrorMessage($SCRIPT_NAME, $strError)
EndIf

;check and use devcon to see if ID exists and if does, if driver already installed
$pid = Run(@ComSpec & " /c devcon status HDAUDIO*", "", "", $STDOUT_CHILD)
Sleep(500)
Local $line
While 1
	$line = StdoutRead($pid)
	If @error Then ExitLoop

	$intNoDevice = StringInStr($line, "No Matching Devices found.")
	WriteLog($SCRIPT_FILENAME & " DebugMSG: $intNoDevice = " & $intNoDevice)
	$intDriverLoaded = StringInStr($line, "Driver is running")
	WriteLog($SCRIPT_FILENAME & " DebugMSG: $intDriverLoaded = " & $intDriverLoaded)
	$intNoDriver = StringInStr($line, "Device has a problem")
	WriteLog($SCRIPT_FILENAME & " DebugMSG: $intNoDriver = " & $intNoDriver)

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

; Local available or need to get from server
$strSourceDir = "\Drivers\Azalia\Audio\Realtek\Win2k-XP-XP64"
$strDestDir = "\drivers\Realtek\Win2k-XP-XP64"
$strFile = $strDestDir & "\setup.exe"

If Not FileExists($strFile) Then
	; Copy from server.
	If Not CopyFromServer($SCRIPT_NAME, $strSourceDir, $strDestDir) Then
		$strError = "ERROR: Could not copy from network."
		ExitWithErrorMessage($SCRIPT_NAME, $strError)
		Exit
	EndIf
EndIf

; Call RunSetup
If Not RunSetup($strFile) Then
	$strError = "Driver failed to install"
	ExitWithErrorMessage($SCRIPT_NAME, $strError)
	Exit
EndIf

; Calls for *.VMGR file to get name and version name based on name of this file

_VersionNumber($strDestDir)

; Ending
ScriptEnding($SCRIPT_NAME, "INSTALLED")

Exit

;***************************************************************************
;** End Of Program
;***************************************************************************