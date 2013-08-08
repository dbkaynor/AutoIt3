#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_icon=..\ICONS\Glonass_logo.ico
#AutoIt3Wrapper_outfile=t:\temp\ImStdIR.exe
#AutoIt3Wrapper_Res_Description=Win7 Standard Infrared Driver
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_Res_LegalCopyright=Copyright © 2010 Intel Corporation. All rights reserved.
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Field=Compiler|AutoIt Version: %AutoItVer%
#AutoIt3Wrapper_Res_Field=Build Date|04/01/2010
#AutoIt3Wrapper_Res_Field=Author|Chris Shorey
#AutoIt3Wrapper_Res_Field=Updated By|Doug Kaynor
#AutoIt3Wrapper_UseX64=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

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
	;**	- Initial release - Chris Shorey 04/01/2010
	;**
	;******************************************************************************************
#comments-end PROGRAM HEADER

; Script/File name
Global Const $SCRIPT_NAME = "Win7 Standard Infrared Driver"
Global Const $SCRIPT_FILENAME = "ImStdIR.au3"
Global Const $SCRIPT_VERSION = "V1.0.0"
Global $Devcon

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
;**************************************************************************

Func RunSetup($strINFPath)
	WriteLog(@ScriptLineNumber & " RunSetup " & $SCRIPT_FILENAME & "~" & $Devcon & "~" & $strINFPath)
	If FileExists($Devcon) Then writelog(@ScriptLineNumber & " " & $Devcon & " ok")
	If FileExists($strINFPath) Then writelog(@ScriptLineNumber & " " & $strINFPath & " ok")

	; Load driver
	If StringInStr(@OSArch, "X86") > 0 Then
		RunWait(@ComSpec & " /c " & $Devcon & " update " & $strINFPath & " *PNP0510", "", "", $STDOUT_CHILD)
	ElseIf StringInStr(@OSArch, "X64") > 0 Then
		Run(@ComSpec & " /c " & $Devcon & " update " & $strINFPath & " *PNP0510", "", "", $STDOUT_CHILD)
		WinWaitActive("Windows Security")
		Send("{TAB}{ENTER}")
	Else
		MsgBox(16, "Unexpected OSArch", @OSArch)
		WriteLog(@ScriptLineNumber & "  Unexpected OSArch " & @OSArch)
	EndIf


	Local $count = 0
	While 1
		Sleep(1000)
		$count += 1
		If $count > 50 Then Return False
		$pid = Run(@ComSpec & " /c " & $Devcon & " status ACPI*PNP0510", "", "", $STDOUT_CHILD)

		Local $line
		While 1
			$line = StdoutRead($pid)
			If @error Then ExitLoop
			If StringLen($line) > 4 Then
				;WriteLog(@ScriptLineNumber & "  " & $SCRIPT_FILENAME & "  " & $line)
				If StringInStr($line, "Driver is running") > 0 Then
					WriteLog(@ScriptLineNumber & "  " & $SCRIPT_FILENAME & " Win7 Standard Infrared Driver")
					Return True
				EndIf
			EndIf
		WEnd
	WEnd
	Return False
EndFunc   ;==>RunSetup

;***************************************************************************
;** Main Program
;***************************************************************************

; Prepare for running
ScriptStarting($SCRIPT_NAME, $SCRIPT_FILENAME, $SCRIPT_VERSION)
Opt("WinTitleMatchMode", 2)

If Not IsSupportedOS() Then
	$strError = "ERROR: Unsupported OS"
	ExitWithErrorMessage($SCRIPT_NAME, $strError)
	Exit
EndIf


;"X86", "IA64", "X64"
If StringInStr(@OSArch, "X86") > 0 Then
	$strDriverPath = "\Drivers\Infrared_Port\Win7_32"
	$Devcon = @HomeDrive & "\bin\devcon.exe"
ElseIf StringInStr(@OSArch, "X64") > 0 Then
	$strDriverPath = "\Drivers\Infrared_Port\Win7_64"
	$Devcon = @HomeDrive & "\bin\devcon_x64.exe"
Else
	$strError = "ERROR: Unsupported Operating System Architecture."
	ExitWithErrorMessage($SCRIPT_NAME, $strError)
EndIf
If Not FileExists($Devcon) Then
	MsgBox(16, "File not found error", $Devcon)
	ExitWithErrorMessage($SCRIPT_NAME, "File not found error:  " & $Devcon)
EndIf

;check and use devcon to see if ID exists and if does, if driver already installed
$pid = Run(@ComSpec & " /c " & $Devcon & " status ACPI*PNP0510", "", "", $STDOUT_CHILD)
Sleep(500)
Local $line
While 1
	$line = StdoutRead($pid)
	If @error Then
		$strError = " No expected results from devcon recieved"
		WriteLog(@ScriptLineNumber & "  " & $SCRIPT_FILENAME & $strError)
		ExitWithErrorMessage($SCRIPT_NAME, $strError)
		Exit
	EndIf
	;WriteLog(@ScriptLineNumber & "  " & $SCRIPT_FILENAME & "  " & $line)
	If StringInStr($line, "No matching devices found.") <> 0 Then
		$strError = " ABORTED: Standard Infrared hardware Device not found"
		WriteLog(@ScriptLineNumber & "  " & $SCRIPT_FILENAME & $strError)
		ExitWithErrorMessage($SCRIPT_NAME, $strError)
		Exit
	EndIf

	If StringInStr($line, "Driver is running") <> 0 Then
		$strError = " ABORTED: Win7 Standard Infrared Driver is already loaded"
		WriteLog(@ScriptLineNumber & "  " & $SCRIPT_FILENAME & $strError)
		ExitWithErrorMessage($SCRIPT_NAME, $strError)
		Exit
	EndIf

	If StringInStr($line, "Device has a problem") <> 0 Then
		WriteLog(@ScriptLineNumber & "  " & $SCRIPT_FILENAME & "Driver has a problem. Install proceding.")
		ExitLoop
	EndIf
WEnd

; Allow user press ESC to abort the program
AbortProgram($SCRIPT_NAME)

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
	$strError = " Standard Infrared driver did not install correctly"
	WriteLog(@ScriptLineNumber & "  " & $SCRIPT_FILENAME & $strError)
	ExitWithErrorMessage($SCRIPT_NAME, $strError)
	Exit
EndIf

; Calls for *.VMGR file to get name and version name based on name of this file
$strDestDir = $strDriverPath ; app's destination folder to make it valid in any conditions
_VersionNumber($strDestDir)

; Ending
ScriptEnding($SCRIPT_NAME, "INSTALLED")
WriteLog(@ScriptLineNumber & "  " & $SCRIPT_FILENAME & " Success! Win7 Standard Infrared Driver installed")
Exit

;***************************************************************************
;** End Of Program
;***************************************************************************