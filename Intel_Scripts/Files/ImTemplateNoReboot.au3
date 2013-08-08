#comments-start PROGRAM HEADER
;******************************************************************************************
;** Intel Corporation, MPG MPAD
;** Title			:  		ImTemplateNoReboot.au3
;** Description	:
;**		DESCRIPTION OF THE SCRIPT.
;**
;** Revision: 	Rev 2.0.0
;******************************************************************************************
;******************************************************************************************
;** Revision History:
;**
;** Update for Rev 2.0.0		- YOUR NAME 03/01/2006
;**	- Initial release
;**
;******************************************************************************************
#comments-end PROGRAM HEADER

; Script/File name
Dim Const $SCRIPT_NAME 		= "DriverInstall"
Dim Const $SCRIPT_FILENAME 	= "DriverInstall.exe"
Dim Const $SCRIPT_VERSION  	= "0.0.1.1"

; Include files
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
Func RunSetup()
	; Write some msg.
	WriteLog("Running " & $SCRIPT_NAME & " RunSetup() function.")

	; Set WinTitleMatchMode to 2 - 1=start, 2=subStr, 3=exact, 4=advanced
	Opt("WinTitleMatchMode", 2)

	; Set RunErrorsFatal to 1 - 1=fatal, 0=silent set @error
	Opt("RunErrorsFatal", 1)

	; Run setup
;~ 	$strFile = @HomeDrive & "\Apps\CPU_Freq_Display\Frequency Display.exe"
;~ 	$strDir  = @HomeDrive & "\Apps\CPU_Freq_Display"
	If Not FileExists($strFile) Then
		Return False
	EndIf
	;Run($strFile, $strDir)
	Run(@ComSpec & " /c " & '"' & $strFile & '"', "", @SW_HIDE)

;~ 	; Welcome to the InstallShield Wizard for Frequency Display
;~ 	$strTitle = "Intel(R) Frequency Display Setup"
;~ 	$strText  = "Welcome to the InstallShield Wizard for Frequency Display"
;~ 	If Not WinWait($strTitle, $strText, 30) Then
;~ 		WriteLog("Step 1 of 4")
;~ 		Return False
;~ 	EndIf
;~ 	ControlClick($strTitle, "&Next >", "&Next >")

;~ 	; License Agreement
;~ 	$strText = "License Agreement"
;~ 	If Not WinWait($strTitle, $strText, 10) Then
;~ 		WriteLog("Step 2 of 4")
;~ 		Return False
;~ 	EndIf
;~ 	ControlClick($strTitle, "&Yes", "&Yes")

;~ 	; Choose Destination Location
;~ 	$strText = "Choose Destination Location"
;~ 	If Not WinWait($strTitle, $strText, 10) Then
;~ 		WriteLog("Step 3 of 4")
;~ 		Return False
;~ 	EndIf
;~ 	ControlClick($strTitle, "&Next >", "&Next >")

;~ 	; Intel(R) Frequency Display Utility Setup Complete
;~ 	$strText = "Intel(R) Frequency Display Utility Setup Complete"
;~ 	If Not WinWait($strTitle, $strText, 150) Then
;~ 			WriteLog("Step 4 of 4 ")
;~ 			Return False
;~ 	EndIf
;~ 	ControlClick($strTitle, "Finish", "Finish")

	Return True
EndFunc

;***************************************************************************
;** Main Program
;***************************************************************************

	; Prepre for running
	ScriptStarting($SCRIPT_NAME, $SCRIPT_FILENAME, $SCRIPT_VERSION)

	; Only run on Win2K and beyond.
	If Not IsSupportedOS() Then
		$strError = "ERROR: Unsupported OS"
		ExitWithErrorMessage($SCRIPT_NAME, $strError)
		Exit
	EndIf

	; If app already installed, abort this install script. Write error to log and registry.
	If IsAppletInstalled($SCRIPT_NAME) Then
  		$strError = "ABORTED: " & $SCRIPT_NAME & " is installed already."
		ExitWithErrorMessage($SCRIPT_NAME, $strError)
		Exit
	Endif

	; Allow user press ESC to abort the program
	AbortProgram($SCRIPT_NAME)

	; Local available or need to get from server
;~ 	$strFile = @HomeDrive & "\APPS\CPU_Freq_Display\Frequency Display.exe"
	If Not FileExists($strFile) Then
		; Assign source/destination dir for copy.
;~ 		$strSourceDir = "\APPS\CPU_Freq_Display"
;~ 		$strDestDir   = "\APPS\CPU_Freq_Display"

		; Copy from server.
		If Not CopyFromServer($SCRIPT_NAME, $strSourceDir, $strDestDir) Then
  			$strError = "ERROR: Could not copy from netowrk."
			ExitWithErrorMessage($SCRIPT_NAME, $strError)
			Exit
		EndIf
	EndIf

	; Call RunSetup
	If Not RunSetup() Then
		$strError = "ERROR: RunSetup() funcation failed."
		ExitWithErrorMessage($SCRIPT_NAME, $strError)
		Exit
	EndIf

	; Ending
	ScriptEnding($SCRIPT_NAME, "INSTALLED")

	Exit

;***************************************************************************
;** End Of Program
;***************************************************************************
