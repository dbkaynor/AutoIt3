#comments-start PROGRAM HEADER
;******************************************************************************************
;** Intel Corporation, MPG MPCS
;** Title: ImUAC.au3
;** Description: 	
;**	   Config script to disable UAC for non-preload OSes.  It will detect current settings
;**    and disply a UI with an option to change the state to the opposite.  Will not launch
;**    if Preload enviornment is detected.
;**
;** Revision: 	Rev 1.0.0
;******************************************************************************************
;******************************************************************************************
;** Revision History:
;** 
;** Update for Rev 1.0.0 - Chris Shorey 07/24/2007
;**	- Initial release
;**
;******************************************************************************************
#comments-end PROGRAM HEADER

#RequireAdmin

; Script/File name
Dim Const $SCRIPT_NAME 		= "UAC Configurator"
Dim Const $SCRIPT_FILENAME 	= "CmUAC.exe"
Dim Const $SCRIPT_VERSION  	= "V1.0.0"

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
	
	;Variables
	$ButtonTitle = "Abort"
	$UIStatus = "Unknown"

;~ 	If @OSVersion <> "WIN_VISTA" Then
;~ 		MsgBox(0, "Complete", "This script only runs on Vista", 3)
;~ 		Exit
;~ 	EndIf

	If FileExists(@HomeDrive & "\logs\Preload.txt") AND @UserName = "Administrator" Then
;~  	MsgBox(0, "Complete", "This script does not need to be run " & "as you already have an administator rights", 3)
		WriteLog("Running " & $SCRIPT_NAME & " RunSetup() function - Admin and Preload detected.")
		Return True ; script is done so exit nicely
	EndIf
		
	;Read UAC Registry Variable
	$UACState = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "EnableLUA")

	If $UACState = 0 Then
		$UIStatus = "Disabled"
		$ButtonTitle = "Enable"
	ElseIf $UACState = 1 Then
		$UIStatus = "Enabled"
		$ButtonTitle = "Disable"
	Else
		MsgBox(4096, "Error", "UAC variable: " & $UACState)
		WriteLog("Error in " & $SCRIPT_NAME & " RunSetup() function - UAC variable: " & $UACState)
		Return 0
	EndIf

	#Region ### START Koda GUI section ### Form=C:\Documents and Settings\ceshorey\Desktop\UACconfig.kxf
	$Form1 = GUICreate("UAC Configurator", 170, 74, 222, 115, BitOR($WS_SYSMENU,$WS_CAPTION,$WS_POPUPWINDOW,$WS_BORDER,$WS_CLIPSIBLINGS))
	$Cancel = GUICtrlCreateButton("Cancel", 8, 40, 75, 25, 0)
	$Input1 = GUICtrlCreateInput($UIStatus, 74, 8, 89, 21, BitOR($ES_CENTER,$ES_AUTOHSCROLL,$ES_READONLY))
	$Apply = GUICtrlCreateButton($ButtonTitle, 88, 40, 75, 25, 0)
	$Label1 = GUICtrlCreateLabel("UAC Status:", 8, 12, 62, 17)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				WriteLog("Running " & $SCRIPT_NAME & " RunSetup() function - User Exit.")
				Return True
			
			Case $Cancel
				WriteLog("Running " & $SCRIPT_NAME & " RunSetup() function - User Exit.")
				Return True
			
			Case $Apply
				If $UACState = 0 Then
					WriteLog("Running " & $SCRIPT_NAME & " RunSetup() function - Enabling UAC")
					Run(@ComSpec & " /k %windir%\System32\reg.exe ADD HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System /v EnableLUA /t REG_DWORD /d 1 /f", "", "")
					Shutdown(6)  ;Force a reboot
				ElseIf $UACState = 1 Then
					WriteLog("Running " & $SCRIPT_NAME & " RunSetup() function - Disabling UAC")
					Run(@ComSpec & " /k %windir%\System32\reg.exe ADD HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System /v EnableLUA /t REG_DWORD /d 0 /f", "", "")
					Shutdown(6)  ;Force a reboot
				EndIf
				
				Return True
				
		EndSwitch
	WEnd

	Return True
	
EndFunc

;***************************************************************************
;** Main Program
;***************************************************************************

	; Prepre for running
	ScriptStarting($SCRIPT_NAME, $SCRIPT_FILENAME, $SCRIPT_VERSION)
	
	; Only run on Win2K and beyond.
	if Not IsWinVistaFamily() Then
		$strError = "ERROR: Unsupported OS"
		ExitWithErrorMessage($SCRIPT_NAME, $strError)
		Exit
	EndIf
	
	; Allow user press ESC to abort the program
	AbortProgram($SCRIPT_NAME)
	
	; Call RunSetup
	If Not RunSetup() Then
		$strError = "ERROR: RunSetup() funcation failed."
		ExitWithErrorMessage($SCRIPT_NAME, $strError)
		exit 
	EndIf
	
	; Ending with Reboot
	ScriptEndingRebootVManager($SCRIPT_NAME, "EXECUTED")	
	
	; Ending
	;($SCRIPT_NAME, "EXECUTED")
		
	Exit
	
;***************************************************************************
;** End Of Program
;***************************************************************************
	