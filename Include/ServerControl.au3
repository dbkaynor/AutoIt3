#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=ServerControl.ico
#AutoIt3Wrapper_Res_Comment=AuCGI Server Controller
#AutoIt3Wrapper_Res_Description=AuCGI Server Controller
#AutoIt3Wrapper_Res_Fileversion=1.0.1.2
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#AutoIt3Wrapper_AU3Check_Parameters=-q -d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Run_Obfuscator=y
#Obfuscator_Parameters=/sf /sv /om /cs=0 /cn=0
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;=======================================
; Original Author: Shafayat
; Updated by: Erik Pilsits (wraithdu)
;=======================================

#include <Constants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <WinAPI.au3>
#include <GuiEdit.au3>

Opt("GUIOnEventMode", 1)
Opt("GUIEventOptions", 1) ; just send min/max/close notifications
Opt("TrayOnEventMode", 1)
Opt("TrayMenuMode", 3) ; default tray menu items (Script Paused/Exit) will not be shown

If ProcessExists("LightTPD.exe") Then
	; server already running
	MsgBox(16, "Error", "Server already running.")
	Exit
EndIf

TraySetState()
TraySetIcon(@ScriptDir & "\LightTPD.exe")
TrayCreateItem("Start")
TrayItemSetOnEvent(-1, "gStartClick")
TrayCreateItem("Stop")
TrayItemSetOnEvent(-1, "gEndClick")
TrayCreateItem("")
TrayCreateItem("Exit")
TrayItemSetOnEvent(-1, "gAppClose")
TraySetOnEvent($TRAY_EVENT_PRIMARYUP, "gRestore")
TraySetClick(8) ; press secondary mouse to show menu
TraySetToolTip("Status: Not Running")

GUICreate("AuCGI Server Controller 1.01", 640, 480)
GUISetOnEvent($GUI_EVENT_CLOSE, "gAppClose")
GUISetOnEvent($GUI_EVENT_MINIMIZE, "gMinimize")
Global $gEdit = GUICtrlCreateEdit("", 2, 48, 636, 430, BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $ES_READONLY, $ES_WANTRETURN, $WS_HSCROLL, $WS_VSCROLL))
GUICtrlSetBkColor($gEdit, 0xFFFFFF)
GUICtrlSetData($gEdit, "AuCGI Server Controller 1.01" & @CRLF)
GUICtrlCreateButton("Start", 4, 4, 80, 40)
GUICtrlSetFont(-1, 11)
GUICtrlSetOnEvent(-1, "gStartClick")
GUICtrlCreateButton("Stop", 88, 4, 80, 40)
GUICtrlSetFont(-1, 11)
GUICtrlSetOnEvent(-1, "gEndClick")
GUICtrlCreateButton("Restart", 172, 4, 80, 40)
GUICtrlSetFont(-1, 11)
GUICtrlSetOnEvent(-1, "gRestartClick")
GUICtrlCreateButton("Clear Log", 276, 4, 80, 40)
GUICtrlSetFont(-1, 11)
GUICtrlSetOnEvent(-1, "gClearLog")
Global $gStatus = GUICtrlCreateLabel("Status: Not Running", 440, 26, 200, 22)
GUICtrlSetFont(-1, 12)
GUICtrlSetColor($gStatus, 0xFF0000)

Global $PID = 0

If $CmdLine[0] > 0 Then
	Switch $CmdLine[1]
		Case "-start"
			; minimized, start server
			gStartClick()
		Case "-startshow"
			; show gui, start server
			GUISetState()
			gStartClick()
		Case "-min"
			; minimized, server not started
		Case Else
			; show GUI, server not started
			GUISetState()
	EndSwitch
Else
	GUISetState()
EndIf

While 1
	Sleep(1000)
WEnd

Func gAppClose()
	If $PID Then ProcessClose($PID)
	Exit
EndFunc   ;==>gAppClose

Func gMinimize()
	GUISetState(@SW_HIDE)
EndFunc   ;==>gMinimize

Func gRestore()
	GUISetState(@SW_SHOW)
EndFunc   ;==>gRestore

Func gEndClick()
	If $PID Then
		ProcessClose($PID)
		$PID = 0
	Else
		gUpdateLog(@CRLF & " Error: Server is not running... ")
	EndIf
EndFunc   ;==>gEndClick

Func gStartClick()
	If Not $PID Then
		$PID = Run('"' & @ScriptDir & '\LightTPD.exe" -f conf/lighttpd-inc.conf -m lib -D', @ScriptDir, @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
		AdlibRegister("CheckStdOut", 1000)
		gUpdateLog(@CRLF & " ===================== SERVER STARTED ===================== ")
		GUICtrlSetData($gStatus, "Status: Running [" & $PID & "]")
		GUICtrlSetColor($gStatus, 0x00FF00)
		TraySetToolTip("Status: Running [" & $PID & "]")
	Else
		gUpdateLog(@CRLF & " Error: Server is already running... ")
	EndIf
EndFunc   ;==>gStartClick

Func gRestartClick()
	If $PID Then
		gTerminateLib()
		gEndClick()
	EndIf
	gStartClick()
EndFunc

Func CheckStdOut()
	Local $err = 0
	Local $line1 = StdoutRead($PID)
	If @error Then $err += 1
	Local $line2 = StderrRead($PID)
	If @error Then $err += 1
	If $err = 2 Then
		gTerminateLib()
		Return
	EndIf

	If $line1 <> '' Then
		$line1 = StringReplace(StringReplace($line1, @LF, @CRLF), @CR, @CRLF)
		gUpdateLog(@CRLF & $line1)
	EndIf
	If $line2 <> '' Then
		$line2 = StringReplace(StringReplace($line2, @LF, @CRLF), @CR, @CRLF)
		gUpdateLog(@CRLF & $line2)
	EndIf
EndFunc   ;==>CheckStdOut

Func gTerminateLib()
	gUpdateLog(@CRLF & " ===================== SERVER STOPPED ===================== ")
	GUICtrlSetData($gStatus, "Status: Not Running")
	GUICtrlSetColor($gStatus, 0xFF0000)
	TraySetToolTip("Status: Not Running")
	AdlibUnRegister("CheckStdOut")
EndFunc   ;==>gTerminateLib

Func gUpdateLog($sText)
	_GUICtrlEdit_AppendText(GUICtrlGetHandle($gEdit), $sText)
EndFunc

Func gClearLog()
	GUICtrlSetData($gEdit, "")
EndFunc