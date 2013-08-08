#Region
#AutoIt3Wrapper_Run_Au3check=y
#AutoIt3Wrapper_Au3Check_Stop_OnWarning=y
#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Tidy_Stop_OnError=y
;#Tidy_Parameters=/gd /sf
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Fileversion=2.0.1.14
#AutoIt3Wrapper_Res_FileVersion_AutoIncrement=Y
#AutoIt3Wrapper_Res_Description=GetIPConfig
#AutoIt3Wrapper_Res_LegalCopyright=GNU-PL
#AutoIt3Wrapper_Res_Comment=A program to get Network info
#AutoIt3Wrapper_Res_Field=Developer|Douglas Kaynor
#AutoIt3Wrapper_Res_LegalCopyright=Copyright © 2009 Douglas B Kaynor
#AutoIt3Wrapper_Res_Field=AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Res_Field=Compile date|%longdate% %time%
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Run_Before=
#AutoIt3Wrapper_Icon=./icons/head_question.ico
#EndRegion

Opt("MustDeclareVars", 1)

If _Singleton(@ScriptName, 1) = 0 Then
	Debug(@ScriptName & " is already running!", 0x40, 5)
	Exit
EndIf

#include <Array.au3>
#include <Date.au3>
#include <GUIConstants.au3>
#include <GuiListView.au3>
#include <GuiTreeView.au3>
#include <GuiConstantsEx.au3>
#include <GuiImageList.au3>
#include <iNet.au3>
#include <Misc.au3>
#include <String.au3>
#include <WindowsConstants.au3>
#include "_DougFunctions.au3"

Global $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")

Global $SystemS = @ScriptName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & @OSServicePack & @CRLF & @OSType & @CRLF & @OSArch & @IPAddress1
Global $mainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)

Global $FontSize = 11

Global $PC = "c:\GetIpConfig.txt"
Global $EXEName

;GUICtrlCreateLabel ( "text", left, top [, width [, height [, style [, exStyle]]]] )

Global $MainForm = GUICreate(@ScriptName & " " & $FileVersion, 500, 225, 10, 10, $mainFormOptions)
GUISetHelp("notepad .\GetIPConfig.au3", $MainForm) ; Need a help file to call here

Global $ButtonTest = GUICtrlCreateButton("Test", 10, 15, 50, 20, 0)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetFont(-1, $FontSize, 400, 0, "Courier Ne1w")
GUICtrlSetTip(-1, "Start testing")

Global $CheckStop = GUICtrlCreateCheckbox("Stop", 70, 10, 60, 30)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetFont(-1, $FontSize, 400, 0, "Courier New")
GUICtrlSetTip(-1, "Stop  looping.")

Global $LabelDelay = GUICtrlCreateLabel("Delay", 135, 15, 55, 20, 0x1000)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetFont(-1, $FontSize, 400, 0, "Courier New")
GUICtrlSetTip(-1, "Delay between loops in milliseconds.")

Global $InputDelay = GUICtrlCreateInput("1000", 200, 15, 50, 20, 0x081)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetFont(-1, $FontSize, 400, 0, "Courier New")
GUICtrlSetTip(-1, "Delay between loops in milliseconds.")

Global $LabelLoops = GUICtrlCreateLabel("Loops", 260, 15, 60, 20, 0x1000)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetFont(-1, $FontSize, 400, 0, "Courier New")
GUICtrlSetTip(-1, "Number of loop to do. Enter 0 for continous.")

Global $InputLoops = GUICtrlCreateInput("1", 320, 15, 50, 20, 0x081)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetFont(-1, $FontSize, 400, 0, "Courier New")
GUICtrlSetTip(-1, "Number of loop to do. Enter 0 for continous.")

Global $ButtonAbout = GUICtrlCreateButton("About", 380, 15, 55, 20, 0)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetFont(-1, $FontSize, 400, 0, "Courier New")
GUICtrlSetTip(-1, "About this program")

Global $ButtonExit = GUICtrlCreateButton("Exit", 440, 15, 50, 20, 0)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetFont(-1, $FontSize, 400, 0, "Courier New")
GUICtrlSetTip(-1, "Exit the program")

Global $CheckShowIPv4 = GUICtrlCreateCheckbox("IPv4", 10, 40, 50, 30)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetFont(-1, $FontSize, 400, 0, "Courier New")
GUICtrlSetTip(-1, "Show IPv4")
GUICtrlSetState($CheckShowIPv4, $GUI_CHECKED)

Global $CheckShowIPv6 = GUICtrlCreateCheckbox("IPv6", 70, 40, 50, 30)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetFont(-1, $FontSize, 400, 0, "Courier New")
GUICtrlSetTip(-1, "Show IPv6")
GUICtrlSetState($CheckShowIPv6, $GUI_UNCHECKED)

Global $CheckShowMAC = GUICtrlCreateCheckbox("MAC", 130, 40, 50, 30)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetFont(-1, $FontSize, 400, 0, "Courier New")
GUICtrlSetTip(-1, "Show MAC")
GUICtrlSetState($CheckShowMAC, $GUI_UNCHECKED)

Global $ButtonRenew = GUICtrlCreateButton("Renew", 180, 45, 50, 20, 0)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetFont(-1, $FontSize, 400, 0, "Courier New")
GUICtrlSetTip(-1, "Renew IP")

Global $ButtonFWoff = GUICtrlCreateButton("FW off", 240, 45, 50, 20, 0)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetFont(-1, $FontSize, 400, 0, "Courier New")
GUICtrlSetTip(-1, "Firewall off")

Global $ButtonFWon = GUICtrlCreateButton("FW on", 300, 45, 50, 20, 0)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetFont(-1, $FontSize, 400, 0, "Courier New")
GUICtrlSetTip(-1, "Firewall on")

Global $LabelCount = GUICtrlCreateLabel("Count", 370, 45, 100, 20, 0x1000)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetFont(-1, $FontSize, 400, 0, "Courier New")
GUICtrlSetTip(-1, "Lopp count")

Global $hTreeView = GUICtrlCreateTreeView(15, 80, 470, 130, BitOR($TVS_LINESATROOT, $TVS_INFOTIP, $WS_GROUP, $WS_TABSTOP), $WS_EX_CLIENTEDGE)
GUICtrlSetResizing(-1, $GUI_DOCKBOTTOM + $GUI_DOCKTOP + $GUI_DOCKLEFT)
GUICtrlSetFont(-1, $FontSize, 1, 0, "Courier New")
GUICtrlSetTip(-1, "This is where the results are listed")

Debug("DBGVIEWCLEAR")

GUISetState(@SW_SHOW)

GetIpInfo()
GetFirewallStatus()

While 1
	Global $nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $ButtonExit
			Exit
		Case $ButtonTest
			GetIpInfoTest()
			GetFirewallStatus()
		Case $ButtonAbout
			About()
		Case $CheckShowMAC
			GetIpInfo()
		Case $CheckShowIPv4
			GetIpInfo()
		Case $CheckShowIPv6
			GetIpInfo()
		Case $ButtonRenew
			RenewIp()
		Case $ButtonFWoff
			$EXEName = "netsh.exe"
			RunWait(@ComSpec & " /c " & $EXEName & " firewall set opmode disable > " & $PC, ".", @SW_HIDE)
			GetFirewallStatus()
			;netsh firewall set opmode disable
		Case $ButtonFWon
			$EXEName = "netsh.exe"
			RunWait(@ComSpec & " /c " & $EXEName & " firewall set opmode enable > " & $PC, ".", @SW_HIDE)
			GetFirewallStatus()
			;netsh firewall set opmode enable
	EndSwitch
WEnd

;-----------------------------------------------
; This routine runs ipconfig with renew
Func RenewIp()
	_GUICtrlTreeView_DeleteAll($hTreeView)
	$EXEName = EnvGet("windir") & "\system32\ipconfig.exe"
	RunWait(@ComSpec & " /c " & $EXEName & " /renew > " & $PC, ".", @SW_HIDE)
	GetIpInfo()
EndFunc   ;==>RenewIp
;-----------------------------------------------
; This routine runs ipconfig and then parses the resulting ipout.txt file.
Func GetIpInfo()
	_GUICtrlTreeView_DeleteAll($hTreeView)
	$EXEName = EnvGet("windir") & "\system32\ipconfig.exe"
	RunWait(@ComSpec & " /c " & $EXEName & " /all > " & $PC, ".", @SW_HIDE)

	Local $file = FileOpen($PC, 0)
	; Check if file opened for reading OK
	If $file = -1 Then
		Debug(" GetIpInfo: Unable to open file for reading: " & $PC, 0x10, 5)
		Return
	EndIf

	; Now parse the ipout file
	While 1
		Local $LineIn = FileReadLine($file)
		If @error = -1 Then ExitLoop
		Local $t1 = StringStripWS($LineIn, 8)
		Local $t2 = StringSplit($t1, "..:", 1)

		Local $tmpS
		If StringInStr($LineIn, "IPv4 Address") > 1 Or StringInStr($LineIn, "IP Address") > 1 Or StringInStr($LineIn, "IPv6 Address") > 1 Then
			$tmpS = $t2[2]
			debug(">>" & $tmpS & "<<")
			If GUICtrlRead($CheckShowIPv4) = $GUI_CHECKED And StringInStr($tmpS, ":") = 0 Then
				_GUICtrlTreeView_Add($hTreeView, 0, "IPv4: " & $tmpS)
			ElseIf GUICtrlRead($CheckShowIPv6) = $GUI_CHECKED And StringInStr($tmpS, ":") Then
				_GUICtrlTreeView_Add($hTreeView, 0, "IPv6: " & $tmpS)
			EndIf
		EndIf

		If GUICtrlRead($CheckShowMAC) = $GUI_CHECKED Then
			If StringInStr($LineIn, "Physical Address") > 1 Then
				Local $D = StringSplit($LineIn, '-') ;Test for valid MAC addresses only
				;_ArrayDisplay($D, "????")
				If $D[0] = 6 Then
					_GUICtrlTreeView_Add($hTreeView, 0, "MAC: " & $t2[2])
				EndIf
			EndIf
		EndIf
	WEnd

	FileClose($file)

EndFunc   ;==>GetIpInfo
;-----------------------------------------------
Func GetFirewallStatus()
	$EXEName = "netsh.exe"
	RunWait(@ComSpec & " /c " & $EXEName & " firewall show state > " & $PC, ".", @SW_HIDE)
	; netsh.exe firewall show state | findstr /i "Operational"
	; netsh firewall set opmode disable
	; netsh firewall set opmode enable

	Local $file = FileOpen($PC, 0)
	; Check if file opened for reading OK
	If $file = -1 Then
		Debug(" GetIpInfo: Unable to open file for reading: " & $PC, 0x10, 5)
		Return
	EndIf

	While 1
		Local $LineIn = FileReadLine($file)
		If @error = -1 Then ExitLoop
		If StringInStr($LineIn, "Operational mode") > 0 Then
			;debug(@ScriptLineNumber, $LineIn)
			Local $A = StringSplit($LineIn, "=")
			;_ArrayDisplay($A)
			Local $R = "Firewall: " & StringStripWS($A[2], 7)
			_GUICtrlTreeView_Add($hTreeView, 0, $R)
			;GUICtrlSetData($LabelFirewall, $R)
		EndIf
	WEnd

	FileClose($file)
EndFunc   ;==>GetFirewallStatus
;-----------------------------------------------
Func GetIpInfoTest()
	GuiDisable("disable")

	Debug("GetIpInfo")
	GUICtrlSetState($CheckStop, $GUI_UNCHECKED)

	Local $MaxLoops = GUICtrlRead($InputLoops)
	Local $Loop = 0
	While $Loop < $MaxLoops Or $MaxLoops <= 0
		$MaxLoops = GUICtrlRead($InputLoops)
		$Loop = $Loop + 1

		GetIpInfo()

		If GUICtrlRead($CheckStop) = $GUI_CHECKED Then ExitLoop

		GUICtrlSetData($LabelCount, StringFormat("Loop %4d", $Loop))

		Sleep(GUICtrlRead($InputDelay))
	WEnd
	GuiDisable("enable")
EndFunc   ;==>GetIpInfoTest
;-----------------------------------------------
Func GuiDisable($choice) ;@SW_ENABLE @SW_disble
	Debug("GuiDisable  " & $choice)

	If $choice = "Enable" Then
		GUICtrlSetState($ButtonTest, $GUI_ENABLE)
		;GUICtrlSetState($ButtonAbout, $GUI_ENABLE)
	ElseIf $choice = "Disable" Then
		GUICtrlSetState($ButtonTest, $GUI_DISABLE)
		;GUICtrlSetState($ButtonAbout, $GUI_DISABLE)
	EndIf
EndFunc   ;==>GuiDisable
;-----------------------------------------------
Func About()
	Local $D = WinGetPos(@ScriptName)
	Local $WinPos
	If IsArray($D) = 1 Then
		$WinPos = StringFormat("%s" & @CRLF & "WinPOS: %d  %d " & @CRLF & "WinSize: %d %d " & @CRLF & "Desktop: %d %d ", _
				$FormID, $D[0], $D[1], $D[2], $D[3], @DesktopWidth, @DesktopHeight)
	Else
		$WinPos = "About ERROR, Check the window name"
	EndIf
	Debug(@CRLF & $SystemS & @CRLF & $WinPos & @CRLF & "Written by Doug Kaynor because I wanted to!", 0x40, 5)
EndFunc   ;==>About

;-----------------------------------------------