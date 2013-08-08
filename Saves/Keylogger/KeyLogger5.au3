#Region
#AutoIt3Wrapper_Run_Au3check=y
#AutoIt3Wrapper_Au3Check_Stop_OnWarning=y
#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Tidy_Stop_OnError=y
;#Tidy_Parameters=/gd /sf
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_Res_FileVersion_AutoIncrement=Y
#AutoIt3Wrapper_Res_Description=KeyLogger
#AutoIt3Wrapper_Res_LegalCopyright=GNU-PL
#AutoIt3Wrapper_Res_Comment=A program to process photos
#AutoIt3Wrapper_Res_Field=Developer|Douglas Kaynor
#AutoIt3Wrapper_Res_LegalCopyright=Copyright ? 2009 Douglas B Kaynor
#AutoIt3Wrapper_Res_Field= AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Res_Field=Compile date|%longdate% %time%
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Run_Before=
#AutoIt3Wrapper_Icon=eagle.ico
#EndRegion

Opt("MustDeclareVars", 1)

If _Singleton(@ScriptName, 1) = 0 Then
	Debug(@ScriptName & " is already running!", 0x40, 5)
	Exit
EndIf

#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <TreeViewConstants.au3>
#include <GuiTreeView.au3>
#include <Date.au3>
#include <Misc.au3>
#include <String.au3>
#include <EditConstants.au3>
#include <GUIListBox.au3>
#include <Misc.au3>
#include <_DougFunctions.au3>

Global $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")

Global $Tmp = StringSplit(@ScriptName, ".")
Global $Project_filename = @ScriptDir & "\AUXFiles\" & $Tmp[1] & ".prj"
Global $LOG_filename = @ScriptDir & "\AUXFiles\" & $Tmp[1] & ".log"
Global $WorkingFolder = @ScriptDir

Global $MainForm = GUICreate(@ScriptName & $FileVersion, 670, 60, 10, 10, BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS))
GUICtrlSetTip($MainForm, "This is the main form of the program")
GUISetFont(10, 400, 0, "Courier New")

Global $ButtonGetRaw = GUICtrlCreateButton("Get raw", 10, 10, 100, 40, $WS_GROUP)
GUICtrlSetTip($ButtonGetRaw, "Get raw data")

Global $ButtonProcess = GUICtrlCreateButton("Process raw", 120, 10, 100, 40, $WS_GROUP)
GUICtrlSetTip($ButtonProcess, "Process raw data")

Global $ButtonSaveList = GUICtrlCreateButton("Save List", 230, 10, 100, 40, $WS_GROUP)
GUICtrlSetTip($ButtonSaveList, "Save the list")

Global $ButtonLoadList = GUICtrlCreateButton("Load List", 340, 10, 100, 40, $WS_GROUP)
GUICtrlSetTip($ButtonLoadList, "Load a list")

Global $ButtonShow = GUICtrlCreateButton("Show list", 450, 10, 100, 40, $WS_GROUP)
GUICtrlSetTip($ButtonShow, "Show raw list")

Global $ButtonExit = GUICtrlCreateButton("Exit", 560, 10, 100, 40, $WS_GROUP)
GUICtrlSetTip($ButtonExit, "Exit the program")

Global $UserDll = DllOpen("user32.dll")

GUISetState(@SW_SHOW)
;-----------------------------------------------
While 1
	Global $nMsg = GUIGetMsg()
	Switch $nMsg
		Case $ButtonGetRaw

		Case $ButtonProcess

		Case $ButtonSaveList

		Case $ButtonLoadList

		Case $GUI_EVENT_CLOSE
			Exit
		Case $ButtonExit
			Exit
	EndSwitch

	If _IsPressed(0xBA) = 1 Then _LogKeyPress(';')
	If _IsPressed(0xBB) = 1 Then _LogKeyPress('=')
	If _IsPressed(0xBC) = 1 Then _LogKeyPress(',')
	If _IsPressed(0xBD) = 1 Then _LogKeyPress('-')
	If _IsPressed(0xBE) = 1 Then _LogKeyPress('.')
	If _IsPressed(0xBF) = 1 Then _LogKeyPress('/')
	If _IsPressed(0xC0) = 1 Then _LogKeyPress('`')
	If _IsPressed(0xDB) = 1 Then _LogKeyPress('[')
	If _IsPressed(0xDC) = 1 Then _LogKeyPress('\')
	If _IsPressed(0xDD) = 1 Then _LogKeyPress(']')
	If _IsPressed(0xDE) = 1 Then _LogKeyPress("'")
	If _IsPressed(0x08) = 1 Then _LogKeyPress('{BACKSPACE}')
	If _IsPressed(0x09) = 1 Then _LogKeyPress('{TAB}')
	If _IsPressed(0x0D) = 1 Then _LogKeyPress('{ENTER}')
	If _IsPressed(0x13) = 1 Then _LogKeyPress('{PAUSE}')
	If _IsPressed(0x14) = 1 Then _LogKeyPress('{CAPSLOCK}')
	If _IsPressed(0x1B) = 1 Then _LogKeyPress('{ESC}')
	If _IsPressed(0x20) = 1 Then _LogKeyPress('{ }')
	If _IsPressed(0x21) = 1 Then _LogKeyPress('{PAGE UP}')
	If _IsPressed(0x22) = 1 Then _LogKeyPress('{PAGE DOWN}')
	If _IsPressed(0x23) = 1 Then _LogKeyPress('{END}')
	If _IsPressed(0x24) = 1 Then _LogKeyPress('{HOME}')
	If _IsPressed(0x25) = 1 Then _LogKeyPress('{LEFT ARROW}')
	If _IsPressed(0x26) = 1 Then _LogKeyPress('{UP ARROW}')
	If _IsPressed(0x27) = 1 Then _LogKeyPress('{RIGHT ARROW}')
	If _IsPressed(0x28) = 1 Then _LogKeyPress('{DOWN ARROW}')
	If _IsPressed(0x2C) = 1 Then _LogKeyPress('{PRINT SCREEN}')
	If _IsPressed(0x2D) = 1 Then _LogKeyPress('{INS}')
	If _IsPressed(0x2E) = 1 Then _LogKeyPress('{DEL}')
	If _IsPressed(0x30) = 1 Then _LogKeyPress('0')
	If _IsPressed(0x31) = 1 Then _LogKeyPress('1')
	If _IsPressed(0x32) = 1 Then _LogKeyPress('2')
	If _IsPressed(0x33) = 1 Then _LogKeyPress('3')
	If _IsPressed(0x34) = 1 Then _LogKeyPress('4')
	If _IsPressed(0x35) = 1 Then _LogKeyPress('5')
	If _IsPressed(0x36) = 1 Then _LogKeyPress('6')
	If _IsPressed(0x37) = 1 Then _LogKeyPress('7')
	If _IsPressed(0x38) = 1 Then _LogKeyPress('8')
	If _IsPressed(0x39) = 1 Then _LogKeyPress('9')
	If _IsPressed(0x41) = 1 Then _LogKeyPress('a')
	If _IsPressed(0x42) = 1 Then _LogKeyPress('b')
	If _IsPressed(0x43) = 1 Then _LogKeyPress('c')
	If _IsPressed(0x44) = 1 Then _LogKeyPress('d')
	If _IsPressed(0x45) = 1 Then _LogKeyPress('e')
	If _IsPressed(0x46) = 1 Then _LogKeyPress('f')
	If _IsPressed(0x47) = 1 Then _LogKeyPress('g')
	If _IsPressed(0x48) = 1 Then _LogKeyPress('h')
	If _IsPressed(0x49) = 1 Then _LogKeyPress('i')
	If _IsPressed(0x4A) = 1 Then _LogKeyPress('j')
	If _IsPressed(0x4B) = 1 Then _LogKeyPress('k')
	If _IsPressed(0x4C) = 1 Then _LogKeyPress('l')
	If _IsPressed(0x4D) = 1 Then _LogKeyPress('m')
	If _IsPressed(0x4E) = 1 Then _LogKeyPress('n')
	If _IsPressed(0x4F) = 1 Then _LogKeyPress('o')
	If _IsPressed(0x50) = 1 Then _LogKeyPress('p')
	If _IsPressed(0x51) = 1 Then _LogKeyPress('q')
	If _IsPressed(0x52) = 1 Then _LogKeyPress('r')
	If _IsPressed(0x53) = 1 Then _LogKeyPress('s')
	If _IsPressed(0x54) = 1 Then _LogKeyPress('t')
	If _IsPressed(0x55) = 1 Then _LogKeyPress('u')
	If _IsPressed(0x56) = 1 Then _LogKeyPress('v')
	If _IsPressed(0x57) = 1 Then _LogKeyPress('w')
	If _IsPressed(0x58) = 1 Then _LogKeyPress('x')
	If _IsPressed(0x59) = 1 Then _LogKeyPress('y')
	If _IsPressed(0x5A) = 1 Then _LogKeyPress('z')
	If _IsPressed(0x5B) = 1 Then _LogKeyPress('{LEFT WIN}')
	If _IsPressed(0x5C) = 1 Then _LogKeyPress('{RIGHT WIN}')
	If _IsPressed(0x60) = 1 Then _LogKeyPress('Num 0')
	If _IsPressed(0x61) = 1 Then _LogKeyPress('Num 1')
	If _IsPressed(0x62) = 1 Then _LogKeyPress('Num 2')
	If _IsPressed(0x63) = 1 Then _LogKeyPress('Num 3')
	If _IsPressed(0x64) = 1 Then _LogKeyPress('Num 4')
	If _IsPressed(0x65) = 1 Then _LogKeyPress('Num 5')
	If _IsPressed(0x66) = 1 Then _LogKeyPress('Num 6')
	If _IsPressed(0x67) = 1 Then _LogKeyPress('Num 7')
	If _IsPressed(0x68) = 1 Then _LogKeyPress('Num 8')
	If _IsPressed(0x69) = 1 Then _LogKeyPress('Num 9')
	If _IsPressed(0x6A) = 1 Then _LogKeyPress('{MULTIPLY}')
	If _IsPressed(0x6B) = 1 Then _LogKeyPress('{ADD}')
	If _IsPressed(0x6C) = 1 Then _LogKeyPress('Separator')
	If _IsPressed(0x6D) = 1 Then _LogKeyPress('{SUBTRACT}')
	If _IsPressed(0x6E) = 1 Then _LogKeyPress('{DECIMAL}')
	If _IsPressed(0x6F) = 1 Then _LogKeyPress('{DIVIDE}')
	If _IsPressed(0x70) = 1 Then _LogKeyPress('F1')
	If _IsPressed(0x71) = 1 Then _LogKeyPress('F2')
	If _IsPressed(0x72) = 1 Then _LogKeyPress('F3')
	If _IsPressed(0x73) = 1 Then _LogKeyPress('F4')
	If _IsPressed(0x74) = 1 Then _LogKeyPress('F5')
	If _IsPressed(0x75) = 1 Then _LogKeyPress('F6')
	If _IsPressed(0x76) = 1 Then _LogKeyPress('F7')
	If _IsPressed(0x77) = 1 Then _LogKeyPress('F8')
	If _IsPressed(0x78) = 1 Then _LogKeyPress('F9')
	If _IsPressed(0x79) = 1 Then _LogKeyPress('F10')
	If _IsPressed(0x77) = 1 Then _LogKeyPress('F8')
	If _IsPressed(0x78) = 1 Then _LogKeyPress('F9')
	If _IsPressed(0x79) = 1 Then _LogKeyPress('F10')
	If _IsPressed(0x7A) = 1 Then _LogKeyPress('F11')
	If _IsPressed(0x7B) = 1 Then _LogKeyPress('F12')
	If _IsPressed(0x7C) = 1 Then _LogKeyPress('F13')
	If _IsPressed(0x7D) = 1 Then _LogKeyPress('F14')
	If _IsPressed(0x7E) = 1 Then _LogKeyPress('F15')
	If _IsPressed(0x7F) = 1 Then _LogKeyPress('F16')
	If _IsPressed(0x80) = 1 Then _LogKeyPress('F17')
	If _IsPressed(0x81) = 1 Then _LogKeyPress('F18')
	If _IsPressed(0x82) = 1 Then _LogKeyPress('F19')
	If _IsPressed(0x83) = 1 Then _LogKeyPress('F20')
	If _IsPressed(0x84) = 1 Then _LogKeyPress('F21')
	If _IsPressed(0x85) = 1 Then _LogKeyPress('F22')
	If _IsPressed(0x86) = 1 Then _LogKeyPress('F23')
	If _IsPressed(0x87) = 1 Then _LogKeyPress('F24')
	If _IsPressed(0x90) = 1 Then _LogKeyPress('{NUM LOCK}')
	If _IsPressed(0x91) = 1 Then _LogKeyPress('{SCROLL LOCK}')
	If _IsPressed(0xA0) = 1 Then _LogKeyPress('{LEFT SHIFT}')
	If _IsPressed(0xA1) = 1 Then _LogKeyPress('{RIGHT SHIFT}')
	If _IsPressed(0xA2) = 1 Then _LogKeyPress('{LEFT CTRL}')
	If _IsPressed(0xA3) = 1 Then _LogKeyPress('{RIGHT CTRL}')
	If _IsPressed(0xA4) = 1 Then _LogKeyPress('{LEFT ALT}')
	If _IsPressed(0xA5) = 1 Then _LogKeyPress('{RIGHT ALT}')
	Sleep(100)
WEnd

;-----------------------------------------------


Func _LogKeyPress($what2log)
	ConsoleWrite($what2log)
	MsgBox(262144, 'Debug line ~' & @ScriptLineNumber, $what2log) ;### Debug MSGBOX
EndFunc   ;==>_LogKeyPress

;-------------------------------------------aa----