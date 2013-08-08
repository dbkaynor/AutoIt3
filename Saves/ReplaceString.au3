#Region
#AutoIt3Wrapper_Run_Au3check=y
#AutoIt3Wrapper_Au3Check_Stop_OnWarning=y
#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Tidy_Stop_OnError=y
;#Tidy_Parameters=/gd /sf
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Fileversion=0.0.0.3
#AutoIt3Wrapper_Res_FileVersion_AutoIncrement=Y
#AutoIt3Wrapper_Res_Description=ReplaceString
#AutoIt3Wrapper_Res_LegalCopyright=GNU-PL
#AutoIt3Wrapper_Res_Comment=Replace strings in a file
#AutoIt3Wrapper_Res_Field=Developer|Douglas Kaynor
#AutoIt3Wrapper_Res_LegalCopyright=Copyright 2010 Douglas B Kaynor
#AutoIt3Wrapper_Res_Field= AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Res_Field=Compile date|%longdate% %time%
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Run_Before=
#AutoIt3Wrapper_Run_Obfuscator=n
#Obfuscator_Parameters= /Convert_Strings=0 /Convert_Numerics=0 /showconsoleinfo=9
#AutoIt3Wrapper_Icon="../icons/Canopus.ico"
#EndRegion

#include <Array.au3>
#include <Date.au3>
#include <file.au3>
#include <Misc.au3>
#include <String.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <StaticConstants.au3>
#include "_DougFunctions.au3"

Opt("MustDeclareVars", 1)
TraySetIcon("../icons/Canopus.ico")

Global Const $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")
Global Const $tmp = StringSplit(@ScriptName, ".")
Global Const $ProgramName = $tmp[1]

Global $SystemS = $ProgramName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & @OSServicePack & @CRLF & @OSType & @CRLF & @OSArch

Global Const $XMLFolder = "\\chakotay\Softval\iAMT\AT7.0\Build\XMLSaves"
Global $WorkingFile = ''
Global $InStr
Global $OutStr

If _Singleton(@ScriptName, 1) = 0 Then
	_Debug(@ScriptName & " is already running!", 0x40, 5)
	Exit
EndIf

#Region ### START Koda GUI section ### Form=
Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $MainForm = GUICreate(@ScriptName & " " & $FileVersion, 520, 300, 10, 10, $MainFormOptions)
GUISetFont(10, 400, 0, "Courier New")

Global $ButtonFileLoad = GUICtrlCreateButton("File load", 10, 10, 100, 25)
Global $ButtonFileSave = GUICtrlCreateButton("File save", 10, 40, 100, 25)
Global $ButtonReplace = GUICtrlCreateButton("Replace", 10, 70, 100, 25)
Global $ButtonEdit = GUICtrlCreateButton("Edit", 120, 10, 80, 25)
Global $ButtonAbout = GUICtrlCreateButton("About", 220, 10, 80, 25)
GUICtrlSetTip(-1, "About the program and some Debug stuff")
Global $ButtonHelp = GUICtrlCreateButton("Help", 220, 40, 80, 25)
GUICtrlSetTip(-1, "Display help information")
Global $ButtonExit = GUICtrlCreateButton("Exit", 220, 70, 80, 25)
GUICtrlSetTip(-1, "Exit the program")
Global $EditString2Find = GUICtrlCreateEdit("\firmware$\", 10, 100, 350, 40)
GUICtrlSetResizing(-1, 802)
Global $EditString2Replace = GUICtrlCreateEdit("\Preload\Firmware\", 10, 140, 350, 40)
GUICtrlSetResizing(-1, 802)
Global $LabelStatus = GUICtrlCreateLabel("", 10, 200, 500, 80, $ss_sunken)

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

HotKeySet("{F11}", "GUI_Enable")

While 1
	Global $t = GUIGetMsg()
	Switch $t
		Case $GUI_EVENT_CLOSE
			Exit
		Case $ButtonExit
			Exit
		Case $ButtonEdit
			Edit()
		Case $ButtonFileLoad
			FileChoose()
		Case $ButtonFileSave
			FileSave()
		Case $ButtonReplace
			Replace()
		Case $EditString2Replace
		Case $EditString2Find
		Case $ButtonHelp
			MsgBox(0, "Help", "Someday maybe", 5)
		Case $ButtonAbout
			About($ProgramName)
		Case Else
			;ConsoleWrite($t & '  ')
	EndSwitch
WEnd
;-----------------------------------------------
Func Replace()
	$OutStr = StringReplace($InStr, GUICtrlRead($EditString2Find), GUICtrlRead($EditString2Replace))
EndFunc   ;==>Replace
;-----------------------------------------------
Func FileChoose()
	_Debug("Edit")
	$WorkingFile = FileOpenDialog("View or Edit a file", $XMLFolder, "XML (*.xml)|All (*.*)", 1)
	ConsoleWrite(@ScriptLineNumber & " +++ " & $WorkingFile & @CRLF)
	GUICtrlSetData($LabelStatus, "File loaded: " & $WorkingFile)

	$InStr = FileRead($WorkingFile)

EndFunc   ;==>FileChoose

;-----------------------------------------------
Func FileSave()
	FileWrite("tmp.xml", $OutStr)
EndFunc   ;==>FileSave
;-----------------------------------------------
;This function allows the user to edit or view any file, useful for changing the config file
Func Edit()
	_Debug("Edit")
	Local $Filename = FileOpenDialog("View or Edit a file", $XMLFolder, "XML (*.xml)|All (*.*)", 1)
	ConsoleWrite(@ScriptLineNumber & " +++ " & $Filename & @CRLF)
	Const $edit1 = "c:\program files\notepad++\notepad++.exe"
	Const $edit2 = "c:\program files (x86)\notepad++\notepad++.exe"
	Const $edit3 = "notepad.exe"
	Local $editor = ""

	If FileExists($edit1) = 1 Then
		$editor = $edit1
	ElseIf FileExists($edit2) = 1 Then
		$editor = $edit2
	Else
		$editor = $edit3
	EndIf
	ShellExecute($editor, $Filename)
EndFunc   ;==>Edit
;-----------------------------------------------
Func About(Const $FormID)
	GuiDisable($GUI_DISABLE)
	_Debug("About")
	Local $D = WinGetPos($FormID)
	Local $WinPos
	If IsArray($D) = True Then
		ConsoleWrite(@ScriptLineNumber & $FormID & @CRLF)
		$WinPos = StringFormat("%s" & @CRLF & "WinPOS: %d  %d " & @CRLF & "WinSize: %d %d " & @CRLF & "Desktop: %d %d ", _
				$FormID, $D[0], $D[1], $D[2], $D[3], @DesktopWidth, @DesktopHeight)
	Else
		$WinPos = ">>>About ERROR, Check the window name<<<"
	EndIf
	_Debug(@CRLF & $SystemS & @CRLF & $WinPos & @CRLF & "Written by Doug Kaynor!", 0x40)
	GuiDisable($GUI_ENABLE)
EndFunc   ;==>About
;-----------------------------------------------

; hotkey F11
Func GUI_Enable()
	GuiDisable($GUI_ENABLE)
EndFunc   ;==>GUI_Enable
;-----------------------------------------------
Func GuiDisable($choice) ;$GUI_ENABLE $GUI_DISABLE

	For $X = 1 To 100
		GUICtrlSetState($X, $choice)
	Next
EndFunc   ;==>GuiDisable