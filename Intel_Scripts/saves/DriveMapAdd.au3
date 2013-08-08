#Region
#AutoIt3Wrapper_Run_Au3check=y
#AutoIt3Wrapper_Au3Check_Stop_OnWarning=y
#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Tidy_Stop_OnError=y
;#Tidy_Parameters=/gd /sf
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Fileversion=1.0.0.18
#AutoIt3Wrapper_Res_FileVersion_AutoIncrement=Y
#AutoIt3Wrapper_Res_Description=DriveMapAdd
#AutoIt3Wrapper_Res_LegalCopyright=GNU-PL
#AutoIt3Wrapper_Res_Comment=A program to map my favorite drives
#AutoIt3Wrapper_Res_Field=Developer|Douglas Kaynor
#AutoIt3Wrapper_Res_LegalCopyright=Copyright ? 2009 Douglas B Kaynor
#AutoIt3Wrapper_Res_Field= AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Res_Field=Compile date|%longdate% %time%
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Run_Before=
#AutoIt3Wrapper_Run_Obfuscator=n
#Obfuscator_Parameters= /Convert_Strings=0 /Convert_Numerics=0 /showconsoleinfo=9
#AutoIt3Wrapper_Icon=./icons/10.ico
#EndRegion

#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <misc.au3>
#include <_DougFunctions.au3>

Global $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")
Global $SystemS = @ScriptName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & @OSServicePack & @CRLF & @OSType & @CRLF & @OSArch

Opt("MustDeclareVars", 1)

If _Singleton(@ScriptName, 1) = 0 Then
	Debug(@ScriptName & " is already running!", 0x40, 5)
	Exit
EndIf

Global $WinType = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)

Global $MainForm = GUICreate("Map Drives", 350, 135, 10, 10, 0)
GUISetFont(10, 400, 0, "Courier New")

Global $LabelUser = GUICtrlCreateLabel("User", 15, 15, 70, 50)
Global $InputUser = GUICtrlCreateInput("amr\dbkaynox", 80, 10, 185, 21, $ES_AUTOHSCROLL)
Global $LabelPassword = GUICtrlCreateLabel("Password", 10, 40, 70, 50)
Global $InputPassword = GUICtrlCreateInput("", 80, 40, 185, 21, BitOR($ES_AUTOHSCROLL, $ES_PASSWORD))
Global $LabelShowPassword = GUICtrlCreateLabel("", 80, 70, 200, 50, $ES_AUTOHSCROLL)

Global $ButtonGo = GUICtrlCreateButton("Go", 275, 10, 65, 30)
Global $ButtonAbout = GUICtrlCreateButton("About", 275, 40, 65, 30)
Global $ButtonExit = GUICtrlCreateButton("Exit", 275, 70, 65, 30)

Func ToggleShowPWD()
	GUICtrlRead($InputPassword)
	ToolTip(GUICtrlRead($InputUser) & @CRLF & GUICtrlRead($InputPassword), 0, 0)
	Sleep(3000)
	ToolTip("")
EndFunc   ;==>ToggleShowPWD

HotKeySet("{ESC}", "ToggleShowPWD")

GUISetState(@SW_SHOW)

Send("{TAB}") ; set the focus to the password

While 1
	Switch GUIGetMsg()
		Case $ButtonExit
			Exit
		Case $GUI_EVENT_CLOSE
			Exit
		Case $ButtonAbout
			About()
		Case $ButtonGo
			If StringLen(GUICtrlRead($InputUser)) > 0 And StringLen(GUICtrlRead($InputPassword)) > 0 Then
				Mapit("S:", "\\chakotay\softval", GUICtrlRead($InputUser), GUICtrlRead($InputPassword))
				Mapit("T:", "\\chakotay\temp\dbkaynox", GUICtrlRead($InputUser), GUICtrlRead($InputPassword))
				Mapit("u:", "\\chakotay\temp\dbkaynox\doug\DBKApps", GUICtrlRead($InputUser), GUICtrlRead($InputPassword))
			Else
				MsgBox(0, "Warning!", "User or password field is empty", 3)
			EndIf

	EndSwitch
WEnd

Func Mapit($Drive, $Path, $User, $PWD)
	Local $T
	DriveMapDel($Drive)
	If DriveMapAdd($Drive, $Path, 9, $User, $PWD) <> 1 Then
		Switch @error
			Case 1
				$T = "Undefined / Other error. " & @extended
			Case 2
				$T = "Access to the remote share was denied"
			Case 3
				$T = "The device is already assigned"
			Case 4
				$T = "Invalid device name"
			Case 5
				$T = "Invalid remote share"
			Case 6
				$T = "Invalid password"
			Case Else
				$T = "Unknown error returned " & @error
		EndSwitch
		MsgBox(0, "DriveMap failed", $Drive & @CRLF & $Path & @CRLF & $T, 3)
	Else
		MsgBox(0, "Drive map success", $Drive & " is mapped to" & DriveMapGet($Drive), 1)
	EndIf
EndFunc   ;==>Mapit

Func About()
	Local $D = WinGetPos("Map Drives")
	Local $WinPos
	If IsArray($D) = 1 Then
		$WinPos = StringFormat("%s" & @CRLF & "WinPOS: %d  %d " & @CRLF & "WinSize: %d %d " & @CRLF & "Desktop: %d %d ", _
				$FormID, $D[0], $D[1], $D[2], $D[3], @DesktopWidth, @DesktopHeight)
	Else
		$WinPos = ">>>About ERROR, Check the window name<<<"
	EndIf
	Debug(@CRLF & $SystemS & @CRLF & $WinPos & @CRLF & "Written by Doug Kaynor because I wanted to!", 0x40, 5)
	Debug("HotKeySet {ESC}  ToggleShowPWD")
EndFunc   ;==>About