#Region
#AutoIt3Wrapper_Run_Au3check=y
#AutoIt3Wrapper_Au3Check_Stop_OnWarning=y
#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Tidy_Stop_OnError=y
;#Tidy_Parameters=/gd /sf
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Fileversion=1.0.0.14
#AutoIt3Wrapper_Res_FileVersion_AutoIncrement=Y
#AutoIt3Wrapper_Res_Description=GetSysInfo
#AutoIt3Wrapper_Res_LegalCopyright=GNU-PL
#AutoIt3Wrapper_Res_Comment=Displays some system information
#AutoIt3Wrapper_Res_Field=Developer|Douglas Kaynor
#AutoIt3Wrapper_Res_LegalCopyright=Copyright ? 2009 Douglas B Kaynor
#AutoIt3Wrapper_Res_Field= AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Res_Field=Compile date|%longdate% %time%
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Run_Before=
#AutoIt3Wrapper_Run_Obfuscator=n
#Obfuscator_Parameters= /Convert_Strings=0 /Convert_Numerics=0 /showconsoleinfo=9
#AutoIt3Wrapper_Icon=./icons/HotSun.ico
#EndRegion

#include <Array.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiTreeView.au3>
#include <WindowsConstants.au3>
#include <misc.au3>
#include <_DougFunctions.au3>

Opt("MustDeclareVars", 1)

Global $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")
Global $SystemS = @ScriptName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & @OSServicePack & @CRLF & @OSType & @CRLF & @OSArch & @IPAddress1

If _Singleton(@ScriptName, 1) = 0 Then
	Debug(@ScriptName & " is already running!", 0x40, 5)
	Exit
EndIf

;Global $WinType = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $MainForm = GUICreate(@ScriptName & " " & $FileVersion, 400, 200, 10, 10, $MainFormOptions)

GUISetFont(10, 400, 0, "Courier New")

GUISetHelp("notepad .\GetSysInfo.au3", $MainForm) ; Need a help file to call here

Global $ButtonRefresh = GUICtrlCreateButton("Refresh", 10, 10, 65, 30)
GUICtrlSetResizing(-1, 802)
Global $ButtonAbout = GUICtrlCreateButton("About", 80, 10, 65, 30)
GUICtrlSetResizing(-1, 802)
Global $ButtonExit = GUICtrlCreateButton("Exit", 150, 10, 65, 30)
GUICtrlSetResizing(-1, 802)

Global $CheckHide = GUICtrlCreateCheckbox("Hide extras", 250, 10, 52, 30)
GUICtrlSetResizing(-1, 802)
GUICtrlSetState($CheckHide, $GUI_CHECKED)

Global $hTreeView = GUICtrlCreateTreeView(10, 50, 800, 140, BitOR($TVS_LINESATROOT, $TVS_INFOTIP, $WS_GROUP, $WS_TABSTOP), $WS_EX_CLIENTEDGE)
;GUICtrlSetResizing($hTreeView, 0)
GUICtrlSetResizing($hTreeView, 98)
;GUICtrlSetTip(-1, "This is where the results are listed")

GUISetState(@SW_SHOW)
GetIt()

While 1
	Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE
			Exit
		Case $ButtonExit
			Exit
		Case $ButtonAbout
			About()
		Case $ButtonRefresh
			GetIt()
	EndSwitch
WEnd
;-----------------------------------------------
Func GetIt()
	_GUICtrlTreeView_DeleteAll($hTreeView)
	Local $RegStringsArray[1]
	_ArrayAdd($RegStringsArray, "HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System~SystemBiosVersion")
	_ArrayAdd($RegStringsArray, "HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System~SystemBiosDate")
	_ArrayAdd($RegStringsArray, "HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\BIOS~BIOSVendor")
	_ArrayAdd($RegStringsArray, "HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\BIOS~BIOSVersion")
	_ArrayAdd($RegStringsArray, "HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\BIOS~SystemProductName")
	_ArrayAdd($RegStringsArray, "HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\BIOS~BaseBoardProduct")

	If GUICtrlRead($CheckHide) = $GUI_UNCHECKED Then
		_ArrayAdd($RegStringsArray, "HKEY_LOCAL_MACHINE\SOFTWARE\Intel\MPG~OEMTablId")
		_ArrayAdd($RegStringsArray, "HKEY_LOCAL_MACHINE\SOFTWARE\Intel\MPG~Os")
		_ArrayAdd($RegStringsArray, "HKEY_LOCAL_MACHINE\SOFTWARE\Intel\MPG~OsFamily")
	EndIf
	_ArrayAdd($RegStringsArray, "HKEY_LOCAL_MACHINE\SOFTWARE\Intel\MPG~OsName")
	_ArrayAdd($RegStringsArray, "HKEY_LOCAL_MACHINE\SOFTWARE\Intel\MPG~PlatformName")

	;_ArrayAdd($RegStringsArray, "HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\BIOS~SystemManufacturer")
	;_ArrayAdd($RegStringsArray, "HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\BIOS~BaseBoardManufacturer")
	_ArrayDelete($RegStringsArray, 0)
	;_ArrayDisplay($RegStringsArray)

	For $A In $RegStringsArray
		Local $B = StringSplit($A, "~")
		;_ArrayDisplay($B)
		Local $C = RegRead($B[1], $B[2])
		If StringLen($C) > 0 And StringInStr($C, "To Be filled by O.E.M") = 0 Then _GUICtrlTreeView_Add($hTreeView, 0, $C)
	Next

	_GUICtrlTreeView_Add($hTreeView, 0, @IPAddress1 & "  " & @IPAddress2 & "  " & @IPAddress3 & "  " & @IPAddress4)

EndFunc   ;==>GetIt
;-----------------------------------------------

Func About()
	Local $D = WinGetPos(@ScriptName)
	Local $WinPos
	If IsArray($D) = 1 Then
		$WinPos = StringFormat("%s" & @CRLF & "WinPOS: %d  %d " & @CRLF & "WinSize: %d %d " & @CRLF & "Desktop: %d %d ", _
				$FormID, $D[0], $D[1], $D[2], $D[3], @DesktopWidth, @DesktopHeight)
	Else
		$WinPos = ">>>About ERROR, Check the window name<<<"
	EndIf
	Debug(@CRLF & $SystemS & @CRLF & $WinPos & @CRLF & "Written by Doug Kaynor because I wanted to!", 0x40, 5)
EndFunc   ;==>About