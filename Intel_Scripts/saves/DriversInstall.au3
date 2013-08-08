#region
#AutoIt3Wrapper_Run_Au3check=y
#AutoIt3Wrapper_Au3Check_Stop_OnWarning=y
#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Tidy_Stop_OnError=y
;#Tidy_Parameters=/gd /sf
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Fileversion=0.0.1.7
#AutoIt3Wrapper_Res_FileVersion_AutoIncrement=Y
#AutoIt3Wrapper_Res_Description=Pingtool
#AutoIt3Wrapper_Res_LegalCopyright=GNU-PL
#AutoIt3Wrapper_Res_Comment=Drivers
#AutoIt3Wrapper_Res_Field=Developer|Douglas Kaynor
#AutoIt3Wrapper_Res_LegalCopyright=Copyright 2010 Douglas B Kaynor
#AutoIt3Wrapper_Res_Field= AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Res_Field=Compile date|%longdate% %time%
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Run_Before=
#AutoIt3Wrapper_Icon=./icons/Glonass_logo.ico
#endregion
TraySetIcon("./icons/Glonass_logo.ico")

Opt("MustDeclareVars", 1)

#include <Array.au3>
#include <Date.au3>
#include <Misc.au3>
#include <String.au3>
#include <ButtonConstants.au3>
#include <WindowsConstants.au3>
#include <TreeViewConstants.au3>
#include <GuiTreeView.au3>
#include <StaticConstants.au3>
#include <GUIConstants.au3>
#include <GUIConstantsEx.au3>
#include "_DougFunctions.au3"


Global $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")
Global $SystemS = @ScriptName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & @OSServicePack & @CRLF & @OSType & @CRLF & @OSArch

Global $tmp = StringSplit(@ScriptName, ".")
Global $ProgramName = $tmp[1]

If _Singleton($ProgramName, 1) = 0 Then
	_Debug($ProgramName & " is already running!", 0x40, 5)
	Exit
EndIf

Global $cfg_filename = FileGetShortName(@ScriptDir & "\AUXFiles\") & "DriversInstall.cfg"
Global $WorkingFolder = @ScriptDir
Global $HandleArray[1]

; "text", left, top [, width [, height [, style [, exStyle]]]]
Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $MainForm = GUICreate($ProgramName & " " & $FileVersion, 550, 350, 10, 10, $MainFormOptions)
GUICtrlSetFont(-1, 12)
Global $ButtonGetDriverList = GUICtrlCreateButton("Get driver list", 10, 10, 100, 30)
GUICtrlSetTip(-1, "Get the list of drivers")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ButtonInstallChecked = GUICtrlCreateButton("Install Checked", 120, 10, 100, 30)
GUICtrlSetTip(-1, "Install checked items")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ButtonChooseConfig = GUICtrlCreateButton("Choose Config", 230, 10, 100, 30)
GUICtrlSetTip(-1, "Choose a config file")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ButtonHelp = GUICtrlCreateButton("Help", 360, 10, 50, 30)
GUICtrlSetTip($ButtonHelp, "Help button")
GUICtrlSetResizing(-1, 800)

Global $ButtonAbout = GUICtrlCreateButton("About", 420, 10, 50, 30)
GUICtrlSetTip($ButtonAbout, "About button")
GUICtrlSetResizing(-1, 800)

Global $ButtonExit = GUICtrlCreateButton("Exit", 480, 10, 50, 30)
GUICtrlSetTip(-1, "Exit the program")
GUICtrlSetResizing(-1, 800)

Global $LabelWorkingFolder = GUICtrlCreateLabel("Working folder", 10, 50, 530, 20, $SS_SUNKEN)
GUICtrlSetTip(-1, "Working folder")
GUICtrlSetResizing(-1, BitOR($GUI_DOCKHEIGHT, $GUI_DOCKTOP))
GUICtrlSetFont(-1, 12)

Global $LabelSysInfo = GUICtrlCreateLabel("System information", 10, 80, 530, 20, $SS_SUNKEN)
GUICtrlSetTip(-1, "System information")
GUICtrlSetResizing(-1, BitOR($GUI_DOCKHEIGHT, $GUI_DOCKTOP))
GUICtrlSetFont(-1, 12)

Global $LabelStatusPath = GUICtrlCreateLabel("Status", 10, 110, 530, 20, $SS_SUNKEN)
GUICtrlSetTip(-1, "Status")
GUICtrlSetResizing(-1, BitOR($GUI_DOCKHEIGHT, $GUI_DOCKTOP))
GUICtrlSetFont(-1, 12)



Global $HTreeView = GUICtrlCreateTreeView(10, 140, 530, 200, BitOR($TVS_HASBUTTONS, $TVS_HASLINES, $TVS_LINESATROOT, $TVS_DISABLEDRAGDROP, $TVS_SHOWSELALWAYS, $TVS_CHECKBOXES, $WS_TABSTOP), $WS_EX_CLIENTEDGE)
GUICtrlSetTip(-1, "This is a list of drivers to install")
GUICtrlSetResizing(-1, BitOR($GUI_DOCKTOP, $GUI_DOCKBOTTOM))
GUICtrlSetFont(-1, 12)

GUISetHelp("notepad", $MainForm) ; Need a help file to call here

GUISetState(@SW_SHOW)

GetSystemInfo()
GetDriverList()
;TestNetwork()

While 1
	Global $nMsg = GUIGetMsg(1)
	Switch $nMsg[0]
		Case $ButtonInstallChecked
			InstallChecked()
		Case $ButtonGetDriverList
			GetDriverList()
		Case $ButtonChooseConfig
			ChooseConfigFile()
		Case $ButtonAbout
			About()
		Case $ButtonHelp
			MsgBox(0, "Help", "Todo list, ???", 0)
		Case $GUI_EVENT_CLOSE
			Exit
		Case $ButtonExit
			Exit
	EndSwitch
WEnd
;-------------------------------------------------------------------------------------
Func InstallChecked()
	_Debug(@ScriptLineNumber & " InstallChecked")
	GUICtrlSetBkColor($LabelStatusPath, Default)
	GUICtrlSetColor($LabelStatusPath, Default)
	Local $HandleArray[1]
	Local $Item = _GUICtrlTreeView_GetFirstItem($HTreeView)
	While True
		If $Item = False Then ExitLoop
		_ArrayAdd($HandleArray, $Item)
		$Item = _GUICtrlTreeView_GetNext($HTreeView, $Item)
	WEnd

	Local $Drivers2Install[1]
	Local $CheckedCount = 0
	For $X = 1 To UBound($HandleArray) - 1
		Local $ItemChecked = _GUICtrlTreeView_GetChecked($HTreeView, $HandleArray[$X])
		Local $ItemText = _GUICtrlTreeView_GetText($HTreeView, $HandleArray[$X])
		If StringInStr($ItemText, '.exe') And $ItemChecked = True Then
			ConsoleWrite($ItemText & "  " & $ItemChecked & @CRLF)
			_ArrayAdd($Drivers2Install, GUICtrlRead($LabelWorkingFolder, $WorkingFolder) & $ItemText)
			$CheckedCount += 1
			;$Text = ''
		EndIf
	Next

	;_ArrayDisplay($Drivers2Install,'splat')
	_ArrayDelete($Drivers2Install, 0)

	GUICtrlSetData($LabelStatusPath, $CheckedCount & " Drivers are selected for install")
	If $CheckedCount = 0 Then
		GUICtrlSetBkColor($LabelStatusPath, 0xf00606)
		GUICtrlSetColor($LabelStatusPath, 0xFFFFFF)
		Return
	EndIf
	TestNetwork()
	Switch MsgBox(1, "Install drivers", $CheckedCount & " drivers are selected For install." & @CRLF & " Are you sure?")
		Case 1 ; OK
			For $EXEName In $Drivers2Install
				ConsoleWrite(@ScriptLineNumber & " " & $EXEName & " " & FileGetShortName($EXEName) & @CRLF)
				RunWait(@ComSpec & " /c " & FileGetShortName($EXEName)) ;, @SW_HIDE)
			Next
		Case 2 ; Cancel
			Return
	EndSwitch

EndFunc   ;==>InstallChecked
;-------------------------------------------------------------------------------------
Func GetDriverList()
	_Debug(@ScriptLineNumber & " GetDriverList " & $cfg_filename)
	Local $file = FileOpen($cfg_filename, 0)
	; Check if file opened for reading OK
	If $file = -1 Then
		_Debug("LoadCFG: Unable to open file for reading: " & $cfg_filename, 0x10, 5)
		Return
	EndIf

	; Read in the first line to verify the file is of the correct type
	If StringCompare(FileReadLine($file, 1), "Valid for DriversInstall config") <> 0 Then
		_Debug("Not valid for DriversInstall config", 0x20, 5)
		FileClose($file)
		Return
	EndIf

	_GUICtrlTreeView_DeleteAll($HTreeView)
	GUICtrlSetData($LabelStatusPath, "")

	; Read in lines of text until the EOF is reached
	While 1
		Local $LineIn = FileReadLine($file)
		If @error = -1 Then ExitLoop
		If StringInStr($LineIn, "WorkingFolder:") = 1 Then
			$WorkingFolder = StringMid($LineIn, StringInStr($LineIn, ":") + 1)
			GUICtrlSetData($LabelWorkingFolder, $WorkingFolder)
		EndIf
		;	Local $temp[1]

		Local $Handles[1]
		If StringInStr($LineIn, "%:") = 1 Then
			_ArrayAdd($Handles, _GUICtrlTreeView_AddChild($HTreeView, 0, StringMid($LineIn, StringInStr($LineIn, ":") + 1)))
		EndIf

		GUICtrlSetColor(-1, 0x0000C0)
		GUICtrlSetFont(-1, 12)
	WEnd
	FileClose($file)

	TestNetwork()

EndFunc   ;==>GetDriverList
;------------------------------------------------------------------------------------
Func GetSystemInfo()
	_Debug(@ScriptLineNumber & " GetSystemInfo")
	Local $SystemProductName = RegRead("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\BIOS", "SystemProductName")
	Local $OsName = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Intel\MPG", "OsName")
	GUICtrlSetData($LabelSysInfo, "System: " & $OsName & "     " & $SystemProductName)
	;GUICtrlSetData($LabelSysInfo, $SystemProductName)
EndFunc   ;==>GetSystemInfo
;------------------------------------------------------------------------------------

Func TestNetwork()
	_Debug(@ScriptLineNumber & " TestNetwork")
	If FileExists($WorkingFolder) = False Then
		Local $Result = MsgBox(4, "Unable to access working folder:", $WorkingFolder & @CRLF & @CRLF & "Do want to log in?")
		;DriveMapDel($Drive)
		Switch $Result
			Case 6
				DriveMapAdd("Z:", "\\chakotay\temp", 8, "amr\dbkaynox")
			Case Else
				Return
		EndSwitch
	EndIf
	## Tidy Error -> "endfunc" is closing previous "if" on line 225
EndFunc   ;==>TestNetwork
;------------------------------------------------------------------------------------
Func ChooseConfigFile()
	_Debug(@ScriptLineNumber & " ChooseConfigFile")
EndFunc   ;==>ChooseConfigFile
;------------------------------------------------------------------------------------
Func About()
	_Debug(@ScriptLineNumber & " About")
	Local $D = WinGetPos($ProgramName)
	Local $WinPos
	If IsArray($D) = 1 Then
		$WinPos = StringFormat("%s" & @CRLF & "WinPOS: %d  %d " & @CRLF & "WinSize: %d %d " & @CRLF & "Desktop: %d %d ", _
				$MainForm, $D[0], $D[1], $D[2], $D[3], @DesktopWidth, @DesktopHeight)
	Else
		$WinPos = ">>>About ERROR, Check the window name<<<"
	EndIf
	_Debug(@CRLF & $SystemS & @CRLF & $WinPos & @CRLF & "Written by Doug Kaynor", 0x40, 5)
EndFunc   ;==>About
;------------------------------------------------------------------------------------