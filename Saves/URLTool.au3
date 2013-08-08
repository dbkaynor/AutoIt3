#region
#AutoIt3Wrapper_Run_Au3check=y
#AutoIt3Wrapper_Au3Check_Stop_OnWarning=y
#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Tidy_Stop_OnError=y
;#Tidy_Parameters=/gd /sf
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Fileversion=1.0.0.10
#AutoIt3Wrapper_Res_FileVersion_AutoIncrement=Y
#AutoIt3Wrapper_Res_Description=URLtool
#AutoIt3Wrapper_Res_LegalCopyright=GNU-PL
#AutoIt3Wrapper_Res_Comment=A program to test URL response
#AutoIt3Wrapper_Res_Field=Developer|Douglas Kaynor
#AutoIt3Wrapper_Res_LegalCopyright=Copyright ? 2009 Douglas B Kaynor
#AutoIt3Wrapper_Res_Field= AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Res_Field=Compile date|%longdate% %time%
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Run_Before=
#AutoIt3Wrapper_Icon=./icons/small_snake.ico
#endregion

Opt("MustDeclareVars", 1)

If _Singleton(@ScriptName, 1) = 0 Then
	_Debug(@ScriptName & " is already running!", 0x40, 5)
	Exit
EndIf

#include <Array.au3>
#include <Date.au3>
#include <iNet.au3>
#include <Misc.au3>
#include <String.au3>

#include <GUIConstants.au3>
#include <GuiListView.au3>
#include <GUIConstantsEx.au3>

#include <GuiTreeView.au3>
#include <GuiImageList.au3>

#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <ListViewConstants.au3>
#include <StaticConstants.au3>
#include <TreeViewConstants.au3>
#include <WindowsConstants.au3>

#include "_DougFunctions.au3"

DirCreate(@ScriptDir & "\AUXFiles")
Global $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")

Global $tmp = StringSplit(@ScriptName, ".")
Global $Project_filename = @ScriptDir & "\AUXFiles\" & $tmp[1] & ".prj"
Global $LOG_filename = @ScriptDir & "\AUXFiles\" & $tmp[1] & ".log"

Global $FireFoxFullPath = "???"
Global $ChromeFullPath = "???"
Global $IEFullPath = "???"
Global $FireFoxPath = "???"
Global $ChromePath = "???"
Global $IEPath = "???"
Global $SystemS = @ScriptName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & @OSServicePack & @CRLF & @OSType & @CRLF & @OSArch & @IPAddress1
Global $mainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)

#region ### START Koda GUI section ### Form=
Global $MainForm = GUICreate(@ScriptName & "  " & $FileVersion, 1010, 410, 0, 0, $mainFormOptions)
GUISetFont(10, 400, 0, "Courier New")
Global $ButtonTest = GUICtrlCreateButton("Test", 8, 8, 50, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "This will test the URLs.")
Global $CheckStop = GUICtrlCreateCheckbox("Stop", 10, 34, 52, 30)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Stop testing")
Global $ButtonData = GUICtrlCreateButton("Data", 8, 70, 50, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Clicking this will display the addresses")
Global $ButtonSaveProject = GUICtrlCreateButton("Save project", 64, 8, 107, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Save the current settings")
Global $ButtonLoadProject = GUICtrlCreateButton("Load project", 64, 30, 107, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Load saved settings")
Global $ButtonSaveLog = GUICtrlCreateButton("Save log", 72, 87, 80, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Save a log of test results")
Global $ButtonEdit = GUICtrlCreateButton("Edit\View", 64, 55, 80, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Edit or view a file")
Global $ButtonAbout = GUICtrlCreateButton("About", 624, 7, 47, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "About button")
Global $ButtonHelp = GUICtrlCreateButton("Help", 624, 30, 47, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Help button")
Global $ButtonExit = GUICtrlCreateButton("Exit", 624, 55, 47, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Exit button")
Global $CheckPublicIP = GUICtrlCreateCheckbox("Public IP", 496, 40, 90, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Include Public IP")
Global $LabelLoops = GUICtrlCreateLabel("Loops", 276, 7, 50, 20, $SS_SUNKEN)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Number of loop to do")
Global $InputLoops = GUICtrlCreateInput("****", 328, 7, 50, 24, BitOR($ES_CENTER, $ES_AUTOHSCROLL))
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Number of loop to do")
Global $LabelDelay = GUICtrlCreateLabel("Delay", 276, 38, 50, 20, $SS_SUNKEN)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Delay between loops in milliseconds")
Global $InputDelay = GUICtrlCreateInput("500", 328, 38, 50, 24, BitOR($ES_CENTER, $ES_AUTOHSCROLL))
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Delay between loops in milliseconds")
Global $InputURL = GUICtrlCreateInput("WEB Address", 380, 7, 240, 24, BitOR($ES_CENTER, $ES_AUTOHSCROLL))
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "String describing the URL addresses to test")
Global $LabelLocalIP = GUICtrlCreateLabel("Local", 276, 88, 50, 20, $SS_SUNKEN)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "This shows the systems local IP addresses")
Global $InputLocalIP = GUICtrlCreateInput("*****", 328, 88, 280, 24, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "This shows the systems local IP addresses")
Global $LabelPublicIP = GUICtrlCreateLabel("Public", 276, 62, 50, 20, $SS_SUNKEN)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "This shows the systems public IP")
Global $InputPublicIP = GUICtrlCreateInput("****", 328, 62, 280, 24, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "This shows the systems public IP")
Global $CheckPing = GUICtrlCreateCheckbox("Ping", 382, 38, 52, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Toogle ping")
Global $CheckWeb = GUICtrlCreateCheckbox("Web", 437, 38, 52, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Toggle Web")
Global $ButtonFetch = GUICtrlCreateButton("Fetch", 676, 94, 50, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Fetch selected value to address input")
Global $ButtonLaunch = GUICtrlCreateButton("Launch", 180, 86, 60, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Launch the selected URL")

Global $GroupLaunch = GUICtrlCreateGroup("Launch", 173, 4, 90, 81)
GUICtrlSetResizing(-1, 802)
Global $RadioFireFox = GUICtrlCreateRadio("FireFox", 178, 19, 80, 20)
GUICtrlSetResizing(-1, 802)
Global $RadioChrome = GUICtrlCreateRadio("Chrome", 178, 39, 80, 20)
GUICtrlSetResizing(-1, 802)
Global $RadioIE = GUICtrlCreateRadio("IE", 178, 59, 80, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlCreateGroup("", -99, -99, 1, 1)

Global $ButtonAdd = GUICtrlCreateButton("Add", 684, 8, 45, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Add data to list")
Global $ButtonClr = GUICtrlCreateButton("Clr", 684, 28, 45, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Clear the list")
Global $ButtonDel = GUICtrlCreateButton("Del", 684, 48, 45, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Delete checked boxes")
Global $ButtonTog = GUICtrlCreateButton("TOG", 684, 68, 45, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Toggle check boxes")
Global $hTreeView = GUICtrlCreateTreeView(744, 5, 265, 396, BitOR($TVS_HASBUTTONS, $TVS_HASLINES, $TVS_LINESATROOT, $TVS_DISABLEDRAGDROP, $TVS_SHOWSELALWAYS, $TVS_CHECKBOXES, $WS_TABSTOP), $WS_EX_CLIENTEDGE)
Global $hListView = GUICtrlCreateListView("Address|Loop|Ping|Web page title", 8, 120, 734, 280, -1, BitOR($WS_EX_CLIENTEDGE, $LVS_EX_GRIDLINES))
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 0, 250)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 1, 110)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 2, 110)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 3, 250)
GUISetState(@SW_SHOW)
#endregion ### END Koda GUI section ###

GUICtrlSetResizing(-1, 98)
GUISetState(@SW_SHOW)

_Debug("DBGVIEWCLEAR")

Global $HandleArray[1]
Global $DataArray[1]
Global $ToggleState = False

DefaultSettings()

LoadProject("start")

GUISetState(@SW_SHOW)
While 1
	Global $nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $ButtonExit
			Exit
		Case $ButtonHelp
			MsgBox(0, "Help", "Someday maybe", 5)
		Case $ButtonTest
			TestTheURLs()
		Case $ButtonData
			DataToListView()
		Case $ButtonAdd
			AddData()
		Case $ButtonDel
			DelData()
		Case $ButtonClr
			ClrData()
		Case $ButtonTog
			ToggleData()
		Case $ButtonFetch
			Fetch()
		Case $ButtonLaunch
			Launch()
		Case $InputURL
			_Debug("Case $URL")
		Case $ButtonSaveProject
			SaveProject()
		Case $ButtonLoadProject
			LoadProject("Menu")
		Case $ButtonSaveLog
			SaveLog()
		Case $ButtonEdit
			Edit()
		Case $ButtonAbout
			About()
		Case $CheckPublicIP
			_Debug("CheckPublicIP")
			If GUICtrlRead($CheckPublicIP) = $GUI_CHECKED Then GetPublicIp()
		Case $CheckStop
			_Debug("CheckStop")
		Case $CheckPing
			_Debug("CheckPing")
		Case $CheckWeb
			_Debug("CheckWeb")
		Case $LabelPublicIP
			GetPublicIp()
		Case $LabelLocalIP
			GetLocalIp()
	EndSwitch
WEnd
;-----------------------------------------------
;This function copies checked data to ListView
Func DataToListView()
	_Debug("DataToListView")
	GuiDisable("disable")

	_GUICtrlListView_DeleteAllItems(GUICtrlGetHandle($hListView))
	Local $DataArraySave = $DataArray
	TreeViewCheckedToArray() ;$hListView

	_RemoveBlankLines($DataArray)
	_RemoveDuplicateLines($DataArray)
	_ArraySort($DataArray)

	For $X In $DataArray
		If StringLen($X) > 3 Then
			If StringLeft($X, 7) <> 'http://' And StringLeft($X, 8) <> 'https://' Then $X = 'http://' & $X
			$X = StringReplace($X, ':\\', '://', 1)
			_GUICtrlListView_AddItem($hListView, $X)
		EndIf
	Next

	$DataArray = $DataArraySave
	GuiDisable("enable")
EndFunc   ;==>DataToListView

;-----------------------------------------------
;This function puts all of the data from the array into treeview
Func ArrayToTreeView()
	_Debug("ArrayToTreeView")
	ReDim $HandleArray[1]
	_GUICtrlTreeView_DeleteAll($hTreeView)

	_RemoveBlankLines($DataArray)
	_RemoveDuplicateLines($DataArray)
	_ArraySort($DataArray)

	;_ArrayDisplay($DataArray, " ArrayToTreeView 2")

	For $X In $DataArray
		If StringLen($X) > 3 Then
			_ArrayAdd($HandleArray, _GUICtrlTreeView_Add($hTreeView, 0, $X))
		EndIf
	Next
EndFunc   ;==>ArrayToTreeView
;-----------------------------------------------
;This function gets all of the data from the treeview into an array
Func TreeViewToArray()
	_Debug("TreeViewToArray")
	ReDim $DataArray[1]
	For $X = 1 To UBound($HandleArray) - 1
		_ArrayAdd($DataArray, _GUICtrlTreeView_GetText($hTreeView, $HandleArray[$X]))
	Next
	;_ArrayDisplay($HandleArray)
EndFunc   ;==>TreeViewToArray
;-----------------------------------------------
;This function gets the checked data from the treeview into an array
Func TreeViewCheckedToArray()
	_Debug("TreeViewCheckedToArray")
	ReDim $DataArray[1]
	For $X = 1 To UBound($HandleArray) - 1
		Local $AA = _GUICtrlTreeView_GetChecked($hTreeView, $HandleArray[$X])
		Local $BB = _GUICtrlTreeView_GetText($hTreeView, $HandleArray[$X])
		_Debug("Checked  " & $AA & "  " & $BB & "  " & $HandleArray[$X])
		If $AA = True Then
			_ArrayAdd($DataArray, $BB)
		EndIf
	Next
	;_ArrayDisplay($hTreeView)
EndFunc   ;==>TreeViewCheckedToArray
;-----------------------------------------------
;This function gets the un-checked data from the treeview into an array
Func TreeViewUnCheckedToArray()
	_Debug("TreeViewUnCheckedToArray")
	ReDim $DataArray[1]
	For $X = 1 To UBound($HandleArray) - 1
		Local $AA = _GUICtrlTreeView_GetChecked($hTreeView, $HandleArray[$X])
		Local $BB = _GUICtrlTreeView_GetText($hTreeView, $HandleArray[$X])
		_Debug("UNChecked  " & $AA & "  " & $BB & "  " & $HandleArray[$X])
		If $AA = False Then
			_ArrayAdd($DataArray, $BB)
		EndIf
	Next
	;_ArrayDisplay($hTreeView)
EndFunc   ;==>TreeViewUnCheckedToArray
;-----------------------------------------------
Func ToggleData()
	_Debug("ToggleData")
	GuiDisable("disable")

	;_ArrayDisplay($HandleArray)
	$ToggleState = Not $ToggleState

	For $X = 1 To UBound($HandleArray) - 1
		_GUICtrlTreeView_SetChecked($hTreeView, $HandleArray[$X], $ToggleState)
		_Debug(StringFormat("%2u   %6x   %d", $X, $HandleArray[$X], $ToggleState))
	Next

	GuiDisable("enable")
EndFunc   ;==>ToggleData
;-----------------------------------------------
Func Launch()
	Local $T
	For $A = 0 To _GUICtrlListView_GetItemCount($hListView) - 1
		If _GUICtrlListView_GetItemSelected($hListView, $A) Then
			$T = _GUICtrlListView_GetItemText($hListView, $A, 0)
		EndIf
	Next

	If StringLen($T) < 3 Then Return

	If (GUICtrlRead($RadioFireFox) = $GUI_CHECKED) Then
		_Debug("Launch Firefox " & $T & "  " & @ScriptLineNumber & "  " & $FireFoxPath)
		ShellExecute($FireFoxPath, $T)
	EndIf
	If (GUICtrlRead($RadioChrome) = $GUI_CHECKED) Then
		_Debug("Launch Chrome " & $T & "  " & @ScriptLineNumber & "  " & $ChromePath)
		ShellExecute($ChromePath, $T)
	EndIf
	If (GUICtrlRead($RadioIE) = $GUI_CHECKED) Then
		_Debug("Launch IE " & $T & "  " & @ScriptLineNumber & "  " & $IEPath)
		ShellExecute($IEPath, $T)
	EndIf
EndFunc   ;==>Launch
;-----------------------------------------------
Func Fetch()
	Local $Data[1]
	Local $hItem = _GUICtrlTreeView_GetFirstItem($hTreeView)
	_ArrayAdd($Data, $hItem)
	While $hItem <> 0
		$hItem = _GUICtrlTreeView_GetNext($hTreeView, $hItem)
		If $hItem <> 0 Then _ArrayAdd($Data, $hItem)
	WEnd

	_ArrayDelete($Data, 0)

	For $X = 0 To UBound($Data) - 1
		If _GUICtrlTreeView_GetSelected($hTreeView, $Data[$X]) == True Then
			GUICtrlSetData($InputURL, _GUICtrlTreeView_GetText($hTreeView, $Data[$X]))
		EndIf
	Next
EndFunc   ;==>Fetch
;-----------------------------------------------
;This function parses the InputBox String and creates the addresses
Func AddData()
	_Debug("AddData")
	GuiDisable("disable")
	$ToggleState = False

	;TreeViewToArray()
	_GUICtrlListView_DeleteAllItems($hListView)

	Local $D = GUICtrlRead($InputURL)
	ConsoleWrite("help>>>" & $D & "<<<   " & @ScriptLineNumber & @CRLF)
	If StringInStr($D, '.') > 0 Then
		$D = StringReplace($D, ':\\', '://', 1)
		If StringLeft($D, 7) <> 'http://' And StringLeft($D, 8) <> 'https://' Then $D = 'http://' & $D
		_ArrayAdd($DataArray, $D)
	EndIf

	ArrayToTreeView()
	GuiDisable("enable")
EndFunc   ;==>AddData
;-----------------------------------------------
Func DelData()
	_Debug("DelData")
	GuiDisable("disable")
	$ToggleState = False

	TreeViewUnCheckedToArray()
	_GUICtrlTreeView_DeleteAll($hTreeView)
	;_ArrayDisplay($DataArray)
	ArrayToTreeView()

	GuiDisable("enable")
EndFunc   ;==>DelData
;-----------------------------------------------
Func ClrData()
	_Debug("ClrData")
	$ToggleState = False

	ReDim $DataArray[1]
	ReDim $HandleArray[1]
	_GUICtrlTreeView_DeleteAll($hTreeView)
	_GUICtrlListView_DeleteAllItems($hListView)
	Global $HandleArray[1]

EndFunc   ;==>ClrData
;-----------------------------------------------
Func DefaultSettings()
	_Debug("DefaultSettings")
	$ToggleState = False

	GUICtrlSetData($InputURL, "127.0.0.1;")
	GUICtrlSetData($InputLoops, 4)
	GUICtrlSetData($InputDelay, 500)
	GUICtrlSetState($CheckStop, $GUI_UNCHECKED)
	GUICtrlSetState($CheckPing, $GUI_CHECKED)
	GUICtrlSetState($CheckWeb, $GUI_CHECKED)
	GUICtrlSetState($CheckPublicIP, $GUI_CHECKED)
EndFunc   ;==>DefaultSettings
;-----------------------------------------------
Func GuiDisable($choice) ;@SW_ENABLE @SW_disble
	_Debug("GuiDisable  " & $choice)
	Global $LastState
	Local $setting

	If $choice = "Enable" Then
		$setting = $GUI_ENABLE
	ElseIf $choice = "Disable" Then
		$setting = $GUI_DISABLE
	ElseIf $choice = "Toggle" Then
		If $LastState = $GUI_DISABLE Then
			$setting = $GUI_ENABLE
		Else
			$setting = $GUI_DISABLE
		EndIf
	Else
		_Debug("Invalid choice at GuiDisable" & $choice, 0x40)
	EndIf

	GUICtrlSetState($ButtonTest, $setting)
	GUICtrlSetState($ButtonData, $setting)
	GUICtrlSetState($ButtonLoadProject, $setting)
	GUICtrlSetState($ButtonSaveProject, $setting)
	GUICtrlSetState($ButtonSaveLog, $setting)
	GUICtrlSetState($ButtonEdit, $setting)
	GUICtrlSetState($ButtonFetch, $setting)
	GUICtrlSetState($ButtonLaunch, $setting)
	GUICtrlSetState($ButtonAdd, $setting)
	GUICtrlSetState($ButtonClr, $setting)
	GUICtrlSetState($ButtonDel, $setting)
	GUICtrlSetState($ButtonTog, $setting)
	GUICtrlSetState($RadioFireFox, $setting)
	GUICtrlSetState($RadioChrome, $setting)
	GUICtrlSetState($RadioIE, $setting)

EndFunc   ;==>GuiDisable
;-----------------------------------------------
Func GetPublicIp()
	_Debug("GetPublicIp")
	GUICtrlSetData($InputPublicIP, "Getting Public IP") ; Getting public IP-1
	GUICtrlSetColor($InputPublicIP, 0xff0000)
	Local $T = _GetIP()
	GUICtrlSetColor($InputPublicIP, 0x000000)
	If testip($T) = "ERROR0" Then
		GUICtrlSetData($InputPublicIP, $T)
	Else
		GUICtrlSetData($InputPublicIP, "**** Failed")
	EndIf

EndFunc   ;==>GetPublicIp
;-----------------------------------------------
Func GetLocalIp()
	_Debug("Getting the local IP address")
	Local $T = ''
	GUICtrlSetColor($InputPublicIP, 0xff0000)
	GUICtrlSetData($InputLocalIP, "GetLocalIp")
	If @IPAddress1 <> '0.0.0.0' Then $T = $T & ';' & IPPad(@IPAddress1)
	If @IPAddress2 <> '0.0.0.0' Then $T = $T & ';' & IPPad(@IPAddress2)
	If @IPAddress3 <> '0.0.0.0' Then $T = $T & ';' & IPPad(@IPAddress3)
	If @IPAddress4 <> '0.0.0.0' Then $T = $T & ';' & IPPad(@IPAddress4)
	GUICtrlSetData($InputLocalIP, $T)
	GUICtrlSetColor($InputPublicIP, 0x000000)
EndFunc   ;==>GetLocalIp
;-----------------------------------------------
Func TestTheURLs()
	_Debug("TestTheURLs()")
	DataToListView()
	GuiDisable("disable")
	GUICtrlSetState($CheckStop, $GUI_UNCHECKED)

	Local $MaxLoops = GUICtrlRead($InputLoops)
	Local $Loop = 0
	While $Loop < $MaxLoops Or $MaxLoops <= 0
		$MaxLoops = GUICtrlRead($InputLoops)
		$Loop = $Loop + 1

		If GUICtrlRead($CheckStop) = $GUI_CHECKED Then ExitLoop

		;now do the testing
		For $A = 0 To _GUICtrlListView_GetItemCount($hListView) - 1
			If GUICtrlRead($CheckStop) = $GUI_CHECKED Then ExitLoop

			Local $tmp = _GUICtrlListView_GetItemText($hListView, $A, 1)
			$tmp = _GUICtrlListView_GetItemText($hListView, $A, 1)
			$tmp = $tmp + 1
			_GUICtrlListView_AddSubItem($hListView, $A, $tmp, 1)
			_GUICtrlListView_AddSubItem($hListView, $A, StringFormat("%d of %d", $Loop, $MaxLoops), 1)

			Local $B = _GUICtrlListView_GetItemText($hListView, $A)
			_Debug("TestTheURLs  " & $B)
			Local $E = StringReplace($B, "http://", '')
			Local $F = StringSplit($E, ":")

			;$CheckPing
			If GUICtrlRead($CheckPing) = $GUI_UNCHECKED Then
				_GUICtrlListView_AddSubItem($hListView, $A, "Not tested", 2) ; result
			Else
				For $X = 0 To 3 ; try four time before giving up
					If GUICtrlRead($CheckStop) = $GUI_CHECKED Then ExitLoop
					_GUICtrlListView_AddSubItem($hListView, $A, "Working....", 2) ; result
					Local $Value = Ping($F[1], 1000)
					If $Value = 0 Then
						_GUICtrlListView_AddSubItem($hListView, $A, 'No response', 2)
					Else
						_GUICtrlListView_AddSubItem($hListView, $A, $Value & " ms", 2)
						ExitLoop
					EndIf
				Next
			EndIf

			If GUICtrlRead($CheckStop) = $GUI_CHECKED Then ExitLoop

			;$CheckWeb
			If GUICtrlRead($CheckWeb) = $GUI_UNCHECKED Then
				_GUICtrlListView_AddSubItem($hListView, $A, "Not tested", 3) ; result
			Else
				_GUICtrlListView_AddSubItem($hListView, $A, "Working....", 3) ; result
				If GUICtrlRead($CheckStop) = $GUI_CHECKED Then ExitLoop
				Local $C = _INetGetSource($B)
				Local $D = StringSplit($C, @CRLF)
				Local $TS = 'No response'

				If $D[0] > 1 Then ;dbk
					Local $Start = False

					For $E In $D
						If StringInStr($E, '<title>') Then $Start = True
						If $Start = True Then $TS = $TS & $E
						If StringInStr($E, '</title>') Then $Start = False
					Next

					$TS = StringRegExpReplace($TS, "[\t]", " ")
					$TS = StringReplace($TS, "&reg;", " ")

					ConsoleWrite('1>>>>' & $TS & '<<<<' & @CRLF)
					Local $Begin = StringInStr($TS, "<title>") + StringLen("<title>")
					Local $Count = StringInStr($TS, "</title>") - $Begin

					$TS = StringMid($TS, $Begin, $Count)
					$TS = StringStripWS($TS, 7)
					ConsoleWrite('2>>>>' & $TS & '<<<<' & @CRLF)
					_Debug('>>' & $TS & '<<')
				EndIf
				_GUICtrlListView_AddSubItem($hListView, $A, $TS, 3) ; result
			EndIf

		Next
		Sleep(GUICtrlRead($InputDelay))
	WEnd

	GuiDisable("enable")
EndFunc   ;==>TestTheURLs
;-----------------------------------------------
Func SaveLog()
	_Debug("SaveLog")
	$LOG_filename = FileSaveDialog("Save log file", @ScriptDir & "\AUXFiles\", _
			"URLTool logs (U*.log)|All logs (*.log)|All files (*.*)", 18, @ScriptDir & "\AUXFiles\URLTool.log")

	Local $file = FileOpen($LOG_filename, 2)
	; Check if file opened for writing OK
	If $file = -1 Then
		_Debug("SaveLog: Unable to open file for writing: " & $LOG_filename, 0x10, 5)
		Return
	EndIf

	_Debug("SaveLog  " & $LOG_filename)
	FileWriteLine($file, "Log file for " & @ScriptName & "  " & _DateTimeFormat(_NowCalc(), 0))

	;FileWriteLine($file, "IPAddress: " & GUICtrlRead($InputURL))
	FileWriteLine($file, "Help: 1 is enabled, 4 is disabled for checkboxes")
	FileWriteLine($file, "CheckPublicIP:" & GUICtrlRead($CheckPublicIP))
	FileWriteLine($file, "InputLoops:" & GUICtrlRead($InputLoops))
	FileWriteLine($file, "InputDelay:" & GUICtrlRead($InputDelay))
	FileWriteLine($file, "RadioFireFox:" & (GUICtrlRead($RadioFireFox)))
	FileWriteLine($file, "RadioChrome:" & (GUICtrlRead($RadioChrome)))
	FileWriteLine($file, "RadioIE:" & (GUICtrlRead($RadioIE)))
	FileWriteLine($file, "FireFoxPath: " & $FireFoxFullPath)
	FileWriteLine($file, "ChromePath: " & $ChromeFullPath)
	FileWriteLine($file, "IEPath: " & $IEFullPath)
	FileWriteLine($file, "LocalIP: " & GUICtrlRead($InputLocalIP))
	FileWriteLine($file, "PublicIP: " & GUICtrlRead($InputPublicIP))
	FileWriteLine($file, "CheckPing:" & GUICtrlRead($CheckPing))
	FileWriteLine($file, "CheckWeb:" & GUICtrlRead($CheckWeb))
	FileWriteLine($file, "InputIPAddressEdit:" & GUICtrlRead($InputURL))
	FileWriteLine($file, '')
	FileWriteLine($file, StringFormat("%3s %-30s %5s %7s  %-10s", "Cnt", "Address", "Loops", "Ping", "Page title)"))

	For $X = 0 To _GUICtrlListView_GetItemCount($hListView) - 1
		Local $T = _GUICtrlListView_GetItemTextString($hListView, $X);, 0)
		Local $array = StringSplit($T, "|")
		Local $NewString = StringFormat("%3d %-30s %5d %7s  %-10s  ", $X + 1, $array[1], $array[2], $array[3], $array[4])
		_Debug("SaveLog:  " & $NewString)
		FileWriteLine($file, $NewString)
	Next

	FileClose($file)
EndFunc   ;==>SaveLog
;-----------------------------------------------
Func SaveProject()
	_Debug("SaveProject")
	$Project_filename = FileSaveDialog("Save project file", @ScriptDir & "\AUXFiles\", _
			"URLTool projects (U*.prj)|All projects (*.prj)|All files (*.*)", 18, @ScriptDir & "\AUXFiles\URLTool.prj")

	Local $file = FileOpen($Project_filename, 2)
	; Check if file opened for writing OK
	If $file = -1 Then
		_Debug("SaveProject: Unable to open file for writing: " & $Project_filename, 0x10, 5)
		Return
	EndIf
	_Debug("SaveProject  " & $Project_filename)
	FileWriteLine($file, "Valid for URLTool project")
	FileWriteLine($file, "Project file for " & @ScriptName & "  " & _DateTimeFormat(_NowCalc(), 0))
	FileWriteLine($file, "Help 1 is enabled, 4 is disabled for checkboxes")
	FileWriteLine($file, "CheckPublicIP:" & GUICtrlRead($CheckPublicIP))
	FileWriteLine($file, "CheckPing:" & GUICtrlRead($CheckPing))
	FileWriteLine($file, "CheckWeb:" & GUICtrlRead($CheckWeb))
	FileWriteLine($file, "InputLoops:" & GUICtrlRead($InputLoops))
	FileWriteLine($file, "InputDelay:" & GUICtrlRead($InputDelay))
	FileWriteLine($file, "RadioFireFox:" & (GUICtrlRead($RadioFireFox)))
	FileWriteLine($file, "RadioChrome:" & (GUICtrlRead($RadioChrome)))
	FileWriteLine($file, "RadioIE:" & (GUICtrlRead($RadioIE)))
	FileWriteLine($file, "FireFoxPath: " & $FireFoxFullPath)
	FileWriteLine($file, "ChromePath: " & $ChromeFullPath)
	FileWriteLine($file, "IEPath: " & $IEFullPath)
	FileWriteLine($file, "InputIPAddressEdit:" & GUICtrlRead($InputURL))

	For $X = 1 To UBound($HandleArray) - 1
		Local $Y = _GUICtrlTreeView_GetText($hTreeView, $HandleArray[$X])
		Local $Z = StringInStr($Y, "*")
		ConsoleWrite('@@ _Debug(' & @ScriptLineNumber & '): ' & $X & "   " & $Y & "   " & $Z & @CRLF) ;### Debug Console

		If $Z = 0 Then
			FileWriteLine($file, "InputIPAddress:" & $Y)
		EndIf
	Next

	FileClose($file)
EndFunc   ;==>SaveProject
;-----------------------------------------------
;This loads the project file but into the tree control but not into the list
Func LoadProject($type)
	_Debug("LoadProject  " & $type)

	If StringCompare($type, "menu") = 0 Then
		$Project_filename = FileOpenDialog("Load project file", @ScriptDir & "\AUXFiles\", _
				"URLTool projects (U*.prj)|All projects (*.prj)|All files (*.*)", 18, @ScriptDir & "\AUXFiles\URLTool.prj")
	EndIf

	Local $file = FileOpen($Project_filename, 0)
	; Check if file opened for reading OK
	If $file = -1 Then
		_Debug("LoadProject: Unable to open file for reading: " & $Project_filename, 0x10, 5)
		Return
	EndIf

	_Debug("LoadProject   " & $Project_filename)
	; Read in the first line to verify the file is of the correct type
	If StringCompare(FileReadLine($file, 1), "Valid for URLTool project") <> 0 Then
		_Debug("Not a valid project file for URLTool", 0x20, 5)
		FileClose($file)
		Return
	EndIf

	ClrData()
	GUICtrlSetData($InputURL, '')

	; Read in lines of text until the EOF is reached
	While 1
		Local $LineIn = FileReadLine($file)
		If @error = -1 Then ExitLoop
		_Debug("LoadProject   " & $LineIn)
		If StringInStr($LineIn, "InputLoops:") Then GUICtrlSetData($InputLoops, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
		If StringInStr($LineIn, "InputDelay:") Then GUICtrlSetData($InputDelay, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
		If StringInStr($LineIn, "RadioFireFox:") Then GUICtrlSetState($RadioFireFox, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
		If StringInStr($LineIn, "RadioChrome:") Then GUICtrlSetState($RadioChrome, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
		If StringInStr($LineIn, "RadioIE:") Then GUICtrlSetState($RadioIE, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
		If StringInStr($LineIn, "FireFoxPath:") Then $FireFoxFullPath = StringMid($LineIn, StringInStr($LineIn, ":") + 1)
		If StringInStr($LineIn, "ChromePath:") Then $ChromeFullPath = StringMid($LineIn, StringInStr($LineIn, ":") + 1)
		If StringInStr($LineIn, "IEPath:") Then $IEFullPath = StringMid($LineIn, StringInStr($LineIn, ":") + 1)
		If StringInStr($LineIn, "CheckPublicIP:") Then GUICtrlSetState($CheckPublicIP, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
		If StringInStr($LineIn, "CheckPing:") Then GUICtrlSetState($CheckPing, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
		If StringInStr($LineIn, "CheckWeb:") Then GUICtrlSetState($CheckWeb, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
		If StringInStr($LineIn, "InputIPAddressEdit:") Then GUICtrlSetData($InputURL, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
		If StringInStr($LineIn, "InputIPAddress:") Then
			Local $D = StringSplit($LineIn, "InputIPAddress:", 3)
			$D[1] = StringReplace($D[1], ':\\', '://', 1)
			If StringLeft($D[1], 7) <> 'http://' And StringLeft($D[1], 8) <> 'https://' Then $D[1] = 'http://' & $D[1]
			_ArrayAdd($DataArray, $D[1])
		EndIf
	WEnd

	FileClose($file)

	Local $A = StringSplit($FireFoxFullPath, ';')
	For $B In $A
		$B = StringReplace(StringStripWS($B, 3), '"', "")
		If FileExists($B) = 1 Then
			$FireFoxPath = $B
		EndIf
	Next

	$A = StringSplit($ChromeFullPath, ';')
	For $B In $A
		$B = StringReplace(StringStripWS($B, 3), '"', "")
		If FileExists($B) = 1 Then
			$ChromePath = $B
			ExitLoop
		EndIf
	Next

	$A = StringSplit($IEFullPath, ';')
	For $B In $A
		$B = StringReplace(StringStripWS($B, 3), '"', "")
		If FileExists($B) = 1 Then
			$IEPath = $B
			ExitLoop
		EndIf
	Next

	ConsoleWrite('F:::>' & $FireFoxPath & @CRLF)
	ConsoleWrite('C:::>' & $ChromePath & @CRLF)
	ConsoleWrite('C:::>' & $IEPath & @CRLF)

	GetLocalIp()
	If GUICtrlRead($CheckPublicIP) = $GUI_CHECKED Then GetPublicIp()

	AddData()
	;_ArrayDisplay($DataArray, "LoadProject")

EndFunc   ;==>LoadProject
;-----------------------------------------------
Func Edit()
	_Debug("Edit")
	Local $Filename = FileOpenDialog("View a file", @ScriptDir & "\AUXFiles\", _
			"All (*.*)", 1, @ScriptDir & "\AUXFiles\URLTool.prj")

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

	_Debug("Edit  " & $Filename)
	ShellExecuteWait($editor, $Filename)

EndFunc   ;==>Edit
;-----------------------------------------------
Func About()
	Local $D = WinGetPos(@ScriptName)
	Local $WinPos
	If IsArray($D) = 1 Then
		$WinPos = StringFormat("%s" & @CRLF & "WinPOS: %d  %d " & @CRLF & "WinSize: %d %d " & @CRLF & "Desktop: %d %d ", _
				$MainForm, $D[0], $D[1], $D[2], $D[3], @DesktopWidth, @DesktopHeight)
	Else
		$WinPos = ">>>About ERROR, Check the window name<<<"
	EndIf
	_Debug(@CRLF & $SystemS & @CRLF & $WinPos & @CRLF & "Written by Doug Kaynor because I wanted to!", 0x40, 5)
EndFunc   ;==>About
;-----------------------------------------------