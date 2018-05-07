#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_icon=../icons/eagle.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Comment=A program to ping network addresses
#AutoIt3Wrapper_Res_Description=Pingtool
#AutoIt3Wrapper_Res_Fileversion=2.0.3.25
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=Y
#AutoIt3Wrapper_Res_LegalCopyright=Copyright 2011 Douglas B Kaynor
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Field=Developer|Douglas Kaynor
#AutoIt3Wrapper_Res_Field=AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Res_Field=Compile date|%longdate% %time%
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=y
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Run_Tidy=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#Region
;#Tidy_Parameters=/gd /sf
#EndRegion

#cs
	Fixed:
	This area is used to store things todo, bugs, and other notes

	Todo:
	Make DEL and CLR buttons harder to hit or prompt before doing
	Handle duplicates in lists
	Work on 1024x768
	Check for window position and fit to screen
#ce

Opt("MustDeclareVars", 1)

If _Singleton(@ScriptName, 1) = 0 Then
	_Debug(@ScriptName & " is already running!" & @CRLF, 0x40, 5)
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

#include <Process.au3>

#include "_DougFunctions.au3"

DirCreate(@ScriptDir & "\AUXFiles")
Global $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")

;Global $version = " Version: 2.0.0.4 "


Global $tmp = StringSplit(@ScriptName, ".")
Global $ProgramName = $tmp[1]
Global $Project_filename = @ScriptDir & "\AUXFiles\" & $ProgramName & ".prj"
Global $LOG_filename = @ScriptDir & "\AUXFiles\" & $ProgramName & ".log"

Global $SystemS = @ScriptName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & @OSServicePack & @CRLF & @OSType & @CRLF & @OSArch & @IPAddress1
Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)

Global $MainForm = GUICreate(@ScriptName & " " & $FileVersion, 1000, 420, 10, 10, $MainFormOptions)
GUISetFont(10, 400, 0, "Courier New")
;GUICtrlSetResizing(-1, 256 + 512)

GUISetHelp("notepad", $MainForm) ; Need a help file to call here

Global $ButtonPing = GUICtrlCreateButton("Ping", 10, 10, 50, 20)
GUICtrlSetTip(-1, "This will do the pings. (What did you expect?)")
GUICtrlSetResizing(-1, 802)

Global $CheckStop = GUICtrlCreateCheckbox("Stop", 10, 30, 52, 30)
GUICtrlSetTip(-1, "Stop pinging")
GUICtrlSetResizing(-1, 802)

Global $ButtonData = GUICtrlCreateButton("Data", 10, 70, 50, 20)
GUICtrlSetTip(-1, "Clicking this will display the addresses")
GUICtrlSetResizing(-1, 802)

Global $ButtonSaveProject = GUICtrlCreateButton("Save project", 60, 10, 110, 20)
GUICtrlSetTip(-1, "Save the current settings")
GUICtrlSetResizing(-1, 802)

Global $ButtonLoadProject = GUICtrlCreateButton("Load project", 60, 30, 110, 20)
GUICtrlSetTip(-1, "Load saved settings")
GUICtrlSetResizing(-1, 802)

Global $ButtonLoadDefaults = GUICtrlCreateButton("Load defaults", 60, 50, 110, 20)
GUICtrlSetTip(-1, "Load default settings")
GUICtrlSetResizing(-1, 802)

Global $ButtonSaveLog = GUICtrlCreateButton("Save log", 60, 70, 110, 20)
GUICtrlSetTip(-1, "Save a log of test results")
GUICtrlSetResizing(-1, 802)

Global $ButtonAbout = GUICtrlCreateButton("About", 170, 10, 55, 20)
GUICtrlSetTip(-1, "About button")
GUICtrlSetResizing(-1, 802)

Global $ButtonHelp = GUICtrlCreateButton("Help", 170, 30, 55, 20)
GUICtrlSetTip(-1, "Help button")
GUICtrlSetResizing(-1, 802)

Global $ButtonEdit = GUICtrlCreateButton("Edit", 170, 50, 55, 20)
GUICtrlSetTip(-1, "Edit or view a file")
GUICtrlSetResizing(-1, 802)

Global $ButtonExit = GUICtrlCreateButton("Exit", 170, 70, 55, 20)
GUICtrlSetTip(-1, "Exit button")
GUICtrlSetResizing(-1, 802)

Global $CheckDNS = GUICtrlCreateCheckbox("DNS", 230, 5, 50, 30)
GUICtrlSetTip(-1, "Display DNS results of alive hosts")
GUICtrlSetResizing(-1, 802)

Global $CheckClass = GUICtrlCreateCheckbox("Class", 230, 30, 60, 30)
GUICtrlSetTip(-1, "Display IP Class")
GUICtrlSetResizing(-1, 802)

Global $CheckPublicIP = GUICtrlCreateCheckbox("Pub", 230, 55, 50, 22)
GUICtrlSetTip(-1, "Include Public IP")
GUICtrlSetResizing(-1, 802)

Global $CheckMAC = GUICtrlCreateCheckbox("MAC", 230, 80, 50, 22)
GUICtrlSetTip(-1, "Display MAC")
GUICtrlSetResizing(-1, 802)

Global $LabelTimeout = GUICtrlCreateLabel("Timeout", 300, 10, 60, 20, $SS_SUNKEN)
GUICtrlSetTip(-1, "Timeout value in milliseconds")
GUICtrlSetResizing(-1, 802)

Global $InputTimeout = GUICtrlCreateInput("****", 365, 10, 50, 20, BitOR($ES_CENTER, $ES_AUTOHSCROLL))
GUICtrlSetTip(-1, "Timeout value in milliseconds")
GUICtrlSetResizing(-1, 802)

Global $LabelLoops = GUICtrlCreateLabel("Loops", 300, 40, 60, 20, $SS_SUNKEN)
GUICtrlSetTip(-1, "Number of loop to do")
GUICtrlSetResizing(-1, 802)

Global $InputLoops = GUICtrlCreateInput("****", 365, 40, 50, 20, BitOR($ES_CENTER, $ES_AUTOHSCROLL))
GUICtrlSetTip(-1, "Number of loop to do")
GUICtrlSetResizing(-1, 802)

Global $LabelDelay = GUICtrlCreateLabel("Delay", 300, 70, 60, 20, $SS_SUNKEN)
GUICtrlSetTip(-1, "Delay between loops in milliseconds")
GUICtrlSetResizing(-1, 802)

Global $InputDelay = GUICtrlCreateInput("100", 365, 70, 50, 20, BitOR($ES_CENTER, $ES_AUTOHSCROLL))
GUICtrlSetTip(-1, "Delay between loops in milliseconds")
GUICtrlSetResizing(-1, 802)

Global $InputIPAddress = GUICtrlCreateInput("InputIPAddress", 430, 10, 260, 20, BitOR($ES_CENTER, $ES_AUTOHSCROLL))
GUICtrlSetTip(-1, "String describing the addresses to test")
GUICtrlSetResizing(-1, 802)

Global $LabelLocalIP = GUICtrlCreateLabel("Local", 430, 35, 55, 20, $SS_SUNKEN)
GUICtrlSetTip(-1, "This is the systems local IP address")
GUICtrlSetResizing(-1, 802)

Global $InputLocalIP = GUICtrlCreateInput("*****", 490, 35, 200, 20, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetTip(-1, "This is the systems local IP address")
GUICtrlSetResizing(-1, 802)

Global $LabelPublicIP = GUICtrlCreateLabel("Public", 430, 60, 55, 20, $SS_SUNKEN)
GUICtrlSetTip(-1, "This is the systems public IP")
GUICtrlSetResizing(-1, 802)

Global $InputPublicIP = GUICtrlCreateInput("****", 490, 60, 200, 20, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetTip(-1, "This is the systems public IP")
GUICtrlSetResizing(-1, 802)

Global $ButtonAddLocal = GUICtrlCreateButton("Add", 430, 85, 35, 20)
GUICtrlSetTip(-1, "Add local data to list")
GUICtrlSetResizing(-1, 802)

Global $ButtonFetch = GUICtrlCreateButton("Fetch", 500, 85, 50, 20)
GUICtrlSetTip(-1, "Fetch selected value to address input")
GUICtrlSetResizing(-1, 802)

Global $ButtonAdd = GUICtrlCreateButton("Add", 700, 5, 45, 20)
GUICtrlSetTip(-1, "Add data to list")
GUICtrlSetResizing(-1, 802)

Global $ButtonClr = GUICtrlCreateButton("Clr", 700, 25, 45, 20)
GUICtrlSetTip($ButtonClr, "Clear the list")
GUICtrlSetResizing(-1, 802)

Global $ButtonDel = GUICtrlCreateButton("Del", 700, 45, 45, 20)
GUICtrlSetTip(-1, "Delete checked boxes")
GUICtrlSetResizing(-1, 802)

Global $ButtonTog = GUICtrlCreateButton("TOG", 700, 65, 45, 20)
GUICtrlSetTip(-1, "Toggle check boxes")
GUICtrlSetResizing(-1, 802)

Global $ButtonTest = GUICtrlCreateButton("Test", 700, 85, 45, 20)
GUICtrlSetTip(-1, "Test")
GUICtrlSetResizing(-1, 802)

Global $hTreeView = GUICtrlCreateTreeView(800, 8, 190, 80, BitOR($TVS_DISABLEDRAGDROP, $TVS_SHOWSELALWAYS, $TVS_CHECKBOXES, $WS_GROUP, $WS_TABSTOP), $WS_EX_CLIENTEDGE)
GUICtrlSetTip(-1, "Select the addresses to ping")
GUICtrlSetResizing(-1, 4)

Const $Address = 0
Const $Status = 1
Const $IPClass = 2
Const $ResponseTime = 3
Const $Worsttime = 4
Const $Loop = 5
Const $Alive = 6
Const $Dead = 7
Const $NameError = 8
Const $MAC = 9

Global $ListView = GUICtrlCreateListView("Address|Status|Class|MS|W-MS|Loop|Alive|Dead|Name\Error|MAC", 8, 110, 980, 300, -1, BitOR($WS_EX_CLIENTEDGE, $LVS_EX_GRIDLINES))
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $Address, 130) ; Address
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $Status, 65) ; Status
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $IPClass, 90) ; IP Class
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $ResponseTime, 50) ; Ms response time
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $Worsttime, 50) ; Worst time
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $Loop, 90) ; Loop
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $Alive, 60) ; Alive
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $Dead, 60) ; Dead
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $NameError, 155) ; Name\Error
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $MAC, 160) ; MAC
GUICtrlSetTip(-1, "Display the test results")
GUICtrlSetResizing(-1, 98)
_GUICtrlListView_SetBkColor($ListView, $CLR_WHITE)
_GUICtrlListView_SetTextColor($ListView, $CLR_BLACK)
_GUICtrlListView_SetTextBkColor($ListView, $CLR_Silver)

_Debug("DBGVIEWCLEAR")
_Debug(@CRLF)

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
		Case $ButtonPing
			PingTheHosts()
		Case $ButtonData
			DataToListView()
		Case $ButtonAddLocal
			AddLocalData()
		Case $ButtonAdd
			AddData()
		Case $ButtonDel
			DelData()
		Case $ButtonClr
			ClrData()
		Case $ButtonTog
			ToggleData()
		Case $ButtonTest
			GetMAC("192.168.0.1")
		Case $ButtonFetch
			Fetch()
		Case $InputIPAddress
			_Debug(@ScriptLineNumber & " Case $InputIPAddress" & @CRLF)
		Case $ButtonSaveProject
			SaveProject()
		Case $ButtonLoadProject
			LoadProject("Menu")
		Case $ButtonLoadDefaults
			DefaultSettings()
		Case $ButtonSaveLog
			SaveLog()
		Case $ButtonEdit
			Edit()
		Case $ButtonAbout
			About($MainForm)
		Case $CheckClass
			_Debug(@ScriptLineNumber & " CheckClass" & @CRLF)
		Case $CheckDNS
			_Debug(@ScriptLineNumber & " CheckDNS" & @CRLF)
		Case $CheckPublicIP
			_Debug(@ScriptLineNumber & " CheckPublicIP" & @CRLF)
			If GUICtrlRead($CheckPublicIP) = $GUI_CHECKED Then GetPublicIp()
		Case $CheckMAC
			_Debug(@ScriptLineNumber & " CheckMAC" & @CRLF)
		Case $CheckStop
			_Debug(@ScriptLineNumber & " CheckStop" & @CRLF)
		Case $LabelPublicIP
			GetPublicIp()
		Case $LabelLocalIP
			GetLocalIp()
	EndSwitch
WEnd
;-----------------------------------------------
;This function copies checked data to ListView
Func DataToListView()
	_Debug(@ScriptLineNumber & " DataToListView" & @CRLF)
	GuiDisable("disable")

	_GUICtrlListView_DeleteAllItems(GUICtrlGetHandle($ListView))
	Local $DataArraySave = $DataArray
	TreeViewCheckedToArray() ;$ListView

	_RemoveBlankLines($DataArray)
	_ArrayUnique($DataArray)
	;_RemoveDuplicateLines($DataArray)
	_ArraySort($DataArray)

	For $X = 0 To UBound($DataArray) - 1
		Local $T = StringReplace($DataArray[$X], '* ', '')
		If StringInStr(_TestIP($T), "ERROR0", 2) Then _GUICtrlListView_AddItem($ListView, _IPUnPad($T))
	Next

	$DataArray = $DataArraySave
	GuiDisable("enable")
EndFunc   ;==>DataToListView

;-----------------------------------------------
;This function puts all of the data from the array into treeview
Func ArrayToTreeView()
	_Debug(@ScriptLineNumber & " ArrayToTreeView" & @CRLF)
	ReDim $HandleArray[1]
	_GUICtrlTreeView_DeleteAll($hTreeView)
	_RemoveBlankLines($DataArray)
	_ArrayUnique($DataArray)
	;_RemoveDuplicateLines($DataArray)
	_ArraySort($DataArray)

	;_ArrayDisplay($DataArray, " ArrayToTreeView")

	For $X In $DataArray
		If _TestIP($X) = "ERROR0" Then
			_ArrayAdd($HandleArray, _GUICtrlTreeView_Add($hTreeView, 0, $X))
		EndIf
	Next
EndFunc   ;==>ArrayToTreeView
;-----------------------------------------------
Func Fetch()
	For $X = 1 To UBound($HandleArray) - 1
		If _GUICtrlTreeView_GetSelected($hTreeView, $HandleArray[$X]) == True Then
			GUICtrlSetData($InputIPAddress, _GUICtrlTreeView_GetText($hTreeView, $HandleArray[$X]))
		EndIf
	Next
EndFunc   ;==>Fetch
;-----------------------------------------------
;This function gets all of the data from the treeview into an array
Func TreeViewToArray()
	_Debug(@ScriptLineNumber & " TreeViewToArray" & @CRLF)
	ReDim $DataArray[1]
	For $X = 1 To UBound($HandleArray) - 1
		_ArrayAdd($DataArray, _GUICtrlTreeView_GetText($hTreeView, $HandleArray[$X]))
	Next
	;_ArrayDisplay($HandleArray)
EndFunc   ;==>TreeViewToArray
;-----------------------------------------------
;This function gets the checked data from the treeview into an array
Func TreeViewCheckedToArray()
	_Debug(@ScriptLineNumber & " TreeViewCheckedToArray" & @CRLF)
	ReDim $DataArray[1]
	For $X = 1 To UBound($HandleArray) - 1
		Local $AA = _GUICtrlTreeView_GetChecked($hTreeView, $HandleArray[$X])
		Local $BB = _GUICtrlTreeView_GetText($hTreeView, $HandleArray[$X])
		_Debug(@ScriptLineNumber & " Checked  " & $AA & "  " & $BB & "  " & $HandleArray[$X] & @CRLF)
		If $AA = True Then
			_ArrayAdd($DataArray, $BB)
		EndIf
	Next
	;_ArrayDisplay($hTreeView)
EndFunc   ;==>TreeViewCheckedToArray
;-----------------------------------------------
;This function gets the un-checked data from the treeview into an array
Func TreeViewUnCheckedToArray()
	_Debug(@ScriptLineNumber & " TreeViewUnCheckedToArray" & @CRLF)
	ReDim $DataArray[1]
	For $X = 1 To UBound($HandleArray) - 1
		Local $AA = _GUICtrlTreeView_GetChecked($hTreeView, $HandleArray[$X])
		Local $BB = _GUICtrlTreeView_GetText($hTreeView, $HandleArray[$X])
		_Debug(@ScriptLineNumber & " UNChecked  " & $AA & "  " & $BB & "  " & $HandleArray[$X] & @CRLF)
		If $AA = False Then
			_ArrayAdd($DataArray, $BB)
		EndIf
	Next
	;_ArrayDisplay($hTreeView)
EndFunc   ;==>TreeViewUnCheckedToArray
;-----------------------------------------------
Func AddLocalData() ; doug   $InputLocalIP
	_Debug(@ScriptLineNumber & " AddLocalData" & @CRLF)
	GuiDisable("disable")
	If GUICtrlRead($CheckPublicIP) = $GUI_CHECKED And StringInStr(GUICtrlRead($InputPublicIP), "**") = 0 Then
		_ArrayAdd($DataArray, '* ' & _IPPad(GUICtrlRead($InputPublicIP)))
	EndIf
	_ArrayAdd($DataArray, '* ' & _IPPad("127.0.0.1"))
	; Parse the addresses
	Local $HostList = StringSplit(GUICtrlRead($InputLocalIP), ';,:')

	;ConsoleWrite('@@ _Debug(@ScriptLineNumber & ' & @ScriptLineNumber & ') : GUICtrlRead($InputLocalIP) = ' & GUICtrlRead($InputLocalIP) & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	_ArrayDelete($HostList, 0) ;get rid of the count value
	;Add all of the addresses to the $hTreeView
	;_ArrayDisplay($HostList, "test")
	For $A In $HostList
		$A = StringStripWS($A, 8) ;remove spaces
		If $A <> "" Then ;ignore blank values
			Local $B = _IPAddress($A) ;expands the address values
			For $C In $B
				Local $D = _IPPad($C)
				_ArrayAdd($DataArray, '* ' & _IPPad($D))
			Next
		EndIf
	Next
	ArrayToTreeView()
	GuiDisable("enable")
EndFunc   ;==>AddLocalData
;-----------------------------------------------
;This function parses the InputBox String and creates the addresses
Func AddData()
	_Debug(@ScriptLineNumber & " AddData" & @CRLF)
	GuiDisable("disable")
	$ToggleState = False
	_GUICtrlListView_DeleteAllItems($ListView)

	; Parse the addresses
	Local $HostList = StringSplit(GUICtrlRead($InputIPAddress), ";:,")
	_ArrayDelete($HostList, 0) ;get rid of the count value
	;Add all of the addresses to the $hTreeView

	;_ArrayDisplay($HostList)

	For $A In $HostList
		$A = StringStripWS($A, 8) ;remove spaces
		If $A <> "" Then ;ignore blank values
			Local $B = _IPAddress($A) ;expands the address values
			For $C In $B
				Local $D = _IPPad($C)
				_ArrayAdd($DataArray, _IPPad($D))
			Next
		EndIf
	Next
	ArrayToTreeView()
	GuiDisable("enable")
EndFunc   ;==>AddData
;-----------------------------------------------
;Global $HandleArray[1]
;Global $DataArray[1]

Func DelData()
	_Debug(@ScriptLineNumber & " DelData" & @CRLF)
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
	_Debug(@ScriptLineNumber & " ClrData" & @CRLF)
	$ToggleState = False

	ReDim $DataArray[1]
	ReDim $HandleArray[1]
	_GUICtrlTreeView_DeleteAll($hTreeView)
	_GUICtrlListView_DeleteAllItems($ListView)
	Global $HandleArray[1]

EndFunc   ;==>ClrData
;-----------------------------------------------
Func DefaultSettings()
	_Debug(@ScriptLineNumber & " DefaultSettings" & @CRLF)
	$ToggleState = False
	GUICtrlSetData($InputIPAddress, "127.0.0.1;")
	GUICtrlSetData($InputTimeout, 1000)
	GUICtrlSetData($InputLoops, 4)
	GUICtrlSetData($InputDelay, 500)
	GUICtrlSetState($CheckStop, $GUI_UNCHECKED)
	GUICtrlSetState($CheckDNS, $GUI_CHECKED)
	GUICtrlSetState($CheckClass, $GUI_CHECKED)
	GUICtrlSetState($CheckPublicIP, $GUI_CHECKED)
	GUICtrlSetState($CheckMAC, $GUI_CHECKED)
	; This should match the initail defined values for main
	WinMove("PingTool", "", 10, 10, 1000, 420)
EndFunc   ;==>DefaultSettings
;-----------------------------------------------
Func GuiDisable($choice) ;@SW_ENABLE @SW_disble
	_Debug(@ScriptLineNumber & " GuiDisable  " & $choice & "   " & @CRLF)
	Static Local $LastState
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
		_Debug(@ScriptLineNumber & " Invalid choice at GuiDisable" & $choice & @CRLF, 0x40)
	EndIf

	GUICtrlSetState($ButtonPing, $setting)
	GUICtrlSetState($ButtonData, $setting)
	GUICtrlSetState($ButtonSaveProject, $setting)
	GUICtrlSetState($ButtonLoadProject, $setting)
	GUICtrlSetState($ButtonLoadDefaults, $setting)
	GUICtrlSetState($ButtonSaveLog, $setting)

	GUICtrlSetState($ButtonAbout, $setting)
	GUICtrlSetState($ButtonHelp, $setting)
	GUICtrlSetState($ButtonEdit, $setting)
	GUICtrlSetState($ButtonExit, $setting)

	GUICtrlSetState($ButtonAddLocal, $setting)
	GUICtrlSetState($ButtonFetch, $setting)
	GUICtrlSetState($ButtonAdd, $setting)
	GUICtrlSetState($ButtonClr, $setting)
	GUICtrlSetState($ButtonDel, $setting)
	GUICtrlSetState($ButtonTog, $setting)
	GUICtrlSetState($ButtonTest, $setting)

EndFunc   ;==>GuiDisable
;-----------------------------------------------
Func GetPublicIp()
	_Debug(@ScriptLineNumber & " GetPublicIp" & @CRLF)
	GUICtrlSetData($InputPublicIP, "Getting Public IP") ; Getting public IP-1
	GUICtrlSetColor($InputPublicIP, 0xff0000)
	Local $T = _GetIP()
	GUICtrlSetColor($InputPublicIP, 0x000000)
	If _TestIp($T) = "ERROR0" Then
		GUICtrlSetData($InputPublicIP, $T)
	Else
		GUICtrlSetData($InputPublicIP, "**** Failed")
	EndIf

EndFunc   ;==>GetPublicIp

;-----------------------------------------------
Func GetLocalIp()
	_Debug(@ScriptLineNumber & " Getting the local IP address" & @CRLF)
	Local $T = ''
	GUICtrlSetColor($InputPublicIP, 0xff0000)
	GUICtrlSetData($InputLocalIP, "GetLocalIp")
	If @IPAddress1 <> '0.0.0.0' Then $T = $T & ';' & _IPPad(@IPAddress1)
	If @IPAddress2 <> '0.0.0.0' Then $T = $T & ';' & _IPPad(@IPAddress2)
	If @IPAddress3 <> '0.0.0.0' Then $T = $T & ';' & _IPPad(@IPAddress3)
	If @IPAddress4 <> '0.0.0.0' Then $T = $T & ';' & _IPPad(@IPAddress4)
	GUICtrlSetData($InputLocalIP, $T)
	GUICtrlSetColor($InputPublicIP, 0x000000)
EndFunc   ;==>GetLocalIp
;-----------------------------------------------
Func PingTheHosts()
	_Debug(@ScriptLineNumber & " PingTheHosts" & @CRLF)
	DataToListView()
	GuiDisable("disable")
	GUICtrlSetState($CheckStop, $GUI_UNCHECKED)

	Local $MaxLoops = GUICtrlRead($InputLoops)
	Local $Loop = 0
	While $Loop < $MaxLoops Or $MaxLoops <= 0
		$MaxLoops = GUICtrlRead($InputLoops)
		$Loop = $Loop + 1

		If GUICtrlRead($CheckStop) = $GUI_CHECKED Then
			ExitLoop
		EndIf

		#cs
			Const $Address = 0
			Const $Status = 1
			Const $IPClass = 2
			Const $ResponseTime = 3
			Const $Worsttime = 4
			Const $Loop = 5
			Const $Alive = 6
			Const $Dead = 7
			Const $NameError = 8
			Const $MAC = 9
		#ce

		;now do the testing
		For $X = 0 To _GUICtrlListView_GetItemCount($ListView)
			If GUICtrlRead($CheckStop) = $GUI_CHECKED Then
				ExitLoop
			EndIf

			Local $T = _GUICtrlListView_GetItemText($ListView, $X)
			_Debug(@ScriptLineNumber & " PingTheHosts  " & $T & @CRLF)
			Local $result = Ping($T, GUICtrlRead($InputTimeout))
			Local $error = @error
			_Debug(@ScriptLineNumber & " PingTheHosts " & $T & "  " & $result & "  " & $error & @CRLF)


			If GUICtrlRead($CheckMAC) = $GUI_CHECKED Then _GUICtrlListView_AddSubItem($ListView, $X, GetMAC($T), $MAC)

			_GUICtrlListView_AddSubItem($ListView, $X, StringFormat("%d", $result), 3)
			Local $tmp = _GUICtrlListView_GetItemText($ListView, $X, 4)
			If $tmp <= $result Then _GUICtrlListView_AddSubItem($ListView, $X, StringFormat("%d", $result), 4)

			;If $result = 9999 Then ;Dead
			If $result = 0 Then ;Dead
				Local $ErrorCause
				Switch $error
					Case 1
						$ErrorCause = "Host is offline"
					Case 2
						$ErrorCause = "Host is unreachable"
					Case 3
						$ErrorCause = "Bad destination"
					Case 4
						$ErrorCause = "Other errors"
					Case Else
						$ErrorCause = @error
				EndSwitch
				_GUICtrlListView_AddSubItem($ListView, $X, $ErrorCause, $NameError)
				$tmp = _GUICtrlListView_GetItemText($ListView, $X, 7)
				$tmp = $tmp + 1
				_GUICtrlListView_AddSubItem($ListView, $X, $tmp, 7)
				_GUICtrlListView_AddSubItem($ListView, $X, "Dead", 1)
			Else ;Alive
				_GUICtrlListView_AddSubItem($ListView, $X, "Alive", 1)

				$tmp = _GUICtrlListView_GetItemText($ListView, $X, 6)
				$tmp = $tmp + 1
				_GUICtrlListView_AddSubItem($ListView, $X, $tmp, 6)
				Local $DNSResult = ""
				If GUICtrlRead($CheckDNS) = $GUI_CHECKED Then
					TCPStartup()
					$DNSResult = _TCPIpToName($T, 0)
					TCPShutdown()
				Else
					$DNSResult = "*** Alive ***"
				EndIf
				_GUICtrlListView_AddSubItem($ListView, $X, $DNSResult, 8)
			EndIf

			Local $ClassResult = ""
			If GUICtrlRead($CheckClass) = $GUI_CHECKED Then
				$ClassResult = _CheckIPClass($T)
			Else
				$ClassResult = ""
			EndIf
			_GUICtrlListView_AddSubItem($ListView, $X, $ClassResult, 2)

			_GUICtrlListView_AddSubItem($ListView, $X, StringFormat("%d of %d", $Loop, $MaxLoops), 5)
		Next
		Sleep(GUICtrlRead($InputDelay))
	WEnd

	GuiDisable("enable")
EndFunc   ;==>PingTheHosts
;-----------------------------------------------
; @ScriptDir & "\AUXFiles\"
Func GetMAC($IPAddress)
	Local $out = ".\auxfiles\arpout.txt"
	Local $TA[1]
	Local $X = _RunDos("arp.exe -a > " & $out)
	ConsoleWrite(@ScriptLineNumber & " " & $X & " " & @error & @CRLF)

	_FileReadToArray($out, $TA)
	ConsoleWrite(@ScriptLineNumber & " " & $TA & @CRLF)
	For $A In $TA
		If StringInStr($A, $IPAddress) = 3 Then
			;ConsoleWrite(@ScriptLineNumber & " " & $IPAddress & " " & @CRLF)
			;ConsoleWrite(@ScriptLineNumber & " " & StringMid($A, 25, 42 - 25) & " " & @CRLF)
			Return StringMid($A, 25, 42 - 25)
		EndIf
	Next

	Return "Not found"
EndFunc   ;==>GetMAC
;-----------------------------------------------
Func SaveLog()
	_Debug(@ScriptLineNumber & " SaveLog" & @CRLF)
	$LOG_filename = FileSaveDialog("Save log file", @ScriptDir & "\AUXFiles\", _
			"PingTool logs (P*.log)|All logs (*.log)|All files (*.*)", 18, @ScriptDir & "\AUXFiles\PingTool.log")

	Local $file = FileOpen($LOG_filename, 2)
	; Check if file opened for writing OK
	If $file = -1 Then
		_Debug(@ScriptLineNumber & " SaveLog: Unable to open file for writing: " & $LOG_filename & @CRLF, 0x10, 5)
		Return
	EndIf

	_Debug(@ScriptLineNumber & " SaveLog  " & $LOG_filename & @CRLF)
	FileWriteLine($file, "Log file for " & @ScriptName & "  " & _DateTimeFormat(_NowCalc(), 0))

	FileWriteLine($file, "IPAddress: " & GUICtrlRead($InputIPAddress))
	FileWriteLine($file, "DNS: " & GUICtrlRead($CheckDNS))
	FileWriteLine($file, "Class: " & GUICtrlRead($CheckClass))
	FileWriteLine($file, "Pub IP: " & GUICtrlRead($CheckPublicIP))
	FileWriteLine($file, "MAC: " & GUICtrlRead($CheckMAC))
	FileWriteLine($file, "Timeout: " & GUICtrlRead($InputTimeout))
	FileWriteLine($file, "Loops: " & GUICtrlRead($InputLoops))
	FileWriteLine($file, "Delay: " & GUICtrlRead($InputDelay))

	FileWriteLine($file, "LocalIP: " & GUICtrlRead($InputLocalIP))
	FileWriteLine($file, "PublicIP: " & GUICtrlRead($InputPublicIP))

	FileWriteLine($file, "Cnt          Address Status      Class    MS  W-MS       Loop Alive Dead      Name\Error")

	For $X = 0 To _GUICtrlListView_GetItemCount($ListView) - 1
		Local $T = _GUICtrlListView_GetItemTextString($ListView, $X) ;, 0)
		Local $array = StringSplit($T, "|")
		Local $NewString = StringFormat("%3d %16s %6s %10s %5d %5d %10s  %4d %4d      %-14s ", $X + 1, $array[1], $array[2], $array[3], $array[4], $array[5], $array[6], $array[7], $array[8], $array[9])
		_Debug(@ScriptLineNumber & " SaveLog:  " & $NewString & @CRLF)
		FileWriteLine($file, $NewString)
	Next

	FileClose($file)
EndFunc   ;==>SaveLog
;-----------------------------------------------
Func SaveProject()
	_Debug(@ScriptLineNumber & " SaveProject" & @CRLF)
	$Project_filename = FileSaveDialog("Save project file", @ScriptDir & "\AUXFiles\", _
			"PingTool projects (P*.prj)|All projects (*.prj)|All files (*.*)", 18, @ScriptDir & "\AUXFiles\PingTool.prj")

	Local $file = FileOpen($Project_filename, 2)
	; Check if file opened for writing OK
	If $file = -1 Then
		_Debug(@ScriptLineNumber & " SaveProject: Unable to open file for writing: " & $Project_filename & @CRLF, 0x10, 5)
		Return
	EndIf
	_Debug(@ScriptLineNumber & " SaveProject  " & $Project_filename & @CRLF)
	FileWriteLine($file, "Valid for PingTool project")
	FileWriteLine($file, "Project file for " & @ScriptName & "  " & _DateTimeFormat(_NowCalc(), 0))
	FileWriteLine($file, "Help 1 is enabled, 4 is disabled for checkboxes")
	FileWriteLine($file, "CheckDNS:" & GUICtrlRead($CheckDNS))
	FileWriteLine($file, "CheckClass:" & GUICtrlRead($CheckClass))
	FileWriteLine($file, "CheckPublicIP:" & GUICtrlRead($CheckPublicIP))
	FileWriteLine($file, "CheckMAC:" & GUICtrlRead($CheckMAC))
	FileWriteLine($file, "InputTimeout:" & GUICtrlRead($InputTimeout))
	FileWriteLine($file, "InputLoops:" & GUICtrlRead($InputLoops))
	FileWriteLine($file, "InputDelay:" & GUICtrlRead($InputDelay))
	FileWriteLine($file, "InputIPAddressEdit:" & GUICtrlRead($InputIPAddress))

	Local $F = WinGetPos("PingTool", "")
	FileWriteLine($file, "MainWinpos:" & $F[0] & " " & $F[1] & " " & $F[2] & " " & $F[3])

	For $X = 1 To UBound($HandleArray) - 1
		Local $Y = _GUICtrlTreeView_GetText($hTreeView, $HandleArray[$X])
		Local $Z = StringInStr($Y, "*")

		If $Z = 0 Then
			FileWriteLine($file, "InputIPAddress:" & $Y)
		EndIf
	Next

	FileClose($file)
EndFunc   ;==>SaveProject
;-----------------------------------------------
;This loads the project file but into the tree control but not into the list
Func LoadProject($type)
	_Debug(@ScriptLineNumber & " LoadProject " & $type & @CRLF)

	If StringCompare($type, "menu") = 0 Then
		$Project_filename = FileOpenDialog("Load project file", @ScriptDir & "\AUXFiles\", _
				"PingTool projects (P*.prj)|All projects (*.prj)|All files (*.*)", 18, @ScriptDir & "\AUXFiles\PingTool.prj")
	EndIf

	Local $file = FileOpen($Project_filename, 0)
	; Check if file opened for reading OK
	If $file = -1 Then
		_Debug(@ScriptLineNumber & " LoadProject: Unable to open file for reading: " & $Project_filename & @CRLF, 0x10, 5)
		Return
	EndIf

	_Debug(@ScriptLineNumber & " LoadProject " & $Project_filename & @CRLF)
	; Read in the first line to verify the file is of the correct type
	If StringCompare(FileReadLine($file, 1), "Valid for PingTool project") <> 0 Then
		_Debug(@ScriptLineNumber & " Not a valid project file for PingTool" & @CRLF, 0x20, 5)
		FileClose($file)
		Return
	EndIf

	ClrData()
	GUICtrlSetData($InputIPAddress, '')

	; Read in lines of text until the EOF is reached
	While 1
		Local $LineIn = FileReadLine($file)
		If @error = -1 Then ExitLoop

		_Debug(@ScriptLineNumber & " " & $LineIn & " " & @CRLF)
		If StringInStr($LineIn, ";") = 1 Then ContinueLoop

		Local $F
		If StringInStr($LineIn, "MainWinpos:") Then
			$F = StringMid($LineIn, StringInStr($LineIn, ":") + 1)
			$F = StringSplit($F, " ", 2)
			WinMove("PingTool", "", $F[0], $F[1], $F[2], $F[3])
		EndIf

		If StringInStr($LineIn, "CheckDNS:") Then GUICtrlSetState($CheckDNS, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
		If StringInStr($LineIn, "CheckClass:") Then GUICtrlSetState($CheckClass, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
		If StringInStr($LineIn, "CheckPublicIP:") Then GUICtrlSetState($CheckPublicIP, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
		If StringInStr($LineIn, "CheckMAC:") Then GUICtrlSetState($CheckMAC, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
		If StringInStr($LineIn, "InputTimeout:") Then GUICtrlSetData($InputTimeout, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
		If StringInStr($LineIn, "InputLoops:") Then GUICtrlSetData($InputLoops, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
		If StringInStr($LineIn, "InputDelay:") Then GUICtrlSetData($InputDelay, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
		If StringInStr($LineIn, "InputIPAddressEdit:") Then GUICtrlSetData($InputIPAddress, StringMid($LineIn, StringInStr($LineIn, ":") + 1))

		If StringInStr($LineIn, "InputIPAddress:") Then
			Local $tmp = StringMid($LineIn, StringInStr($LineIn, ":") + 1)
			If StringInStr($tmp, "*") = 0 Then _ArrayAdd($DataArray, _IPPad($tmp))
		EndIf
	WEnd

	FileClose($file)
	;_GUICtrlListView_DeleteAllItems($ListView)
	GetLocalIp()
	If GUICtrlRead($CheckPublicIP) = $GUI_CHECKED Then GetPublicIp()

	AddData()

	; If the main window is not visible, make it visible
	$F = WinGetPos("PingTool", "")
	_debug(@ScriptLineNumber & " DesktopWidth: " & $F[0] & " " & @DesktopWidth & @CRLF)
	_debug(@ScriptLineNumber & " DesktopHeight: " & $F[1] & " " & @DesktopHeight & @CRLF)
	If $F[0] > @DesktopWidth Or $F[1] > @DesktopHeight Then WinMove("PingTool", "", 10, 10, 1000, 420)
	If $F[0] < 0 Or $F[1] < 0 Then WinMove("PingTool", "", 10, 10, 1000, 420)

EndFunc   ;==>LoadProject
;-----------------------------------------------
Func ToggleData()
	_Debug(@ScriptLineNumber & " ToggleData" & @CRLF)
	GuiDisable("disable")

	;_ArrayDisplay($HandleArray)
	$ToggleState = Not $ToggleState

	For $X = 1 To UBound($HandleArray) - 1
		_GUICtrlTreeView_SetChecked($hTreeView, $HandleArray[$X], $ToggleState)
		_Debug(@ScriptLineNumber & StringFormat(" %2u   %6x   %d", $X, $HandleArray[$X], $ToggleState) & @CRLF)
	Next

	GuiDisable("enable")
EndFunc   ;==>ToggleData
;-----------------------------------------------

Func Edit()
	_Debug(@ScriptLineNumber & " Edit" & @CRLF)
	Local $Filename = FileOpenDialog("View a file", @ScriptDir & "\AUXFiles\", _
			"All (*.*)", 1, @ScriptDir & "\AUXFiles\PingTool.prj")

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

	_Debug(@ScriptLineNumber & " Edit  " & $Filename & @CRLF)
	ShellExecuteWait($editor, $Filename)

EndFunc   ;==>Edit
;-----------------------------------------------
Func About($FormID)
	Local $D = WinGetPos($ProgramName)
	Local $WinPos
	If IsArray($D) = 1 Then
		$WinPos = StringFormat("%s" & @CRLF & "WinPOS: %d  %d " & @CRLF & "WinSize: %d %d " & @CRLF & "Desktop: %d %d ", _
				$FormID, $D[0], $D[1], $D[2], $D[3], @DesktopWidth, @DesktopHeight)
	Else
		$WinPos = ">>>About ERROR, Check the window name<<<"
	EndIf
	_Debug(@CRLF & $SystemS & @CRLF & $WinPos & @CRLF & "Written by Doug Kaynor because I wanted to!" & @CRLF, 0x40, 0)
EndFunc   ;==>About
