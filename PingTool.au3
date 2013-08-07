#region ;**** Directives created by AutoIt3Wrapper_GUI ****
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
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****
#region
;#Tidy_Parameters=/gd /sf
#endregion

#cs
	Fixed:
	This area is used to store things todo, bugs, and other notes

	Todo:
	Hide dead devices in display
	Make DEL and CLR buttons harder to hit or prompt before doing
	Handle duplicates in lists
	Work on 1024x768
	Check for window position and fit to screen
#ce

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
#include <GuiListBox.au3>
#include <GUIConstantsEx.au3>
#include <GuiTreeView.au3>
;#include <GuiImageList.au3>
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
GUISetHelp("notepad", $MainForm) ; Need a help file to call here
Global $ButtonPing = GUICtrlCreateButton("Ping", 10, 10, 50, 20)
GUICtrlSetTip(-1, "This will do the pings. (What did you expect?)")
Global $CheckStop = GUICtrlCreateCheckbox("Stop", 10, 30, 52, 30)
GUICtrlSetTip(-1, "Stop pinging")
Global $ButtonData = GUICtrlCreateButton("Data", 10, 70, 50, 20)
GUICtrlSetTip(-1, "Clicking this will display the addresses")
Global $ButtonSaveProject = GUICtrlCreateButton("Save project", 60, 10, 110, 20)
GUICtrlSetTip(-1, "Save the current settings")
Global $ButtonLoadProject = GUICtrlCreateButton("Load project", 60, 30, 110, 20)
GUICtrlSetTip(-1, "Load saved settings")
Global $ButtonLoadDefaults = GUICtrlCreateButton("Load defaults", 60, 50, 110, 20)
GUICtrlSetTip(-1, "Load default settings")
Global $ButtonSaveLog = GUICtrlCreateButton("Save log", 60, 70, 110, 20)
GUICtrlSetTip(-1, "Save a log of test results")
Global $ButtonAbout = GUICtrlCreateButton("About", 170, 10, 55, 20)
GUICtrlSetTip(-1, "About button")
Global $ButtonHelp = GUICtrlCreateButton("Help", 170, 30, 55, 20)
GUICtrlSetTip(-1, "Help button")
Global $ButtonEdit = GUICtrlCreateButton("Edit", 170, 50, 55, 20)
GUICtrlSetTip(-1, "Edit or view a file")
Global $ButtonExit = GUICtrlCreateButton("Exit", 170, 70, 55, 20)
GUICtrlSetTip(-1, "Exit button")
Global $CheckDNS = GUICtrlCreateCheckbox("DNS", 230, 5, 50, 30)
GUICtrlSetTip(-1, "Display DNS results of alive hosts")
Global $CheckClass = GUICtrlCreateCheckbox("Class", 230, 30, 60, 30)
GUICtrlSetTip(-1, "Display IP Class")
Global $CheckPublicIP = GUICtrlCreateCheckbox("Pub", 230, 55, 50, 22)
GUICtrlSetTip(-1, "Include Public IP")
Global $CheckMAC = GUICtrlCreateCheckbox("MAC", 230, 80, 50, 22)
GUICtrlSetTip(-1, "Display MAC")
Global $LabelTimeout = GUICtrlCreateLabel("Timeout", 300, 10, 60, 20, $SS_SUNKEN)
GUICtrlSetTip(-1, "Timeout value in milliseconds")
Global $InputTimeout = GUICtrlCreateInput("****", 365, 10, 50, 20, BitOR($ES_CENTER, $ES_AUTOHSCROLL))
GUICtrlSetTip(-1, "Timeout value in milliseconds")
Global $LabelLoops = GUICtrlCreateLabel("Loops", 300, 40, 60, 20, $SS_SUNKEN)
GUICtrlSetTip(-1, "Number of loop to do")
Global $InputLoops = GUICtrlCreateInput("****", 365, 40, 50, 20, BitOR($ES_CENTER, $ES_AUTOHSCROLL))
GUICtrlSetTip(-1, "Number of loop to do")
Global $LabelDelay = GUICtrlCreateLabel("Delay", 300, 70, 60, 20, $SS_SUNKEN)
GUICtrlSetTip(-1, "Delay between loops in milliseconds")
Global $InputDelay = GUICtrlCreateInput("100", 365, 70, 50, 20, BitOR($ES_CENTER, $ES_AUTOHSCROLL))
GUICtrlSetTip(-1, "Delay between loops in milliseconds")
Global $InputIPAddress = GUICtrlCreateInput("InputIPAddress", 430, 10, 260, 20, BitOR($ES_CENTER, $ES_AUTOHSCROLL))
GUICtrlSetTip(-1, "String describing the addresses to test")
Global $LabelLocalIP = GUICtrlCreateLabel("Local", 430, 35, 55, 20, $SS_SUNKEN)
GUICtrlSetTip(-1, "This is the systems local IP address")
Global $InputLocalIP = GUICtrlCreateInput("*****", 490, 35, 200, 20, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetTip(-1, "This is the systems local IP address")
Global $LabelPublicIP = GUICtrlCreateLabel("Public", 430, 60, 55, 20, $SS_SUNKEN)
GUICtrlSetTip(-1, "This is the systems public IP")
Global $InputPublicIP = GUICtrlCreateInput("****", 490, 60, 200, 20, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetTip(-1, "This is the systems public IP")
Global $ButtonAddLocal = GUICtrlCreateButton("Add", 430, 85, 35, 20)
GUICtrlSetTip(-1, "Add local data to list")
Global $ButtonFetch = GUICtrlCreateButton("Fetch", 500, 85, 50, 20)
GUICtrlSetTip(-1, "Fetch selected value to address input")
;Global $CheckHideDead = GUICtrlCreateCheckbox("Hide dead", 570, 85, 50, 22)
;GUICtrlSetTip(-1, "Display MAC")
Global $ButtonHideDead = GUICtrlCreateButton("Hide dead", 570, 85, 100, 20)
GUICtrlSetTip(-1, "Hide dead IP's")
Global $ButtonAdd = GUICtrlCreateButton("Add", 700, 5, 45, 20)
GUICtrlSetTip(-1, "Add data to list")
Global $ButtonClr = GUICtrlCreateButton("Clr", 700, 25, 45, 20)
GUICtrlSetTip($ButtonClr, "Clear the list")
Global $ButtonDel = GUICtrlCreateButton("Del", 700, 45, 45, 20)
GUICtrlSetTip(-1, "Delete checked boxes")
Global $ButtonTog = GUICtrlCreateButton("TOG", 700, 65, 45, 20)
GUICtrlSetTip(-1, "Toggle check boxes")
Global $ButtonTest = GUICtrlCreateButton("Test", 700, 85, 45, 20)
GUICtrlSetTip(-1, "Test")
Global $hTreeView = GUICtrlCreateTreeView(800, 8, 190, 80, BitOR($TVS_DISABLEDRAGDROP, $TVS_SHOWSELALWAYS, $TVS_CHECKBOXES, $WS_GROUP, $WS_TABSTOP), $WS_EX_CLIENTEDGE)
GUICtrlSetTip(-1, "Select the addresses to ping")


For $x = 0 To 100
	GUICtrlSetResizing($x, $GUI_DOCKALL)
Next

GUICtrlSetResizing($hTreeView, $GUI_DOCKRIGHT)

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
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $Address, 130); Address
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $Status, 65) ; Status
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $IPClass, 90) ; IP Class
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $ResponseTime, 50) ; Ms response time
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $Worsttime, 50) ; Worst time
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $Loop, 90); Loop
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $Alive, 60) ; Alive
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $Dead, 60) ; Dead
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $NameError, 155); Name\Error
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $MAC, 160); MAC
GUICtrlSetTip(-1, "Display the test results")
GUICtrlSetResizing(-1, 98)
_GUICtrlListView_SetBkColor($ListView, $CLR_WHITE)
_GUICtrlListView_SetTextColor($ListView, $CLR_BLACK)
_GUICtrlListView_SetTextBkColor($ListView, $CLR_Silver)

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
			_Debug("Case $InputIPAddress")
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
			_Debug("CheckClass")
		Case $CheckDNS
			_Debug("CheckDNS")
		Case $CheckPublicIP
			_Debug("CheckPublicIP")
			If GUICtrlRead($CheckPublicIP) = $GUI_CHECKED Then GetPublicIp()
		Case $CheckMAC
			_Debug("CheckMAC")
		Case $CheckStop
			_Debug("CheckStop")
		Case $LabelPublicIP
			GetPublicIp()
		Case $LabelLocalIP
			GetLocalIp()
		Case $ButtonHideDead
			; $ListView

			;For $x = 0 To _GUICtrlListView_GetItemCount($ListView)
			;	Global $T = _GUICtrlListView_GetItemText($ListView, $x, $Status) ;retrive the status subitem text
			;	ConsoleWrite(@ScriptLineNumber & " >>" & $T & '<<>>' & $x & "<< " & @CRLF)
			;	If StringInStr($T, 'Dead') > 0 Then
			Global $j = _GUICtrlListView_DeleteAllItems($ListView)
			;ConsoleWrite(@ScriptLineNumber & " " & $j & @CRLF)
			;	EndIf
			;Next
	EndSwitch
WEnd
;-----------------------------------------------
;This function copies checked data to ListView
Func DataToListView()
	_Debug("DataToListView")
	GuiDisable("disable")

	_GUICtrlListView_DeleteAllItems(GUICtrlGetHandle($ListView))
	Local $DataArraySave = $DataArray
	TreeViewCheckedToArray() ;$ListView

	_RemoveBlankLines($DataArray)
	_ArrayUnique($DataArray)
	_ArraySort($DataArray)

	For $x = 0 To UBound($DataArray) - 1
		Local $T = StringReplace($DataArray[$x], '* ', '')
		If StringInStr(_TestIP($T), "ERROR0", 2) Then _GUICtrlListView_AddItem($ListView, _IPUnPad($T))
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
	_ArrayUnique($DataArray)
	_ArraySort($DataArray)

	For $x In $DataArray
		If _TestIP($x) = "ERROR0" Then
			_ArrayAdd($HandleArray, _GUICtrlTreeView_Add($hTreeView, 0, $x))
		EndIf
	Next
EndFunc   ;==>ArrayToTreeView
;-----------------------------------------------
Func Fetch()
	For $x = 1 To UBound($HandleArray) - 1
		If _GUICtrlTreeView_GetSelected($hTreeView, $HandleArray[$x]) == True Then
			GUICtrlSetData($InputIPAddress, _GUICtrlTreeView_GetText($hTreeView, $HandleArray[$x]))
		EndIf
	Next
EndFunc   ;==>Fetch
;-----------------------------------------------
;This function gets all of the data from the treeview into an array
Func TreeViewToArray()
	_Debug("TreeViewToArray")
	ReDim $DataArray[1]
	For $x = 1 To UBound($HandleArray) - 1
		_ArrayAdd($DataArray, _GUICtrlTreeView_GetText($hTreeView, $HandleArray[$x]))
	Next
EndFunc   ;==>TreeViewToArray
;-----------------------------------------------
;This function gets the checked data from the treeview into an array
Func TreeViewCheckedToArray()
	_Debug("TreeViewCheckedToArray")
	ReDim $DataArray[1]
	For $x = 1 To UBound($HandleArray) - 1
		Local $AA = _GUICtrlTreeView_GetChecked($hTreeView, $HandleArray[$x])
		Local $BB = _GUICtrlTreeView_GetText($hTreeView, $HandleArray[$x])
		_Debug("Checked  " & $AA & "  " & $BB & "  " & $HandleArray[$x])
		If $AA = True Then
			_ArrayAdd($DataArray, $BB)
		EndIf
	Next
EndFunc   ;==>TreeViewCheckedToArray
;-----------------------------------------------
;This function gets the un-checked data from the treeview into an array
Func TreeViewUnCheckedToArray()
	_Debug("TreeViewUnCheckedToArray")
	ReDim $DataArray[1]
	For $x = 1 To UBound($HandleArray) - 1
		Local $AA = _GUICtrlTreeView_GetChecked($hTreeView, $HandleArray[$x])
		Local $BB = _GUICtrlTreeView_GetText($hTreeView, $HandleArray[$x])
		_Debug("UNChecked  " & $AA & "  " & $BB & "  " & $HandleArray[$x])
		If $AA = False Then
			_ArrayAdd($DataArray, $BB)
		EndIf
	Next
	;_ArrayDisplay($hTreeView)
EndFunc   ;==>TreeViewUnCheckedToArray
;-----------------------------------------------
Func AddLocalData() ; doug   $InputLocalIP
	_Debug("AddLocalData")
	GuiDisable("disable")
	If GUICtrlRead($CheckPublicIP) = $GUI_CHECKED And StringInStr(GUICtrlRead($InputPublicIP), "**") = 0 Then
		_ArrayAdd($DataArray, '* ' & _IPPad(GUICtrlRead($InputPublicIP)))
	EndIf
	_ArrayAdd($DataArray, '* ' & _IPPad("127.0.0.1"))
	; Parse the addresses
	Local $HostList = StringSplit(GUICtrlRead($InputLocalIP), ';,:')

	ConsoleWrite('@@ _Debug(' & @ScriptLineNumber & ') : GUICtrlRead($InputLocalIP) = ' & GUICtrlRead($InputLocalIP) & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

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
	_Debug("AddData")
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
	_GUICtrlListView_DeleteAllItems($ListView)
	Global $HandleArray[1]

EndFunc   ;==>ClrData
;-----------------------------------------------
Func DefaultSettings()
	_Debug("DefaultSettings")
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
	_Debug("GetPublicIp")
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
	_Debug("Getting the local IP address")
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
	_Debug("PingTheHosts")
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

		;now do the testing
		For $x = 0 To _GUICtrlListView_GetItemCount($ListView)
			If GUICtrlRead($CheckStop) = $GUI_CHECKED Then
				ExitLoop
			EndIf

			Local $T = _GUICtrlListView_GetItemText($ListView, $x)
			_Debug("PingTheHosts  " & $T)
			Local $result = Ping($T, GUICtrlRead($InputTimeout))
			Local $error = @error
			_Debug("PingTheHosts  " & $T & "   " & $result & "   " & $error)

			If GUICtrlRead($CheckMAC) = $GUI_CHECKED Then _GUICtrlListView_AddSubItem($ListView, $x, GetMAC($T), $MAC)

			_GUICtrlListView_AddSubItem($ListView, $x, StringFormat("%d", $result), 3)
			Local $tmp = _GUICtrlListView_GetItemText($ListView, $x, 4)
			If $tmp <= $result Then _GUICtrlListView_AddSubItem($ListView, $x, StringFormat("%d", $result), 4)

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
				_GUICtrlListView_AddSubItem($ListView, $x, $ErrorCause, $NameError)
				$tmp = _GUICtrlListView_GetItemText($ListView, $x, 7)
				$tmp = $tmp + 1
				_GUICtrlListView_AddSubItem($ListView, $x, $tmp, 7)
				_GUICtrlListView_AddSubItem($ListView, $x, "Dead", 1)

			Else ;Alive
				_GUICtrlListView_AddSubItem($ListView, $x, "Alive", 1)

				$tmp = _GUICtrlListView_GetItemText($ListView, $x, 6)
				$tmp = $tmp + 1
				_GUICtrlListView_AddSubItem($ListView, $x, $tmp, 6)
				Local $DNSResult = ""
				If GUICtrlRead($CheckDNS) = $GUI_CHECKED Then
					TCPStartup()
					$DNSResult = _TCPIpToName($T, 0)
					TCPShutdown()
				Else
					$DNSResult = "*** Alive ***"
				EndIf
				_GUICtrlListView_AddSubItem($ListView, $x, $DNSResult, 8)
			EndIf

			Local $ClassResult = ""
			If GUICtrlRead($CheckClass) = $GUI_CHECKED Then
				$ClassResult = _CheckIPClass($T)
			Else
				$ClassResult = ""
			EndIf
			_GUICtrlListView_AddSubItem($ListView, $x, $ClassResult, 2)

			_GUICtrlListView_AddSubItem($ListView, $x, StringFormat("%d of %d", $Loop, $MaxLoops), 5)
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
	;ConsoleWrite(@ScriptLineNumber & " " & _RunDOS("arp.exe -a > " & $out) & @CRLF)
	_FileReadToArray($out, $TA)

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
	_Debug("SaveLog")
	$LOG_filename = FileSaveDialog("Save log file", @ScriptDir & "\AUXFiles\", _
			"PingTool logs (P*.log)|All logs (*.log)|All files (*.*)", 18, @ScriptDir & "\AUXFiles\PingTool.log")

	Local $file = FileOpen($LOG_filename, 2)
	; Check if file opened for writing OK
	If $file = -1 Then
		_Debug("SaveLog: Unable to open file for writing: " & $LOG_filename, 0x10, 5)
		Return
	EndIf

	_Debug("SaveLog  " & $LOG_filename)
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

	For $x = 0 To _GUICtrlListView_GetItemCount($ListView) - 1
		Local $T = _GUICtrlListView_GetItemTextString($ListView, $x);, 0)
		Local $array = StringSplit($T, "|")
		Local $NewString = StringFormat("%3d %16s %6s %10s %5d %5d %10s  %4d %4d      %-14s ", $x + 1, $array[1], $array[2], $array[3], $array[4], $array[5], $array[6], $array[7], $array[8], $array[9])
		_Debug("SaveLog:  " & $NewString)
		FileWriteLine($file, $NewString)
	Next

	FileClose($file)
EndFunc   ;==>SaveLog
;-----------------------------------------------
Func SaveProject()
	_Debug("SaveProject")
	$Project_filename = FileSaveDialog("Save project file", @ScriptDir & "\AUXFiles\", _
			"PingTool projects (P*.prj)|All projects (*.prj)|All files (*.*)", 18, @ScriptDir & "\AUXFiles\PingTool.prj")

	Local $file = FileOpen($Project_filename, 2)
	; Check if file opened for writing OK
	If $file = -1 Then
		_Debug("SaveProject: Unable to open file for writing: " & $Project_filename, 0x10, 5)
		Return
	EndIf
	_Debug("SaveProject  " & $Project_filename)
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

	For $x = 1 To UBound($HandleArray) - 1
		Local $Y = _GUICtrlTreeView_GetText($hTreeView, $HandleArray[$x])
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
	_Debug("LoadProject  " & $type)

	If StringCompare($type, "menu") = 0 Then
		$Project_filename = FileOpenDialog("Load project file", @ScriptDir & "\AUXFiles\", _
				"PingTool projects (P*.prj)|All projects (*.prj)|All files (*.*)", 18, @ScriptDir & "\AUXFiles\PingTool.prj")
	EndIf

	Local $file = FileOpen($Project_filename, 0)
	; Check if file opened for reading OK
	If $file = -1 Then
		_Debug("LoadProject: Unable to open file for reading: " & $Project_filename, 0x10, 5)
		Return
	EndIf

	_Debug("LoadProject   " & $Project_filename)
	; Read in the first line to verify the file is of the correct type
	If StringCompare(FileReadLine($file, 1), "Valid for PingTool project") <> 0 Then
		_Debug("Not a valid project file for PingTool", 0x20, 5)
		FileClose($file)
		Return
	EndIf

	ClrData()
	GUICtrlSetData($InputIPAddress, '')

	; Read in lines of text until the EOF is reached
	While 1
		Local $LineIn = FileReadLine($file)
		If @error = -1 Then ExitLoop

		_Debug("LoadProject   " & $LineIn)
		If StringInStr($LineIn, ";") = 1 Then ContinueLoop
		Local $F
		If StringInStr($LineIn, "MainWinpos:") Then
			$F = StringMid($LineIn, StringInStr($LineIn, ":") + 1)
			$F = StringSplit($F, " ", 2)
			WinMove("PingTool", "", $F[0], $F[1], $F[2], $F[3])
		EndIf

		; If the main window is not visible, make it visible
		$F = WinGetPos("PingTool", "")
		ConsoleWrite(@ScriptLineNumber & " " & $F[0] & " " & @DesktopWidth & @CRLF)
		ConsoleWrite(@ScriptLineNumber & " " & $F[1] & " " & @DesktopHeight & @CRLF)
		If $F[0] > @DesktopWidth Or $F[1] > @DesktopHeight Then WinMove("PingTool", "", 10, 10, 1000, 420)
		If $F[0] < 0 Or $F[1] < 0 Then WinMove("PingTool", "", 10, 10, 1000, 420)

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

EndFunc   ;==>LoadProject
;-----------------------------------------------
Func ToggleData()
	_Debug("ToggleData")
	GuiDisable("disable")

	;_ArrayDisplay($HandleArray)
	$ToggleState = Not $ToggleState

	For $x = 1 To UBound($HandleArray) - 1
		_GUICtrlTreeView_SetChecked($hTreeView, $HandleArray[$x], $ToggleState)
		_Debug(StringFormat("%2u   %6x   %d", $x, $HandleArray[$x], $ToggleState))
	Next

	GuiDisable("enable")
EndFunc   ;==>ToggleData
;-----------------------------------------------

Func Edit()
	_Debug("Edit")
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

	_Debug("Edit  " & $Filename)
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
	_Debug(@CRLF & $SystemS & @CRLF & $WinPos & @CRLF & "Written by Doug Kaynor because I wanted to!", 0x40, 0)
EndFunc   ;==>About

