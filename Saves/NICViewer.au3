#RequireAdmin
#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=../icons/Cryptkeeper.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Comment=A program to list Device (NIC) Viewer
#AutoIt3Wrapper_Res_Description=Device (NIC) Viewer
#AutoIt3Wrapper_Res_Fileversion=2.0.0.104
#AutoIt3Wrapper_Res_FileVersion_AutoIncrement=Y
#AutoIt3Wrapper_Res_ProductVersion=666
#AutoIt3Wrapper_Res_LegalCopyright=Copyright © 2012 Douglas B Kaynor
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Field=Developer|Douglas Kaynor
#AutoIt3Wrapper_Res_Field=AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Res_Field=Compile date|%longdate% %time%
#AutoIt3Wrapper_Au3Check_Stop_OnWarning=y
#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Run_Tidy=y
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****

#CS
	Add right click for properties
	Cleanup, remove dead code
	Enable removal of non existant devices
#CE

Opt("MustDeclareVars", 1) ; require pre-declared varibles
If _Singleton(@ScriptName, 1) = 0 Then
	_Debug(@ScriptName & " is already running!", 0x40)
	Exit
EndIf
#include <Array.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <File.au3>
#include <Constants.au3>
#include <GUIMenu.au3>
#include <GUIConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiListView.au3>
#include <ListViewConstants.au3>
#include <Misc.au3>
#include <Process.au3>
#include <String.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <_DougFunctions.au3>

TraySetIcon("./icons/Cryptkeeper.ico")


Global $NICViewerFiles = $AUXPath & "\NICViewerFiles\"
DirCreate($NICViewerFiles)

Global $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")
Global $tmp = StringSplit(@ScriptName, ".")
Global $ProgramName = $tmp[1]
Global $ResultLocation = ""
Global $SystemS = $ProgramName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & _
		@OSServicePack & @CRLF & @OSType & @CRLF & @OSArch
Global $SystemProductName
Global $OsName
Global $LOGFILE = FileGetShortName($AUXPath & $ProgramName & ".log")
Global $DEVIDFILE = FileGetShortName($AUXPath & $ProgramName & ".dev")
Global $NICVIEWHELPFILE = FileGetShortName($AUXPath & $ProgramName & ".txt")
Global $DEBUG = False
Global $DEVID_array

Global $hWnd, $iMsg, $iwParam, $ilParam
GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")

For $X = 1 To $CmdLine[0]
	;ConsoleWrite($X & " >> " & $CmdLine[$X] & @CRLF)
	Select
		Case StringInStr($CmdLine[$X], "help") > 0 Or _
				StringInStr($CmdLine[$X], "?") > 0
			Help()
			Exit
		Case StringInStr($CmdLine[$X], "logfile") > 0
			Global $Y = StringSplit($CmdLine[$X], "=")
			$LOGFILE = FileGetShortName(StringStripWS($Y[2], 3))
			_Debug(@ScriptLineNumber & " > >" & $LOGFILE & " < <" & @CRLF)
		Case StringInStr($CmdLine[$X], "debug") > 0
			$DEBUG = True
		Case Else
			_Debug("Unknown cmdline option found: > >" & $CmdLine[$X] & " < <", True)
			Exit
	EndSelect
Next
FileWriteLine($LOGFILE, "-----------------------------------------------------------")
FileWriteLine($LOGFILE, StringFormat("%s Startup %s %s %s %s %s %s", _SystemLocalTime(), $ProgramName, $FileVersion, @OSVersion, @OSServicePack, @OSType, @OSArch))

FileWriteLine($LOGFILE, " Command line arguments: " & $CmdLineRaw)

Global $ArrayRAll[1]
Global $ArrayRActive[1]
Global $ArrayS[1]
Global $ArrayC[1]
Global $ArrayZ[1][2]
Global $ShortPath = FileGetShortName(@ScriptDir)
Global $DEVCON

If @OSArch = "X86" Then
	$DEVCON = $UtilPath & "devcon_x86.exe"
	If Not FileExists($DEVCON) Then
		MsgBox(48, "Devcon error X86", $DEVCON & " must exist in script directory")
		Exit
	EndIf
ElseIf @OSArch = "X64" Then
	$DEVCON = $UtilPath & "devcon_x64.exe"
	If Not FileExists($DEVCON) Then
		MsgBox(48, "Devcon error X64", $DEVCON & " must exist in script directory")
		Exit
	EndIf
Else
	MsgBox(48, " OS Arch type not supported ", @OSArch)
	Exit
EndIf
$DEVCON = FileGetShortName($DEVCON)

If Not FileExists($DEVCON) Then
	_Debug(@ScriptLineNumber & "  " & $DEVCON & " must exist in AUXFiles ", 0x010)
	Exit
EndIf


Const $DEVCONHELPFILENAME = FileGetShortName($NICViewerFiles & "\DevconHelp.txt")
If FileExists($DEVCONHELPFILENAME) = True Then FileDelete($DEVCONHELPFILENAME)

Global $Z[1]
$Z = GetDevconData('/help')
_FileWriteFromArray($DEVCONHELPFILENAME, $Z)
If FileGetSize($DEVCONHELPFILENAME) < 1000 Then
	_Debug(@ScriptLineNumber & "  Devcon not found", True)
	Exit
EndIf

If _FileReadToArray($DEVIDFILE, $DEVID_array) <> 1 Then
	MsgBox(48, "Device ID file read Error", "Device ID file read Error" & @CRLF & $DEVIDFILE)
	Exit
EndIf

Const $RESULTFILEALL = FileGetShortName($NICViewerFiles) & "resultsAll.txt"
Const $RESULTFILEACTIVE = FileGetShortName($NICViewerFiles) & "resultsActive.txt"
Const $STATUSFILE = FileGetShortName($NICViewerFiles) & "status.txt"
Const $COMBINEFILE = FileGetShortName($NICViewerFiles) & "combine.txt"
Const $DUMPFILE = FileGetShortName($NICViewerFiles) & $ProgramName & ".dmp"
Const $RESCANFILE = FileGetShortName($NICViewerFiles) & "rescan.txt"
Const $GETMOREDATA = FileGetShortName($NICViewerFiles) & "GetMoreData.txt"
Const $NODEINFOFILE = FileGetShortName($NICViewerFiles) & "nodeinfo.txt"
Global $ArrayConfigData[1]
Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, _
		$WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
; Device Main --------------------------
Global $MainForm = GUICreate(@ScriptName & $FileVersion, 800, 500, 10, 10, $MainFormOptions)
GUISetFont(10, 400, -1, "Courier new")

Global $ButtonRunDevcon = GUICtrlCreateButton("Run", 10, 10, 100)
GUICtrlSetTip(-1, "Run devcon.exe to build a list of all devices")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ButtonRescan = GUICtrlCreateButton("Rescan", 120, 10, 100)
GUICtrlSetTip(-1, "Rescan for hardware changes and then run devcon.exe to build a list of all devices")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonDump = GUICtrlCreateButton("Dump", 230, 10, 100)
GUICtrlSetTip(-1, "Dump lists to file " & $DUMPFILE)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ButtonViewEditMain = GUICtrlCreateButton("View\Edit", 340, 10, 100)
GUICtrlSetTip(-1, "View\Edit the config  or other text file")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonAbout = GUICtrlCreateButton("About", 450, 10, 100)
GUICtrlSetTip(-1, "About the program and some Debug stuff")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonHelp = GUICtrlCreateButton("Help", 560, 10, 100)
GUICtrlSetTip(-1, "Display help information")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonExit = GUICtrlCreateButton("Exit", 700, 10, 100)
GUICtrlSetTip(-1, "Exit the program")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ButtonAwatch = GUICtrlCreateButton("Awatch", 700, 50, 100)
GUICtrlSetTip(-1, "TEST WHATEVER")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ButtonSearch = GUICtrlCreateButton("Search", 10, 50, 100)
GUICtrlSetTip(-1, "Search the list")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $InputSearch = GUICtrlCreateInput("", 120, 50, 100)
GUICtrlSetTip(-1, "String to search for")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $CheckShowIntelOnly = GUICtrlCreateCheckbox("Show Intel only", 250, 50)
GUICtrlSetTip(-1, "Show devices with 'VEN_8086' only")
GUICtrlSetState(-1, $GUI_CHECKED)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $CheckShowRunning = GUICtrlCreateCheckbox("Show running only", 250, 75)
GUICtrlSetTip(-1, "Show running devices only")
GUICtrlSetState(-1, $GUI_CHECKED)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $CheckHideNoMatch = GUICtrlCreateCheckbox("Hide 'no match'", 450, 50)
GUICtrlSetTip(-1, "Hide devices with no match in device ID file")
GUICtrlSetState(-1, $GUI_CHECKED)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $CheckShowPCIOnly = GUICtrlCreateCheckbox("Show PCI/PCIe only", 450, 75)
GUICtrlSetTip(-1, "Show devices with 'PCI' only")
GUICtrlSetState(-1, $GUI_CHECKED)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ListView = GUICtrlCreateListView("ID|Device manager name|Status|Other", 10, 100, 780, 360, $LVS_REPORT, BitOR($LVS_EX_FULLROWSELECT, $WS_EX_CLIENTEDGE, $LVS_EX_GRIDLINES))
GUICtrlSetTip(-1, "This is the list box")
GUICtrlSetResizing(-1, BitOR($GUI_DOCKTOP, $GUI_DOCKBOTTOM))
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 0, 250)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 1, 250)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 2, 150)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 3, 200)

Global $LabelStatusMain = GUICtrlCreateLabel("Status main", 10, 145 + 330, 780, 20, $SS_SUNKEN) ; 800, 500, 10, 10
GUICtrlSetTip(-1, "Display main window status")
GUICtrlSetResizing(-1, $GUI_DOCKSTATEBAR)

Global $DeviceFormInformation = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, _
		$WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
; Device Install --------------------------
Global $DeviceInformation = GUICreate("Device information", 800, 500, 10, 10, $DeviceFormInformation)
GUISetFont(10, 400, 0, "Courier New")
Global $ButtonRefresh = GUICtrlCreateButton("Refresh", 90, 10, 80, 30)
GUICtrlSetTip(-1, "Refresh the display")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonHelpDEV = GUICtrlCreateButton("Help", 170, 10, 80, 30)
GUICtrlSetTip(-1, "Someday I will put help here")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonEditDeviceInstall = GUICtrlCreateButton("View\Edit", 250, 10, 100, 30)
GUICtrlSetTip(-1, "View\Edit the config  or other text file")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonAboutDEV = GUICtrlCreateButton("About", 350, 10, 80, 30)
GUICtrlSetTip(-1, "About the program and some Debug stuff")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonDone = GUICtrlCreateButton("Done", 430, 10, 80, 30)
GUICtrlSetTip(-1, "Close this window and return to the main window")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $LabelDeviceIDLabel = GUICtrlCreateLabel("Device ID", 30, 50, 80, 20, $SS_SUNKEN)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $LabelDeviceID = GUICtrlCreateLabel("", 120, 50, 600, 20, $SS_SUNKEN)
GUICtrlSetTip(-1, "Device ID label")
GUICtrlSetResizing(-1, $GUI_DOCKALL - $GUI_DOCKWIDTH)

Global $EditDeviceResults = GUICtrlCreateEdit("Results", 10, 110, 780, 360, $WS_THICKFRAME + $WS_VSCROLL + $WS_HSCROLL)
GUICtrlSetTip(-1, "Results Edit box") ; 800, 500, 10, 10,
GUICtrlSetResizing(-1, $GUI_DOCKBOTTOM + $GUI_DOCKTOP + $GUI_DOCKLEFT)

;Global $LabelStatusMain
Global $LabelStatusInstall = GUICtrlCreateLabel("Status install", 10, 145 + 330, 780, 20, $SS_SUNKEN) ; 800, 500, 10, 1
GUICtrlSetTip(-1, "Display install window status")
GUICtrlSetResizing(-1, $GUI_DOCKSTATEBAR)

GUISetState(@SW_HIDE, $DeviceInformation)

GUISetHelp("notepad.exe " & @ScriptDir & "\AUXFiles\DeviceViewer.txt", $MainForm)

_Debug("DBGVIEWCLEAR")

HotKeySet("{F10}", "PositionWindows")
HotKeySet("{F11}", "Toggle")

GUISetState(@SW_SHOW, $MainForm)
WinMove("DeviceViewer", "", (@DesktopWidth - 900) / 2, 5, 900, 500) ;   800, 500, 10, 10
PositionWindows()
GUISetHelp("notepad.exe " & @ScriptDir & $NICVIEWHELPFILE, $MainForm)

If $DEBUG = False Then
	GetData()
	FilterData()
EndIf

;-----------------------------------------------
While 1
	Global $msg = GUIGetMsg(True)
	Switch $msg[0]
		Case $ButtonRunDevcon
			_GUICtrlListView_DeleteAllItems(GUICtrlGetHandle($ListView))
			GetData()
			FilterData()
			_ListView_Sort()
		Case $ButtonRescan
			_GUICtrlListView_DeleteAllItems(GUICtrlGetHandle($ListView))
			Rescan()
			GetData()
			FilterData()
			_ListView_Sort()
		Case $ButtonDump
			DumpToFile()
		Case $ButtonSearch
			Search()
		Case $InputSearch
			Search()
		Case $ButtonHelp
			Help()
		Case $ButtonDone
			GUISetState(@SW_HIDE, $DeviceInformation)
			GUISetState(@SW_SHOW, $MainForm)
		Case $ButtonViewEditMain
			ViewEdit()
		Case $ButtonEditDeviceInstall
			ViewEdit()
		Case $ButtonAbout
			About($ProgramName)
		Case $ButtonRefresh
			GetNodeInformation()
		Case $ButtonAboutDEV
			About("Device Install")
		Case $ButtonHelpDEV
			_Debug(@ScriptName & @CRLF & $FileVersion & @CRLF & " Might put help here someday!", 0x40, 15)

		Case $LabelDeviceID
			ClipPut(GUICtrlRead($LabelDeviceID))
		Case $EditDeviceResults
			ClipPut(GUICtrlRead($EditDeviceResults))
		Case $GUI_EVENT_CLOSE
			If $msg[1] = $MainForm Then Exit
			If $msg[1] = $DeviceInformation Then
				GUISetState(@SW_HIDE, $DeviceInformation)
				GUISetState(@SW_SHOW, $MainForm)
			EndIf
		Case $ButtonExit
			ExitLoop
		Case $ButtonAwatch
			ShellExecute($UtilPath & 'awatch.exe')
			;_ListView_Sort()

		Case $ListView
	EndSwitch
WEnd

;================================================================================
;Listview column sort and double click
Func WM_NOTIFY($hWnd, $iMsg, $iwParam, $ilParam)
	#forceref $hWnd, $iMsg, $iwParam, $ilParam
	Local $hWndFrom, $iCode, $tNMHDR

	$tNMHDR = DllStructCreate($tagNMHDR, $ilParam)
	$hWndFrom = HWnd(DllStructGetData($tNMHDR, "hWndFrom"))
	$iCode = DllStructGetData($tNMHDR, "Code")
	If @error Then Return

	Local $tInfo
	Local $ColumnIndex

	Switch $hWndFrom
		Case GUICtrlGetHandle($ListView)
			Switch $iCode
				Case $LVN_COLUMNCLICK
					$tInfo = DllStructCreate($tagNMLISTVIEW, $ilParam)
					$ColumnIndex = DllStructGetData($tInfo, "SubItem")
					_Debug(@ScriptLineNumber & "  " & $ColumnIndex)
					_ListView_Sort($ColumnIndex)
				Case Else
					$tInfo = DllStructCreate($tagNMLISTVIEW, $ilParam)
					$ColumnIndex = DllStructGetData($tInfo, "SubItem")
					Local $result = DllStructGetData($tInfo, 3)
					If $result = -114 Then
						_Debug(@ScriptLineNumber & "  " & $iCode & "  " & $result)
						DeviceInformation()
					EndIf
			EndSwitch
	EndSwitch

	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_NOTIFY

;-----------------------------------------------
;Listview column sort part 2
Func _ListView_Sort($cIndex = 0)
	Local $iColumnsCount, $iDimension, $iItemsCount, $aItemsTemp, $aItemsText, $iCurPos, $iImgSummand, $i, $j
	$iColumnsCount = _GUICtrlListView_GetColumnCount($ListView)
	$iDimension = $iColumnsCount * 2
	$iItemsCount = _GUICtrlListView_GetItemCount($ListView)

	If $iItemsCount < 1 Then Return

	Local $aItemsTemp[1][$iDimension]

	For $i = 0 To $iItemsCount - 1
		$aItemsTemp[0][0] += 1
		ReDim $aItemsTemp[$aItemsTemp[0][0] + 1][$iDimension]
		$aItemsText = _GUICtrlListView_GetItemTextArray($ListView, $i)
		$iImgSummand = $aItemsText[0] - 1
		For $j = 1 To $aItemsText[0]
			$aItemsTemp[$aItemsTemp[0][0]][$j - 1] = $aItemsText[$j]
			$aItemsTemp[$aItemsTemp[0][0]][$j + $iImgSummand] = _GUICtrlListView_GetItemImage($ListView, $i, $j - 1)
		Next
	Next

	$iCurPos = $aItemsTemp[1][$cIndex]
	_ArraySort($aItemsTemp, 0, 1, 0, $cIndex)
	If StringInStr($iCurPos, $aItemsTemp[1][$cIndex]) Then _ArraySort($aItemsTemp, 1, 1, 0, $cIndex)
	For $i = 1 To $aItemsTemp[0][0]
		For $j = 1 To $iColumnsCount
			_GUICtrlListView_SetItemText($ListView, $i - 1, $aItemsTemp[$i][$j - 1], $j - 1)
			_GUICtrlListView_SetItemImage($ListView, $i - 1, $aItemsTemp[$i][$j + $iImgSummand], $j - 1)
		Next
	Next
EndFunc   ;==>_ListView_Sort
;-----------------------------------------------
; hotkey F10
Func PositionWindows()
	WinMove("NICViewer", "", (@DesktopWidth - 900) / 2, 5, 900, 500)
	;MoveWindows()
EndFunc   ;==>PositionWindows
;-----------------------------------------------
; hotkey F11
Func Toggle()
	_Debug(@ScriptLineNumber & " Toggle")
	GuiDisable("Toggle")
EndFunc   ;==>Toggle
;-----------------------------------------------
; Takes in a devcon command and returns an array of results
Func GetDevconData($command)
	Local $foo = Run(@ComSpec & " /c " & $DEVCON & " " & $command, " .", @SW_HIDE, $STDOUT_CHILD)

	;ConsoleWrite(@ScriptLineNumber & " " & $DEVCON & "  " & $command & @CRLF)

	Local $A[1]
	Local $line
	While 1
		$line = StdoutRead($foo)
		If @error Then ExitLoop
		If StringLen($line) > 0 Then _ArrayAdd($A, $line)
	WEnd

	Local $C[1]
	Local $E[1]
	For $B In $A
		$C = StringSplit($B, @CRLF, 2)
		For $D In $C
			If StringLen($D) > 0 Then _ArrayAdd($E, $D)
		Next
	Next

	_RemoveBlankLines($E)
	Return $E
EndFunc   ;==>GetDevconData
;-----------------------------------------------
;This function gets the data from DEVCON.EXE
Func GetData()
	SetStatus("Getdata")
	;_Debug(@ScriptLineNumber & "   " & _GUICtrlListView_GetItemCount($ListView) & @CRLF)
	For $X = 0 To _GUICtrlListView_GetItemCount($ListView)
		_GUICtrlListView_SetItemText($ListView, $X, "", 0)
		_GUICtrlListView_SetItemText($ListView, $X, "", 1)
		_GUICtrlListView_SetItemText($ListView, $X, "", 2)
		_GUICtrlListView_SetItemText($ListView, $X, "", 3)
	Next
	FileDelete($STATUSFILE)
	FileDelete($RESULTFILEALL)
	FileDelete($RESULTFILEACTIVE)
	FileDelete($COMBINEFILE)
	FileDelete($GETMOREDATA)
	FileDelete($NODEINFOFILE)

	GuiDisable("Disable")

	Local $Z[1]
	$Z = GetDevconData('FIND *')
	_FileWriteFromArray($RESULTFILEACTIVE, $Z)
	$Z = GetDevconData('FindAll *')
	_FileWriteFromArray($RESULTFILEALL, $Z)
	$Z = GetDevconData('status *')
	_FileWriteFromArray($STATUSFILE, $Z)
	$Z = GetDevconData('driverfiles *')
	_FileWriteFromArray($GETMOREDATA, $Z)

	; The following gets and parses the data that was output from devcon
	ReDim $ArrayC[1]
	ReDim $ArrayZ[1][2]
	ResultsAll()
	ResultsActive()
	Status()
	;This builds a hash of the items. The name is the key and the count is the data.
	;The hash is a one dimensional array
	Local $TmpArray[2] ;Create 2 slots so that a crash with no data is avoided
	Local $InArray[1]
	_ArrayConcatenate($InArray, $ArrayRAll)
	_ArrayConcatenate($InArray, $ArrayRActive)
	_ArraySort($InArray)
	_ArrayDelete($InArray, 0)
	For $X = 0 To UBound($InArray) - 1
		Local $T = StringStripWS($InArray[$X], 3)
		Local $POS = _ArraySearch($TmpArray, $T)
		If $POS = -1 Then
			_ArrayAdd($TmpArray, $T)
			_ArrayAdd($TmpArray, 1)
		Else
			$TmpArray[$POS + 1] += 1
		EndIf
	Next
	;Convert the hash to a two dimensional array
	Local $count = 0
	While UBound($TmpArray) > 0
		$ArrayZ[$count][0] = _ArrayPop($TmpArray)
		$ArrayZ[$count][1] = _ArrayPop($TmpArray)
		$count += 1
		ReDim $ArrayZ[UBound($ArrayZ) + 1][2]
	WEnd
	;Local $ACombine[1]
	For $X = 0 To UBound($ArrayZ) - 1; this is the active list
		If $ArrayZ[$X][0] = 1 Then
			_ArrayAdd($ArrayC, $ArrayZ[$X][1] & ": Device not present")
		EndIf
	Next

	; The final data is in ArrayC
	GuiDisable("Enable")
	SetStatus("Getdata complete")
EndFunc   ;==>GetData
;-----------------------------------------------
;This function filters the combined data per user choices
Func FilterData()
	SetStatus("FilterData")
	GuiDisable("Disable")
	FileWriteLine($LOGFILE, _SystemLocalTime() & "Filtering data")
	If UBound($ArrayC) < 5 Then
		GetData()
	EndIf
	_GUICtrlListView_DeleteAllItems(GUICtrlGetHandle($ListView))

	_Debug("DBGVIEWCLEAR")

	Local $NICViewer_dev
	_FileReadToArray($DEVIDFILE, $NICViewer_dev)

	For $DevStringIndex = 0 To UBound($ArrayC) - 1
		Local $DevStringArray = StringSplit($ArrayC[$DevStringIndex], ":")
		If $DevStringArray[0] <> 3 Then
			_ArrayAdd($DevStringArray, "?????")
		EndIf

		Local $DeviceName = $DevStringArray[1]
		Local $DevStringArray_SplitA = _ArrayToString($DevStringArray, "~")

		Local $DevStringArray_SplitB = StringSplit($DevStringArray[1], '&')
		If $DevStringArray_SplitB[0] = 1 Then ContinueLoop
		Local $NICViewer_Name = ""

		For $NICViewer_devString In $NICViewer_dev
			If StringInStr(StringReplace($NICViewer_devString, "&", ""), $DevStringArray_SplitB[2]) > 0 Then
				Local $N = StringSplit($NICViewer_devString, "~")
				If $N[0] = 2 Then
					$NICViewer_Name = StringStripWS($N[2], 3)
				Else
					ConsoleWrite(@ScriptLineNumber & " Bad array value: " & _ArrayToString($N) & @CRLF)
				EndIf
				ExitLoop
			EndIf
		Next

		If GUICtrlRead($CheckShowPCIOnly) = $GUI_CHECKED And StringInStr($DeviceName, 'PCI\VEN') = 0 Then ContinueLoop
		If GUICtrlRead($CheckShowIntelOnly) = $GUI_CHECKED And StringInStr($DevStringArray_SplitA, 'PCI\VEN_8086') = 0 Then ContinueLoop
		If GUICtrlRead($CheckShowRunning) = $GUI_CHECKED And StringInStr($DevStringArray_SplitA, 'Driver is running') = 0 Then ContinueLoop
		If GUICtrlRead($CheckHideNoMatch) = $GUI_CHECKED And $NICViewer_Name = '' Then ContinueLoop

		Local $ListView_item = _GUICtrlListView_AddItem($ListView, StringStripWS($DevStringArray[1], 3))
		_GUICtrlListView_AddSubItem($ListView, $ListView_item, StringStripWS($DevStringArray[2], 3), 1)
		_GUICtrlListView_AddSubItem($ListView, $ListView_item, StringStripWS($DevStringArray[3], 3), 2)
		_GUICtrlListView_AddSubItem($ListView, $ListView_item, $NICViewer_Name, 3)

		;ConsoleWrite(@ScriptLineNumber & " " & _ArrayToString($DevStringArray) & " " & @CRLF)
	Next

	GuiDisable("Enable")
	SetStatus("FilterData complete")
EndFunc   ;==>FilterData
;-----------------------------------------------
;This function displays the device Install window where information may be viewed
Func DeviceInformation()
	GuiDisable("Disable")
	SetStatus("DeviceInformation")
	_Debug(@ScriptLineNumber & " DeviceInformation")
	GUISetState(@SW_HIDE, $MainForm)
	GUISetState(@SW_SHOW, $DeviceInformation)

	Local $DevID
	Local $DevName
	Local $DevStatus
	For $A = 0 To _GUICtrlListView_GetItemCount($ListView) - 1
		If _GUICtrlListView_GetItemSelected($ListView, $A) Then
			$DevID = _GUICtrlListView_GetItemText($ListView, $A, 0)
			$DevName = _GUICtrlListView_GetItemText($ListView, $A, 1)
			$DevStatus = _GUICtrlListView_GetItemText($ListView, $A, 2)
		EndIf
	Next
	Local $TL = ''
	_Debug(@ScriptLineNumber & " Data to work with:" & @CRLF & $DevID & @CRLF & $DevName & @CRLF & $DevStatus)
	If StringInStr($DevStatus, "Device not present") = 0 Then
		Local $GMD
		_FileReadToArray($GETMOREDATA, $GMD)
		ConsoleWrite(@ScriptLineNumber & ": " & _ArrayToString($GMD) & @CRLF)

		Local $V
		;Local $Start = False
		Local $count = 1
		While 1
			If $count >= UBound($GMD) Then ExitLoop
			$V = $GMD[$count]
			$count += 1

			If StringInStr($DevID, $V) = 1 Then
				;$Start = True
				$TL = StringStripWS($V, 3)
				GUICtrlSetData($LabelDeviceID, $TL)
			EndIf
		WEnd
	Else
		GUICtrlSetData($LabelDeviceID, $DevID)
	EndIf

	$TL = GetNodeInformation()
	GUICtrlSetData($EditDeviceResults, "Current device information: " & @CRLF & $TL)

	GuiDisable("Enable")
	SetStatus("InstallDrivers complete")
EndFunc   ;==>DeviceInformation
;----------------------------------------------

Func GetNodeInformation()
	SetStatus(" GetNodeInformation")
	GuiDisable("Disable")
	FileDelete($NODEINFOFILE)

	Local $DeviceID = GUICtrlRead($LabelDeviceID)
	If StringLen($DeviceID) < 5 Then
		MsgBox(48, "GetNodeInformation error", "Invalid device ID" & @CRLF & $DeviceID)
		GuiDisable("Enable")
		Return
	EndIf

	Local $RR

	If StringInStr($DeviceID, "DEV_") > 0 Then
		$RR = StringSplit($DeviceID, "&")
	Else
		$RR = StringSplit($DeviceID, "\")
	EndIf

	If UBound($RR) > 2 Then
		_Debug(@ScriptLineNumber & "   " & $RR[2])
		Local $parms1 = ' drivernodes *' & $RR[2] & '* > ' & $NODEINFOFILE
		_Debug(@ScriptLineNumber & " " & $parms1)
		_Debug(@ScriptLineNumber & " Devcon 5 " & _RunDOS($DEVCON & $parms1))
	EndIf
	Local $TA
	_FileReadToArray($NODEINFOFILE, $TA)
	_ArrayDelete($TA, 0)

	Local $TL
	If IsArray($TA) Then
		For $T In $TA
			$TL = $TL & @CRLF & $T
		Next
	Else
		$TL = $TL & @CRLF & "No node information found"
	EndIf
	_Debug(@ScriptLineNumber & "  " & $TL)
	GuiDisable("Enable")
	SetStatus(" GetNodeInformation complete")
	Return $TL
EndFunc   ;==>GetNodeInformation

;-----------------------------------------------
;This function processes the all device file from DEVCON.EXE
Func ResultsAll()
	SetStatus("ResultsAll")
	Local $file = FileOpen($RESULTFILEALL, 0)
	; Check if file opened for reading OK
	If $file = -1 Then
		_Debug("Unable to open file for reading: " & $RESULTFILEALL, 0x10, 5)
		Return
	EndIf
	Local $tmp = FileRead($file)
	FileClose($file)
	$ArrayRAll = StringSplit($tmp, @CRLF, 3)
	_TrimArray($ArrayRAll)
	_TrimArray($ArrayRAll)
	_ArraySort($ArrayRAll)
	_ArrayDelete($ArrayRAll, 0)
	_ArrayDelete($ArrayRAll, 0)
	For $RR = 0 To UBound($ArrayRAll) - 1
		Local $SS = StringSplit($ArrayRAll[$RR], ":")
		If $SS[0] = 1 Then
			$ArrayRAll[$RR] = $ArrayRAll[$RR] & " : No name found "
		EndIf
	Next
	FileWriteLine($LOGFILE, _SystemLocalTime() & "All data found: " & UBound($ArrayRAll))
	SetStatus("ResultsAll complete")
EndFunc   ;==>ResultsAll
;-----------------------------------------------
;This function processes the active device file from DEVCON.EXE
Func ResultsActive()
	SetStatus("ResultsActive")
	; We open the results info here
	Local $file = FileOpen($RESULTFILEACTIVE, 0)
	; Check if file opened for reading OK
	If $file = -1 Then
		_Debug("Unable to open file for reading: " & $RESULTFILEACTIVE, 0x10, 5)
		Return
	EndIf
	Local $tmp = FileRead($file)
	FileClose($file)
	$ArrayRActive = StringSplit($tmp, @CRLF, 3)
	_TrimArray($ArrayRActive)
	_ArraySort($ArrayRActive)
	_ArrayDelete($ArrayRActive, 0)
	_ArrayDelete($ArrayRActive, 0)
	;_ArrayDisplay($ArrayRActive, "$ArrayRActive " & @ScriptLineNumber)
	For $RR = 0 To UBound($ArrayRActive) - 1
		Local $SS = StringSplit($ArrayRActive[$RR], ":")
		If $SS[0] = 1 Then
			$ArrayRActive[$RR] = $ArrayRActive[$RR] & " : No name found "
		EndIf
	Next
	;_ArrayDisplay($ArrayRActive, "$ArrayRActive " & @ScriptLineNumber)
	FileWriteLine($LOGFILE, _SystemLocalTime() & "Active data found: " & UBound($ArrayRActive))
	SetStatus("ResultsActive complete")
EndFunc   ;==>ResultsActive
;-----------------------------------------------
;This function processes the status file from DEVCON.EXE
Func Status()
	SetStatus("Status")
	; We open the status info here
	Local $file = FileOpen($STATUSFILE, 0)
	; Check if file opened for reading OK
	If $file = -1 Then
		_Debug("Unable to open file for reading: " & $STATUSFILE, 0x10, 5)
		Return
	EndIf
	Local $tmp = FileRead($file)
	;ConsoleWrite(@ScriptLineNumber & " " & $tmp & @CRLF)
	FileClose($file)
	$ArrayS = StringSplit($tmp, @CRLF, 3)
	; This routine fixes a problem with 'Device has a problem: 28.'  removes the : and the .
	Local $count = 0
	While True
		$count += 1
		If $count >= UBound($ArrayS) Then ExitLoop
		$ArrayS[$count] = StringReplace($ArrayS[$count], "Device has a problem: 28.", "Device has a problem = 28")
		_Debug(@ScriptLineNumber & "  " & $count & "  " & $ArrayS[$count] & @CRLF)
	WEnd
	; This should fix a problem that occurs when no name is found in the status listing
	$count = 0
	While True
		$count += 1
		If Not IsArray($ArrayS) Then ExitLoop
		;_ArrayDisplay($ArrayS, @ScriptLineNumber)
		If StringInStr($ArrayS[$count], 'matching device') <> 0 Then ExitLoop
		If StringInStr($ArrayS[$count], "   ") = 0 And StringInStr($ArrayS[$count + 1], "Name:") = 0 Then
			_ArrayInsert($ArrayS, $count + 1, "Name: Name not found")
			$count += 1
			_Debug(@ScriptLineNumber & "  " & $count & "  " & $ArrayS[$count])
		EndIf
		;Debugviewer(@ScriptLineNumber & "  " & $Count)
	WEnd
	;_ArrayDisplay($ArrayS, "$ArrayS " & @ScriptLineNumber)
	$count = 0
	While True
		$count += 1
		Local $TString ; This is a temporary storage string
		If $count >= UBound($ArrayS) Then ExitLoop
		$TString = $ArrayS[$count]
		If StringInStr($TString, "matching device") > 0 Then ExitLoop
		If StringInStr($TString, "   ") = 0 Then
			$count += 1
			If StringInStr($ArrayS[$count], "Name:") = 0 Then
				$TString = $TString & ": Name: Name not found"
			Else
				$TString = $TString & ":" & $ArrayS[$count]
			EndIf
			$count += 1
			$TString = $TString & ":" & $ArrayS[$count]
			$TString = StringReplace($TString, "Name:", "")
			_ArrayAdd($ArrayC, $TString)
		EndIf
	WEnd
	_ArrayDelete($ArrayC, 0)
	_ArraySort($ArrayC)
	;_ArrayDisplay($ArrayC, "$ArrayC " & @ScriptLineNumber)
	FileWriteLine($LOGFILE, _SystemLocalTime() & "Combined data found: " & UBound($ArrayC))
	SetStatus("Status complete")
EndFunc   ;==>Status
;-----------------------------------------------
;This function allows the user to edit or view any file, useful for changing the config file
Func ViewEdit($FileName = $DEVIDFILE)
	SetStatus("ViewEdit " & $FileName)
	GuiDisable("Disable")
	If $FileName = '' Then
		$FileName = FileOpenDialog("View or Edit a file", @ScriptDir, "All (*.*)", 1);, $CONFIGFILENAME)
	EndIf

	Local $editor = _ChoseTextEditor()

	SetStatus("ViewEdit " & $FileName & " complete")
	ShellExecute($editor, $FileName)
	GuiDisable("Enable")
	SetStatus("ViewEdit " & $FileName & " complete")
EndFunc   ;==>ViewEdit
;-----------------------------------------------
;This function re-scans the hardware looking for hardware changes
Func Rescan()
	SetStatus("Rescan")
	GuiDisable("Disable")
	_GUICtrlListView_DeleteAllItems($ListView)
	;Local $result = _RunDOS($DEVCON & ' rescan >> ' & $RESCANFILE)
	_RunDOS($DEVCON & ' rescan >> ' & $RESCANFILE)
	;_Debug("Rescan for new hardware complete  " & $result, True)
	FileWriteLine($LOGFILE, _SystemLocalTime() & "Rescan completed")
	GuiDisable("Enable")
	SetStatus("Rescan complete")
EndFunc   ;==>Rescan
;-----------------------------------------------
Func Search()
	Static Local $iStart = -1
	SetStatus(" Search " & $iStart)

	Local $SS = GUICtrlRead($InputSearch)
	Local $T = _GUICtrlListView_FindInText($ListView, $SS, $iStart, True)
	$iStart = $T

	If $iStart = -1 Then
		SetStatus(" Search for " & $SS & " not found.")
		_GUICtrlListView_ClickItem($ListView, 0)
		Return
	EndIf
	_GUICtrlListView_EnsureVisible($ListView, $T)
	_GUICtrlListView_ClickItem($ListView, $T)
	SetStatus(" Search complete. Found on line " & $iStart)
EndFunc   ;==>Search
;-----------------------------------------------
;This function dumps the combined array to a disk file
Func DumpToFile()
	SetStatus("DumpToFile")
	Local $file = FileOpen($DUMPFILE, 2)
	_FileWriteFromArray($file, $ArrayC)
	FileClose($file)
	_Debug("Dump complete" & @CRLF & "All devcon data sent to: " & $DUMPFILE, 5)
	FileWriteLine($LOGFILE, _SystemLocalTime() & "Data dumped to file " & $DUMPFILE)
	SetStatus(" DumpToFile complete")
EndFunc   ;==>DumpToFile
;-----------------------------------------------
;This function enables or disables the GUI items
Func GuiDisable($choice) ;@SW_ENABLE @SW_disble
	;_Debug(@ScriptLineNumber & " GuiDisable " & $choice)
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
		_Debug("Invalid choice at GuiDisable" & $choice & "   " & $setting, 0x40)
		Return
	EndIf
	$LastState = $setting
	For $X = 1 To 100
		GUICtrlSetState($X, $setting)
	Next
EndFunc   ;==>GuiDisable
;-----------------------------------------------
;This function displays About info and some Debug stuff
Func About(Const $FormID)
	SetStatus("About " & $FormID)
	Local $D = WinGetPos($FormID)
	Local $WinPos
	If IsArray($D) = True Then
		;ConsoleWrite(@ScriptLineNumber & $FormID & @CRLF)
		$WinPos = StringFormat("%s" & @CRLF & "WinPOS: %d  %d " & @CRLF & "WinSize: %d %d " & @CRLF & "Desktop: %d %d ", _
				$FormID, $D[0], $D[1], $D[2], $D[3], @DesktopWidth, @DesktopHeight)
	Else
		$WinPos = ">>>About ERROR, Check the window name<<<"
	EndIf
	_Debug(@CRLF & $SystemS & @CRLF & $WinPos & @CRLF & "Written by Doug Kaynor!", 0x40)
EndFunc   ;==>About
;-----------------------------------------------
Func SetStatus($msg)
	GUICtrlSetData($LabelStatusMain, $msg)
	GUICtrlSetData($LabelStatusInstall, $msg)
	;_Debug(@ScriptLineNumber & " " & $msg)
EndFunc   ;==>SetStatus
;-----------------------------------------------
Func Help()
	_Debug(@ScriptLineNumber & " Help")
	Local $helpstr = 'Command line startup options:' & @CRLF & @CRLF & _
			"help or ? Help information is displayed" & @CRLF & _
			"debug   Various debug functions" & @CRLF & _
			"logfile (the log file path must be wrapped in quotes)" & @CRLF & @CRLF & _
			"Hotkeys" & @CRLF & @CRLF & _
			"F10  = Windows resize and position" & @CRLF & _
			"F11 = Enable or Disable the GUI" & @CRLF
	_Debug(@ScriptName & @CRLF & $FileVersion & @CRLF & @CRLF & $helpstr, 0x40)
EndFunc   ;==>Help
;-----------------------------------------------

