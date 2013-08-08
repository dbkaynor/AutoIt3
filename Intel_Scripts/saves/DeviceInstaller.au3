#RequireAdmin
#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=../icons/Cryptkeeper.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Comment=A program to list system devices
#AutoIt3Wrapper_Res_Description=DeviceInstaller
#AutoIt3Wrapper_Res_Fileversion=2.0.0.87
#AutoIt3Wrapper_Res_FileVersion_AutoIncrement=Y
#AutoIt3Wrapper_Res_ProductVersion=666
#AutoIt3Wrapper_Res_LegalCopyright=Copyright © 2010 Douglas B Kaynor
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Field=Developer|Douglas Kaynor
#AutoIt3Wrapper_Res_Field=AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Res_Field=Compile date|%longdate% %time%
#AutoIt3Wrapper_Au3Check_Stop_OnWarning=y
#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Run_Tidy=y
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****
#CS
	Written by Doug Kaynor 03/10/2010
	Added support for path variables in the config file
	Fixed problem with no matching system name
	Added a message box for search results not found
	Fixed config validate error message
	Fixed getting non PCI device driver info
	Add windows resizing and positioning
	"Device Install box" show in the same location and size as the main window.
	LoadConfig will run when ever filter data
	Added status bars to both windows
	Added command line option "debug"
	Fixed refresh button error
	Added an install function for dot net 3.5
	Changed in to is
	Fixed Gui enables after several error conditions
	Added display of driver version info
	>>> Fix problem with blank config driver version
	>>> Check for .Net (HKLM\SOFTWARE|Microsoft)
#CE
#region
;#Tidy_Parameters=/gd /sf
#endregion
Opt("MustDeclareVars", 1) ; require pre-declared varibles
If _Singleton(@ScriptName, 1) = 0 Then
	_Debug(@ScriptName & " is already running!", 0x40)
	Exit
EndIf
#include <Array.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
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
DirCreate(@ScriptDir & "\DevInstallerFiles")
DirCreate(@ScriptDir & "\AUXFiles")
Global $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")
Global $tmp = StringSplit(@ScriptName, ".")
Global $ProgramName = $tmp[1]
Global $ResultLocation = ""
Global $SystemS = $ProgramName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & _
		@OSServicePack & @CRLF & @OSType & @CRLF & @OSArch
Global $SystemProductName
Global $OsName
Global $LOGFILE = FileGetShortName(@ScriptDir & "\AUXFiles\") & $ProgramName & ".log"
Global $NOGUI = False
Global $QUIET = False
Global $SILENT = False
Global $DEBUG = False
Global $DOTNETPATH

Global $hWnd, $iMsg, $iwParam, $ilParam
GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")

For $X = 1 To $CmdLine[0]
	ConsoleWrite($X & " >> " & $CmdLine[$X] & @CRLF)
	Select
		Case StringInStr($CmdLine[$X], "help") > 0 Or _
				StringInStr($CmdLine[$X], "?") > 0
			Help()
			Exit
		Case StringInStr($CmdLine[$X], "logfile") > 0
			Global $Y = StringSplit($CmdLine[$X], "=")
			$LOGFILE = FileGetShortName(StringStripWS($Y[2], 3))
			_debug(@ScriptLineNumber & " > >" & $LOGFILE & " < <" & @CRLF)
		Case StringInStr($CmdLine[$X], "nogui") > 0
			$NOGUI = True
		Case StringInStr($CmdLine[$X], "quiet") > 0
			$QUIET = True
		Case StringInStr($CmdLine[$X], "debug") > 0
			$DEBUG = True
		Case Else
			_debug("Unknown cmdline option found: > >" & $CmdLine[$X] & " < <", True)
			Exit
	EndSelect
Next
FileWriteLine($LOGFILE, "-----------------------------------------------------------")
FileWriteLine($LOGFILE, StringFormat("%s Startup %s %s %s %s %s %s", _SystemLocalTime(), $ProgramName, $FileVersion, @OSVersion, @OSServicePack, @OSType, @OSArch))
FileWriteLine($LOGFILE, _SystemLocalTime() & StringFormat("NOGUI:%s   Quiet:%s   Logfile:%s", $NOGUI, $QUIET, $LOGFILE))

FileWriteLine($LOGFILE, " Command line arguments: " & $CmdLineRaw)

Global $ArrayRAll[1]
Global $ArrayRActive[1]
Global $ArrayS[1]
Global $ArrayC[1]
Global $ArrayZ[1][2]
Global $ShortPath = FileGetShortName(@ScriptDir)
Const $DEVCON = "devcon_x86.exe"
Const $DEVCONHELP = FileGetShortName(@ScriptDir & "\DevInstallerFiles\") & "DevconHelp.txt"
FileDelete($DEVCONHELP)

_debug(@ScriptLineNumber & " Devcon " & _RunDOS($DEVCON & ' /help > ' & $DEVCONHELP))
If FileGetSize($DEVCONHELP) < 1e3 Then
	_debug($DEVCON & " must exist in path or current folder", 0x010)
	Exit
EndIf
Const $CONFIGFILENAME = FileGetShortName(@ScriptDir & "\AUXFiles\") & $ProgramName & ".cfg"
Const $RESULTFILEALL = FileGetShortName(@ScriptDir & "\DevInstallerFiles\") & "resultsAll.txt"
Const $RESULTFILEACTIVE = FileGetShortName(@ScriptDir & "\DevInstallerFiles\") & "resultsActive.txt"
Const $STATUSFILE = FileGetShortName(@ScriptDir & "\DevInstallerFiles\") & "status.txt"
Const $COMBINEFILE = FileGetShortName(@ScriptDir & "\DevInstallerFiles\") & "combine.txt"
Const $DUMPFILE = FileGetShortName(@ScriptDir & "\DevInstallerFiles\") & $ProgramName & ".dmp"
Const $RESCANFILE = FileGetShortName(@ScriptDir & "\DevInstallerFiles\") & "rescan.txt"
Const $GETMOREDATA = FileGetShortName(@ScriptDir & "\DevInstallerFiles\") & "GetMoreData.txt"
Const $NODEINFOFILE = FileGetShortName(@ScriptDir & "\DevInstallerFiles\") & "nodeinfo.txt"
Global $ArrayConfigData[1]
Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, _
		$WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
; Device Main --------------------------
Global $MainForm = GUICreate(@ScriptName & $FileVersion, 800, 500, 10, 10, $MainFormOptions)
GUISetFont(10, 400, -1, "Courier new")

Global $GroupPosition = GUICtrlCreateGroup("", 2, 0, 200, 135)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $RadioStopped = GUICtrlCreateRadio("Device is stopped", 10, 10)
GUICtrlSetTip(-1, "List currently stopped devices")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $RadioRunning = GUICtrlCreateRadio("Driver is running", 10, 30)
GUICtrlSetTip(-1, "List running drivers")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $RadioProblem = GUICtrlCreateRadio("Device has a problem", 10, 50)
GUICtrlSetTip(-1, "List devices that have problems")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $RadioNotPresent = GUICtrlCreateRadio("Device not present", 10, 70)
GUICtrlSetTip(-1, "List devices that are not present")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $RadioAllDevices = GUICtrlCreateRadio("All", 10, 90)
GUICtrlSetTip(-1, "List all devices")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $RadioConfig = GUICtrlCreateRadio("Config", 10, 110)
GUICtrlSetTip(-1, "List by config file")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlSetState($RadioConfig, $GUI_CHECKED)
Global $ButtonRunDevcon = GUICtrlCreateButton("Run devcon", 210, 5, 120)
GUICtrlSetTip(-1, "Run devcon.exe to build a list of all devices")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ButtonInstall = GUICtrlCreateButton("Install", 210, 35, 120)
GUICtrlSetTip(-1, "Install the selected device")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonRescan = GUICtrlCreateButton("Rescan", 210, 65, 120)
GUICtrlSetTip(-1, "Rescan for hardware changes")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonInstallDotNet = GUICtrlCreateButton("Install .net", 210, 95, 120)
GUICtrlSetTip(-1, "Install .net 3.5")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ComboFilter = GUICtrlCreateCombo("", 340, 10, 90) ; create first item
GUICtrlSetTip(-1, "Filter by device type")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetData(-1, "ALL|ACPI\|DISPLAY\|IDE\|PCI\|PCIIDE\|ROOT\|STORAGE\|UMB\|USB\|USBSTOR\|SW\|WPDBUSENUMROOT\|HTREE\", "ALL")
Global $ButtonSearch = GUICtrlCreateButton("Search", 340, 35, 90)
GUICtrlSetTip(-1, "Search the list")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $InputSearch = GUICtrlCreateInput("", 340, 65, 90)
GUICtrlSetTip(-1, "Data to search for")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonDump = GUICtrlCreateButton("Dump", 340, 95, 90)
GUICtrlSetTip(-1, "Dump lists to file " & $DUMPFILE)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $LabelSysInfo = GUICtrlCreateLabel("System", 440, 10, 350, 45, $SS_SUNKEN)
GUICtrlSetTip(-1, "Display system information")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonLoadConfig = GUICtrlCreateButton("Load Config", 440, 65, 100)
GUICtrlSetTip(-1, "Load a config file")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonViewEditMain = GUICtrlCreateButton("View\Edit", 540, 65, 80)
GUICtrlSetTip(-1, "View\Edit the config  or other text file")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonAbout = GUICtrlCreateButton("About", 620, 65, 50)
GUICtrlSetTip(-1, "About the program and some Debug stuff")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonHelp = GUICtrlCreateButton("Help", 670, 65, 40)
GUICtrlSetTip(-1, "Display help information")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonExit = GUICtrlCreateButton("Exit", 730, 65, 40)
GUICtrlSetTip(-1, "Exit the program")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
;  left, top [, width [, height [,
Global $ListView = GUICtrlCreateListView("ID|Device manager name|Status|Config file name", 10, 145, 780, 320, $LVS_REPORT, BitOR($LVS_EX_FULLROWSELECT, $WS_EX_CLIENTEDGE, $LVS_EX_GRIDLINES))

GUICtrlSetTip(-1, "This is the list box")
GUICtrlSetResizing(-1, BitOR($GUI_DOCKTOP, $GUI_DOCKBOTTOM))
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 0, 150)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 1, 175)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 2, 160)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 3, 260)

Global $LabelStatusMain = GUICtrlCreateLabel("Status main", 10, 145 + 330, 780, 20, $SS_SUNKEN) ; 800, 500, 10, 10
GUICtrlSetTip(-1, "Display main window status")
GUICtrlSetResizing(-1, $GUI_DOCKSTATEBAR)

Global $DeviceFormInstall = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, _
		$WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
; Device Install --------------------------
Global $DeviceInstall = GUICreate("Device Install", 800, 500, 10, 10, $DeviceFormInstall)
GUISetFont(10, 400, 0, "Courier New")
Global $ButtonInstallOK = GUICtrlCreateButton("Install", 10, 10, 80, 30)
GUICtrlSetTip(-1, "Do the driver install")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonRefresh = GUICtrlCreateButton("Refresh", 90, 10, 80, 30)
GUICtrlSetTip(-1, "Refresh the display")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonHelpDEV = GUICtrlCreateButton("Help", 170, 10, 80, 30)
GUICtrlSetTip(-1, "Someday I will put help here")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonEditDeviceInstall = GUICtrlCreateButton("View\Edit", 250, 10, 80, 30)
GUICtrlSetTip(-1, "View\Edit the config  or other text file")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonAboutDEV = GUICtrlCreateButton("About", 330, 10, 80, 30)
GUICtrlSetTip(-1, "About the program and some Debug stuff")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonDone = GUICtrlCreateButton("Done", 410, 10, 80, 30)
GUICtrlSetTip(-1, "Close this window and return to the main window")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $LabelDeviceIDLabel = GUICtrlCreateLabel("Device ID", 30, 50, 80, 20, $SS_SUNKEN)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $LabelDeviceID = GUICtrlCreateLabel("", 120, 50, 600, 20, $SS_SUNKEN)
GUICtrlSetTip(-1, "Device ID label")
GUICtrlSetResizing(-1, $GUI_DOCKALL - $GUI_DOCKWIDTH)

Global $NewDriverVersionLabel = GUICtrlCreateLabel("Driver version", 30, 80, 80, 20, $SS_SUNKEN)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $NewDriverVersionValue = GUICtrlCreateLabel("", 120, 80, 600, 20, $SS_SUNKEN)
GUICtrlSetResizing(-1, $GUI_DOCKALL - $GUI_DOCKWIDTH)


Global $EditDeviceResults = GUICtrlCreateEdit("Results", 10, 110, 780, 360, $WS_THICKFRAME + $WS_VSCROLL + $WS_HSCROLL)
GUICtrlSetTip(-1, "Results Edit box") ; 800, 500, 10, 10,
GUICtrlSetResizing(-1, $GUI_DOCKBOTTOM + $GUI_DOCKTOP + $GUI_DOCKLEFT)

;Global $LabelStatusMain
Global $LabelStatusInstall = GUICtrlCreateLabel("Status install", 10, 145 + 330, 780, 20, $SS_SUNKEN) ; 800, 500, 10, 1
GUICtrlSetTip(-1, "Display install window status")
GUICtrlSetResizing(-1, $GUI_DOCKSTATEBAR)

GUISetState(@SW_HIDE, $DeviceInstall)
#endregion ### END Koda GUI section ###

GUISetHelp("notepad.exe " & @ScriptDir & "\AUXFiles\DeviceInstaller.txt", $MainForm)

_Debug("DBGVIEWCLEAR")

HotKeySet("{F6}", "PositionWindows")
HotKeySet("{F7}", "ShowNodes")
HotKeySet("{F8}", "AutoRunSilentNoGUI")
HotKeySet("{F9}", "AutoRunSilent")
HotKeySet("{F10}", "AutoRunQuiet")
HotKeySet("{F11}", "Toggle")

If $NOGUI = True Then
	GUISetState(@SW_HIDE, $MainForm)
	AutoRunSilentNoGUI()
Else
	GUISetState(@SW_SHOW, $MainForm)
	WinMove("DeviceInstaller", "", (@DesktopWidth - 800) / 2, 5, 800, 500) ;   800, 500, 10, 10
	PositionWindows()
EndIf

GetSystemInfo()
LoadConfig()
If $DEBUG = False Then
	TestNetwork()
	GetData()
	FilterData()
EndIf

;-----------------------------------------------
While 1
	Global $msg = GUIGetMsg(True)
	Switch $msg[0]
		Case $ButtonRunDevcon
			GetData()
			FilterData()
		Case $ButtonInstallDotNet
			InstallDotNet()
		Case $ButtonRescan
			Rescan()
			GetData()
			FilterData()
		Case $ButtonDump
			DumpToFile()
		Case $ButtonSearch
			Search()
		Case $InputSearch
			Search()
		Case $ButtonHelp
			Help()
		Case $RadioStopped
			FilterData()
		Case $RadioRunning
			FilterData()
		Case $RadioProblem
			FilterData()
		Case $RadioNotPresent
			FilterData()
		Case $RadioAllDevices
			FilterData()
		Case $RadioConfig
			FilterData()
		Case $ComboFilter
			FilterData()
		Case $ButtonDone
			GUISetState(@SW_HIDE, $DeviceInstall)
			GUISetState(@SW_SHOW, $MainForm)
		Case $ButtonLoadConfig
			LoadConfig()
		Case $ButtonViewEditMain
			ViewEdit()
		Case $ButtonEditDeviceInstall
			ViewEdit()
		Case $ButtonAbout
			About($ProgramName)
		Case $ButtonInstall
			InstallDrivers()
		Case $ButtonRefresh
			GetNodeInformation()
		Case $ButtonInstallOK
			DeviceInstall()
		Case $ButtonAboutDEV
			About("Device Install")
		Case $ButtonHelpDEV
			_debug(@ScriptName & @CRLF & $FileVersion & @CRLF & " Might put help here someday!", 0x40, 15)
		Case $LabelSysInfo
			ClipPut(GUICtrlRead($LabelSysInfo))
		Case $LabelDeviceID
			ClipPut(GUICtrlRead($LabelDeviceID))
		Case $EditDeviceResults
			ClipPut(GUICtrlRead($EditDeviceResults))
		Case $GUI_EVENT_CLOSE
			If $msg[1] = $MainForm Then Exit
			If $msg[1] = $DeviceInstall Then
				GUISetState(@SW_HIDE, $DeviceInstall)
				GUISetState(@SW_SHOW, $MainForm)
			EndIf
		Case $ButtonExit
			ExitLoop
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
						InstallDrivers()
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
Func MoveWindows()
	Local $F = WinGetPos("DeviceInstaller", "")
	WinMove("Device Install", "", $F[0], $F[1], $F[2], 500)
EndFunc   ;==>MoveWindows
;-----------------------------------------------
; hotkey F6
Func PositionWindows()
	WinMove("DeviceInstaller", "", (@DesktopWidth - 800) / 2, 5, 800, 500)
	MoveWindows()
EndFunc   ;==>PositionWindows
;-----------------------------------------------
; Hotkey f7
Func ShowNodes()
	ViewEdit($NODEINFOFILE)
EndFunc   ;==>ShowNodes
;-----------------------------------------------
; hotkey F8
Func AutoRunSilentNoGUI()
	SetStatus("AutoRunSilentNoGUI")
	GUISetState(@SW_HIDE, $MainForm)
	$QUIET = True
	$SILENT = False
	AutoRunQuiet()
	$QUIET = False
	SetStatus("AutoRunSilentNoGUI complete")
EndFunc   ;==>AutoRunSilentNoGUI
;-----------------------------------------------
; hotkey F9
Func AutoRunSilent()
	SetStatus("AutoRunSilent")
	$QUIET = True
	$SILENT = False
	AutoRunQuiet()
	$QUIET = False
	SetStatus("AutoRunSilent complete")
EndFunc   ;==>AutoRunSilent
;-----------------------------------------------
; hotkey F10
Func AutoRunQuiet()
	SetStatus("AutoRunQuiet")
	FileWriteLine($LOGFILE, _SystemLocalTime() & "Autorun beginning")
	GetData()
	GUICtrlSetState($RadioConfig, $GUI_CHECKED)
	LoadConfig()
	For $X = 0 To _GUICtrlListView_GetItemCount($ListView) - 1
		Local $ID = _GUICtrlListView_GetItemText($ListView, $X, 0)
		_debug(@ScriptLineNumber & " > " & $ID)
		Local $Name = _GUICtrlListView_GetItemText($ListView, $X, 1)
		_Debug(@ScriptLineNumber & "  >   " & $Name)
		Local $Status = _GUICtrlListView_GetItemText($ListView, $X, 2)
		_Debug(@ScriptLineNumber & "  >   " & $Status)
		Local $Name2 = _GUICtrlListView_GetItemText($ListView, $X, 3)
		_Debug(@ScriptLineNumber & "  >   " & $Status)
		FileWriteLine($LOGFILE, _SystemLocalTime() & "AutoRun. ID: " & $ID & "  " & $Name & "  " & $Status & "  " & $Name2)
		DeviceInstall($ID)
	Next
	GetData()
	FilterData(1)
	FileWriteLine($LOGFILE, _SystemLocalTime() & "Autorun complete")
	Exit
	SetStatus("AutoRunQuiet complete")
EndFunc   ;==>AutoRunQuiet
;-----------------------------------------------
; hotkey F11
Func Toggle()
	_Debug(@ScriptLineNumber & " Toggle")
	GuiDisable("Toggle")
EndFunc   ;==>Toggle
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
	Local $parms1 = ' FIND *  > ' & $RESULTFILEACTIVE
	Local $parms2 = ' FindAll * > ' & $RESULTFILEALL
	Local $parms3 = ' status * > ' & $STATUSFILE
	Local $parms4 = ' driverfiles * > ' & $GETMOREDATA
	_Debug(@ScriptLineNumber & " Devcon 1 " & _RunDOS($DEVCON & $parms1))
	_Debug(@ScriptLineNumber & " Devcon 2 " & _RunDOS($DEVCON & $parms2))
	_Debug(@ScriptLineNumber & " Devcon 3 " & _RunDOS($DEVCON & $parms3))
	_Debug(@ScriptLineNumber & " Devcon 4 " & _RunDOS($DEVCON & $parms4))
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
Func FilterData($pass = 0)
	SetStatus("FilterData")
	GuiDisable("Disable")

	LoadConfig()
	FileWriteLine($LOGFILE, _SystemLocalTime() & "Filtering data")
	If UBound($ArrayC) < 5 Then
		GetData()
	EndIf
	_GUICtrlListView_DeleteAllItems(GUICtrlGetHandle($ListView))

	_Debug("DBGVIEWCLEAR")
	;Remove what we don't care about (filter) ;  $ComboFilter
	Local $ComboFilterData = GUICtrlRead($ComboFilter);
	For $X = 0 To UBound($ArrayC) - 1
		Local $UU = StringSplit($ArrayC[$X], ":")
		Local $YY

		If GUICtrlRead($RadioAllDevices) = $GUI_CHECKED Then
			If StringInStr($UU[1], $ComboFilterData) = 1 Or StringCompare($ComboFilterData, "All") = 0 Then
				$YY = _GUICtrlListView_AddItem($ListView, StringStripWS($UU[1], 3))
				_GUICtrlListView_AddSubItem($ListView, $YY, StringStripWS($UU[2], 3), 1)
				_GUICtrlListView_AddSubItem($ListView, $YY, StringStripWS($UU[3], 3), 2)
				GetNameFromConfig($UU[1], $YY)
			EndIf
		ElseIf GUICtrlRead($RadioConfig) = $GUI_CHECKED Then
			For $Y = 0 To UBound($ArrayConfigData) - 1
				Local $Junk = StringSplit($ArrayConfigData[$Y], ":")
				Local $Junk2 = StringStripWS($Junk[1], 3)
				If StringInStr($UU[1], $Junk2) > 0 Then
					$YY = _GUICtrlListView_AddItem($ListView, StringStripWS($UU[1], 3))
					_GUICtrlListView_AddSubItem($ListView, $YY, StringStripWS($UU[2], 3), 1)
					_GUICtrlListView_AddSubItem($ListView, $YY, StringStripWS($UU[3], 3), 2)
					GetNameFromConfig($UU[1], $YY)
				EndIf
			Next
		ElseIf GUICtrlRead($RadioStopped) = $GUI_CHECKED And StringInStr($UU[3], "stopped") <> 0 And _
				(StringInStr($UU[1], $ComboFilterData) = 1 Or StringCompare($ComboFilterData, "All") = 0) Then
			$YY = _GUICtrlListView_AddItem($ListView, StringStripWS($UU[1], 3))
			_GUICtrlListView_AddSubItem($ListView, $YY, StringStripWS($UU[2], 3), 1)
			_GUICtrlListView_AddSubItem($ListView, $YY, StringStripWS($UU[3], 3), 2)
			GetNameFromConfig($UU[1], $YY)

		ElseIf GUICtrlRead($RadioRunning) = $GUI_CHECKED And StringInStr($UU[3], "running") <> 0 And _
				(StringInStr($UU[1], $ComboFilterData) = 1 Or StringCompare($ComboFilterData, "All") = 0) Then
			$YY = _GUICtrlListView_AddItem($ListView, StringStripWS($UU[1], 3))
			_GUICtrlListView_AddSubItem($ListView, $YY, StringStripWS($UU[2], 3), 1)
			_GUICtrlListView_AddSubItem($ListView, $YY, StringStripWS($UU[3], 3), 2)
			GetNameFromConfig($UU[1], $YY)
		ElseIf GUICtrlRead($RadioProblem) = $GUI_CHECKED And StringInStr($UU[3], "problem") <> 0 And _
				(StringInStr($UU[1], $ComboFilterData) = 1 Or StringCompare($ComboFilterData, "All") = 0) Then
			$YY = _GUICtrlListView_AddItem($ListView, StringStripWS($UU[1], 3))
			_GUICtrlListView_AddSubItem($ListView, $YY, StringStripWS($UU[2], 3), 1)
			_GUICtrlListView_AddSubItem($ListView, $YY, StringStripWS($UU[3], 3), 2)
			Local $ResultString = StringFormat("%s  %s  %s", $UU[1], $UU[2], $UU[3])
			GetNameFromConfig($UU[1], $YY)
			FileWriteLine($LOGFILE, _SystemLocalTime() & $ResultString)
			If $SILENT = False And $pass <> 0 Then MsgBox(48, "Device has a problem", $ResultString, 10)
		ElseIf GUICtrlRead($RadioNotPresent) = $GUI_CHECKED And StringInStr($UU[3], "Not Present") <> 0 And _
				(StringInStr($UU[1], $ComboFilterData) = 1 Or StringCompare($ComboFilterData, "All") = 0) Then
			$YY = _GUICtrlListView_AddItem($ListView, StringStripWS($UU[1], 3))
			_GUICtrlListView_AddSubItem($ListView, $YY, StringStripWS($UU[2], 3), 1)
			_GUICtrlListView_AddSubItem($ListView, $YY, StringStripWS($UU[3], 3), 2)
			GetNameFromConfig($UU[1], $YY)
		EndIf
	Next
	GuiDisable("Enable")
	SetStatus("FilterData complete")
EndFunc   ;==>FilterData
;-----------------------------------------------
Func GetNameFromConfig($DEV_VEN1, $YY)

	For $Y = 0 To UBound($ArrayConfigData) - 1
		Local $CFD = StringSplit($ArrayConfigData[$Y], ":")
		Local $DEV_VEN2 = StringStripWS($CFD[1], 3)
		$DEV_VEN1 = StringStripWS($DEV_VEN1, 3)
		$DEV_VEN2 = StringStripWS($DEV_VEN2, 3)
		Local $Name = StringStripWS($CFD[2], 3)
		If StringInStr($DEV_VEN1, $DEV_VEN2) > 0 Then
			_Debug(@ScriptLineNumber & "  >>" & $DEV_VEN1 & "<<>>" & $Name & "<<>>" & $YY & "<< ")
			_GUICtrlListView_AddSubItem($ListView, $YY, $Name, 3)
			Return
		Else
			_GUICtrlListView_AddSubItem($ListView, $YY, "------", 3)
		EndIf
	Next

EndFunc   ;==>GetNameFromConfig
;-----------------------------------------------
;This function displays the device Install window where installing may be done
Func InstallDrivers()
	GuiDisable("Disable")
	SetStatus("InstallDrivers")
	_Debug(@ScriptLineNumber & " InstallDrivers")
	GUISetState(@SW_HIDE, $MainForm)
	MoveWindows()
	GUISetState(@SW_SHOW, $DeviceInstall)

	Local $DevID
	; Local $DevName
	Local $DevStatus
	For $a = 0 To _GUICtrlListView_GetItemCount($ListView) - 1
		If _GUICtrlListView_GetItemSelected($ListView, $a) Then
			$DevID = _GUICtrlListView_GetItemText($ListView, $a, 0)
			; $DevName = _GUICtrlListView_GetItemText($ListView, $a, 1)
			$DevStatus = _GUICtrlListView_GetItemText($ListView, $a, 2)
		EndIf
	Next
	Local $TL = ''
	;_Debug(@ScriptLineNumber & " Data to work with:" & @CRLF & $DevID & @CRLF & $DevName & @CRLF & $DevStatus)
	If StringInStr($DevStatus, "Device not present") = 0 Then
		Local $GMD
		_FileReadToArray($GETMOREDATA, $GMD)

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
		GUICtrlSetState($ButtonInstallOK, $GUI_ENABLE)
	Else
		GUICtrlSetState($ButtonInstallOK, $GUI_DISABLE)
		GUICtrlSetData($LabelDeviceID, $DevID)
	EndIf

	$TL = GetNodeInformation()
	GUICtrlSetData($EditDeviceResults, "Current device information: " & @CRLF & $TL)

	GUICtrlSetData($NewDriverVersionValue, GetNewDriverVersion($DevID))

	GuiDisable("Enable")
	SetStatus("InstallDrivers complete")
EndFunc   ;==>InstallDrivers
;-----------------------------------------------
Func GetNewDriverVersion($DevID)
	Local $NewDriverVersion = "New driver version is not defined"
	_Debug(@ScriptLineNumber & "  " & $DevID)

	;_ArrayDisplay($ArrayConfigData, @ScriptLineNumber)
	For $i In $ArrayConfigData
		Local $j = StringSplit($i, ":")
		;_ArrayDisplay($j, @ScriptLineNumber)

		$j[1] = StringStripWS($j[1], 3)
		$DevID = StringStripWS($DevID, 3)
		_Debug(@ScriptLineNumber & "  >>" & $j[1] & "<<>>" & $DevID & "<<  ")

		If StringInStr($DevID, $j[1]) <> 0 Then
			If $j[0] = 5 Then
				$NewDriverVersion = $j[5]
				ExitLoop
			Else
				$NewDriverVersion = "New driver version is not defined"
				ExitLoop
			EndIf
			$NewDriverVersion = "??????"
		EndIf
	Next

	_Debug(@ScriptLineNumber & "  " & $NewDriverVersion)
	Return $NewDriverVersion
EndFunc   ;==>GetNewDriverVersion
;-----------------------------------------------

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
;This function processes the installed devices file from DEVCON.EXE
Func DeviceInstall($ID = "")
	SetStatus(" DeviceInstall")
	GuiDisable("Disable")
	;This holds the ID of the device that needs to be installed
	Local $InstallData[1]
	If $ID = "" Then $ID = GUICtrlRead($LabelDeviceID)

	If $ID = "" Then
		GuiDisable("Enable")
		Return
	EndIf

	; This loops through the config data and tries to get an ID match
	For $FF In $ArrayConfigData
		$InstallData = StringSplit($FF, ":")
		If StringInStr($ID, StringStripWS($InstallData[1], 3)) > 0 Then ExitLoop ;Got a match
	Next
	FileWriteLine($LOGFILE, _SystemLocalTime() & "DeviceInstall. ID: " & $ID)
	If StringInStr($FF, "Dummy") <> 0 Then
		If $QUIET = False Then MsgBox(48, "Install problem", "No matching install information found", 10)
		FileWriteLine($LOGFILE, _SystemLocalTime() & "No matching install information found")
		GuiDisable("Enable")
		Return
	EndIf

	Local $MRES
	If $QUIET = False Then
		$MRES = MsgBox(4 + 32 + 256, "Install this device? ", $InstallData[2] & "  " & $InstallData[3], 10)
		If $MRES = 7 Or $MRES = -1 Then
			FileWriteLine($LOGFILE, _SystemLocalTime() & "Install aborted.")
			GuiDisable("Enable")
			Return
		EndIf
	Else
		$MRES = 6 ; yes
	EndIf
	Switch $MRES
		Case 6 ; yes
			Local $file2run = FileGetShortName(StringStripWS($InstallData[3], 3))
			If FileExists($file2run) = True Then
				Local $Res = "Run " & $file2run & " Result: " & RunWait($file2run & " " & $InstallData[4])
				FileWriteLine($LOGFILE, _SystemLocalTime() & "Device name: " & $InstallData[2] & " Install file: " & $file2run & " Results: " & $Res)
			Else
				FileWriteLine($LOGFILE, _SystemLocalTime() & "Install file does not exist. " & $InstallData[2] & " " & $file2run)
				If $QUIET = False Then MsgBox(16, "Install file does not exist.", $InstallData[2] & @CRLF & $file2run, 10)
				GuiDisable("Enable")
				Return
			EndIf
		Case 7 ; no
			FileWriteLine($LOGFILE, _SystemLocalTime() & "Install canceled by user " & $InstallData[2] & "    " & $InstallData[3] & @CRLF)
			GuiDisable("Enable")
			Return
	EndSwitch
	GuiDisable("Enable")
	SetStatus(" DeviceInstall complete")
EndFunc   ;==>DeviceInstall
;-----------------------------------------------
;This function processes the all device file from DEVCON.EXE
Func ResultsAll()
	SetStatus("ResultsAll")
	Local $file = FileOpen($RESULTFILEALL, 0)
	; Check if file opened for reading OK
	If $file = -1 Then
		_debug("Unable to open file for reading: " & $RESULTFILEALL, 0x10, 5)
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
		_debug("Unable to open file for reading: " & $RESULTFILEACTIVE, 0x10, 5)
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
		_debug("Unable to open file for reading: " & $STATUSFILE, 0x10, 5)
		Return
	EndIf
	Local $tmp = FileRead($file)
	FileClose($file)
	$ArrayS = StringSplit($tmp, @CRLF, 3)
	; This routine fixes a problem with 'Device has a problem: 28.'  removes the : and the .
	Local $count = 0
	While True
		$count += 1
		If $count >= UBound($ArrayS) Then ExitLoop
		$ArrayS[$count] = StringReplace($ArrayS[$count], "Device has a problem: 28.", "Device has a problem = 28")
		_Debugviewer(@ScriptLineNumber & "  " & $count & "  " & $ArrayS[$count] & @CRLF)
	WEnd
	; This should fix a problem that occurs when no name is found in the status listing
	$count = 0
	While True
		$count += 1
		If StringInStr($ArrayS[$count], 'matching device') <> 0 Then ExitLoop
		If StringInStr($ArrayS[$count], "   ") = 0 And StringInStr($ArrayS[$count + 1], "Name:") = 0 Then
			_ArrayInsert($ArrayS, $count + 1, "Name: Name not found")
			$count += 1
			_Debugviewer(@ScriptLineNumber & "  " & $count & "  " & $ArrayS[$count])
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
;validate the congif file (look for extra : or duplicate DEV values or missing data
Func ValidateConfigData($MyArray)
	SetStatus("ValidateConfigData")
	_ArraySort($MyArray)
	Local $TA1[1]

	If IsArray($MyArray) = False Then
		ClipPut(GUICtrlRead($LabelSysInfo))
		Switch MsgBox(20, "System ID error", "Match for System ID not found " & @CRLF & $SystemProductName & @CRLF & "Exit program?")
			Case 6
				Exit
			Case 7
				;GuiDisable("Enable")
				Return
		EndSwitch
	EndIf

	For $T In $MyArray
		Local $TA2 = StringSplit($T, ":")
		If $TA2[0] <> 4 And $TA2[0] <> 5 Then
			MsgBox(48, "Config validate error", "Must have 3 or 4 ':' in line" & @CRLF & $T)
		EndIf
		_ArrayAdd($TA1, $TA2[1])
	Next

	; The following removes duplicate entries
	_ArrayDelete($TA1, 0)
	Local $TA3 = _ArrayUnique($TA1)
	_ArrayDelete($TA3, 0)

	If UBound($TA1) <> UBound($TA3) Then
		MsgBox(48, "A dulicate VEN or DEV found in config file", "A dulicate VEN or DEV found in config file" & @CRLF & $CONFIGFILENAME)
	EndIf
	SetStatus("ValidateConfigData complete")
EndFunc   ;==>ValidateConfigData
;-----------------------------------------------
;This function loads (or reloads)the config file
Func LoadConfig()
	SetStatus("LoadConfig " & $CONFIGFILENAME)
	Local $file = FileOpen($CONFIGFILENAME, 0)
	; Check if file opened for reading OK
	If $file = -1 Then
		FileWriteLine($LOGFILE, "LoadCFG: Unable to open file for reading: " & $CONFIGFILENAME)
		_Debug("LoadCFG: Unable to open file for reading: " & $CONFIGFILENAME, 0x10, 5)
		Return
	EndIf
	$ArrayConfigData = 0
	Global $ArrayConfigData[1]
	; Read in the first line to verify the file is of the correct type
	If StringCompare(FileReadLine($file, 1), "Valid for " & $ProgramName & " config") <> 0 Then
		FileWriteLine($LOGFILE, $CONFIGFILENAME & " Not valid for " & $ProgramName)
		_debug($CONFIGFILENAME & " Not valid for " & $ProgramName & " config", 0x20, 5)
		FileClose($file)
		Exit
	EndIf
	;Read in lines of text until the EOF is reached
	;Lines begining with ; are comments and are ignored
	;Lines less than 3 characters in length are ignored
	;Lines begining with set are for variables

	Local $SetArray[1]
	Local $Start = False
	While 1
		Local $LineIn = FileReadLine($file)
		If @error = -1 Then ExitLoop
		$LineIn = StringStripWS($LineIn, 3)

		If StringInStr($LineIn, ";") = 1 Or StringLen($LineIn) < 3 Then ContinueLoop

		If StringInStr($LineIn, "DotNetPath") = 1 Then $DOTNETPATH = $LineIn

		If StringInStr($LineIn, "set ") = 1 Then ; get the variables
			Local $T = StringSplit($LineIn, " =")
			_ArrayAdd($SetArray, "%" & $T[2] & "%=" & $T[3])
		Else
			If StringInStr($LineIn, $SystemProductName) <> 0 Then ; get the platform type
				$Start = True
			Else
				If StringInStr($LineIn, "[") = 1 Then $Start = False
				If $Start Then _ArrayAdd($ArrayConfigData, $LineIn)
			EndIf
			; $OsName
		EndIf

	WEnd
	FileClose($file)

	_ArrayDelete($SetArray, 0)

	For $T In $SetArray ; loop through the lines and substitute set values
		For $i = 1 To UBound($ArrayConfigData) - 1
			Local $U = StringSplit($T, "=")
			$ArrayConfigData[$i] = StringReplace($ArrayConfigData[$i], $U[1], $U[2])
			$DOTNETPATH = StringReplace($DOTNETPATH, $U[1], $U[2])
		Next
	Next

	_ArrayDelete($ArrayConfigData, 0) ; this removes the first entry (blank)
	_ArrayAdd($ArrayConfigData, "Dummy:Dummy:This is a dummy entry:Dummy") ; add a dummy entry
	ValidateConfigData($ArrayConfigData)

	;_ArrayDisplay($ArrayConfigData, @ScriptLineNumber)

	FileWriteLine($LOGFILE, _SystemLocalTime() & "Config file loaded: " & $CONFIGFILENAME)
	SetStatus("LoadConfig " & $CONFIGFILENAME & " complete")
EndFunc   ;==>LoadConfig

;-----------------------------------------------
;This function allows the user to edit or view any file, useful for changing the config file
Func ViewEdit($FileName = '')
	SetStatus("ViewEdit " & $FileName)
	GuiDisable("Disable")
	If $FileName = '' Then
		$FileName = FileOpenDialog("View or Edit a file", @ScriptDir, "All (*.*)", 1);, $CONFIGFILENAME)
	EndIf
	ConsoleWrite(@ScriptLineNumber & " +++ " & $FileName & @CRLF)
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
;This function retrives some system info from the registry
Func GetSystemInfo()
	SetStatus("GetSystemInfo")
	;$SystemProductName = RegRead("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\BIOS", "SystemProductName")
	;$OsName = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Intel\MPG", "OsName")

	$SystemProductName = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Intel\MPG", "PlatformName")
	$OsName = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Intel\MPG", "OsName")

	GUICtrlSetData($LabelSysInfo, $OsName & @CRLF & $SystemProductName)
	FileWriteLine($LOGFILE, _SystemLocalTime() & "Get system info completed")
	SetStatus("GetSystemInfo complete")
EndFunc   ;==>GetSystemInfo
;-----------------------------------------------
;This function tests the network to be sure that it is accessable. If not it prompts the user to login.
Func TestNetwork()
	SetStatus("TestNetwork")
	Local $TestLocaltion = "\\chakotay\temp"
	If FileChangeDir($TestLocaltion) = 0 Then
		FileWriteLine($LOGFILE, _SystemLocalTime() & "Network test unable to access " & $TestLocaltion & @CRLF)
		SetStatus("TestNetwork failed  " & $TestLocaltion & @CRLF)
		MsgBox(48, "TestNetwork failed", "Network test unable to access " & $TestLocaltion)
	Else
		FileWriteLine($LOGFILE, _SystemLocalTime() & "Network test completed. " & $TestLocaltion & @CRLF)
		SetStatus("TestNetwork success  " & $TestLocaltion & @CRLF)
	EndIf
EndFunc   ;==>TestNetwork
;-----------------------------------------------
Func Search()
	Static Local $iStart = 0
	SetStatus(" Search " & $iStart)
	Local $T = _GUICtrlListView_FindInText($ListView, GUICtrlRead($InputSearch), $iStart, True)
	$iStart = $T
	_Debug(@ScriptLineNumber & " Found search " & $iStart)
	If $iStart = -1 Then
		MsgBox(48, "Search complete", "Search item not found" & @CRLF & GUICtrlRead($InputSearch))
	EndIf
	_GUICtrlListView_EnsureVisible($ListView, $T)
	_GUICtrlListView_ClickItem($ListView, $T)
	SetStatus(" Search " & $iStart & " complete")
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
		ConsoleWrite(@ScriptLineNumber & $FormID & @CRLF)
		$WinPos = StringFormat("%s" & @CRLF & "WinPOS: %d  %d " & @CRLF & "WinSize: %d %d " & @CRLF & "Desktop: %d %d ", _
				$FormID, $D[0], $D[1], $D[2], $D[3], @DesktopWidth, @DesktopHeight)
	Else
		$WinPos = ">>>About ERROR, Check the window name<<<"
	EndIf
	_Debug(@CRLF & $SystemS & @CRLF & $WinPos & @CRLF & "Written by Doug Kaynor!", 0x40)
EndFunc   ;==>About
;-----------------------------------------------
; HKLM\SOFTWARE|Microsoft
Func InstallDotNet()
	SetStatus("Install DotNet3.5")
	GuiDisable("Disable")

	Local $DOT3_5Name = RegRead("HKLM\SOFTWARE\Microsoft\.NETFramework\", "*")
	_Debug(@ScriptLineNumber & " " & _SystemLocalTime() & " >>" & @error & "<<>>" & $DOT3_5Name & "<<")

	Switch MsgBox(36, "Install DotNet 3.5", "Are you sure?")
		Case 6 ; yes
			Local $temp = StringSplit($DOTNETPATH, "=")
			Local $FileName = FileGetShortName($temp[2])
			If FileExists($FileName) = True Then
				SetStatus("Installing DotNet 3.5")
				_Debug(@ScriptLineNumber & " DotNetPath install: " & _RunDOS($FileName) & "   " & $FileName)
				SetStatus("Install DotNet 3.5 complete")
			Else
				MsgBox(16, "Dotnet 3.5 install error", "File does not exist" & @CRLF & $FileName)
				SetStatus("Install DotNet 3.5 failed")
			EndIf
		Case 7 ; no
			SetStatus("Install DotNet 3.5 aborted")
	EndSwitch

	GuiDisable("Enable")

EndFunc   ;==>InstallDotNet
;-----------------------------------------------
Func SetStatus($msg)
	GUICtrlSetData($LabelStatusMain, $msg)
	GUICtrlSetData($LabelStatusInstall, $msg)
	_Debug(@ScriptLineNumber & " " & $msg)
EndFunc   ;==>SetStatus
;-----------------------------------------------
Func Help()
	_Debug(@ScriptLineNumber & " Help")
	Local $helpstr = 'Command line startup options:' & @CRLF & @CRLF & _
			"help or ? Help information is displayed" & @CRLF & _
			"nogui   No GUI is displayed" & @CRLF & _
			"quiet   Most user prompts are supressed" & @CRLF & _
			"silent  All user prompts are supressed" & @CRLF & _
			"debug   Various debug functions" & @CRLF & _
			"logfile (the log file path must be wrapped in quotes)" & @CRLF & @CRLF & _
			"Hotkeys" & @CRLF & @CRLF & _
			"F6  = Windows resize and position" & @CRLF & _
			"F7  = Show nodeinfo.txt" & @CRLF & _
			"F8  = AutoRunSilentNoGUI" & @CRLF & _
			"F9  = AutorunSilent" & @CRLF & _
			"F10 = AutoRunQuiet" & @CRLF & _
			"F11 = Enable or Disable the GUI" & @CRLF
	_Debug(@ScriptName & @CRLF & $FileVersion & @CRLF & @CRLF & $helpstr, 0x40)
EndFunc   ;==>Help
;-----------------------------------------------