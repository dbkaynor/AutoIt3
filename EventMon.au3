#AutoIt3Wrapper_icon=../icons/Dice.ico
#AutoIt3Wrapper_outfile=EventMon.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Comment=A EventLog tool
#AutoIt3Wrapper_Res_Description=EventLog tool
#AutoIt3Wrapper_Res_Fileversion=1.0.0.107
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=Y
#AutoIt3Wrapper_Res_ProductVersion=666
#AutoIt3Wrapper_Res_LegalCopyright=Copyright 2016, 2017 Douglas B Kaynor
#AutoIt3Wrapper_Res_SaveSource=y
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Field=AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Res_Field=Compile date|%longdate% %time%
#AutoIt3Wrapper_Res_Field=Developer|Douglas Kaynor
#AutoIt3Wrapper_Res_Field=Email|doug@kaynor.net
#AutoIt3Wrapper_Res_Field=Made By|Douglas Kaynor
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=y
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Run_Tidy=y
#Tidy_Parameters=/tc 4 /kv 2
#AutoIt3Wrapper_Run_Au3Stripper=n
#Au3Stripper_Parameters=/mo

Opt("MustDeclareVars", 1)
Opt("TrayIconDebug", 1)
Opt("TrayAutoPause", 0)
Opt("WinTitleMatchMode", 2)
Opt("GUIResizeMode", 0)

#cs
    This area is used to store things

    clipget
    Fix help (F1)

    Filtered file count
    Button to check for dmp files

#CE

#RequireAdmin

#include <Array.au3>
#include <ButtonConstants.au3>
#include <Constants.au3>
#include <ColorConstants.au3>
#include <Date.au3>
#include <EventLog.au3>
#include <file.au3>
#include <GUIConstantsEx.au3>
#include <GuiListView.au3>
#include <ListViewConstants.au3>
#include <Misc.au3>
#include <MsgBoxConstants.au3>
#include <SliderConstants.au3>
#include <StaticConstants.au3>
#include <StringConstants.au3>
#include <String.au3>
#include <WindowsConstants.au3>
#include <_DougFunctions.au3>

DirCreate(@ScriptDir & "\AUXFiles\")


#CS
    If FileInstall("Working.jpg", @ScriptDir & "\AUXFiles\Working.jpg") = 0 Then
    MsgBox($MB_ICONINFORMATION +$MB_TOPMOST , "FileInstall failure", "Working.jpg not found")
    EndIf
#CE

Global $DebugString = ""
Global Const $ProgramName = "Event log monitor"
Global $Project_filename = @ScriptDir & "\AUXFiles\" & $ProgramName & ".prj"
Global $LOG_filename = @ScriptDir & "\AUXFiles\" & $ProgramName & ".log"
Global $LoopCount
Global $hEventLog

Global Const $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")
Global $SystemS = $ProgramName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & @OSServicePack & @CRLF & @OSType & @CRLF & @OSArch

For $x = 1 To $CmdLine[0]
    _debug($x & " >> " & $CmdLine[$x] & @CRLF)
    Select
        Case StringCompare($CmdLine[$x], "help", $STR_NOCASESENSEBASIC) == 0 Or StringCompare($CmdLine[$x], "?") == 0
            Help()
            Exit
        Case StringCompare($CmdLine[$x], "clear", $STR_NOCASESENSEBASIC) == 0
            ClearLogs()
            Exit
        Case Else
            MsgBox($MB_ICONWARNING + $MB_TOPMOST, "Parse command line", " Unknown command line option " & @CRLF & $CmdLine[$x], 10)
            Exit
    EndSelect
Next

If _Singleton($ProgramName, 1) = 0 Then
    MsgBox($MB_ICONWARNING + $MB_TOPMOST, "Already running", $ProgramName & " is already running!", 10)
    Exit
EndIf


Global Const $TrayIconRunning = '../icons/Dice.ico'
TraySetIcon($TrayIconRunning)
TrayTip($ProgramName, $FileVersion, 3)

Opt("TrayMenuMode", 1)
Opt("TrayOnEventMode", 0)
;Opt("TrayIconDebug", 1)

; Mainform 10, 10, 675, 530) WinMove ( "title", "text", x, y [, width [, height [, speed]]] )
Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_MAXIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $MainForm = GUICreate("Event log monitor  " & $FileVersion, 650, 530, 10, 10, $MainFormOptions)
;---- Top row
Global $ButtonGetData = GUICtrlCreateButton("Get data", 30, 10, 55, 25)
GUICtrlSetTip(-1, "Fetch the data from selected logs")
GUICtrlSetResizing(-1, 802)

Global $ButtonEventLogStats = GUICtrlCreateButton("Stats", 90, 10, 55, 25)
GUICtrlSetTip(-1, "Display event log statistics")
GUICtrlSetResizing(-1, 802)

Global $ButtonClearLogs = GUICtrlCreateButton("Clear logs", 150, 10, 55, 25)
GUICtrlSetTip(-1, "Clear all logs. Check 'Save' if you want to save before clearing")
GUICtrlSetResizing(-1, 802)

Global $ButtonEventViewer = GUICtrlCreateButton("Event viewer", 210, 10, 70, 25)
GUICtrlSetTip(-1, "Open the event control viewer")
GUICtrlSetResizing(-1, 802)
Global $ButtonOpenSaveFolder = GUICtrlCreateButton("Open save", 285, 10, 75, 25)
GUICtrlSetTip(-1, "Open the log save location.")
GUICtrlSetResizing(-1, 802)
Global $ButtonAbout = GUICtrlCreateButton("About", 370, 10, 75, 25)
GUICtrlSetTip(-1, "Display application information.")
GUICtrlSetResizing(-1, 802)
Global $ButtonHelp = GUICtrlCreateButton("Help", 450, 10, 75, 25)
GUICtrlSetTip(-1, "Display application help")
GUICtrlSetResizing(-1, 802)
Global $ButtonExit = GUICtrlCreateButton("Exit", 540, 10, 75, 25)
GUICtrlSetTip(-1, "Exit this application")
GUICtrlSetResizing(-1, 802)

Global $GroupOptions = GUICtrlCreateGroup("Options", 30, 40, 130, 40)
GUICtrlSetResizing(-1, 802)

Global $CheckSave = GUICtrlCreateCheckbox("Save", 40, 55, 50, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Save logs before clearing")

Global $CheckAbort = GUICtrlCreateCheckbox("Abort", 100, 55, 50, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Abort operations")

GUICtrlCreateGroup("", -99, -99, 1, 1)

Global $ButtonSaveToClipboard = GUICtrlCreateButton("Copy to clip", 170, 40, 75, 25)
GUICtrlSetTip(-1, "Save selected data to clipboard")
GUICtrlSetResizing(-1, 802)

Global $ButtonSaveDataAsText = GUICtrlCreateButton("Save as text", 250, 40, 75, 25)
GUICtrlSetTip(-1, "Save the displayed data as text")
GUICtrlSetResizing(-1, 802)

Global $ButtonSaveProject = GUICtrlCreateButton("Save project", 330, 40, 75, 25)
GUICtrlSetTip(-1, "Save the current settings")
GUICtrlSetResizing(-1, 802)

Global $ButtonLoadProject = GUICtrlCreateButton("Load project", 410, 40, 75, 25)
GUICtrlSetTip(-1, "Load saved settings")
GUICtrlSetResizing(-1, 802)

Global $ButtonLoadSettingDefaults = GUICtrlCreateButton("Load defaults", 490, 40, 75, 25)
GUICtrlSetTip(-1, "Load default settings")
GUICtrlSetResizing(-1, 802)

Global $GroupLogType = GUICtrlCreateGroup("Log type", 30, 80, 340, 40)
GUICtrlSetResizing(-1, 802)
Global $CheckSystem = GUICtrlCreateCheckbox("System", 40, 95, 60, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Process system logs")
Global $CheckSecurity = GUICtrlCreateCheckbox("Security", 110, 95, 70, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Process security logs")
Global $CheckApplication = GUICtrlCreateCheckbox("Application", 180, 95, 70, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Process application logs")
Global $ButtonToggleLogType = GUICtrlCreateButton("Toggle", 260, 95, 50, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Toggle log type selections")
Global $ButtonDefaultsLogType = GUICtrlCreateButton("Defaults", 315, 95, 50, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Set log type selections to defaults")
GUICtrlCreateGroup("", -99, -99, 1, 1)

Global $GroupEventType = GUICtrlCreateGroup("Event type", 30, 120, 530, 40)
GUICtrlSetResizing(-1, 802)
Global $CheckInformation = GUICtrlCreateCheckbox("Information", 40, 135, 70, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Show information events")
Global $CheckWarnings = GUICtrlCreateCheckbox("Warnings", 115, 135, 70, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Show warning events")
Global $CheckErrors = GUICtrlCreateCheckbox("Errors", 185, 135, 50, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Show error events")
Global $CheckSuccessAudits = GUICtrlCreateCheckbox("Success Audits", 235, 135, 90, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Show successful audit events")
Global $CheckFailAudits = GUICtrlCreateCheckbox("Fail Audits", 330, 135, 70, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Show failed audit events")
Global $CheckOther = GUICtrlCreateCheckbox("Other", 400, 135, 50, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Other")
Global $ButtonToggleEventType = GUICtrlCreateButton("Toggle", 450, 135, 50, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Toggle event type selections")
Global $ButtonDefaultsEventType = GUICtrlCreateButton("Defaults", 505, 135, 50, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Set log type selections to defaults")
GUICtrlCreateGroup("", -99, -99, 1, 1)

Global $ButtonSearchDMP = GUICtrlCreateButton("DMP files", 580, 135, 50, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Search for dmp (dump) files)")

Global $ButtonDeviceMgr = GUICtrlCreateButton("Dev manager", 640, 135, 80, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Check device manager for errors")

GUICtrlCreateGroup("Filter ", 30, 160, 650, 40)
GUICtrlSetResizing(-1, 802)
Global $InputFilter = GUICtrlCreateInput('', 40, 175, 450, 20)
GUICtrlSetTip(-1, "Input a filter or filters separated by ;")
GUICtrlSetResizing(-1, 802)
Global $CheckExclude = GUICtrlCreateCheckbox("Exclude", 500, 175, 60, 20)
GUICtrlSetTip(-1, "Toggle between exclude or include")
GUICtrlSetResizing(-1, 802)
Global $CheckIgnore = GUICtrlCreateCheckbox("Ignore", 560, 175, 60, 20)
GUICtrlSetTip(-1, "Ignore fiter")
GUICtrlSetResizing(-1, 802)
Global $ButtonClearFilter = GUICtrlCreateButton("Clear Filter", 620, 175, 55, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Clear the filter list")

GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateGroup("Loop", 375, 80, 260, 40)
GUICtrlSetResizing(-1, 802)
Global $CheckLoop = GUICtrlCreateCheckbox("Loop", 380, 95, 40, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Start/Stop looping")
Global $SliderDelay = GUICtrlCreateSlider(430, 95, 125, 20, $TBS_AUTOTICKS)
GUICtrlSetResizing(-1, 802)
GUICtrlSetLimit(-1, 10, 1)
;GUICtrlSetData(-1, 5)
GUICtrlSetTip(-1, "Set loop delay seconds")
Global $LabelSliderDelay = GUICtrlCreateLabel(GUICtrlRead($SliderDelay), 560, 95, 20, 20, $SS_SUNKEN)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Current loop delay setting")
Global $LabelLoopCount = GUICtrlCreateLabel($LoopCount, 590, 95, 40, 20, $SS_SUNKEN)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Current loop count")
GUICtrlCreateGroup("", -99, -99, 1, 1)

Global $LabelStatus = GUICtrlCreateLabel("Status main", 20, 210, 620, 20, $SS_SUNKEN) ;20, 490, 640, 20
GUICtrlSetResizing(-1, 802 - 256)
GUICtrlSetTip(-1, "Status messages")

Global $LabelWorking = GUICtrlCreateLabel('Idle', 650, 10, 100, 100, $SS_SUNKEN) ;20, 490, 640, 20
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Working status")
GUICtrlSetFont($LabelWorking, 16)
GUICtrlSetBkColor($LabelWorking, $CLR_NONE)

Const $Number = 0
Const $Type = 1
Const $Date = 2
Const $Source = 3
Const $Description = 4

;GUICtrlCreateListView ( "text", left, top [, width [, height [, style = -1 [, exStyle = -1]]]] )
Global $ListView = GUICtrlCreateListView("Number|Event type|Date logged|Event source|Event description", 20, 240, 620, 250, $LVS_REPORT, BitOR($LVS_EX_FULLROWSELECT, $WS_EX_CLIENTEDGE, $LVS_EX_GRIDLINES))
GUICtrlSetTip(-1, "This is the list box")
;GUICtrlSetResizing(-1, BitOR($GUI_DOCKTOP, $GUI_DOCKBOTTOM))
GUICtrlSetResizing(-1, 102)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, $Number, 100)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, $Type, 120)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, $Date, 145)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, $Source, 140)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, $Description, 1000)

Static $ToggleState = $GUI_CHECKED
Static $SavedTime = TimerInit()

If FileExists($LOG_filename) Then FileDelete($LOG_filename)

$DebugString = "DBGVIEWCLEAR"
;$DebugString = @ScriptLineNumber & " Debug viewer cleared"
SetStatusMessage($DebugString)

SetDefaults()

LoadProject("start")

GUISetState(@SW_SHOW)

If Not IsAdmin() Then
    $DebugString = @ScriptLineNumber & " Program not started as administrator"
    SetStatusMessage($DebugString)
    MsgBox($MB_ICONWARNING + $MB_TOPMOST, 'Program not started as administrator', 'This program must be run as adminstrator for full functionality.', 3)
Else
    $DebugString = @ScriptLineNumber & " Program started as administrator"
    SetStatusMessage($DebugString)
EndIf

While 1
    Global $t = GUIGetMsg()
    Switch $t
        Case $ListView ;Global $ListView = GUICtrlCreateListView(
            ClipHandler()
        Case $ButtonSaveToClipboard
            ClipHandler()
        Case $ButtonGetData
            _GUICtrlListView_DeleteAllItems(GUICtrlGetHandle($ListView))
            GetData('Verbose')
        Case $ButtonEventLogStats
            _GUICtrlListView_DeleteAllItems(GUICtrlGetHandle($ListView))
            GetData('StatsOnly')
        Case $ButtonEventViewer
            ShellExecuteWait("eventvwr.msc", "/s")
        Case $ButtonClearLogs
            ClearLogs()
            GetData('Verbose')
        Case $ButtonOpenSaveFolder
            ClipPut(EnvGet("temp"))
            RunWait("C:\WINDOWS\EXPLORER.EXE /n,/e," & EnvGet("temp"))
        Case $SliderDelay
            GUICtrlSetData($LabelSliderDelay, GUICtrlRead($SliderDelay))
        Case $ButtonToggleLogType
            Toggle('Log')
        Case $ButtonDefaultsLogType
            SetDefaults('log')

        Case $ButtonSearchDMP
            SearchDMP()
        Case $ButtonDeviceMgr
            CheckDeviceManager()
        Case $ButtonToggleEventType
            Toggle('event')
        Case $ButtonSaveDataAsText
            SaveDataAsText()
        Case $ButtonClearFilter
            GUICtrlSetData($InputFilter, "")
            GUICtrlSetState($CheckExclude, $GUI_UNCHECKED)
            GUICtrlSetState($CheckIgnore, $GUI_UNCHECKED)
        Case $ButtonDefaultsEventType
            SetDefaults('event')
        Case $ButtonSaveProject
            SaveProject()
        Case $ButtonLoadProject
            LoadProject("Menu")
        Case $ButtonLoadSettingDefaults
            SetDefaults('all')
        Case $ButtonAbout
            About($ProgramName)
        Case $ButtonHelp
            Help()
        Case $ButtonExit
            $DebugString = @ScriptLineNumber & " ButtonExit"
            SetStatusMessage($DebugString)
            Exit
        Case $GUI_EVENT_CLOSE
            $DebugString = @ScriptLineNumber & " GUI_EVENT_CLOSE"
            SetStatusMessage($DebugString)
            Exit
    EndSwitch
    CheckChangeCounter()
WEnd
;-----------------------------------------------

Func ClipHandler()
    $DebugString = @ScriptLineNumber & " ClipHandler"
    SetStatusMessage($DebugString)
    Local $XX = _GUICtrlListView_GetItemTextString($ListView, _GUICtrlListView_GetSelectionMark($ListView))
    _ArrayDisplay(StringSplit($XX, "|"), "Selected data", "", 16)
    ;MsgBox($MB_ICONWARNING+ $MB_TOPMOST, 'Copy to clipboard', _GUICtrlListView_GetItemTextString($ListView, _GUICtrlListView_GetSelectionMark($ListView)), 3)
    ;ClipPut($XX)

EndFunc   ;==>ClipHandler
;-----------------------------------------------

; This loops the display of
Func CheckChangeCounter()

    If GUICtrlRead($CheckLoop) = $GUI_UNCHECKED Then
        Return
    EndIf

    Local $CurrentTime = TimerDiff($SavedTime) / 1000 ; seconds

    If GUICtrlRead($CheckAbort) = $GUI_CHECKED Then
        GUICtrlSetState($CheckLoop, $GUI_UNCHECKED)
        GUICtrlSetState($CheckAbort, $GUI_UNCHECKED)
        $LoopCount = 0
        GUICtrlSetData($LabelLoopCount, $LoopCount)
        Return
    EndIf

    If Mod(Int($CurrentTime * 400), 10) <> 0 Then Return
    If $CurrentTime > GUICtrlRead($SliderDelay) Then
        $DebugString = @ScriptLineNumber & " " & $CurrentTime & " " & GUICtrlRead($SliderDelay)
        SetStatusMessage($DebugString)
        $SavedTime = TimerInit()
        _GUICtrlListView_DeleteAllItems(GUICtrlGetHandle($ListView))
        GetData('Verbose')
        $LoopCount = $LoopCount + 1

        GUICtrlSetData($LabelLoopCount, $LoopCount)
    EndIf
EndFunc   ;==>CheckChangeCounter
;-----------------------------------------------
Func SaveDataAsText()
    $DebugString = @ScriptLineNumber & " SaveDataAsText"
    SetStatusMessage($DebugString)
    Local $SaveDataAsText_filename = $ProgramName
    $SaveDataAsText_filename = FileSaveDialog("Save data as text file", @ScriptDir, $SaveDataAsText_filename & " SaveDataAsText (*.txt)|All files (*.txt)|All files (*.*)", 18, $SaveDataAsText_filename & ".txt")

    Local $file = FileOpen($SaveDataAsText_filename, 2)
    ; Check if file opened for writing OK
    If $file = -1 Then
        $DebugString = @ScriptLineNumber & " SaveDataAsText: Unable to open file for writing: " & $SaveDataAsText_filename
        SetStatusMessage($DebugString)
        Return
    EndIf
    $DebugString = @ScriptLineNumber & " SaveDataAsText  " & $SaveDataAsText_filename
    SetStatusMessage($DebugString)

    Local $Count = _GUICtrlListView_GetItemCount(GUICtrlGetHandle($ListView))
    For $x = 0 To $Count
        Local $DataString = _GUICtrlListView_GetItemTextString(GUICtrlGetHandle($ListView), $x)
        FileWriteLine($file, $DataString)
        $DebugString = @ScriptLineNumber & " " & $DataString
        SetStatusMessage($DebugString)
    Next
    FileClose($file)

EndFunc   ;==>SaveDataAsText
;-----------------------------------------------
Func ClearLogs()
    GUICtrlSetBkColor($LabelWorking, $COLOR_RED)
    GUICtrlSetData($LabelWorking, "Clear logs")
    $DebugString = @ScriptLineNumber & " ClearLogs"
    SetStatusMessage($DebugString)

    If Not IsAdmin() Then
        MsgBox($MB_ICONWARNING + $MB_TOPMOST, 'Fuction unavailable.', 'Fuction unavailable. Program must be run as administrator')
        Return
    EndIf

    If MsgBox($MB_ICONQUESTION + $MB_OKCANCEL + $MB_TOPMOST, "Clear all logs", "Are you sure? (verify save check box!)") = 2 Then Return

    Local $tCur = _Date_Time_GetLocalTime()
    Local $tStr = _Date_Time_SystemTimeToDateTimeStr($tCur)
    $tStr = StringRegExpReplace($tStr, "[:/ ]", "", 0)

    EnvGet("temp")
    If GUICtrlRead($CheckSave) = $GUI_CHECKED Then
        $hEventLog = _EventLog__Open("", "Application")
        _EventLog__Clear($hEventLog, EnvGet("temp") & "\Application_" & $tStr & ".evtx")
        _EventLog__Close($hEventLog)

        $hEventLog = _EventLog__Open("", "Security")
        _EventLog__Clear($hEventLog, EnvGet("temp") & "\Security_" & $tStr & ".evtx")
        _EventLog__Close($hEventLog)

        $hEventLog = _EventLog__Open("", "Setup")
        _EventLog__Clear($hEventLog, EnvGet("temp") & "\Setup_" & $tStr & ".evtx")
        _EventLog__Close($hEventLog)

        $hEventLog = _EventLog__Open("", "System")
        _EventLog__Clear($hEventLog, EnvGet("temp") & "\System_" & $tStr & ".evtx")
        _EventLog__Close($hEventLog)
        MsgBox($MB_ICONWARNING + $MB_TOPMOST, "Log files saved", "Saved to " & EnvGet("temp"))
    Else
        $hEventLog = _EventLog__Open("", "Application")
        _EventLog__Clear($hEventLog, "")
        _EventLog__Close($hEventLog)
        $hEventLog = _EventLog__Open("", "Security")
        _EventLog__Clear($hEventLog, "")
        _EventLog__Close($hEventLog)
        $hEventLog = _EventLog__Open("", "Setup")
        _EventLog__Clear($hEventLog, "")
        _EventLog__Close($hEventLog)
        $hEventLog = _EventLog__Open("", "System")
        _EventLog__Clear($hEventLog, "")
        _EventLog__Close($hEventLog)
    EndIf
    _GUICtrlListView_DeleteAllItems(GUICtrlGetHandle($ListView))
    GUICtrlSetBkColor($LabelWorking, $CLR_NONE)
    GUICtrlSetData($LabelWorking, 'Idle')
EndFunc   ;==>ClearLogs
;-----------------------------------------------
Func GetData($Type)
    GUICtrlSetBkColor($LabelWorking, $COLOR_RED)
    GUICtrlSetData($LabelWorking, 'Get data')
    $DebugString = @ScriptLineNumber & " GetData " & $Type
    SetStatusMessage($DebugString, True)

    Local $ListView_item
    GUICtrlSetState($CheckAbort, $GUI_UNCHECKED)

    If GUICtrlRead($CheckAbort) = $GUI_CHECKED Then
        SetStatusMessage("GetData aborted")
        Return
    EndIf

    If GUICtrlRead($CheckApplication) = $GUI_CHECKED Then
        $ListView_item = _GUICtrlListView_AddItem($ListView, '')
        _GUICtrlListView_AddSubItem($ListView, $ListView_item, '--Application log--', $Type)
        $hEventLog = _EventLog__Open("", "Application")
        GetInformation($hEventLog, "Application", $Type)
        _EventLog__Close($hEventLog)
        _GUICtrlListView_AddItem($ListView, "  ")
    EndIf
    If GUICtrlRead($CheckAbort) = $GUI_CHECKED Then Return
    _debug(@ScriptLineNumber & " Application complete" & @CRLF)

    If GUICtrlRead($CheckSecurity) = $GUI_CHECKED Then
        If Not IsAdmin() Then
            $ListView_item = _GUICtrlListView_AddItem($ListView, '')
            _GUICtrlListView_AddSubItem($ListView, $ListView_item, 'Security log unavailble.', 1)
            _GUICtrlListView_AddSubItem($ListView, $ListView_item, 'Must be run as administrator', 2)
            _GUICtrlListView_AddItem($ListView, "  ")
        Else
            $ListView_item = _GUICtrlListView_AddItem($ListView, '')
            _GUICtrlListView_AddSubItem($ListView, $ListView_item, '--Security log--', $Type)
            $hEventLog = _EventLog__Open("", "Security")
            GetInformation($hEventLog, "Security", $Type)
            _EventLog__Close($hEventLog)
            _GUICtrlListView_AddItem($ListView, "  ")
        EndIf
    EndIf
    If GUICtrlRead($CheckAbort) = $GUI_CHECKED Then
        GUICtrlSetBkColor($LabelWorking, $CLR_NONE)
        GUICtrlSetData($LabelWorking, 'Idle')
        Return
    EndIf

    If GUICtrlRead($CheckSystem) = $GUI_CHECKED Then
        $ListView_item = _GUICtrlListView_AddItem($ListView, '')
        _GUICtrlListView_AddSubItem($ListView, $ListView_item, '--System log--', $Type)
        $hEventLog = _EventLog__Open("", "System")
        GetInformation($hEventLog, "System", $Type)
        _EventLog__Close($hEventLog)
        _GUICtrlListView_AddItem($ListView, "  ")
    EndIf
    _debug(@ScriptLineNumber & " System complete" & @CRLF)

    GUICtrlSetBkColor($LabelWorking, $CLR_NONE)
    GUICtrlSetData($LabelWorking, 'Idle')
EndFunc   ;==>GetData
;-----------------------------------------------
Func GetInformation($hEventLog, $trace, $Type)
    $DebugString = @ScriptLineNumber & " GetInformation " & $hEventLog & " " & $trace & " " & $Type
    SetStatusMessage($DebugString)
    Local $ListView_item
    $ListView_item = _GUICtrlListView_AddItem($ListView, '')
    _GUICtrlListView_AddSubItem($ListView, $ListView_item, "Log full", 0)
    _GUICtrlListView_AddSubItem($ListView, $ListView_item, _EventLog__Full($hEventLog), 1)
    $ListView_item = _GUICtrlListView_AddItem($ListView, '')
    _GUICtrlListView_AddSubItem($ListView, $ListView_item, "Event count", 0)
    _GUICtrlListView_AddSubItem($ListView, $ListView_item, _EventLog__Count($hEventLog), 1)
    GetStats($hEventLog, $trace, $Type)
EndFunc   ;==>GetInformation
;-----------------------------------------------
Func GetStats($hEventLog, $trace, $Type)
    GUICtrlSetBkColor($LabelWorking, $COLOR_RED)
    GUICtrlSetData($LabelWorking, 'Get stats')
    $DebugString = @ScriptLineNumber & " GetStats " & $hEventLog & " " & $trace & " " & $Type
    SetStatusMessage($DebugString)
    Local $Errors = 0
    Local $Warnings = 0
    Local $Information = 0
    Local $SuccessAudit = 0
    Local $FailureAudit = 0
    Local $Other = 0
    Local $Else = 0
    Local $FilterArray
    Local $Counter = 0
    #cs
        Const $Number = 0
        Const $Type = 1
        Const $Date = 2
        Const $Source = 3
        Const $Description = 4
    #ce

    ;Parse the filter string into as array use ; as the separator
    ;Use the elements of $FilterArray to run StringInStr against the $TmpString

    ;$DebugString = "DBGVIEWCLEAR" ;$DebugString = @ScriptLineNumber & " Debug viewer cleared"
    ;SetStatusMessage($DebugString)

    If StringLen(GUICtrlRead($InputFilter)) > 0 Then ; There is a input filter
        $FilterArray = StringSplit(GUICtrlRead($InputFilter), ";", 2)
        ;SetStatusMessage(@ScriptLineNumber & "  " & _ArrayToString($FilterArray) & "   " & StringLen(GUICtrlRead($InputFilter)) & "  " & GUICtrlRead($InputFilter))

    EndIf

    ;SetStatusMessage(StringLen(GUICtrlRead($InputFilter)) & "  " & GUICtrlRead($InputFilter))
    ;_Debug(@ScriptLineNumber & "++++" & _ArrayToString($FilterArray))
    For $x = 1 To _EventLog__Count($hEventLog)
        If GUICtrlRead($CheckAbort) = $GUI_CHECKED Then Return
        Local $TmpString = ''
        Local $a = _EventLog__Read($hEventLog)

        $Counter = $Counter + 1
        Switch $a[7]
            Case 0
                $Other += 1
                If GUICtrlRead($CheckOther) = $GUI_CHECKED Then $TmpString = String($a[1]) & "~" & "Other:" & String($a[6]) & "~" & String($a[2]) & " " & String($a[3]) & "~" & String($a[10]) & "~" & String($a[13])
            Case 1
                $Errors += 1
                If GUICtrlRead($CheckErrors) = $GUI_CHECKED Then $TmpString = String($a[1]) & "~" & "Error:" & String($a[6]) & "~" & String($a[2]) & " " & String($a[3]) & "~" & String($a[10]) & "~" & String($a[13])
            Case 2
                $Warnings += 1
                If GUICtrlRead($CheckWarnings) = $GUI_CHECKED Then $TmpString = String($a[1]) & "~" & "Warning:" & String($a[6]) & "~" & String($a[2]) & " " & String($a[3]) & "~" & String($a[10]) & "~" & String($a[13])
            Case 4
                $Information += 1
                If GUICtrlRead($CheckInformation) = $GUI_CHECKED Then $TmpString = ($a[1]) & "~" & "Information:" & String($a[6]) & "~" & String($a[2]) & "  " & String($a[3]) & "~" & String($a[10]) & "~" & String($a[13])
            Case 8
                $SuccessAudit += 1
                If GUICtrlRead($CheckSuccessAudits) = $GUI_CHECKED Then $TmpString = String($a[1]) & "~" & "Success Audit:" & String($a[6]) & "~" & String($a[2]) & "  " & String($a[3]) & "~" & String($a[10]) & "~" & String($a[13])
            Case 16
                $FailureAudit += 1
                If GUICtrlRead($CheckFailAudits) = $GUI_CHECKED Then $TmpString = String($a[1]) & "~" & "Fail Audit:" & String($a[6]) & "~" & String($a[2]) & "  " & String($a[3]) & "~" & String($a[10]) & "~" & String($a[13])
            Case Else
                $Else += 1
                If GUICtrlRead($CheckOther) = $GUI_CHECKED Then $TmpString = String($a[1]) & "~" & "Else:" & String($a[6]) & "~" & String($a[2]) & "  " & String($a[3]) & "~" & String($a[10]) & "~" & String($a[13])
        EndSwitch

        If StringInStr($Type, 'Verbose') > 0 Then
            If StringLen($TmpString) > 0 Then ; There is a tmpstring so it needs to be tested
                Local $E
                Local $tag
                $E = StringSplit($TmpString, "~", 1)

                ;If there is NOT a filter always show the string
                ;Local $Ignore = GUICtrlRead($CheckIgnore) = $GUI_CHECKED
                If (StringLen(GUICtrlRead($InputFilter)) = 0) Or (GUICtrlRead($CheckIgnore) = $GUI_CHECKED) Then
                    $tag = "show"

                Else ;TODO if there is a filter string do the following
                    For $a In $FilterArray ;loop throught the array for matches
                        Local $SiSPosition = StringInStr($TmpString, $a)
                        Local $Exclude = GUICtrlRead($CheckExclude) = $GUI_CHECKED

                        If $Exclude = True Then $tag = 'show'
                        If $Exclude = False Then $tag = 'hide'

                        ;SetStatusMessage(@ScriptLineNumber & "<<>>" & $SiSPosition & "<<>>" & $Exclude & "<<>>" & $a & "<<>>" & "<<>>" & $TmpString)
                        If $Exclude = True And $SiSPosition > 0 Then ;If anything matches don't display the line
                            $tag = 'hide'
                            ExitLoop
                        EndIf
                        If $Exclude = False And $SiSPosition > 0 Then ;If anything matches do display the line
                            $tag = 'show'
                            ExitLoop
                        EndIf
                    Next
                EndIf

                DisplayResults($E, $Counter, $tag) ; If no filter do this
            EndIf
        EndIf
    Next

    If StringInStr($Type, 'StatsOnly') > 0 Then
        Local $ListView_item
        $ListView_item = _GUICtrlListView_AddItem($ListView, "Errors")
        _GUICtrlListView_AddSubItem($ListView, $ListView_item, $Errors, 1)
        $ListView_item = _GUICtrlListView_AddItem($ListView, "Warnings")
        _GUICtrlListView_AddSubItem($ListView, $ListView_item, $Warnings, 1)
        $ListView_item = _GUICtrlListView_AddItem($ListView, "Information")
        _GUICtrlListView_AddSubItem($ListView, $ListView_item, $Information, 1)
        $ListView_item = _GUICtrlListView_AddItem($ListView, "SuccessAudit")
        _GUICtrlListView_AddSubItem($ListView, $ListView_item, $SuccessAudit, 1)
        $ListView_item = _GUICtrlListView_AddItem($ListView, "FailureAudit")
        _GUICtrlListView_AddSubItem($ListView, $ListView_item, $FailureAudit, 1)
        $ListView_item = _GUICtrlListView_AddItem($ListView, "Other")
        _GUICtrlListView_AddSubItem($ListView, $ListView_item, $Other, 1)
        $ListView_item = _GUICtrlListView_AddItem($ListView, "Else")
        _GUICtrlListView_AddSubItem($ListView, $ListView_item, $Else, 1)

    EndIf

    $DebugString = @ScriptLineNumber & " Operation complete"
    SetStatusMessage($DebugString)
    GUICtrlSetBkColor($LabelWorking, $CLR_NONE)
    GUICtrlSetData($LabelWorking, 'Idle')
EndFunc   ;==>GetStats
;-----------------------------------------------
; If $tag is "hide" then we exit
Func DisplayResults($E, $Counter, $tag = "")
    Local $ListView_item
    If StringInStr($tag, "hide") Then Return
    If $E[0] >= 1 Then $ListView_item = _GUICtrlListView_AddItem($ListView, StringFormat("%3s  %6s", $Counter, $E[1]))
    If $E[0] >= 2 Then _GUICtrlListView_AddSubItem($ListView, $ListView_item, $E[2], $Type)
    If $E[0] >= 3 Then _GUICtrlListView_AddSubItem($ListView, $ListView_item, $E[3], $Date)
    If $E[0] >= 4 Then _GUICtrlListView_AddSubItem($ListView, $ListView_item, $E[4], $Source)
    If $E[0] >= 5 Then _GUICtrlListView_AddSubItem($ListView, $ListView_item, $E[5], $Description)
EndFunc   ;==>DisplayResults
;-----------------------------------------------

Func CheckDeviceManager()
    SetStatusMessage("Start Device manager")
    GUICtrlSetBkColor($LabelWorking, $COLOR_RED)
    GUICtrlSetData($LabelWorking, 'Start device manager')
    ShellExecute("C:\Windows\System32\devmgmt.msc")
    SetStatusMessage("Start Device manager complete")
    GUICtrlSetBkColor($LabelWorking, $CLR_NONE)
    GUICtrlSetData($LabelWorking, 'Idle')
EndFunc   ;==>CheckDeviceManager

;-----------------------------------------------
Func SearchDMP() ;Search the hard drive for dmp files and optionaly move them to Downloads
    SetStatusMessage(@ScriptLineNumber & " Search for DMP files start")
    GUICtrlSetBkColor($LabelWorking, $COLOR_RED)
    GUICtrlSetData($LabelWorking, 'Search for DMP files')
    Local $results = _FileListToArrayRec("c:\", "*.dmp", 1, 1, 0, 2)
    If @error == 0 Then
        SetStatusMessage(@ScriptLineNumber & " " & UBound($results) - 1 & " DMP files found")

        Local $sString
        _ArrayDelete($results, 0)
        For $vElement In $results
            $sString = $sString & $vElement & @CRLF
        Next

        If MsgBox($MB_ICONQUESTION + $MB_YESNO + $MB_TOPMOST, "DMP files found", $sString & @CRLF & "Move " & UBound($results) & " files to Downloads?") = $IDYES Then
            While UBound($results)
                Local $TString = _ArrayPop($results)
                If StringInStr($TString, ".dmp") > 0 Then
                    Local $tS1 = StringReplace($TString, "\", "~")
                    Local $tS2 = StringReplace($tS1, ":", "!")
                    Local $tS3 = StringReplace($tS2, ".dmp", ".dm^")

                    If FileMove($TString, EnvGet("HOMEPATH") & "\Downloads\" & $tS3, $FC_NOOVERWRITE) = 0 Then
                        MsgBox($MB_ICONQUESTION + $MB_YESNO, "Problem", $TString & @CRLF & $tS1 & @CRLF & $tS3)
                    EndIf
                EndIf
            WEnd
        EndIf

    Else
        MsgBox($MB_ICONINFORMATION + $MB_TOPMOST, "Search DMP", "No DMP files found")
    EndIf

    SetStatusMessage(@ScriptLineNumber & " Search for DMP files complete")
    GUICtrlSetBkColor($LabelWorking, $CLR_NONE)
    GUICtrlSetData($LabelWorking, 'Idle')
EndFunc   ;==>SearchDMP

;-----------------------------------------------

Func SetDefaults($Type = 'all')
    $DebugString = @ScriptLineNumber & " SetDefaults " & $Type
    SetStatusMessage($DebugString)

    If $Type = 'all' Then
        If WinMove($ProgramName, "", 10, 10, 780, 540) <> 0 Then
            GUICtrlSetPos($LabelStatus, -1, -1, 780 - 50, -1)
            GUICtrlSetPos($ListView, -1, -1, 780 - 50, 540 - 300)
        Else
            MsgBox($MB_ICONERROR + $MB_TOPMOST, "WinMove error 1 ", $ProgramName & " window not found")
        EndIf

        GUICtrlSetData($SliderDelay, 5)
        GUICtrlSetData($LabelSliderDelay, GUICtrlRead($SliderDelay))
        GUICtrlSetState($CheckSave, $GUI_UNCHECKED)
        GUICtrlSetState($CheckAbort, $GUI_UNCHECKED)
    EndIf

    If $Type = "log" Or $Type = 'all' Then
        GUICtrlSetState($CheckSystem, $GUI_CHECKED)
        GUICtrlSetState($CheckSecurity, $GUI_UNCHECKED)
        GUICtrlSetState($CheckApplication, $GUI_CHECKED)
    EndIf

    If $Type = 'event' Or $Type = 'all' Then
        GUICtrlSetState($CheckInformation, $GUI_UNCHECKED)
        GUICtrlSetState($CheckWarnings, $GUI_CHECKED)
        GUICtrlSetState($CheckErrors, $GUI_CHECKED)
        GUICtrlSetState($CheckSuccessAudits, $GUI_UNCHECKED)
        GUICtrlSetState($CheckFailAudits, $GUI_CHECKED)
        GUICtrlSetState($CheckOther, $GUI_CHECKED)
    EndIf
EndFunc   ;==>SetDefaults
;-----------------------------------------------
Func SaveProject()
    $DebugString = @ScriptLineNumber & " SaveProject"
    SetStatusMessage($DebugString)
    $Project_filename = FileSaveDialog("Save project file", @ScriptDir & "\AUXFiles\", _
            $ProgramName & " projects (E*.prj)|All projects (*.prj)|All files (*.*)", 18, @ScriptDir & "\AUXFiles\" & $ProgramName & ".prj")

    Local $file = FileOpen($Project_filename, 2)
    ; Check if file opened for writing OK
    If $file = -1 Then
        $DebugString = @ScriptLineNumber & " SaveProject: Unable to open file for writing: " & $Project_filename
        SetStatusMessage($DebugString)
        Return
    EndIf
    $DebugString = @ScriptLineNumber & " SaveProject  " & $Project_filename
    SetStatusMessage($DebugString)
    FileWriteLine($file, "Valid for EventMon project")
    FileWriteLine($file, "Project file for " & @ScriptName & "  " & _DateTimeFormat(_NowCalc(), 0))
    FileWriteLine($file, "Help 1 is enabled, 4 is disabled for checkboxes")
    FileWriteLine($file, "CheckSave:" & GUICtrlRead($CheckSave))

    FileWriteLine($file, "CheckSystem:" & GUICtrlRead($CheckSystem))
    FileWriteLine($file, "CheckSecurity:" & GUICtrlRead($CheckSecurity))
    FileWriteLine($file, "CheckApplication:" & GUICtrlRead($CheckApplication))
    FileWriteLine($file, "CheckInformation:" & GUICtrlRead($CheckInformation))
    FileWriteLine($file, "CheckWarnings:" & GUICtrlRead($CheckWarnings))
    FileWriteLine($file, "CheckErrors:" & GUICtrlRead($CheckErrors))
    FileWriteLine($file, "CheckSuccessAudits:" & GUICtrlRead($CheckSuccessAudits))
    FileWriteLine($file, "CheckFailAudits:" & GUICtrlRead($CheckFailAudits))
    FileWriteLine($file, "CheckOther:" & GUICtrlRead($CheckOther))
    FileWriteLine($file, "LabelSliderDelay:" & GUICtrlRead($LabelSliderDelay))
    FileWriteLine($file, "SliderDelay:" & GUICtrlRead($SliderDelay))
    FileWriteLine($file, "CheckExclude:" & GUICtrlRead($CheckExclude))
    FileWriteLine($file, "CheckIgnore:" & GUICtrlRead($CheckIgnore))
    FileWriteLine($file, "InputFilter:" & GUICtrlRead($InputFilter))

    Local $F = WinGetPos($ProgramName, "")
    If @error = 0 Then
        FileWriteLine($file, "MainWinpos:" & $F[0] & " " & $F[1] & " " & $F[2] & " " & $F[3])
    Else
        $DebugString = @ScriptLineNumber & " SaveProject: Unable to get " & $ProgramName & " window position"
        SetStatusMessage($DebugString)
    EndIf

    FileClose($file)
EndFunc   ;==>SaveProject
;-----------------------------------------------
;This loads the project file but into the tree control but not into the list
Func LoadProject($Type)
    $DebugString = @ScriptLineNumber & " LoadProject " & $Type
    SetStatusMessage($DebugString)

    If StringCompare($Type, "menu") = 0 Then
        $Project_filename = FileOpenDialog("Load project file", @ScriptDir & "\AUXFiles\", _
                $ProgramName & " projects (E*.prj)|All projects (*.prj)|All files (*.*)", 18, @ScriptDir & "\AUXFiles\" & $ProgramName & ".prj")
    EndIf

    Local $file = FileOpen($Project_filename, 0)
    ; Check if file opened for reading OK
    If $file = -1 Then
        $DebugString = @ScriptLineNumber & " LoadProject: Unable to open file for reading: " & $Project_filename
        SetStatusMessage($DebugString)
        Return
    EndIf

    $DebugString = @ScriptLineNumber & " LoadProject   " & $Project_filename
    SetStatusMessage($DebugString)
    ; Read in the first line to verify the file is of the correct type
    If StringCompare(FileReadLine($file, 1), "Valid for EventMon project") <> 0 Then
        MsgBox($MB_ICONERROR + $MB_TOPMOST, "Invalid project file", "Not a valid EventMon project file")
        FileClose($file)
        Return
    EndIf

    ; Read in lines of text until the EOF is reached
    While 1
        Local $LineIn = FileReadLine($file)
        If @error = -1 Then ExitLoop

        $DebugString = @ScriptLineNumber & " LoadProject " & $LineIn
        SetStatusMessage($DebugString)
        If StringInStr($LineIn, ";") = 1 Then ContinueLoop

        If StringInStr($LineIn, "CheckSave:") Then GUICtrlSetState($CheckSave, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckSystem:") Then GUICtrlSetState($CheckSystem, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckSecurity:") Then GUICtrlSetState($CheckSecurity, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckApplication:") Then GUICtrlSetState($CheckApplication, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckInformation:") Then GUICtrlSetState($CheckInformation, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckAbort:") Then GUICtrlSetState($CheckAbort, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckWarnings:") Then GUICtrlSetState($CheckWarnings, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckErrors:") Then GUICtrlSetState($CheckErrors, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckSuccessAudits:") Then GUICtrlSetState($CheckSuccessAudits, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckFailAudits:") Then GUICtrlSetState($CheckFailAudits, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckOther:") Then GUICtrlSetState($CheckOther, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckExclude:") Then GUICtrlSetState($CheckExclude, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckIgnore:") Then GUICtrlSetState($CheckIgnore, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "InputFilter:") Then GUICtrlSetData($InputFilter, StringMid($LineIn, StringInStr($LineIn, ":") + 1))

        Local $F
        If StringInStr($LineIn, "MainWinpos:") Then
            $F = StringMid($LineIn, StringInStr($LineIn, ":") + 1)
            $F = StringSplit($F, " ", 2)
            If WinMove($ProgramName, "", $F[0], $F[1], $F[2], $F[3]) <> 0 Then
                GUICtrlSetPos($LabelStatus, -1, -1, $F[2] - 50, -1)
                GUICtrlSetPos($ListView, -1, -1, $F[2] - 50, $F[3] - 310)
            Else
                MsgBox($MB_ICONERROR + $MB_TOPMOST, "WinMove error 1 ", $ProgramName & " window not found")
            EndIf
        EndIf
    WEnd

    FileClose($file)

    ; If the main window is not visible (off screen), make it visible
    $F = GetWinPos($ProgramName)
    $DebugString = @ScriptLineNumber & " DesktopWidth: " & $F[0] & " " & @DesktopWidth
    SetStatusMessage($DebugString)
    $DebugString = @ScriptLineNumber & " DesktopHeight: " & $F[1] & " " & @DesktopHeight
    SetStatusMessage($DebugString)

    If $F[0] > @DesktopWidth Or $F[1] > @DesktopHeight Then WinMove($ProgramName, "", 10, 10, 700, 515)
    If $F[0] < 0 Or $F[1] < 0 Then WinMove($ProgramName, "", 10, 10, 700, 515)

EndFunc   ;==>LoadProject

;-----------------------------------------------
Func Toggle($Type = 'all')
    $DebugString = @ScriptLineNumber & " Toggle " & $Type & @CRLF
    SetStatusMessage($DebugString)
    If $ToggleState = $GUI_CHECKED Then
        $ToggleState = $GUI_UNCHECKED
    Else
        $ToggleState = $GUI_CHECKED
    EndIf
    ;GUICtrlSetState($CheckSave, $ToggleState)

    If $Type = "log" Or $Type = 'all' Then
        GUICtrlSetState($CheckSystem, $ToggleState)
        GUICtrlSetState($CheckSecurity, $ToggleState)
        GUICtrlSetState($CheckApplication, $ToggleState)
    EndIf
    If $Type = 'event' Or $Type = 'all' Then
        GUICtrlSetState($CheckInformation, $ToggleState)
        GUICtrlSetState($CheckWarnings, $ToggleState)
        GUICtrlSetState($CheckErrors, $ToggleState)
        GUICtrlSetState($CheckSuccessAudits, $ToggleState)
        GUICtrlSetState($CheckFailAudits, $ToggleState)
        GUICtrlSetState($CheckOther, $ToggleState)
    EndIf
EndFunc   ;==>Toggle
;-----------------------------------------------
Func About(Const $FormID)
    $DebugString = @ScriptLineNumber & " About " & $FormID
    SetStatusMessage($DebugString)
    Local $D = GetWinPos($FormID)
    Local $WinPos
    If IsArray($D) = True Then
        $DebugString = @ScriptLineNumber & " " & $FormID
        SetStatusMessage($DebugString)
        $WinPos = StringFormat("%s" & @CRLF & "WinPOS: %d  %d " & @CRLF & "WinSize: %d %d " & @CRLF & "Desktop: %d %d ", _
                $FormID, $D[0], $D[1], $D[2], $D[3], @DesktopWidth, @DesktopHeight)
    Else
        $WinPos = ">>>About ERROR, Check the window name<<<"
    EndIf
    MsgBox($MB_ICONINFORMATION + $MB_TOPMOST, "About", $SystemS & @CRLF & $WinPos & @CRLF & "Written by Doug Kaynor!")
EndFunc   ;==>About
;-----------------------------------------------
Func Help()
    $DebugString = @ScriptLineNumber & " Help"
    SetStatusMessage($DebugString)
    Local $helpstr = "EventMon  " & $FileVersion & @CRLF & _
            "Startup options: " & @CRLF & _
            "help or ?   Display this help file" & @CRLF & _
            "clear       Clear all logs." & @CRLF
    MsgBox($MB_ICONINFORMATION + $MB_TOPMOST, "EventMon startup help", @CRLF & @CRLF & $helpstr)
EndFunc   ;==>Help
;-----------------------------------------------
Func GetWinPos($WinName)
    $DebugString = @ScriptLineNumber & " GetWinPos " & $WinName
    SetStatusMessage($DebugString)
    Local $F
    Local $g
    While Not IsArray($F)
        $F = WinGetPos($WinName)
        Sleep(100)
        $g = $g + 1
        If $g > 100 Then ExitLoop
    WEnd
    Return ($F)
EndFunc   ;==>GetWinPos
;-----------------------------------------------
Func SetStatusMessage($Message, $Color = False)
    _Debug(@ScriptName & " " & $Message, $LOG_filename)
    If StringInStr($Message, "-1") = 1 Then $Message = StringTrimLeft($Message, 3)
    GUICtrlSetData($LabelStatus, $Message)
    If $Color Then
        GUICtrlSetBkColor($LabelStatus, $COLOR_FUCHSIA)
    Else
        GUICtrlSetBkColor($LabelStatus, $CLR_NONE)
    EndIf
EndFunc   ;==>SetStatusMessage

;-----------------------------------------------
