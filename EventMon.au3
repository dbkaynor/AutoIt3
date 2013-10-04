#AutoIt3Wrapper_icon=../icons/openbsd.ico
#AutoIt3Wrapper_outfile=EventMon.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Comment=A EventLog tool
#AutoIt3Wrapper_Res_Description=EventLog tool
#AutoIt3Wrapper_Res_Fileversion=1.0.0.32
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=Y
#AutoIt3Wrapper_Res_ProductVersion=666
#AutoIt3Wrapper_Res_LegalCopyright=Copyright 2011 Douglas B Kaynor
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
#endregion

Opt("MustDeclareVars", 1)
Opt("TrayIconDebug", 1)
Opt("TrayAutoPause", 0)
Opt("WinTitleMatchMode", 2)

;#include <Array.au3>
#include <ButtonConstants.au3>
#include <Constants.au3>
#include <Date.au3>
#include <EventLog.au3>
#include <GUIConstantsEx.au3>
#include <GuiListView.au3>
#include <ListViewConstants.au3>
#include <Misc.au3>
#include <SliderConstants.au3>
#include <StaticConstants.au3>
#include <String.au3>
#include <WindowsConstants.au3>
#include <_DougFunctions.au3>

;DirCreate("AUXFiles")
;Global $AuxPath = @ScriptDir & "\AUXFiles\"
;FileInstall($AuxPath & "Working.jpg", $AuxPath & "Working.jpg", 0)

;Global Const $tmp = StringSplit(@ScriptName, ".")
Global Const $ProgramName = "Event log monitor"
Global $Project_filename = @ScriptDir & "\AUXFiles\" & $ProgramName & ".prj"
Global $LOG_filename = @ScriptDir & "\AUXFiles\" & $ProgramName & ".log"

Global $hEventLog

If _Singleton($ProgramName, 1) = 0 Then
    MsgBox(48, "Already running", $ProgramName & " is already running!", 10)
    Exit
EndIf

Global Const $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")
Global $SystemS = $ProgramName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & @OSServicePack & @CRLF & @OSType & @CRLF & @OSArch

; Mainform
Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $MainForm = GUICreate("Event log monitor  " & $FileVersion, 630, 515, 200, 150, $MainFormOptions)

Global $ButtonGetData = GUICtrlCreateButton("Get data", 10, 10, 75, 25)
GUICtrlSetTip(-1, "Fetch the data from selected logs")
GUICtrlSetResizing(-1, 802)
Global $ButtonClearLogs = GUICtrlCreateButton("Clear logs", 180, 10, 75, 25)
GUICtrlSetTip(-1, "Clear all logs. Check 'Save' if you want to save before clearing")
GUICtrlSetResizing(-1, 802)
Global $ButtonEventViewer = GUICtrlCreateButton("Event viewer", 95, 10, 75, 25)
GUICtrlSetTip(-1, "Open the event control viewer")
GUICtrlSetResizing(-1, 802)
Global $ButtonOpenSaveFolder = GUICtrlCreateButton("Open save", 265, 10, 75, 25)
GUICtrlSetTip(-1, "Open the log save location.")
GUICtrlSetResizing(-1, 802)
Global $ButtonAbout = GUICtrlCreateButton("About", 350, 10, 75, 25)
GUICtrlSetTip(-1, "Display application information.")
GUICtrlSetResizing(-1, 802)
Global $ButtonHelp = GUICtrlCreateButton("Help", 430, 10, 75, 25)
GUICtrlSetTip(-1, "Display application help")
GUICtrlSetResizing(-1, 802)
Global $ButtonExit = GUICtrlCreateButton("Exit", 520, 10, 75, 25)
GUICtrlSetTip(-1, "Exit this application")
GUICtrlSetResizing(-1, 802)

Global $ButtonSaveProject = GUICtrlCreateButton("Save project", 350, 40, 75, 25)
GUICtrlSetTip(-1, "Save the current settings")
GUICtrlSetResizing(-1, 802)

Global $ButtonLoadProject = GUICtrlCreateButton("Load project", 430, 40, 75, 25)
GUICtrlSetTip(-1, "Load saved settings")
GUICtrlSetResizing(-1, 802)

Global $ButtonLoadDefaults = GUICtrlCreateButton("Load defaults", 520, 40, 75, 25)
GUICtrlSetTip(-1, "Load default settings")
GUICtrlSetResizing(-1, 802)

Global $GroupOptions = GUICtrlCreateGroup("Options", 30, 40, 240, 40)
GUICtrlSetResizing(-1, 802)
Global $CheckSave = GUICtrlCreateCheckbox("Save", 40, 55, 50, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Save logs before clearing")
Global $CheckSplash = GUICtrlCreateCheckbox("Splash", 110, 55, 50, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Display splash")
Global $CheckAbort = GUICtrlCreateCheckbox("Abort", 170, 55, 50, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Abort operations")
GUICtrlCreateGroup("", -99, -99, 1, 1)

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

GUICtrlCreateGroup("Filter event type", 30, 160, 240, 40)
GUICtrlSetResizing(-1, 802)
Global $InputFilterNumber = GUICtrlCreateInput('', 40, 175, 150, 20)
GUICtrlSetTip(-1, "Input a data filter. Supports regular expressions")
GUICtrlSetResizing(-1, 802)
Global $CheckExcludeNumber = GUICtrlCreateCheckbox('Exclude', 200, 175, 85, 20)
GUICtrlSetTip(-1, "Exclude or include the filter string")
GUICtrlSetResizing(-1, 802)
GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateGroup("Filter event source", 300, 160, 240, 40)
GUICtrlSetResizing(-1, 802)
Global $InputFilterSource = GUICtrlCreateInput('', 310, 175, 150, 20)
GUICtrlSetTip(-1, "Input a data filter. Supports regular expressions")
GUICtrlSetResizing(-1, 802)
Global $CheckExcludeName = GUICtrlCreateCheckbox('Exclude', 470, 175, 85, 20)
GUICtrlSetTip(-1, "Exclude or include the filter string")
GUICtrlSetResizing(-1, 802)
GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateGroup("Loop", 370, 80, 250, 40)
GUICtrlSetResizing(-1, 802)
Global $CheckLoop = GUICtrlCreateCheckbox("Loop", 380, 95, 40, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Start/Stop looping")
Global $SliderDelay = GUICtrlCreateSlider(440, 95, 140, 20, $TBS_AUTOTICKS);, BitOR($TBS_AUTOTICKS, $TBS_NOTICKS))
GUICtrlSetResizing(-1, 802)
GUICtrlSetLimit(-1, 60, 1)
GUICtrlSetData(-1, 2)
GUICtrlSetTip(-1, "Loop delay setting")
Global $LabelSliderDelay = GUICtrlCreateLabel(GUICtrlRead($SliderDelay), 580, 95, 20, 20, $SS_SUNKEN)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Current loop delay setting")
GUICtrlCreateGroup("", -99, -99, 1, 1)

Const $Number = 0
Const $Type = 1
Const $Date = 2
Const $Source = 3
Const $Description = 4

Global $ListView = GUICtrlCreateListView("Number|Event type|Date logged|Event source|Event description", 20, 210, 590, 280, $LVS_REPORT, BitOR($LVS_EX_FULLROWSELECT, $WS_EX_CLIENTEDGE, $LVS_EX_GRIDLINES))
GUICtrlSetTip(-1, "This is the list box")
GUICtrlSetResizing(-1, BitOR($GUI_DOCKTOP, $GUI_DOCKBOTTOM))
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, $Number, 70)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, $Type, 120)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, $Date, 145)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, $Source, 150)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, $Description, 1000)

Global $LabelStatus = GUICtrlCreateLabel("Status main", 10, 490, 590, 20, $SS_SUNKEN)
GUICtrlSetResizing(-1, 64)
GUICtrlSetTip(-1, "Status message")

Static $ToggleState = $GUI_CHECKED
Static $SavedTime = TimerInit()

SetDefaults()

LoadProject("start")

GUISetState(@SW_SHOW)

While 1
    Global $t = GUIGetMsg()
    Switch $t
        Case $ButtonGetData
            _GUICtrlListView_DeleteAllItems(GUICtrlGetHandle($ListView))
            GetData()
        Case $ButtonEventViewer
            ShellExecuteWait("eventvwr.msc", "/s")
        Case $ButtonClearLogs
            ClearLogs()
        Case $ButtonOpenSaveFolder
            ClipPut(EnvGet("temp"))
            RunWait("C:\WINDOWS\EXPLORER.EXE /n,/e," & EnvGet("temp"))
        Case $SliderDelay
            GUICtrlSetData($LabelSliderDelay, GUICtrlRead($SliderDelay))
        Case $ButtonToggleLogType
            Toggle('Log')
        Case $ButtonDefaultsLogType
            SetDefaults('log')
        Case $ButtonToggleEventType
            Toggle('event')
        Case $ButtonDefaultsEventType
            SetDefaults('event')
        Case $ButtonSaveProject
            SaveProject()
        Case $ButtonLoadProject
            LoadProject("Menu")
        Case $ButtonLoadDefaults
            SetDefaults('all')
        Case $ButtonAbout
            About($ProgramName)
        Case $ButtonHelp
            Help()
        Case $ButtonExit
            _Debug("$ButtonExit")
            Exit
        Case $GUI_EVENT_CLOSE
            _Debug("GUI_EVENT_CLOSE")
            Exit
    EndSwitch
    CheckChangeCounter()
WEnd
;-----------------------------------------------
Func CheckChangeCounter()
    If GUICtrlRead($CheckLoop) = $GUI_UNCHECKED Then
        ; $ShowSplash = True
        Return
    EndIf
    ;$ShowSplash = False
    Local $CurrentTime = TimerDiff($SavedTime) / 1000 ; seconds

    If GUICtrlRead($CheckAbort) = $GUI_CHECKED Then Return

    If Mod(Int($CurrentTime * 400), 10) <> 0 Then Return
    If $CurrentTime > GUICtrlRead($SliderDelay) Then
        ConsoleWrite(@CRLF & @ScriptLineNumber & " " & $CurrentTime & " " & GUICtrlRead($SliderDelay) & " " & @CRLF)
        $SavedTime = TimerInit()
        _GUICtrlListView_DeleteAllItems(GUICtrlGetHandle($ListView))
        GetData()

    EndIf
EndFunc   ;==>CheckChangeCounter
;-----------------------------------------------
Func ClearLogs()
    If MsgBox(32 + 1, "Clear all logs", "Are you sure? (verify save check box!)") = 2 Then Return

    Local $tCur = _Date_Time_GetLocalTime()
    Local $tStr = _Date_Time_SystemTimeToDateTimeStr($tCur)
    $tStr = StringRegExpReplace($tStr, "[:/ ]", "", 0)
    ConsoleWrite(@ScriptLineNumber & " " & $tStr & @CRLF)
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
        MsgBox(48, "Log files saved", "Saved to " & EnvGet("temp"))
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
EndFunc   ;==>ClearLogs
;-----------------------------------------------
Func GetData()
    Local $ListView_item
    GUICtrlSetState($CheckAbort, $GUI_UNCHECKED)

    If GUICtrlRead($CheckSplash) = $GUI_CHECKED Then SplashImageOn("Working", $AuxPath & "Working.jpg", -1, -1, -1, -1, 18)
    If GUICtrlRead($CheckAbort) = $GUI_CHECKED Then Return
    If GUICtrlRead($CheckApplication) = $GUI_CHECKED Then
        $ListView_item = _GUICtrlListView_AddItem($ListView, '')
        _GUICtrlListView_AddSubItem($ListView, $ListView_item, '--Application log--', $Type)
        $hEventLog = _EventLog__Open("", "Application")
        GetInformation($hEventLog)
        _EventLog__Close($hEventLog)
        _GUICtrlListView_AddItem($ListView, "  ")
    EndIf
    If GUICtrlRead($CheckAbort) = $GUI_CHECKED Then Return
    If GUICtrlRead($CheckSecurity) = $GUI_CHECKED Then
        $ListView_item = _GUICtrlListView_AddItem($ListView, '')
        _GUICtrlListView_AddSubItem($ListView, $ListView_item, '--Security log--', $Type)
        $hEventLog = _EventLog__Open("", "Security")
        GetInformation($hEventLog)
        _EventLog__Close($hEventLog)
        _GUICtrlListView_AddItem($ListView, "  ")
    EndIf
    If GUICtrlRead($CheckAbort) = $GUI_CHECKED Then Return
    If GUICtrlRead($CheckSystem) = $GUI_CHECKED Then
        $ListView_item = _GUICtrlListView_AddItem($ListView, '')
        _GUICtrlListView_AddSubItem($ListView, $ListView_item, '--System log--', $Type)
        $hEventLog = _EventLog__Open("", "System")
        GetInformation($hEventLog)
        _EventLog__Close($hEventLog)
        _GUICtrlListView_AddItem($ListView, "  ")
    EndIf

    SplashOff()
EndFunc   ;==>GetData
;-----------------------------------------------
Func GetInformation($hEventLog)
    Local $ListView_item
    $ListView_item = _GUICtrlListView_AddItem($ListView, '')
    _GUICtrlListView_AddSubItem($ListView, $ListView_item, "Log full", $Type)
    _GUICtrlListView_AddSubItem($ListView, $ListView_item, _EventLog__Full($hEventLog), $Date)
    $ListView_item = _GUICtrlListView_AddItem($ListView, '')
    _GUICtrlListView_AddSubItem($ListView, $ListView_item, "Event count", $Type)
    _GUICtrlListView_AddSubItem($ListView, $ListView_item, _EventLog__Count($hEventLog), $Date)
    GetStats($hEventLog)
EndFunc   ;==>GetInformation
;-----------------------------------------------
Func GetStats($hEventLog)
    Local $Errors = 0
    Local $Warnings = 0
    Local $Information = 0
    Local $SuccessAudit = 0
    Local $FailureAudit = 0
    Local $Other = 0
    Local $Else = 0

    Local $Counter = 0
    #cs
        Const $Number = 0
        Const $Type = 1
        Const $Date = 2
        Const $Source = 3
        Const $Description = 4
    #ce

    For $x = 1 To _EventLog__Count($hEventLog)
        If GUICtrlRead($CheckAbort) = $GUI_CHECKED Then Return
        Local $TmpString = ''
        Local $a = _EventLog__Read($hEventLog)
        $Counter = $Counter + 1
        Switch $a[7]
            Case 0
                $Other += 1
                If GUICtrlRead($CheckOther) = $GUI_CHECKED Then $TmpString = String($a[1]) & "~" & "Other:" & String($a[6]) & "~" & String($a[2]) & " " & String($a[3]) & "~" & String($a[10]) & "~" & String($a[13])
                ;$TmpString = "Other:" & String($a[6]) & "  " & String($a[7])
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
                ;$TmpString = "Other:" & String($a[6]) & "  " & String($a[7])
        EndSwitch
        Local $E
        If StringLen($TmpString) > 0 Then ; There is a tmpstring
            If StringLen(GUICtrlRead($InputFilterNumber)) = 0 And StringLen(GUICtrlRead($InputFilterSource)) = 0 Then ; There is no filter
                $E = StringSplit($TmpString, "~", 1)
                DisplayResults($E, $Counter)
            Else ;Here we have at least one filter
                $E = StringSplit($TmpString, "~")
                Local $F = StringSplit($E[2], ":", 2)
                ConsoleWrite(@ScriptLineNumber & ": " & _ArrayToString($F) & @CRLF)
                ConsoleWrite(@ScriptLineNumber & ": " & UBound($F) & @CRLF)
                If UBound($F) > 2 Then
                    Local $result
                    $result = StringRegExp($F[1], GUICtrlRead($InputFilterNumber), 0)
                    If GUICtrlRead($CheckExcludeNumber) = $GUI_CHECKED And $result = 0 Then DisplayResults($E, $Counter)
                    If GUICtrlRead($CheckExcludeNumber) = $GUI_UNCHECKED And $result > 0 Then DisplayResults($E, $Counter)

                    $result = StringRegExp($F[3], GUICtrlRead($InputFilterSource), 0)
                    If GUICtrlRead($CheckExcludeName) = $GUI_CHECKED And $result = 0 Then DisplayResults($E, $Counter)
                    If GUICtrlRead($CheckExcludeName) = $GUI_UNCHECKED And $result > 0 Then DisplayResults($E, $Counter)
                EndIf
            EndIf
        EndIf
    Next

    GUICtrlSetData($LabelStatus, StringFormat("Errors:%s  Warnings:%s  Information:%s  SuccessAudit:%s  FailureAudit:%s  Other:%s   Else:%s" & @CRLF, $Errors, $Warnings, $Information, $SuccessAudit, $FailureAudit, $Other, $Else))

EndFunc   ;==>GetStats
;-----------------------------------------------
Func DisplayResults($E, $Counter)
    Local $ListView_item
    If $E[0] >= 1 Then $ListView_item = _GUICtrlListView_AddItem($ListView, StringFormat("%3s  %6s", $Counter, $E[1]))
    If $E[0] >= 2 Then _GUICtrlListView_AddSubItem($ListView, $ListView_item, $E[2], $Type)
    If $E[0] >= 3 Then _GUICtrlListView_AddSubItem($ListView, $ListView_item, $E[3], $Date)
    If $E[0] >= 4 Then _GUICtrlListView_AddSubItem($ListView, $ListView_item, $E[4], $Source)
    If $E[0] >= 5 Then _GUICtrlListView_AddSubItem($ListView, $ListView_item, $E[5], $Description)
EndFunc   ;==>DisplayResults
;-----------------------------------------------
Func SetDefaults($Type = 'all')
    _Debug($ProgramName)
    WinMove($ProgramName, "", 10, 10, 630, 515)
    GUICtrlSetData($SliderDelay, 2)
    GUICtrlSetData($LabelSliderDelay, GUICtrlRead($SliderDelay))
    GUICtrlSetState($CheckSave, $GUI_UNCHECKED)
    GUICtrlSetState($CheckSplash, $GUI_UNCHECKED)
    GUICtrlSetState($CheckAbort, $GUI_UNCHECKED)
    GUICtrlSetState($CheckExcludeNumber, $GUI_UNCHECKED)
    GUICtrlSetState($CheckExcludeName, $GUI_UNCHECKED)
    GUICtrlSetState($CheckLoop, $GUI_UNCHECKED)
    GUICtrlSetData($InputFilterNumber, "")
    GUICtrlSetData($InputFilterSource, "")
    If $Type = "log" Or $Type = 'all' Then
        GUICtrlSetState($CheckSystem, $GUI_CHECKED)
        GUICtrlSetState($CheckSecurity, $GUI_CHECKED)
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
    _Debug("SaveProject")
    $Project_filename = FileSaveDialog("Save project file", @ScriptDir & "\AUXFiles\", _
            $ProgramName & " projects (E*.prj)|All projects (*.prj)|All files (*.*)", 18, @ScriptDir & "\AUXFiles\" & $ProgramName & ".prj")

    Local $file = FileOpen($Project_filename, 2)
    ; Check if file opened for writing OK
    If $file = -1 Then
        _Debug("SaveProject: Unable to open file for writing: " & $Project_filename, 0x10, 5)
        Return
    EndIf
    _Debug("SaveProject  " & $Project_filename)
    FileWriteLine($file, "Valid for EventMon project")
    FileWriteLine($file, "Project file for " & @ScriptName & "  " & _DateTimeFormat(_NowCalc(), 0))
    FileWriteLine($file, "Help 1 is enabled, 4 is disabled for checkboxes")
    FileWriteLine($file, "CheckSave:" & GUICtrlRead($CheckSave))
    FileWriteLine($file, "CheckSplash:" & GUICtrlRead($CheckSave))
    FileWriteLine($file, "CheckSystem:" & GUICtrlRead($CheckSystem))
    FileWriteLine($file, "CheckSecurity:" & GUICtrlRead($CheckSecurity))
    FileWriteLine($file, "CheckApplication:" & GUICtrlRead($CheckApplication))
    FileWriteLine($file, "CheckInformation:" & GUICtrlRead($CheckInformation))
    FileWriteLine($file, "CheckWarnings:" & GUICtrlRead($CheckWarnings))
    FileWriteLine($file, "CheckErrors:" & GUICtrlRead($CheckErrors))
    FileWriteLine($file, "CheckSuccessAudits:" & GUICtrlRead($CheckSuccessAudits))
    FileWriteLine($file, "CheckFailAudits:" & GUICtrlRead($CheckFailAudits))
    FileWriteLine($file, "CheckOther:" & GUICtrlRead($CheckOther))
    FileWriteLine($file, "CheckExclude:" & GUICtrlRead($CheckExcludeNumber))
    FileWriteLine($file, "CheckExclude:" & GUICtrlRead($CheckExcludeName))
    FileWriteLine($file, "LabelSliderDelay:" & GUICtrlRead($LabelSliderDelay))
    FileWriteLine($file, "SliderDelay:" & GUICtrlRead($SliderDelay))
    FileWriteLine($file, "InputFilterNumber:" & GUICtrlRead($InputFilterNumber))
    FileWriteLine($file, "InputFilterSource:" & GUICtrlRead($InputFilterSource))
    Local $F = GetWinPos($ProgramName)
    FileWriteLine($file, $ProgramName & "_pos:" & $F[0] & " " & $F[1] & " " & $F[2] & " " & $F[3])

    FileClose($file)
EndFunc   ;==>SaveProject
;-----------------------------------------------
;This loads the project file but into the tree control but not into the list
Func LoadProject($Type)
    _Debug("LoadProject  " & $Type)

    If StringCompare($Type, "menu") = 0 Then
        $Project_filename = FileOpenDialog("Load project file", @ScriptDir & "\AUXFiles\", _
                $ProgramName & " projects (E*.prj)|All projects (*.prj)|All files (*.*)", 18, @ScriptDir & "\AUXFiles\" & $ProgramName & ".prj")
    EndIf

    Local $file = FileOpen($Project_filename, 0)
    ; Check if file opened for reading OK
    If $file = -1 Then
        _Debug("LoadProject: Unable to open file for reading: " & $Project_filename, 0x10, 5)
        Return
    EndIf

    _Debug("LoadProject   " & $Project_filename)
    ; Read in the first line to verify the file is of the correct type
    If StringCompare(FileReadLine($file, 1), "Valid for EventMon project") <> 0 Then
        MsgBox(16, "Invalid project file", "Not a valid EventMon project file")
        FileClose($file)
        Return
    EndIf

    ; Read in lines of text until the EOF is reached
    While 1
        Local $LineIn = FileReadLine($file)
        If @error = -1 Then ExitLoop

        _Debug("LoadProject   " & $LineIn)
        If StringInStr($LineIn, ";") = 1 Then ContinueLoop

        Local $F
        If StringInStr($LineIn, $ProgramName & "_pos:") Then
            _Debug("MainWinpos: " & $LineIn)
            $F = StringMid($LineIn, StringInStr($LineIn, ":") + 1)
            $F = StringSplit($F, " ", 2)
            _Debug(ConsoleWrite(@ScriptLineNumber & ": 2 " & _ArrayToString($F) & @CRLF))
            WinMove($ProgramName, "", $F[0], $F[1], $F[2], $F[3])
        EndIf

        If StringInStr($LineIn, "CheckSave:") Then GUICtrlSetState($CheckSave, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckSplash:") Then GUICtrlSetState($CheckSplash, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
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
        If StringInStr($LineIn, "CheckExcludeNumber:") Then GUICtrlSetState($CheckExcludeNumber, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckExcludeName:") Then GUICtrlSetState($CheckExcludeName, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckLoop:") Then GUICtrlSetState($CheckLoop, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "LabelSliderDelay:") Then GUICtrlSetData($LabelSliderDelay, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "SliderDelay:") Then GUICtrlSetData($SliderDelay, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "InputFilterNumber:") Then GUICtrlSetData($InputFilterNumber, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "InputFilterSource:") Then GUICtrlSetData($InputFilterSource, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
    WEnd

    FileClose($file)

    ; If the main window is not visible, make it visible
    $F = GetWinPos($ProgramName)
    _debug(@ScriptLineNumber & " " & $F[0] & " " & @DesktopWidth)
    _debug(@ScriptLineNumber & " " & $F[1] & " " & @DesktopHeight & @CRLF)
    If $F[0] > @DesktopWidth Or $F[1] > @DesktopHeight Then WinMove($ProgramName, "", 10, 10, 630, 515)
    If $F[0] < 0 Or $F[1] < 0 Then WinMove($ProgramName, "", 10, 10, 630, 515)

EndFunc   ;==>LoadProject
;-----------------------------------------------
Func Toggle($Type = 'all')
    If $ToggleState = $GUI_CHECKED Then
        $ToggleState = $GUI_UNCHECKED
    Else
        $ToggleState = $GUI_CHECKED
    EndIf
    ;GUICtrlSetState($CheckSave, $ToggleState)
    ;GUICtrlSetState($CheckSplash, $ToggleState)
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
    Local $D = GetWinPos($FormID)
    Local $WinPos
    If IsArray($D) = True Then
        ConsoleWrite(@ScriptLineNumber & $FormID & @CRLF)
        $WinPos = StringFormat("%s" & @CRLF & "WinPOS: %d  %d " & @CRLF & "WinSize: %d %d " & @CRLF & "Desktop: %d %d ", _
                $FormID, $D[0], $D[1], $D[2], $D[3], @DesktopWidth, @DesktopHeight)
    Else
        $WinPos = ">>>About ERROR, Check the window name<<<"
    EndIf
    MsgBox(64, "About", $SystemS & @CRLF & $WinPos & @CRLF & "Written by Doug Kaynor!")
EndFunc   ;==>About
;-----------------------------------------------
Func Help()
    Local $helpstr = "EventLog  " & $FileVersion & @CRLF & _
            "Startup options: " & @CRLF & _
            "help or ?   Display this help file" & @CRLF & _
            "clear       Clear all logs." & @CRLF
    MsgBox(64, "Help", @CRLF & @CRLF & $helpstr)
EndFunc   ;==>Help
;-----------------------------------------------
Func GetWinPos($WinName)
    Local $F
    Local $G
    While Not IsArray($F)
        $F = WinGetPos($WinName)
        Sleep(100)
        $G = $G + 1
        If $G > 100 Then ExitLoop
        _Debug(ConsoleWrite("Filewrite error: " & $G & "  " & _ArrayToString($F) & @CRLF))
    WEnd
    Return ($F)
EndFunc   ;==>GetWinPos
;-----------------------------------------------

