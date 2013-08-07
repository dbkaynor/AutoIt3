#AutoIt3Wrapper_icon=../icons/openbsd.ico
#AutoIt3Wrapper_outfile=EventMon.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Comment=An EventLog monitoring tool
#AutoIt3Wrapper_Res_Description=An EventLog monitoring tool
#AutoIt3Wrapper_Res_Fileversion=1.0.1.3
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=Y
#AutoIt3Wrapper_Res_ProductVersion=666
#AutoIt3Wrapper_Res_LegalCopyright=Copyright 2012 Douglas B Kaynor
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
#include <GuiFontAndColors.au3>
#include <GuiListView.au3>
#include <ListViewConstants.au3>
#include <Misc.au3>
#include <SliderConstants.au3>
#include <StaticConstants.au3>
#include <String.au3>
#include <WindowsConstants.au3>
#include <_DougFunctions.au3>

;FileInstall($AuxPath & "Working.jpg", $AuxPath & "Working.jpg", 0)

;Global Const $tmp = StringSplit(@ScriptName, ".")
Global Const $ProgramName = "EventMon"
Global $Project_filename = @ScriptDir & "\AUXFiles\" & $ProgramName & ".prj"
Global $LOG_filename = @ScriptDir & "\AUXFiles\" & $ProgramName & ".log"
Global $LastSearchFound = -1
Global $hEventLog

If _Singleton($ProgramName, 1) = 0 Then
    MsgBox(48, "Already running", $ProgramName & " is already running!", 10)
    Exit
EndIf

Global Const $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")
Global $SystemS = $ProgramName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & @OSServicePack & @CRLF & @OSType & @CRLF & @OSArch

Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $MainForm = GUICreate($ProgramName & "  " & $FileVersion, 715, 580, 10, 10, $MainFormOptions)
;GUISetFont(9, 400, -1, "courier")

Global $ButtonGetData = GUICtrlCreateButton("Get data", 20, 10, 75, 25)
GUICtrlSetTip($ButtonGetData, "Fetch the data from selected logs")
Global $ButtonEventViewer = GUICtrlCreateButton("Event viewer", 100, 10, 75, 25)
GUICtrlSetTip($ButtonEventViewer, "Open the event control viewer")
Global $ButtonClearLogs = GUICtrlCreateButton("Clear logs", 180, 10, 75, 25)
GUICtrlSetTip($ButtonClearLogs, "Clear all logs. Check 'Save' if you want to save before clearing")
Global $ButtonOpenSaveFolder = GUICtrlCreateButton("Open save", 260, 10, 75, 25)
GUICtrlSetTip($ButtonOpenSaveFolder, "Open the log save location.")
Global $ButtonAbout = GUICtrlCreateButton("About", 340, 10, 75, 25)
GUICtrlSetTip($ButtonAbout, "Display application information.")
Global $ButtonHelp = GUICtrlCreateButton("Help", 420, 10, 75, 25)
GUICtrlSetTip($ButtonHelp, "Display application help")
Global $ButtonExit = GUICtrlCreateButton("Exit", 500, 10, 75, 25)
GUICtrlSetTip($ButtonExit, "Exit this application")

;left, top , width , height
Global $GroupOptions = GUICtrlCreateGroup("Options", 15, 40, 600, 45)
Global $CheckSave = GUICtrlCreateCheckbox("Save", 25, 55, 60, 25)
GUICtrlSetTip($CheckSave, "Save logs before clearing")
Global $CheckSplash = GUICtrlCreateCheckbox("Splash", 100, 55, 60, 25)
GUICtrlSetTip($CheckSplash, "Display splash")
Global $CheckAbort = GUICtrlCreateCheckbox("Abort", 190, 55, 70, 25)
GUICtrlSetTip($CheckAbort, "Abort operations")
Global $CheckJunkFilter = GUICtrlCreateCheckbox("Junk filter", 262, 55, 70, 25)
GUICtrlSetTip($CheckJunkFilter, "Junk filter for description")
Global $ButtonSaveProject = GUICtrlCreateButton("Save project", 350, 50, 75, 25)
GUICtrlSetTip($ButtonSaveProject, "Save the current settings")
Global $ButtonLoadProject = GUICtrlCreateButton("Load project", 435, 50, 75, 25)
GUICtrlSetTip($ButtonLoadProject, "Load saved settings")
Global $ButtonLoadDefaults = GUICtrlCreateButton("Load defaults", 520, 50, 75, 25)
GUICtrlSetTip($ButtonLoadDefaults, "Load default settings")
GUICtrlCreateGroup("", -99, -99, 1, 1)

; text, left, top, width, height
Global $GroupLogType = GUICtrlCreateGroup("Log type", 15, 90, 400, 45)
Global $CheckSystem = GUICtrlCreateCheckbox("System", 25, 105, 70, 25)
GUICtrlSetTip($CheckSystem, "Process system logs")
Global $CheckSecurity = GUICtrlCreateCheckbox("Security", 95, 105, 70, 25)
GUICtrlSetTip($CheckSecurity, "Process security logs")
Global $CheckApplication = GUICtrlCreateCheckbox("Application", 165, 105, 70, 25)
GUICtrlSetTip($CheckApplication, "Process application logs")
Global $ButtonToggleLogType = GUICtrlCreateButton("Toggle", 240, 105, 75, 25)
GUICtrlSetTip($ButtonToggleLogType, "Toggle log type selections")
Global $ButtonDefaultsLogType = GUICtrlCreateButton("Defaults", 330, 105, 75, 25)
GUICtrlSetTip($ButtonDefaultsLogType, "Set log type selections to defaults")
GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateGroup("Loop", 425, 90, 260, 45)
Global $CheckLoop = GUICtrlCreateCheckbox("Loop", 435, 105, 50, 25)
GUICtrlSetTip($CheckLoop, "Start/Stop looping")
Global $SliderDelay = GUICtrlCreateSlider(490, 105, 150, 25)
GUICtrlSetLimit($SliderDelay, 60, 1)
GUICtrlSetData($SliderDelay, 1)
GUICtrlSetTip($SliderDelay, "Loop delay setting")
Global $LabelSliderDelay = GUICtrlCreateLabel("", 650, 105, 20, 25, $SS_SUNKEN)
GUICtrlSetTip($LabelSliderDelay, "Current loop delay setting")
GUICtrlCreateGroup("", -99, -99, 1, 1)
Global $GroupType = GUICtrlCreateGroup("Message type", 15, 140, 650, 45)
Global $CheckInformation = GUICtrlCreateCheckbox("Information", 25, 155, 70, 25)
GUICtrlSetTip($CheckInformation, "Show information events")
Global $CheckWarnings = GUICtrlCreateCheckbox("Warnings", 105, 155, 70, 25)
GUICtrlSetTip($CheckWarnings, "Show warning events")
Global $CheckErrors = GUICtrlCreateCheckbox("Errors", 190, 155, 50, 25)
GUICtrlSetTip($CheckErrors, "Show error events")
Global $CheckSuccessAudits = GUICtrlCreateCheckbox("Success Audits", 260, 155, 97, 25)
GUICtrlSetTip($CheckSuccessAudits, "Show successful audit events")
Global $CheckFailAudits = GUICtrlCreateCheckbox("Fail Audits", 370, 155, 81, 25)
GUICtrlSetTip($CheckFailAudits, "Show failed audit events")
Global $CheckOther = GUICtrlCreateCheckbox("Other", 465, 155, 50, 25)
GUICtrlSetTip($CheckOther, "Other")
Global $ButtonToggleEventType = GUICtrlCreateButton("Toggle", 525, 155, 59, 25)
GUICtrlSetTip($ButtonToggleEventType, "Toggle event type selections")
Global $ButtonDefaultsEventType = GUICtrlCreateButton("Defaults", 600, 155, 59, 25)
GUICtrlSetTip($ButtonDefaultsEventType, "Set log type selections to defaults")
GUICtrlCreateGroup("", -99, -99, 1, 1)

;"text", left, top [, width [, height
GUICtrlCreateGroup("Filter/Find data", 15, 190, 630, 45)
Global $InputFilter = GUICtrlCreateInput("", 25, 205, 130, 25)
GUICtrlSetTip(-1, "Input a filter string. (Supports regular expressions)")
Global $ButtonFind = GUICtrlCreateButton("Find", 160, 205, 30, 25)
Global $CheckFilterExclude = GUICtrlCreateCheckbox("Exclude", 200, 205, 60, 25)
GUICtrlSetTip(-1, "Exclude or include the filter string")
Global $RadioFilterAll = GUICtrlCreateRadio("All", 270, 205, 60, 25)
Global $RadioFilterNumber = GUICtrlCreateRadio("Number", 315, 205, 60, 25)
Global $RadioFilterType = GUICtrlCreateRadio("Type", 380, 205, 60, 25)
Global $RadioFilterDate = GUICtrlCreateRadio("Date", 440, 205, 60, 25)
Global $RadioFilterSource = GUICtrlCreateRadio("Source", 500, 205, 60, 25)
Global $RadioFilterDescription = GUICtrlCreateRadio("Description", 560, 205, 80, 25)
GUICtrlSetState($RadioFilterAll, $GUI_CHECKED)

GUICtrlSetTip(-1, "Select filter type")
GUICtrlCreateGroup("", -99, -99, 1, 1)

Global $LabelStatus = GUICtrlCreateLabel("Status main", 20, 250, 670, 25, $SS_SUNKEN)
GUICtrlSetTip($LabelStatus, "Status message")

;$ListView constants
Const $Number = 0
Const $Type = 1
Const $Date = 2
Const $Source = 3
Const $Description = 4
Global $ListView = GUICtrlCreateListView("Number|Event type|Date logged|Event source|Event description", 20, 290, 670, 280, -1, BitOR($WS_EX_CLIENTEDGE, $LVS_EX_GRIDLINES, $LVS_EX_FULLROWSELECT))
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $Number, 70)
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $Type, 80)
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $Date, 130)
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $Source, 200)
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $Description, 200)
GUICtrlSetTip($ListView, "This is the list box")

GUISetState(@SW_SHOW)
#endregion ### END Koda GUI section ###

For $x = 0 To 100 ; This sets the resize mode for all GUI items (assumingg 100 items)
    GUICtrlSetResizing($x, $GUI_DOCKALL)
Next

;Change the resize mode to selected items
GUICtrlSetResizing($ListView, $GUI_DOCKTOP + $GUI_DOCKBOTTOM + $GUI_DOCKRIGHT + $GUI_DOCKLEFT)
GUICtrlSetResizing($LabelStatus, $GUI_DOCKHEIGHT + $GUI_DOCKTOP + $GUI_DOCKRIGHT + $GUI_DOCKLEFT)

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
            _About($ProgramName, $SystemS)
        Case $ButtonHelp
            ShellExecute($AuxPath & 'EventMon.htm')
        Case $ButtonExit
            _Debug(@ScriptLineNumber & '  ButtonExit')
            Exit
        Case $ButtonFind
            FindDataInList()

        Case $InputFilter
            $LastSearchFound = -1
        Case $CheckAbort
            GuiDisable($GUI_ENABLE)
            ConsoleWrite(@ScriptLineNumber & ' CheckAbort ' & @CRLF)
        Case $GUI_EVENT_CLOSE
            _Debug(@ScriptLineNumber & '  GUI_EVENT_CLOSE')
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
    GuiDisable($GUI_DISABLE)
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
    GuiDisable($GUI_ENABLE)
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

        If StringLen($TmpString) > 0 Then ;There is a tmpstring
            $E = StringSplit($TmpString, "~", 1)
            If StringLen(GUICtrlRead($InputFilter)) = 0 Then ; There are no filters
                DisplayResults($E, $Counter)
            Else ; Here we have a filter
                Local $F = StringSplit($E[2], ":", 2)
                ;ConsoleWrite(@ScriptLineNumber & ': ' & _ArrayToString($E) & @CRLF)
                If UBound($F) > 1 Then
                    Local $Result
                    Local $Tmp
                    ;dbk

                    ; ConsoleWrite(@ScriptLineNumber & ': ' & _ArrayToString($E) & @CRLF)
                    ; ConsoleWrite(@ScriptLineNumber & ': ' & _ArrayToString($F) & @CRLF)
                    If GUICtrlRead($RadioFilterAll) = $GUI_CHECKED Then $Tmp = $TmpString
                    If GUICtrlRead($RadioFilterNumber) = $GUI_CHECKED Then $Tmp = $E[1]
                    If GUICtrlRead($RadioFilterType) = $GUI_CHECKED Then $Tmp = $E[2]
                    If GUICtrlRead($RadioFilterDate) = $GUI_CHECKED Then $Tmp = $E[3]
                    If GUICtrlRead($RadioFilterSource) = $GUI_CHECKED Then $Tmp = $E[4]
                    If GUICtrlRead($RadioFilterDescription) = $GUI_CHECKED Then $Tmp = $E[5]
                    $Result = StringRegExp(StringUpper($Tmp), StringUpper(GUICtrlRead($InputFilter)), 0) ; EVENT TYPE filter. Returns 1 (matched) or 0 (no match)

                    If GUICtrlRead($CheckFilterExclude) = $GUI_UNCHECKED And $Result = 1 Then DisplayResults($E, $Counter)
                    If GUICtrlRead($CheckFilterExclude) = $GUI_CHECKED And $Result = 0 Then DisplayResults($E, $Counter)

                EndIf
            EndIf
        EndIf
    Next

    GUICtrlSetData($LabelStatus, StringFormat("Errors:%s  Warnings:%s  Information:%s  SuccessAudit:%s  FailureAudit:%s  Other:%s   Else:%s" & @CRLF, $Errors, $Warnings, $Information, $SuccessAudit, $FailureAudit, $Other, $Else))

EndFunc   ;==>GetStats
;-----------------------------------------------
Func FindDataInList()
    Local $t = GUICtrlRead($InputFilter)
    Local $iI = _GUICtrlListView_FindInText($ListView, $t, $LastSearchFound, True, False)
    $LastSearchFound = $iI
    ConsoleWrite(@ScriptLineNumber & ' ' & $t & '   ' & $iI & @CRLF)
    _GUICtrlListView_EnsureVisible($ListView, $iI)
    _GUICtrlListView_SetItemSelected($ListView, $iI)
EndFunc   ;==>FindDataInList
;-----------------------------------------------
Func DisplayResults($E, $Counter)
    Local $ListView_item
    If GUICtrlRead($CheckJunkFilter) = $GUI_CHECKED Then $E[5] = StringRegExpReplace($E[5], '[^a-z A-Z 0-9]', '')

    ;If StringInStr($E[5], 'local') > 0 Then ConsoleWrite(@ScriptLineNumber & " " & $E[5] & @CRLF)
    ;StringRegExpReplace($TmpString, '[^:print:]', '')
    If $E[0] >= 1 Then $ListView_item = _GUICtrlListView_AddItem($ListView, StringFormat("%3s  %6s", $Counter, $E[1]))
    If $E[0] >= 2 Then _GUICtrlListView_AddSubItem($ListView, $ListView_item, $E[2], $Type)
    If $E[0] >= 3 Then _GUICtrlListView_AddSubItem($ListView, $ListView_item, $E[3], $Date)
    If $E[0] >= 4 Then _GUICtrlListView_AddSubItem($ListView, $ListView_item, $E[4], $Source)
    If $E[0] >= 5 Then _GUICtrlListView_AddSubItem($ListView, $ListView_item, $E[5], $Description)
EndFunc   ;==>DisplayResults
;-----------------------------------------------
Func SetDefaults($Type = 'all')
    _Debug(@ScriptLineNumber & '  ' & $ProgramName & '  ' & $Type)
    _CheckWindowLocation($MainForm, 'center')
    GUICtrlSetData($SliderDelay, 2)
    GUICtrlSetData($LabelSliderDelay, GUICtrlRead($SliderDelay))
    GUICtrlSetState($CheckSave, $GUI_UNCHECKED)
    GUICtrlSetState($CheckSplash, $GUI_UNCHECKED)
    GUICtrlSetState($CheckAbort, $GUI_UNCHECKED)
    GUICtrlSetState($CheckJunkFilter, $GUI_CHECKED)
    GUICtrlSetState($CheckLoop, $GUI_UNCHECKED)

    GUICtrlSetData($InputFilter, "")
    GUICtrlSetState($CheckFilterExclude, $GUI_UNCHECKED)
    GUICtrlSetState($RadioFilterAll, $GUI_CHECKED)
    GUICtrlSetState($RadioFilterNumber, $GUI_UNCHECKED)
    GUICtrlSetState($RadioFilterType, $GUI_UNCHECKED)
    GUICtrlSetState($RadioFilterDate, $GUI_UNCHECKED)
    GUICtrlSetState($RadioFilterSource, $GUI_UNCHECKED)
    GUICtrlSetState($RadioFilterDescription, $GUI_UNCHECKED)

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
    _Debug(@ScriptLineNumber & ' SaveProject')
    $Project_filename = FileSaveDialog('Save project file', @ScriptDir & '\AUXFiles\', _
            $ProgramName & ' projects (E*.prj)|All projects (*.prj)|All files (*.*)', 18, @ScriptDir & '\AUXFiles\' & $ProgramName & '.prj')

    Local $file = FileOpen($Project_filename, 2)
    ; Check if file opened for writing OK
    If $file = -1 Then
        _Debug(@ScriptLineNumber & ' SaveProject: Unable to open file for writing: ' & $Project_filename, '', '', True)
        Return
    EndIf
    _Debug(@ScriptLineNumber & "  SaveProject  " & $Project_filename)
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

    FileWriteLine($file, "CheckJunkFilter:" & GUICtrlRead($CheckJunkFilter))
    FileWriteLine($file, "CheckLoop:" & GUICtrlRead($CheckLoop))
    FileWriteLine($file, "LabelSliderDelay:" & GUICtrlRead($LabelSliderDelay))
    FileWriteLine($file, "SliderDelay:" & GUICtrlRead($SliderDelay))

    FileWriteLine($file, "InputFilter:" & GUICtrlRead($InputFilter))
    FileWriteLine($file, "CheckFilterExclude" & GUICtrlRead($CheckFilterExclude))
    FileWriteLine($file, "RadioFilterAll:" & GUICtrlRead($RadioFilterAll))
    FileWriteLine($file, "RadioFilterNumber" & GUICtrlRead($RadioFilterNumber))
    FileWriteLine($file, "RadioFilterType" & GUICtrlRead($RadioFilterType))
    FileWriteLine($file, "$RadioFilterDate" & GUICtrlRead($RadioFilterDate))
    FileWriteLine($file, "$RadioFilterSource" & GUICtrlRead($RadioFilterSource))
    FileWriteLine($file, "$RadioFilterDescription" & GUICtrlRead($RadioFilterDescription))
    Local $F = GetWinPos($ProgramName)
    FileWriteLine($file, $ProgramName & "_pos:" & $F[0] & " " & $F[1] & " " & $F[2] & " " & $F[3])

    FileClose($file)
EndFunc   ;==>SaveProject
;-----------------------------------------------
;This loads the project file but into the tree control but not into the list
Func LoadProject($Type)
    _Debug(@ScriptLineNumber & ' LoadProject  ' & $Type)

    If StringCompare($Type, "menu") = 0 Then
        $Project_filename = FileOpenDialog("Load project file", @ScriptDir & "\AUXFiles\", _
                $ProgramName & " projects (E*.prj)|All projects (*.prj)|All files (*.*)", 18, @ScriptDir & "\AUXFiles\" & $ProgramName & ".prj")
    EndIf

    Local $file = FileOpen($Project_filename, 0)
    ; Check if file opened for reading OK
    If $file = -1 Then
        _Debug(@ScriptLineNumber & "  LoadProject: Unable to open file for reading: " & $Project_filename, '', '', True)
        Return
    EndIf

    _Debug(@ScriptLineNumber & '  LoadProject   ' & $Project_filename)
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

        _Debug(@ScriptLineNumber & '  LoadProject  ' & $LineIn)
        If StringInStr($LineIn, ";") = 1 Then ContinueLoop

        Local $F
        If StringInStr($LineIn, $ProgramName & "_pos:") Then
            _Debug(@ScriptLineNumber & '  MainWinpos: ' & $LineIn)
            $F = StringMid($LineIn, StringInStr($LineIn, ":") + 1)
            $F = StringSplit($F, " ", 2)
            ConsoleWrite(@ScriptLineNumber & ": 2 " & _ArrayToString($F) & @CRLF)
            WinMove($ProgramName, "", $F[0], $F[1], $F[2], $F[3])
        EndIf

        If StringInStr($LineIn, "CheckSave:") Then GUICtrlSetState($CheckSave, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckSplash:") Then GUICtrlSetState($CheckSplash, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckSystem:") Then GUICtrlSetState($CheckSystem, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckSecurity:") Then GUICtrlSetState($CheckSecurity, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckApplication:") Then GUICtrlSetState($CheckApplication, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckInformation:") Then GUICtrlSetState($CheckInformation, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckJunkFilter:") Then GUICtrlSetState($CheckJunkFilter, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckWarnings:") Then GUICtrlSetState($CheckWarnings, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckErrors:") Then GUICtrlSetState($CheckErrors, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckSuccessAudits:") Then GUICtrlSetState($CheckSuccessAudits, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckFailAudits:") Then GUICtrlSetState($CheckFailAudits, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckOther:") Then GUICtrlSetState($CheckOther, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckLoop:") Then GUICtrlSetState($CheckLoop, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "LabelSliderDelay:") Then GUICtrlSetData($LabelSliderDelay, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "SliderDelay:") Then GUICtrlSetData($SliderDelay, StringMid($LineIn, StringInStr($LineIn, ":") + 1))

        If StringInStr($LineIn, "InputFilter:") Then GUICtrlSetData($InputFilter, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckFilterExclude:") Then GUICtrlSetData($CheckFilterExclude, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "RadioFilterAll:") Then GUICtrlSetData($RadioFilterAll, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "RadioFilterNumber:") Then GUICtrlSetData($RadioFilterNumber, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "RadioFilterType:") Then GUICtrlSetData($RadioFilterType, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "RadioFilterDate:") Then GUICtrlSetData($RadioFilterDate, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "RadioFilterSource:") Then GUICtrlSetData($RadioFilterSource, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "RadioFilterDescription:") Then GUICtrlSetData($RadioFilterDescription, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        ; If the main window is not visible, make it visible
        _SetWindowPosition($MainForm, $ProgramName, $LineIn)
        _CheckWindowLocation($MainForm)
    WEnd

    FileClose($file)
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
Func GetWinPos($WinName)
    Local $F
    Local $G
    While Not IsArray($F)
        $F = WinGetPos($WinName)
        Sleep(100)
        $G = $G + 1
        If $G > 100 Then ExitLoop
        _Debug(@ScriptLineNumber & '  Filewrite error: ' & $G & '  ' & _ArrayToString($F) & @CRLF)
    WEnd
    Return ($F)
EndFunc   ;==>GetWinPos
;-----------------------------------------------
Func GuiDisable($choice) ;$GUI_ENABLE $GUI_DISABLE
    For $x = 1 To 200
        GUICtrlSetState($x, $choice)
    Next
    GUICtrlSetState($CheckAbort, $GUI_ENABLE)
EndFunc   ;==>GuiDisable
;-----------------------------------------------

