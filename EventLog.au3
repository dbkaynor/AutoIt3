#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_icon=../icons/log.ico
#AutoIt3Wrapper_outfile=EventLog.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Comment=A EventLog tool
#AutoIt3Wrapper_Res_Description=EventLog tool
#AutoIt3Wrapper_Res_Fileversion=1.0.0.12
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
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****
;#AutoIt3Wrapper_Run_Obfuscator=n
;#AutoIt3Wrapper_Run_cvsWrapper=v
;#Obfuscator_Parameters=/striponly
;#AutoIt3Wrapper_Res_Field=Credits|
;#AutoIt3Wrapper_Run_After=copy "%out%" "..\..\Programs_Updates\EventLog.exe"
;#NoTrayIcon
;#Tidy_Parameters=/gd /sf
;#Tidy_Parameters=/nsdp /sf
;#Tidy_Parameters=/sci=9
#cs
    This area is used to store things todo, bugs, and other notes
    
    Fixed:
    Add scroll to $ListResults
    Tool tips
    
    Todo:
    Functionality with other OS (Winxp)
    Add help (F1)
#CE

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
#include <GUIListBox.au3>
#include <Misc.au3>
#include <String.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <_DougFunctions.au3>

DirCreate("AUXFiles")
Global $AuxPath = @ScriptDir & "\AUXFiles\"
;FileInstall($AuxPath & "Working.jpg", $AuxPath & "Working.jpg", 0)

Global Const $tmp = StringSplit(@ScriptName, ".")
Global Const $ProgramName = $tmp[1]
Global $hEventLog

If _Singleton($ProgramName, 1) = 0 Then
    MsgBox(48, "Already running", $ProgramName & " is already running!", 10)
    Exit
EndIf

Global Const $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")
Global $SystemS = $ProgramName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & @OSServicePack & @CRLF & @OSType & @CRLF & @OSArch

; Mainform
Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $MainForm = GUICreate("Event log", 400, 400, 10, 10, $MainFormOptions)

Global $ButtonGetData = GUICtrlCreateButton("Get data", 10, 10)
GUICtrlSetTip(-1, "Fetch the data from application, security, and system logs. Check 'Statistics' for all data.")
GUICtrlSetResizing(-1, 802)
Global $ButtonClearLogs = GUICtrlCreateButton("Clear logs", 60, 10)
GUICtrlSetTip(-1, "Clear all logs. Check 'Save' if you want to save before clearing.")
GUICtrlSetResizing(-1, 802)
Global $ButtonEventViewer = GUICtrlCreateButton("Event viewer", 120, 10)
GUICtrlSetTip(-1, "Open the system event viewer (eventvwr.msc).")
GUICtrlSetResizing(-1, 802)
Global $ButtonOpenSaveFolder = GUICtrlCreateButton("Open save", 200, 10)
GUICtrlSetTip(-1, "Open the log save location.")
GUICtrlSetResizing(-1, 802)
Global $ButtonAbout = GUICtrlCreateButton("About", 280, 10)
GUICtrlSetTip(-1, "Display application information.")
GUICtrlSetResizing(-1, 802)
Global $ButtonHelp = GUICtrlCreateButton("Help", 320, 10)
GUICtrlSetTip(-1, "Display application help.")
GUICtrlSetResizing(-1, 802)
Global $ButtonExit = GUICtrlCreateButton("Exit", 360, 10)
GUICtrlSetTip(-1, "Exit the application.")
GUICtrlSetResizing(-1, 802)
Global $ListResults = GUICtrlCreateList("", 10, 70, 380, 300, BitOR($LBS_NOTIFY, $WS_BORDER, $WS_TABSTOP, $WS_GROUP, $LBS_DISABLENOSCROLL, $WS_VSCROLL, $WS_HSCROLL))
GUICtrlSetTip(-1, "This is the list of results gathered.")
GUICtrlSetResizing(-1, 802 - 64)
Global $CheckStats = GUICtrlCreateCheckbox("Statistics", 10, 40)
GUICtrlSetTip(-1, "Show detailed log statistics (may be slow)")
GUICtrlSetResizing(-1, 802)
Global $CheckSave = GUICtrlCreateCheckbox("Save", 100, 40)
GUICtrlSetTip(-1, "Save logs before clearing them.")
GUICtrlSetResizing(-1, 802)

GUICtrlSetState($CheckStats, True)
GUICtrlSetState($CheckSave, False)
GUISetState()

HotKeySet("{F1}", "Help")
HotKeySet("{F11}", "GUI_Enable")


For $x = 1 To $CmdLine[0]
    ConsoleWrite($x & " >> " & $CmdLine[$x] & @CRLF)
    Select
        Case StringInStr($CmdLine[$x], "help") > 0 Or StringInStr($CmdLine[$x], "?") > 0
            Help()
            Exit
        Case StringInStr($CmdLine[$x], "clear") > 0
            ClearLogs()
            Exit
        Case StringInStr($CmdLine[$x], "about") > 0
            About("Event log")
            Exit
        Case Else
            MsgBox(48, @ScriptLineNumber & " Unknown cmdline option found:", ">>" & $CmdLine[$x] & "<<")
            Exit
    EndSelect
Next


While 1
    Global $t = GUIGetMsg()
    Switch $t
        Case $ButtonGetData
            _GUICtrlListBox_ResetContent($ListResults)
            GetData()
        Case $ButtonClearLogs
            ClearLogs()
        Case $ButtonEventViewer
            ShellExecuteWait("eventvwr.msc", "/s")
        Case $ButtonOpenSaveFolder
            GuiDisable($GUI_DISABLE)
            ClipPut(EnvGet("temp"))
            RunWait("C:\WINDOWS\EXPLORER.EXE /n,/e," & EnvGet("temp"))
            GuiDisable($GUI_ENABLE)
        Case $ButtonAbout
            About("Event log")
        Case $ButtonHelp
            Help()
        Case $ButtonExit
            Exit
        Case $GUI_EVENT_CLOSE
            Exit
    EndSwitch
WEnd
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
    _GUICtrlListBox_ResetContent($ListResults)
EndFunc   ;==>ClearLogs
;-----------------------------------------------
Func GetData()
    If GUICtrlRead($CheckStats) = $GUI_CHECKED Then SplashImageOn("Working", $AuxPath & "Working.jpg", -1, -1, -1, -1, 18)
    _GUICtrlListBox_AddString($ListResults, "----Application log----")
    $hEventLog = _EventLog__Open("", "Application")
    ConsoleWrite(@ScriptLineNumber & " " & $hEventLog & @CRLF)
    GetInformation($hEventLog)
    _EventLog__Close($hEventLog)
    _GUICtrlListBox_AddString($ListResults, "  ")

    _GUICtrlListBox_AddString($ListResults, "----Security log----")
    $hEventLog = _EventLog__Open("", "Security")
    ConsoleWrite(@ScriptLineNumber & " " & $hEventLog & @CRLF)
    GetInformation($hEventLog)
    _EventLog__Close($hEventLog)
    _GUICtrlListBox_AddString($ListResults, "  ")

    _GUICtrlListBox_AddString($ListResults, "----System log----")
    $hEventLog = _EventLog__Open("", "System")
    ConsoleWrite(@ScriptLineNumber & " " & $hEventLog & @CRLF)
    GetInformation($hEventLog)
    _EventLog__Close($hEventLog)
    _GUICtrlListBox_AddString($ListResults, "  ")
    SplashOff()
EndFunc   ;==>GetData

;-----------------------------------------------
Func GetInformation($hEventLog)
    GuiDisable($GUI_DISABLE)
    _GUICtrlListBox_AddString($ListResults, "Log full: " & _EventLog__Full($hEventLog))
    _GUICtrlListBox_AddString($ListResults, "Log record count: " & _EventLog__Count($hEventLog))

    GetStats($hEventLog)
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>GetInformation
;-----------------------------------------------
Func GetStats($hEventLog)
    If GUICtrlRead($CheckStats) = $GUI_CHECKED Then
        Local $Errors = 0
        Local $Warnings = 0
        Local $Information = 0
        Local $SuccessAudit = 0
        Local $FailureAudit = 0
        Local $Other = 0
        For $x = 1 To _EventLog__Count($hEventLog)
            Local $a = _EventLog__Read($hEventLog)
            Switch $a[7]
                Case 1
                    $Errors += 1
                Case 2
                    $Warnings += 1
                Case 4
                    $Information += 1
                Case 8
                    $SuccessAudit += 1
                Case 16
                    $FailureAudit += 1
                Case Else
                    $Other += 1
                    ConsoleWrite(@ScriptLineNumber & " Other:>" & $a[7] & "<" & @CRLF)
            EndSwitch

            ;ConsoleWrite(@ScriptLineNumber & " "& $a[7] & "  " & $a[8] & @CRLF)
        Next
        _GUICtrlListBox_AddString($ListResults, "Error: " & $Errors)
        _GUICtrlListBox_AddString($ListResults, "Warning: " & $Warnings)
        _GUICtrlListBox_AddString($ListResults, "Information: " & $Information)
        _GUICtrlListBox_AddString($ListResults, "Success Audit: " & $SuccessAudit)
        _GUICtrlListBox_AddString($ListResults, "Failure Audit: " & $FailureAudit)
        _GUICtrlListBox_AddString($ListResults, "Other: " & $Other)

    EndIf

EndFunc   ;==>GetStats

;-----------------------------------------------
Func GuiDisable($choice) ;$GUI_ENABLE $GUI_DISABLE
    For $x = 1 To 200
        GUICtrlSetState($x, $choice)
    Next
EndFunc   ;==>GuiDisable
;-----------------------------------------------
Func About(Const $FormID)
    GuiDisable($GUI_DISABLE)
    Local $d = WinGetPos($FormID)
    Local $WinPos
    If IsArray($d) = True Then
        ConsoleWrite(@ScriptLineNumber & $FormID & @CRLF)
        $WinPos = StringFormat("%s" & @CRLF & "WinPOS: %d  %d " & @CRLF & "WinSize: %d %d " & @CRLF & "Desktop: %d %d ", _
                $FormID, $d[0], $d[1], $d[2], $d[3], @DesktopWidth, @DesktopHeight)
    Else
        $WinPos = ">>>About ERROR, Check the window name<<<"
    EndIf
    MsgBox(64, "About", $SystemS & @CRLF & $WinPos & @CRLF & "Written by Doug Kaynor!")
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>About
;-----------------------------------------------
; hotkey F11
Func GUI_Enable()
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>GUI_Enable
;-----------------------------------------------
Func Help()
    GuiDisable($GUI_DISABLE)
    Local $helpstr = "EventLog  " & $FileVersion & @CRLF & _
            "Startup options: " & @CRLF & _
            "help or ?   Display this help file" & @CRLF & _
            "clear       Clear all logs." & @CRLF & _
            "Hot Keys: " & @CRLF & _
            "F1  = Help" & @CRLF & _
            "F11 = Unlock GUI"
    MsgBox(64, "Help", @CRLF & @CRLF & $helpstr)
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>Help