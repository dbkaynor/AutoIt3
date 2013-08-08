#region
#AutoIt3Wrapper_icon=../icons/Cryptkeeper.ico
#AutoIt3Wrapper_outfile=URW4WinW.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Comment=GUI wrapper for URW4Win
#AutoIt3Wrapper_Res_Description=Wrapper Set up URW4Win
#AutoIt3Wrapper_Res_Fileversion=1.0.0.16
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

#cs
    This area is used to store things todo, bugs, and other notes
    
    Fixed:
    
    Todo:
    pa
#CE

#include <array.au3>
#include <Constants.au3>
#include <file.au3>
#include <GUIConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiListView.au3>
#include <misc.au3>
#include <Process.au3>
#include <WindowsConstants.au3>
#include <StaticConstants.au3>

Opt("MustDeclareVars", 1)

Global $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")
Global $tmp = StringSplit(@ScriptName, ".")
Global Const $ProgramName = $tmp[1]

Global $logfileName = $ProgramName & ".log"
Global $datafileName = $ProgramName & ".dat"
Global $OIDsFolder = '\'
Global $WireSharkFolder = ''

Global $ProcessedLeaseArray[1]

If _Singleton($ProgramName, 1) = 0 Then
    MsgBox(32, "Already running", $ProgramName & " is already running!")
    Exit
EndIf

Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $MainForm = GUICreate($ProgramName & "  " & $FileVersion, 600, 260, 20, 20, $MainFormOptions)
GUISetFont(9, 400, 0, "Courier New")

Global $URW4Win = FileGetShortName(@ScriptDir & "\AUXFiles\URW4win.exe")
If Not FileExists($URW4Win) Then
    MsgBox(16, "Missing file", "URW4win.exe not found" & @CRLF & "must be located in the AUXFiles folder")
    Exit
EndIf


GUICtrlCreateLabel("Block size:", 10, 10, 100, 20)
GUICtrlSetTip(-1, "Enter block size (1024 .. 65536)")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $InputBlockSize = GUICtrlCreateInput("65536", 110, 10, 50, 20)
GUICtrlSetTip(-1, "Enter block size (1024 .. 65536)")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

GUICtrlCreateLabel("Delay (ms):", 10, 35, 100, 20)
GUICtrlSetTip(-1, "Enter delay value")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $InputDelay = GUICtrlCreateInput("0", 110, 35, 50, 20)
GUICtrlSetTip(-1, "Enter delay value")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

GUICtrlCreateLabel("Cycles:", 10, 60, 100, 20)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetTip(-1, "Enter cycle count")
Global $InputCycles = GUICtrlCreateInput("0", 110, 60, 50, 20)
GUICtrlSetTip(-1, "Enter cycle count")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

GUICtrlCreateLabel("Log interval:", 10, 85, 100, 20)
GUICtrlSetTip(-1, "Logging interval")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $InputLogInterval = GUICtrlCreateInput("0", 110, 85, 50, 20)
GUICtrlSetTip(-1, "Logging interval")
GUICtrlSetResizing(-1, $GUI_DOCKALL)


GUICtrlCreateLabel("Data size:", 170, 10, 100, 20)
GUICtrlSetTip(-1, "Enter data file size")
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $InputDataSize = GUICtrlCreateInput("1e9", 250, 10, 50, 20)
GUICtrlSetTip(-1, "Enter file data size")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $CheckDataCheck = GUICtrlCreateCheckbox("Check the data", 320, 10, 120, 20)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetTip(-1, "Do not do data check")

Global $CheckHammer = GUICtrlCreateCheckbox("Hammer it", 320, 30, 120, 20)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetTip(-1, "Hammer it")

Global $CheckOpenRW = GUICtrlCreateCheckbox("Open R/W", 320, 50, 120, 20)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetTip(-1, "open r/w file once")

Global $CheckStartPaused = GUICtrlCreateCheckbox("Start paused", 320, 70, 120, 20)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetTip(-1, "Start in paused state")

Global $CheckReadOnly = GUICtrlCreateCheckbox("Read only", 320, 90, 120, 20)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetTip(-1, "Read only")

Global $ButtonLogFile = GUICtrlCreateButton("Logfile name", 10, 110, 110, 20)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetTip(-1, "Input a log file name")
Global $InputLogFileName = GUICtrlCreateInput($logfileName, 120, 110, 100, 20)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetTip(-1, "Input a log file name")

Global $ButtonDataFile = GUICtrlCreateButton("Data file name", 230, 110, 110, 20)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetTip(-1, "Input a data file name")
Global $InputDataFileName = GUICtrlCreateInput($datafileName, 340, 110, 110, 20)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetTip(-1, "Input a data file name")



Global $ButtonStart = GUICtrlCreateButton("Start", 480, 10, 120, 20)
GUICtrlSetTip(-1, "Start the program")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ButtonResmon = GUICtrlCreateButton("Resource Monitor", 480, 30, 120, 20)
GUICtrlSetTip(-1, "Start resource monitor")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ButtonOIDS = GUICtrlCreateButton("OID's", 480, 50, 120, 20)
GUICtrlSetTip(-1, "Start OIDs")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ButtonEventViewer = GUICtrlCreateButton("Event viewer", 480, 70, 120, 20)
GUICtrlSetTip(-1, "Start Event viewer")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ButtonEventMonitor = GUICtrlCreateButton("Event monitor", 480, 90, 120, 20)
GUICtrlSetTip(-1, "Start Event monitor")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ButtonWireShark = GUICtrlCreateButton("WireShark", 480, 110, 120, 20)
GUICtrlSetTip(-1, "Start WireShark")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ButtonExit = GUICtrlCreateButton("Exit", 480, 130, 120, 20)
GUICtrlSetTip(-1, "Exit the program")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

GUICtrlCreateLabel("Status", 10, 135, 150, 25)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetTip(-1, "Status")
Global $LabelStatus = GUICtrlCreateLabel("Status", 10, 155, 570, 100, $SS_SUNKEN)
GUICtrlSetResizing(-1, $GUI_DOCKBORDERS)
GUICtrlSetTip(-1, "Status")

GUICtrlSetState($CheckDataCheck, $GUI_CHECKED)
GUICtrlSetState($CheckHammer, $GUI_CHECKED)
GUICtrlSetState($CheckStartPaused, $GUI_UNCHECKED)

GUISetState(@SW_SHOW)

Global $F = WinGetPos("URW4Win", "")
If $F[0] > @DesktopWidth Or $F[1] > @DesktopHeight Then WinMove("URW4Win", "", 10, 10, 630, 300)
If $F[0] < 0 Or $F[1] < 0 Then WinMove("URW4Win", "", 10, 10, 630, 300)

GUISetState()

;-----------------------------------------------
While 1
    Switch GUIGetMsg()
        Case $ButtonStart
            Doit()
        Case $ButtonLogFile
            ChoseLogFile()
        Case $ButtonResmon
            ShellExecute("resmon.exe")
        Case $ButtonOIDS
            LaunchOIDs()
        Case $ButtonEventViewer
            ShellExecuteWait("eventvwr.msc", "/s")
        Case $ButtonEventMonitor
            LaunchEventMonitor()
        Case $ButtonWireShark
            LaunchWireShark()
        Case $ButtonExit
            ExitLoop
        Case $GUI_EVENT_CLOSE
            Exit
    EndSwitch
WEnd
;-----------------------------------------------
Func LaunchEventMonitor()

    If Not FileExists(@ScriptDir & "\eventmon.exe") Then
        $OIDsFolder = FileSelectFolder("Select folder for eventmon.exe", "\", 2) & "\"
        If @error = 1 Then Return
    EndIf
    ;MsgBox(48, "eventmonitor.exe ", "eventmonitor.exe location" & @CRLF & $OIDsFolder)
    ShellExecuteWait("eventmon.exe", "")
EndFunc   ;==>LaunchEventMonitor
;-----------------------------------------------
Func LaunchOIDs()
    If Not FileExists($OIDsFolder & "oids.exe") Then
        $OIDsFolder = FileSelectFolder("Select folder for oids.exe", "\", 2) & "\"
        If @error = 1 Then Return
    EndIf
    ;MsgBox(48, "OIDs location", "OIDs location" & @CRLF & $OIDsFolder)
    ShellExecute($OIDsFolder & "oids.exe")
EndFunc   ;==>LaunchOIDs
;-----------------------------------------------
Func LaunchWireShark()
    $WireSharkFolder = 'c:\program files\wireshark\'
    If Not FileExists($WireSharkFolder & "wireshark.exe") Then
        $WireSharkFolder = 'c:\program files (x86)\wireshark\'
        If Not FileExists($WireSharkFolder & "wireshark.exe") Then
            $WireSharkFolder = FileSelectFolder("Select folder for wireshark.exe", "\", 2) & "\"
            If @error = 1 Then Return
        EndIf
    EndIf

    GUICtrlSetData($LabelStatus, $WireSharkFolder & "wireshark.exe")
    ShellExecute($WireSharkFolder & "wireshark.exe")
EndFunc   ;==>LaunchWireShark
;-----------------------------------------------
; verify that the various user input parameters make sense. Return 0 if all is well
Func VerifyParameters()
    Local $error = 0
    If GUICtrlRead($InputBlockSize) < 1024 Or GUICtrlRead($InputBlockSize) > 65536 Then $error = 1

    If $error = 0 Then Return $error

    Local $string
    Switch $error
        Case 1
            $string = "Block size problem"
        Case 2
            $string = "Something else"
    EndSwitch

    MsgBox(16, "Parameter error", "Parameter error" & @CRLF & $string)
    Return $error

EndFunc   ;==>VerifyParameters
;-----------------------------------------------
#cs
    
    -b <size> (block size   1024 .. 65536)
    -c no data Check
    -d <ms> (set the delay)
    -e <cycles> (number of cycles default is forever)
    -h (help)
    -k (hammer it, pace set to zero)
    -l <file>  log to file
    -o (open r/w file once)
    -p (Start in paused state)
    -r (read only)
    -t <sec>  (logging interval)
    -u (no interface)
    
#ce
;-----------------------------------------------
Func ChoseLogFile()
    Local $tmp = FileOpenDialog("Chose a log file", @ScriptDir, "*.log", 10, $ProgramName & ".log")
    If $tmp = "" Then Return

EndFunc   ;==>ChoseLogFile
;-----------------------------------------------
Func DoIt()
    If VerifyParameters() <> 0 Then Return
    Local $CommandString
    If GUICtrlRead($CheckDataCheck) = $GUI_UNCHECKED Then $CommandString = " -c "
    If GUICtrlRead($CheckHammer) = $GUI_CHECKED Then $CommandString = $CommandString & " -k "
    If GUICtrlRead($CheckOpenRW) = $GUI_CHECKED Then $CommandString = $CommandString & " -o "
    If GUICtrlRead($CheckReadOnly) = $GUI_CHECKED Then $CommandString = $CommandString & " -r "
    If GUICtrlRead($CheckStartPaused) = $GUI_CHECKED Then $CommandString = $CommandString & " -p "

    $CommandString = $CommandString & " -b " & GUICtrlRead($InputBlockSize)
    If GUICtrlRead($InputDelay) > 0 Then $CommandString = $CommandString & " -d " & GUICtrlRead($InputDelay)
    If GUICtrlRead($InputCycles) > 0 Then $CommandString = $CommandString & " -e " & GUICtrlRead($InputCycles)

    If Int(GUICtrlRead($InputLogInterval)) > 0 And StringLen(GUICtrlRead($InputLogFileName)) > 5 Then
        $CommandString = $CommandString & " -l '" & FileGetShortName(GUICtrlRead($InputLogFileName)) & "'"
        $CommandString = $CommandString & " -t " & Int(GUICtrlRead($InputLogInterval))
    EndIf

    ConsoleWrite(@ScriptLineNumber & " >>" & Execute(GUICtrlRead($InputDataSize)) & "<<  " & @CRLF)
    ConsoleWrite(@ScriptLineNumber & " >>" & StringLen(GUICtrlRead($InputDataFileName)) & "<<  " & @CRLF)

    If Execute(GUICtrlRead($InputDataSize)) > 10 And StringLen(GUICtrlRead($InputDataFileName)) > 5 Then
        $CommandString = $CommandString & "  " & FileGetShortName(GUICtrlRead($InputDataFileName))
        $CommandString = $CommandString & "  " & Execute(GUICtrlRead($InputDataSize))
    EndIf

    GUICtrlSetData($LabelStatus, $CommandString)

    Local $Result = 0

    $Result = Run(@ComSpec & " /c " & $URW4Win & " " & $CommandString, " .", @SW_HIDE, $STDOUT_CHILD)
    ConsoleWrite(@ScriptLineNumber & " " & $URW4Win & @CRLF)
    ConsoleWrite(@ScriptLineNumber & " " & $CommandString & @CRLF)
    ConsoleWrite(@ScriptLineNumber & " " & $Result & @CRLF)

EndFunc   ;==>DoIt

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
        MsgBox(15, "Invalid choice at GuiDisable", $choice & "   " & $setting)
        Return
    EndIf
    $LastState = $setting
    For $X = 1 To 100
        GUICtrlSetState($X, $setting)
    Next

    ;the following controls are always enabled
    ;GUICtrlSetState($CheckAbort, $GUI_ENABLE)
    ;GUICtrlSetState($LabelStatus, $GUI_ENABLE)
EndFunc   ;==>GuiDisable
;-----------------------------------------------



