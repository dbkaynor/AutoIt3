#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_icon=../icons/alarm_bell.ico
#AutoIt3Wrapper_outfile=C:\Program Files (x86)\AutoIt3\Dougs\PowerCycle.exe
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Comment=PowerCycle
#AutoIt3Wrapper_Res_Description=PowerCycle
#AutoIt3Wrapper_Res_Fileversion=0.0.0.13
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=Y
#AutoIt3Wrapper_Res_ProductVersion=000
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
#AutoIt3Wrapper_Run_Debug_Mode=N
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****

Opt("MustDeclareVars", 1) ; require pre-declared varibles
Opt("WinTitleMatchMode", 2)
If _Singleton(@ScriptName, 1) = 0 Then
    _Debug(@ScriptName & ' is already running!', '', True, 10)
    Exit
EndIf

#include <ButtonConstants.au3>
#include <Constants.au3>
#include <Date.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <GUIListBox.au3>
#include <iNet.au3>
#include <ListBoxConstants.au3>
#include <Misc.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <_DougFunctions.au3>

Global Const $tmp = StringSplit(@ScriptName, ".")
Global Const $ProgramName = $tmp[1]

Global Const $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")
Global $SystemS = @ScriptName & @LF & $FileVersion & @LF & @OSVersion & @LF & @OSServicePack & @LF & @OSType & @LF & @OSArch
Global Const $LogFilename = $ProgramName & ".log"

Global $WinPosition
Global $OnButtonPosition[2]
Global $OffButtonPosition[2]
Global $CyclePuttonPosition[2]
Global $Trained = False
Global $PowerOnFail = 0
Global $PowerOffFail = 0

Global $SavedTime = -999
Global $CurrentTime = 0
Global $TimeDiff = 0
Global $Editor = ''
Global $Projectfilename = $AuxPath & $ProgramName & '.prj'

Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, _
        $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
; Main form --------------------------
Global $MainForm = GUICreate(@ScriptName & $FileVersion, 550, 350, 10, 10, $MainFormOptions)
GUISetFont(10, 400, -1, "Courier new")

Global $ButtonTrain = GUICtrlCreateButton("Train", 10, 10, 60)
GUICtrlSetTip(-1, 'Train the GUI power button positions')
Global $ButtonStartTesting = GUICtrlCreateButton("Start", 80, 10, 60)
GUICtrlSetTip(-1, 'Start the testing')
Global $ButtonOptions = GUICtrlCreateButton("Options", 150, 10, 70)
GUICtrlSetTip(-1, 'Display and set options')

Global $CheckAbort = GUICtrlCreateCheckbox('Abort', 230, 10, 60)
GUICtrlSetTip(-1, 'Abort testing')

Global $ButtonMainHelp = GUICtrlCreateButton("Help", 310, 10, 60)
GUICtrlSetTip(-1, 'Help')
Global $ButtonMainAbout = GUICtrlCreateButton("About", 380, 10, 60)
GUICtrlSetTip(-1, 'About (debug info) for main window')

Global $ButtonExit = GUICtrlCreateButton("Exit", 450, 10, 60)
GUICtrlSetTip(-1, 'Exit the program')

Global $InputTestStatus = GUICtrlCreateInput("", 10, 40, 520, 20, $ES_READONLY)
GUICtrlSetTip(-1, 'Status message')
Global $ListResults = GUICtrlCreateList('', 10, 70, 530, 280, BitOR($LBS_DISABLENOSCROLL, $WS_BORDER, $WS_HSCROLL, $WS_VSCROLL))
GUICtrlSetTip(-1, 'Status messages')

; Options form
Global $OptionsFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, _
        $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $OptionsForm = GUICreate("Select options", 450, 400, 10, 10, $OptionsFormOptions)
GUICtrlCreateLabel("SUT IP Address", 20, 13, 110, 10)
GUICtrlSetTip(-1, 'IP address of the SUT')
Global $InputSUTIPAddress = GUICtrlCreateInput("", 150, 8, 156, 20, BitOR($GUI_SS_DEFAULT_INPUT, $ES_CENTER))
GUICtrlSetTip(-1, 'IP address of the SUT')
GUICtrlCreateLabel("Command window name", 20, 43, 126, 20)
GUICtrlSetTip(-1, 'The title of the window with the power buttons')
Global $InputCommandWindowName = GUICtrlCreateInput("", 150, 40, 156, 20, BitOR($GUI_SS_DEFAULT_INPUT, $ES_CENTER))
GUICtrlSetTip(-1, 'The title of the window with the power buttons')
GUICtrlCreateLabel("Status window name", 20, 76, 110, 20)
GUICtrlSetTip(-1, 'The title of the power status window')
Global $InputStatusWindowName = GUICtrlCreateInput("", 150, 72, 156, 20, BitOR($GUI_SS_DEFAULT_INPUT, $ES_CENTER))
GUICtrlSetTip(-1, 'The title of the power status window')
GUICtrlCreateLabel("Loops", 20, 112, 110, 20)
GUICtrlSetTip(-1, 'Number of loops the test should run for')
Global $InputLoopsTotal = GUICtrlCreateInput("", 150, 112, 50, 20, BitOR($GUI_SS_DEFAULT_INPUT, $ES_CENTER))
GUICtrlSetTip(-1, 'Number of loops the test should run for')
GUICtrlCreateLabel("Ping timeout", 20, 144, 110, 20)
GUICtrlSetTip(-1, 'The time-out for the ping command')
Global $InputPingTimeout = GUICtrlCreateInput("", 150, 144, 50, 20, BitOR($GUI_SS_DEFAULT_INPUT, $ES_CENTER))
GUICtrlSetTip(-1, 'The time-out for the ping command')
GUICtrlSetTip(-1, 'Delay between pings (seconds')
GUICtrlCreateLabel("Delay between pings", 20, 176, 110, 20)
GUICtrlSetTip(-1, 'Delay between pings (seconds')
Global $InputDelayBetweenPings = GUICtrlCreateInput("", 150, 176, 50, 20, BitOR($GUI_SS_DEFAULT_INPUT, $ES_CENTER))
GUICtrlSetTip(-1, 'Delay between pings (seconds')

GUICtrlCreateLabel('Delay after on', 20, 208, 110, 20)
GUICtrlSetTip(-1, 'Seconds to wait from power on command to start of pings')
Global $InputDelayAfterOn = GUICtrlCreateInput("", 150, 208, 50, 20, BitOR($GUI_SS_DEFAULT_INPUT, $ES_CENTER))
GUICtrlSetTip(-1, 'Seconds to wait from power on command to start of pings')

GUICtrlCreateLabel("Delay after off", 20, 240, 110, 20)
GUICtrlSetTip(-1, 'Seconds to wait from power off command to start of pings')
Global $InputDelayAfterOff = GUICtrlCreateInput("", 150, 240, 50, 20, BitOR($GUI_SS_DEFAULT_INPUT, $ES_CENTER))
GUICtrlSetTip(-1, 'Seconds to wait from power off command to start of pings')
GUICtrlCreateLabel("Power on tries", 20, 272, 110, 20)
GUICtrlSetTip(-1, 'Number of retries for power on to be successfull')
Global $InputPowerOnTries = GUICtrlCreateInput("", 150, 270, 50, 20, BitOR($GUI_SS_DEFAULT_INPUT, $ES_CENTER))
GUICtrlSetTip(-1, 'Number of retries for power on to be successfull')
GUICtrlCreateLabel("Power off tries", 20, 304, 110, 20)
GUICtrlSetTip(-1, 'Number of retries for power off to be successfull')
Global $InputPowerOffTries = GUICtrlCreateInput("", 150, 304, 50, 20, BitOR($GUI_SS_DEFAULT_INPUT, $ES_CENTER))
GUICtrlSetTip(-1, 'Number of retries for power off to be successfull')

Global $ButtonPingSUT = GUICtrlCreateButton("Ping SUT", 215, 110, 90)
GUICtrlSetTip(-1, 'Ping the SUT using the SUT IP address')
Global $ButtonStartDebugger = GUICtrlCreateButton("Start debugger", 215, 140, 90)
GUICtrlSetTip(-1, 'Start the Windows debug tool')
Global $ButtonOptionsAbout = GUICtrlCreateButton("About", 215, 170, 90)
GUICtrlSetTip(-1, 'About (debug info) for options window')
Global $CheckTestMode = GUICtrlCreateCheckbox("Test mode", 215, 200, 90)
GUICtrlSetTip(-1, 'Test mode (do not actually cycle power)')
Global $ButtonHideOptionsForm = GUICtrlCreateButton("Hide options form ", 215, 260, 90)
GUICtrlSetTip(-1, 'Hide this (options) window')

Global $ButtonLoadProject = GUICtrlCreateButton('Load project', 315, 110, 90)
GUICtrlSetTip(-1, 'Load a saved project')
Global $ButtonSaveProject = GUICtrlCreateButton('Save project', 315, 140, 90)
GUICtrlSetTip(-1, 'Save current settings as a project')
Global $ButtonloadDefaultOptions = GUICtrlCreateButton('Load default', 315, 170, 90)
GUICtrlSetTip(-1, 'Load a default settings')
Global $ButtonEditProject = GUICtrlCreateButton('Edit project', 315, 200, 90)
GUICtrlSetTip(-1, 'Edit/view a saved project')
Global $ButtonViewLog = GUICtrlCreateButton('View log', 315, 230, 90)
GUICtrlSetTip(-1, 'View the log file')
Global $ButtonDeleteLog = GUICtrlCreateButton('Delete log', 315, 260, 90)
GUICtrlSetTip(-1, 'Delete/Clear the log file')

Global $InputPingStatus = GUICtrlCreateInput("", 20, 330, 320, 20, $ES_READONLY)
GUICtrlSetTip(-1, 'Show the ping command status')

For $x = 0 To 100 ; This sets the resize mode for all GUI items (assumingg 100 items)
    GUICtrlSetResizing($x, $GUI_DOCKALL)
Next

;Change the resize mode to selected items
GUICtrlSetResizing($ListResults, $GUI_DOCKTOP + $GUI_DOCKBOTTOM + $GUI_DOCKRIGHT + $GUI_DOCKLEFT)
GUICtrlSetResizing($InputTestStatus, $GUI_DOCKHEIGHT + $GUI_DOCKTOP + $GUI_DOCKRIGHT + $GUI_DOCKLEFT)


LoadDefaults()
LoadProject("start")
_CheckWindowLocation($MainForm)
_CheckWindowLocation($OptionsForm)
GUISetState(@SW_SHOW, $MainForm)
GUISetState(@SW_HIDE, $OptionsForm)
;-----------------------------------------------
While 1
    Global $nMsg = GUIGetMsg(1)
    Switch $nMsg[0]
        ; Main menu
        Case $ButtonExit
            Exit
        Case $GUI_EVENT_CLOSE
            If $nMsg[1] = $MainForm Then
                Exit
            ElseIf $nMsg[1] = $OptionsForm Then
                GUISetState(@SW_HIDE, $OptionsForm)
                GUISetState(@SW_SHOW, $MainForm)
            EndIf
        Case $ButtonOptions
            GUISetState(@SW_SHOW, $OptionsForm)
            GUISetState(@SW_HIDE, $MainForm)
        Case $ButtonHideOptionsForm
            GUISetState(@SW_HIDE, $OptionsForm)
            GUISetState(@SW_SHOW, $MainForm)
        Case $GUI_EVENT_RESIZED
        Case $ButtonStartTesting
            StartTesting()
        Case $ButtonTrain
            Train()

        Case $ButtonMainHelp
            ShellExecute($AuxPath & 'PowerCycle.htm')
        Case $ButtonMainAbout
            _About($ProgramName, $SystemS)
        Case $ListResults
            If MsgBox(32 + 4, 'Clear status messages', 'Do you want to clear this box?') = 6 Then
                _GUICtrlListBox_ResetContent($ListResults)
            EndIf
            ; Option menu
        Case $ButtonOptionsAbout
            _About('Select options', $SystemS)
        Case $InputSUTIPAddress
        Case $ButtonPingSUT
            PingSUT()
        Case $ButtonStartDebugger
            _StartDebugViewer('')
        Case $ButtonLoadProject
            LoadProject("menu")
        Case $ButtonSaveProject
            SaveProject()
        Case $ButtonloadDefaultOptions
            LoadDefaults()
        Case $ButtonEditProject
            $Editor = _ChoseTextEditor()
            ShellExecute($Editor, $Projectfilename)
        Case $ButtonViewLog
            $Editor = _ChoseTextEditor()
            ShellExecute($Editor, $LogFilename)
        Case $ButtonDeleteLog
            If MsgBox(32 + 4, 'Delete log file', 'Do you want to delete the log file?') = 6 Then
                FileDelete($LogFilename)
            EndIf
        Case $CheckAbort
            GuiDisable($GUI_ENABLE)
            ConsoleWrite(@ScriptLineNumber & ' CheckAbort ' & @CRLF)
    EndSwitch
WEnd

;-----------------------------------------------
Func LoadDefaults()
    GUICtrlSetData($InputSUTIPAddress, '172.31.107.200')
    GUICtrlSetData($InputCommandWindowName, 'Raritan CC-SG Html Client')
    GUICtrlSetData($InputStatusWindowName, 'Power Status Messages')
    GUICtrlSetData($InputLoopsTotal, 3)
    ;Delays and timeouts are in seconds
    GUICtrlSetData($InputPingTimeout, 2)
    GUICtrlSetData($InputDelayBetweenPings, 6)
    GUICtrlSetData($InputDelayAfterOn, 10)
    GUICtrlSetData($InputDelayAfterOff, 5)
    GUICtrlSetData($InputPowerOnTries, 60)
    GUICtrlSetData($InputPowerOffTries, 11)
    GUICtrlSetState($CheckTestMode, $GUI_CHECKED)

    _SetWindowPosition('MainWindow:', $ProgramName, 'MainWindow:10 10 550 350')
    _SetWindowPosition("OptionsWindow:", 'Select options', 'OptionsWindow:10 10 450 400')
    _CheckWindowLocation($MainForm)
    _CheckWindowLocation($OptionsForm)
EndFunc   ;==>LoadDefaults

;-----------------------------------------------
Func LoadProject($type)
    If StringCompare($type, "menu") = 0 Then
        $Projectfilename = FileOpenDialog("Loadproject file", @ScriptDir & ".\AUXFiles\", _
                $ProgramName & ' projects (' & $ProgramName & '.prj)|All projects (*.prj)|All files (*.*)', 18, @ScriptDir & '.\AUXFiles\' & $ProgramName & '.prj')
    EndIf

    Local $file = FileOpen($Projectfilename, 0)
    ; Check if file opened for reading OK
    If $file = -1 Then
        _Debug("Unable to open file for reading: " & $Projectfilename, '', True)
        Return
    EndIf

    ; Read in the first line to verify the file is of the correct type
    If StringInStr(FileReadLine($file, 1), "Project file for " & $ProgramName) <> 1 Then
        ConsoleWrite(@ScriptLineNumber & " >" & FileReadLine($file, 1) & "<" & @CRLF)
        ConsoleWrite(@ScriptLineNumber & " >Project file for " & $ProgramName & "<" & @CRLF)
        ConsoleWrite(@ScriptLineNumber & " " & StringInStr(FileReadLine($file, 1), "Project file for " & $ProgramName) & @CRLF)
        _Debug("Not a Project file for " & $ProgramName, '', True)
        Return
    EndIf

    ; Read in lines of text until the EOF is reached
    While 1
        Local $lineIn = FileReadLine($file)
        If @error = -1 Then ExitLoop
        If StringInStr($lineIn, "SUTIPAddress:") Then GUICtrlSetData($InputSUTIPAddress, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
        If StringInStr($lineIn, "CommandWindowName:") Then GUICtrlSetData($InputCommandWindowName, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
        If StringInStr($lineIn, "StatusWindowName:") Then GUICtrlSetData($InputStatusWindowName, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
        If StringInStr($lineIn, "LoopsTotal:") Then GUICtrlSetData($InputLoopsTotal, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
        If StringInStr($lineIn, "PingTimeout:") Then GUICtrlSetData($InputPingTimeout, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
        If StringInStr($lineIn, "DelayBetweenPings:") Then GUICtrlSetData($InputDelayBetweenPings, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
        If StringInStr($lineIn, "DelayAfterOn:") Then GUICtrlSetData($InputDelayAfterOn, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
        If StringInStr($lineIn, "DelayAfterOff:") Then GUICtrlSetData($InputDelayAfterOff, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
        If StringInStr($lineIn, "PowerOnTries:") Then GUICtrlSetData($InputPowerOnTries, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
        If StringInStr($lineIn, "PowerOffTries:") Then GUICtrlSetData($InputPowerOffTries, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
        If StringInStr($lineIn, "TestMode:") Then GUICtrlSetState($CheckTestMode, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
        _SetWindowPosition('MainWindow:', $ProgramName, $lineIn)
        _SetWindowPosition("OptionsWindow:", 'Select options', $lineIn)
        _CheckWindowLocation($MainForm)
        _CheckWindowLocation($OptionsForm)
    WEnd
    FileClose($file)

EndFunc   ;==>LoadProject
;-----------------------------------------------
Func SaveProject()
    $Projectfilename = FileSaveDialog("Save project file", @ScriptDir & ".\AUXFiles\", _
            $ProgramName & ' projects (' & $ProgramName & '.prj)|All projects (*.prj)|All files (*.*)', 18, @ScriptDir & '.\AUXFiles\' & $ProgramName & '.prj')

    Local $file = FileOpen($Projectfilename, 2)
    ; Check if file opened for writing OK
    If $file = -1 Then
        _Debug("Unable to open file for writing: " & $Projectfilename, '', True)
        Exit
    EndIf

    FileWriteLine($file, "Project file for " & $ProgramName & "  Saved on " & _DateTimeFormat(_NowCalc(), 0))
    ; Write the lines of text to the file
    FileWriteLine($file, "Valid for " & $ProgramName)
    FileWriteLine($file, 'SUTIPAddress:' & GUICtrlRead($InputSUTIPAddress))
    FileWriteLine($file, 'CommandWindowName:' & GUICtrlRead($InputCommandWindowName))
    FileWriteLine($file, 'StatusWindowName:' & GUICtrlRead($InputStatusWindowName))
    FileWriteLine($file, "LoopsTotal:" & GUICtrlRead($InputLoopsTotal))
    FileWriteLine($file, "PingTimeout:" & GUICtrlRead($InputPingTimeout)) ;milli-seconds
    ;Delays are in seconds
    FileWriteLine($file, "DelayBetweenPings:" & GUICtrlRead($InputDelayBetweenPings))
    FileWriteLine($file, "DelayAfterOn:" & GUICtrlRead($InputDelayAfterOn))
    FileWriteLine($file, "DelayAfterOff:" & GUICtrlRead($InputDelayAfterOff))
    FileWriteLine($file, "PowerOnTries:" & GUICtrlRead($InputPowerOnTries))
    FileWriteLine($file, "PowerOffTries:" & GUICtrlRead($InputPowerOffTries))
    FileWriteLine($file, "TestMode:" & GUICtrlRead($CheckTestMode))
    _SaveWindowPosition("MainWindow:", $ProgramName, $file)
    _SaveWindowPosition("OptionsWindow:", 'Select options', $file)

    FileClose($file)
EndFunc   ;==>SaveProject
;-----------------------------------------------
Func CheckWindow()
    If WinExists(GUICtrlRead($InputCommandWindowName), '') = 0 Then
        MsgBox(16, 'Window not found', 'Window not found' & @CRLF & GUICtrlRead($InputCommandWindowName), 10)
        _Debug(@ScriptLineNumber & ' Window not found ' & GUICtrlRead($InputCommandWindowName), $LogFilename, $ListResults)
        Return 1
    EndIf
    Return 0
EndFunc   ;==>CheckWindow
;-----------------------------------------------
Func Train()
    _Debug(@ScriptLineNumber & ' Train', $LogFilename, $ListResults)
    If CheckWindow() <> 0 Then Return
    GUICtrlSetState($CheckAbort, $GUI_UNCHECKED)
    WinSetTrans(GUICtrlRead($InputCommandWindowName), '', 255)
    WinSetState(GUICtrlRead($InputCommandWindowName), '', @SW_RESTORE)
    WinActivate(GUICtrlRead($InputCommandWindowName))

    If GUICtrlRead($CheckTestMode) = $GUI_CHECKED Then
        $WinPosition = WinGetPos(GUICtrlRead($InputCommandWindowName))
        $OnButtonPosition[0] = 50
        $OnButtonPosition[1] = 50
        $OffButtonPosition[0] = 50
        $OffButtonPosition[1] = 100
        $CyclePuttonPosition[0] = 50
        $CyclePuttonPosition[1] = 150
        $Trained = True
        Return
    EndIf

    If MsgBox(49, 'Position window', 'Position window to desired location' & @CR & 'Make sure On, Off, and Cycle buttons are visible' & @CR & 'Press Enter when ready') = 2 Then Return
    $WinPosition = WinGetPos(GUICtrlRead($InputCommandWindowName))
    _Debug(@ScriptLineNumber & ' Window position:' & _ArrayToString($WinPosition), $LogFilename, $ListResults)
    If MsgBox(49, 'Position mouse', 'Position mouse or On button' & @CR & 'Press Enter when ready') = 2 Then Return
    $OnButtonPosition = MouseGetPos()
    _Debug(@ScriptLineNumber & ' On button:' & _ArrayToString($OnButtonPosition), $LogFilename, $ListResults)
    If MsgBox(49, 'Position mouse', 'Position mouse or Off button' & @CR & 'Press enter when ready') = 2 Then Return
    $OffButtonPosition = MouseGetPos()
    _Debug(@ScriptLineNumber & ' Off button:' & _ArrayToString($OffButtonPosition), $LogFilename, $ListResults)
    If MsgBox(49, 'Position mouse', 'Position mouse or Cycle button' & @CR & 'Press enter when ready') = 2 Then Return
    $CyclePuttonPosition = MouseGetPos()
    _Debug(@ScriptLineNumber & ' Cycle button:' & _ArrayToString($CyclePuttonPosition), $LogFilename, $ListResults)
    $Trained = True
EndFunc   ;==>Train
;-----------------------------------------------
Func StartTesting()
    If Not pingsut() Then
        If MsgBox(49, 'SUT is not reponding to ping', _
                'Check that it is powered on and/or check your setup (IP etc)') & _
                'Cancel to abort, OK to continue.' = 2 Then Return
    EndIf

    If GUICtrlRead($CheckTestMode) = $GUI_CHECKED Then
        If MsgBox(48 + 1, 'Test mode is set', 'Use the option menu to disable it ' & _
                'if desired and restart test or press OK to continue in test mode') = 2 Then Return
    EndIf
    GuiDisable($GUI_DISABLE)
    If Not $Trained Then Train()
    _Debug(@ScriptLineNumber & ' Start', $LogFilename, $ListResults)
    GUICtrlSetData($InputTestStatus, ' Start of ' & GUICtrlRead($InputLoopsTotal) & ' total loops')
    GUICtrlSetState($CheckAbort, $GUI_UNCHECKED)

    Local $WindowMoveSpeed = 5
    $PowerOnFail = 0
    $PowerOffFail = 0
    WinSetTrans(GUICtrlRead($InputCommandWindowName), '', 128)
    Local $CurrentLoop = 0
    While $CurrentLoop < GUICtrlRead($InputLoopsTotal)
        If GUICtrlRead($CheckAbort) = $GUI_CHECKED Then
            _Debug(@ScriptLineNumber & ' Abort detected', $LogFilename, $ListResults)
            ExitLoop
        EndIf
        If CheckWindow() Then ExitLoop
        $CurrentLoop = $CurrentLoop + 1

        WinActivate(GUICtrlRead($InputCommandWindowName))
        WinWaitActive(GUICtrlRead($InputCommandWindowName))
        WinMove(GUICtrlRead($InputCommandWindowName), '', $WinPosition[0], $WinPosition[1], $WinPosition[2], $WinPosition[3], $WindowMoveSpeed)
        ClickMouse('off')
        KillPowerStatusWindows()
        _Debug(@ScriptLineNumber & ' Off button click', $LogFilename, $ListResults)
        If GUICtrlRead($CheckAbort) = $GUI_CHECKED Then
            _Debug(@ScriptLineNumber & ' Abort detected', $LogFilename, $ListResults)
            ExitLoop
        EndIf
        TestSUTState('off')

        WinActivate(GUICtrlRead($InputCommandWindowName))
        WinWaitActive(GUICtrlRead($InputCommandWindowName))
        WinMove(GUICtrlRead($InputCommandWindowName), '', $WinPosition[0], $WinPosition[1], $WinPosition[2], $WinPosition[3], $WindowMoveSpeed)
        ClickMouse('on')
        KillPowerStatusWindows()
        _Debug(@ScriptLineNumber & ' On button click', $LogFilename, $ListResults)
        If GUICtrlRead($CheckAbort) = $GUI_CHECKED Then
            _Debug(@ScriptLineNumber & ' Abort detected', $LogFilename, $ListResults)
            ExitLoop
        EndIf
        TestSUTState('on')
        Local $TS = ' Done with ' & $CurrentLoop & ' loops of ' & GUICtrlRead($InputLoopsTotal) & ' Power off fails: ' & $PowerOffFail & ' Power on fails: ' & $PowerOnFail
        _Debug(@ScriptLineNumber & $TS, $LogFilename, $ListResults)
        GUICtrlSetData($InputTestStatus, $TS)
        KillPowerStatusWindows()
    WEnd

    _Debug(@ScriptLineNumber & ' All done', $LogFilename, $ListResults)
    ;All done so turn SUT back on
    WinSetTrans(GUICtrlRead($InputCommandWindowName), '', 255)
    WinActivate(GUICtrlRead($InputCommandWindowName))
    WinMove(GUICtrlRead($InputCommandWindowName), '', $WinPosition[0], $WinPosition[1], $WinPosition[2], $WinPosition[3], $WindowMoveSpeed)
    ClickMouse('on')
    KillPowerStatusWindows()
    $TS = ' Testing done ' & $CurrentLoop & ' loops of ' & GUICtrlRead($InputLoopsTotal) & ' Power off fails: ' & $PowerOffFail & ' Power on fails: ' & $PowerOnFail
    _Debug(@ScriptLineNumber & $TS, $LogFilename, $ListResults)
    GUICtrlSetData($InputTestStatus, $TS)
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>StartTesting

;-----------------------------------------------
;This will loop until SUT responds or doesn't or until timeout
Func TestSUTState($type)
    _Debug(@ScriptLineNumber & ' TestSUTState ' & $type, $LogFilename, $ListResults)
    Local $TryCount = 0
    Local $DelayBetweenPings = GUICtrlRead($InputDelayBetweenPings)
    Local $DelayAfterOn = GUICtrlRead($InputDelayAfterOn)
    Local $DelayAfterOff = GUICtrlRead($InputDelayAfterOff)
    Local $PowerOnTries = GUICtrlRead($InputPowerOnTries)
    Local $PowerOffTries = GUICtrlRead($InputPowerOffTries)

    If GUICtrlRead($CheckTestMode) = $GUI_CHECKED Then
        $DelayBetweenPings = 1
        $DelayAfterOn = 1
        $DelayAfterOff = 1
        $PowerOnTries = 1
        $PowerOffTries = 1
    EndIf

    $SavedTime = TimerInit()
    If $type = 'on' Then
        MySleep($DelayAfterOn)
        While $TryCount < $PowerOnTries And PingSUT() = False
            If GUICtrlRead($CheckAbort) = $GUI_CHECKED Then ExitLoop
            mySleep($DelayBetweenPings)
            $TryCount = $TryCount + 1
        WEnd
        If $TryCount >= $PowerOnTries Then $PowerOnFail = $PowerOnFail + 1

    ElseIf $type = 'off' Then
        MySleep($DelayAfterOff)
        While $TryCount < $PowerOffTries And PingSUT() = True
            If GUICtrlRead($CheckAbort) = $GUI_CHECKED Then ExitLoop
            MySleep($DelayBetweenPings)
            $TryCount = $TryCount + 1
        WEnd
        If $TryCount >= $PowerOffTries Then $PowerOffFail = $PowerOffFail + 1
    Else
        MsgBox(48, 'TestSUTState error', 'TestSUTState error. Bad type' & @CRLF & $type)
    EndIf
    _Debug(@ScriptLineNumber & StringFormat(" SUT response time: (Type) %s %3.1f seconds", $type, TimerDiff($SavedTime) / 1000), $LogFilename, $ListResults)
EndFunc   ;==>TestSUTState
;-----------------------------------------------
Func ClickMouse($State)
    Local $MouseMoveSpeed = 10
    Local $MouseButton

    ; This is for debugging
    If GUICtrlRead($CheckTestMode) = $GUI_CHECKED Then
        $MouseButton = 'right'
    Else
        $MouseButton = 'left'
    EndIf

    Switch $State
        Case 'on'
            MouseClick($MouseButton, $OnButtonPosition[0], $OnButtonPosition[1], 1, $MouseMoveSpeed) ; On button
        Case 'off'
            MouseClick($MouseButton, $OffButtonPosition[0], $OffButtonPosition[1], 1, $MouseMoveSpeed) ; Off button
        Case Else
            MsgBox(48, 'ClickMouse State error', 'ClickMouse State error. Bad state value' & @CRLF & $State)
    EndSwitch

EndFunc   ;==>ClickMouse
;-----------------------------------------------
Func PingSUT()
    Local $PingTimeout = GUICtrlRead($InputPingTimeout)
    If GUICtrlRead($CheckTestMode) = $GUI_CHECKED Then $PingTimeout = 500
    Local $result = Ping(GUICtrlRead($InputSUTIPAddress), $PingTimeout)
    Local $Error = @error
    ;_Debug(@ScriptLineNumber & ' PingIP: ' & GUICtrlRead($InputSUTIPAddress) & '   Time: ' & $result & '   Error: ' & PingError($Error), $LogFilename)
    If $result = 0 Or @error <> 0 Then
        GUICtrlSetData($InputPingStatus, ' PingSUT FAIL. ' & _PingError($Error) & ' Return time: ' & $result)
        _Debug(@ScriptLineNumber & ' PingSUT FAIL. ' & _PingError($Error) & ' Return time: ' & $result, $LogFilename, $ListResults)
        Return False
    Else
        GUICtrlSetData($InputPingStatus, ' PingSUT PASS. ' & _PingError($Error) & ' Return time: ' & $result)
        _Debug(@ScriptLineNumber & ' PingSUT PASS. ' & _PingError($Error) & ' Return time: ' & $result, $LogFilename, $ListResults)
        Return True
    EndIf
EndFunc   ;==>PingSUT
;-----------------------------------------------
Func KillPowerStatusWindows()
    If WinExists(GUICtrlRead($InputStatusWindowName)) = 0 Then Sleep(2000)
    While WinExists(GUICtrlRead($InputStatusWindowName)) = 1
        WinClose(GUICtrlRead($InputStatusWindowName))
        WinKill(GUICtrlRead($InputStatusWindowName))
    WEnd
EndFunc   ;==>KillPowerStatusWindows
;-----------------------------------------------
Func MySleep($Time)
    For $x = 0 To 1000
        Sleep($Time)
        If GUICtrlRead($CheckAbort) = $GUI_CHECKED Then ExitLoop
    Next
EndFunc   ;==>MySleep
;-----------------------------------------------
Func GuiDisable($choice) ;$GUI_ENABLE $GUI_DISABLE
    For $x = 1 To 200
        GUICtrlSetState($x, $choice)
    Next
    GUICtrlSetState($CheckAbort, $GUI_ENABLE)
EndFunc   ;==>GuiDisable
;-----------------------------------------------
