#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_icon=../icons/HotSun.ico
#AutoIt3Wrapper_outfile=C:\Program Files (x86)\AutoIt3\Dougs\Power.exe
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Comment=Start various programs
#AutoIt3Wrapper_Res_Description=Start various programs
#AutoIt3Wrapper_Res_Fileversion=0.0.0.5
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=Y
#AutoIt3Wrapper_Res_ProductVersion=000
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
#AutoIt3Wrapper_Run_Debug_Mode=N
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****

#cs
    COMPLETED
    Save, restore, and default options
    Verify the existance of the power switcher
    Put options in a separate window
    Logging
    Fix validate paths
    Catch return data from switcher
    Add wake on lan feature
    Manually enter SUT MAC
    load default options button
    
    TO DO
    load saved options on startup
    Verbosity level
#ce

#include <ButtonConstants.au3>
#include <Constants.au3>
#include <Date.au3>
#include <EditConstants.au3>
#include <File.au3>
#include <ff.au3>
#include <GUIMenu.au3>
#include <GUIConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiListBox.au3>
#include <Misc.au3>
#include <String.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <Winpcap.au3>
#include <_DougFunctions.au3>

DirCreate("AUXFiles")
AutoItSetOption("MustDeclareVars", 1)

Global Const $tmp = StringSplit(@ScriptName, ".")
Global Const $ProgramName = $tmp[1]

Global $AuxPath = @ScriptDir & "\AUXFiles\"
Global Const $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")
Global $SystemS = @ScriptName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & @OSServicePack & @CRLF & @OSType & @CRLF & @OSArch

Global $Running = False
Global $IterationCount = 0
Global $PassCount = 0
Global $FailCountOn = 0
Global $FailCountOff = 0
Global $SUTMACAddress = ''
Global $PowerMACAddress = ''
Global $cmd = ''

Opt("MustDeclareVars", 1) ; require pre-declared varibles
If _Singleton(@ScriptName, 1) = 0 Then
    MsgBox(16, @ScriptName, @ScriptName & @CRLF & " is already running!")
    Exit
EndIf

; Power Main --------------------------
Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, _
        $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $MainForm = GUICreate($ProgramName & "  " & @ComputerName & "  " & $FileVersion, 620, 375, 10, 10, $MainFormOptions)
GUISetFont(9, 400, -1, "Courier new")

Global $ButtonRun = GUICtrlCreateButton("Run", 10, 10, 60, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Run test")
Global $ButtonOptionsMain = GUICtrlCreateButton("Options", 100, 10, 60, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Display the options menu")
Global $ButtonToolsMain = GUICtrlCreateButton("Tools", 190, 10, 60, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Display the tools menu")

Global $ButtonMainAbout = GUICtrlCreateButton("About", 300, 10, 60, 25, $WS_GROUP)
GUICtrlSetTip(-1, "About the main window")
Global $ButtonExit = GUICtrlCreateButton("Exit", 380, 10, 60, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Exit Program")

Global $CheckBoxAbort = GUICtrlCreateCheckbox("Abort", 550, 10, 60, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Abort the looping Program")

Global $ListResults = GUICtrlCreateList("", 10, 40, 600, 300, _
        BitOR($LBS_NOSEL, $LBS_DISABLENOSCROLL, $WS_BORDER, $WS_HSCROLL, $WS_VSCROLL, $WS_TABSTOP, $LBS_NOTIFY))
Global $LabelStatusMain = GUICtrlCreateLabel("status", 10, 340, 600, 20, $SS_SUNKEN)

; Options form
Global $OptionFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, _
        $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $OptionForm = GUICreate("Options", 550, 400, 10, 10, $OptionFormOptions)

GUICtrlCreateLabel("SUT IP Address", 10, 50, 102, 20)
Global $InputSUTIPAddress = GUICtrlCreateInput("SUT IP Address", 120, 50, 400, 20)
GUICtrlSetTip(-1, "SUT IP Address")
GUICtrlCreateLabel("Switcher IP Address", 10, 80, 102, 20)
Global $InputSwitcherIPAddress = GUICtrlCreateInput("Switcher IP Address", 120, 80, 400, 20)
GUICtrlSetTip(-1, "Switcher IP Address")
GUICtrlCreateLabel("Max Iterations", 10, 110, 102, 20)
Global $InputMaxIterations = GUICtrlCreateInput("Max Iterations", 120, 110, 400, 20)
GUICtrlSetTip(-1, "Max Iterations")
GUICtrlCreateLabel("Ping delay", 10, 140, 74, 20)
Global $InputPingDelay = GUICtrlCreateInput("Ping Delay", 120, 140, 400, 20)
GUICtrlSetTip(-1, "Ping delay")
GUICtrlCreateLabel("Max ping retries", 10, 170, 116, 20)
Global $InputMaxPingRetrys = GUICtrlCreateInput("Max Ping Retrys", 120, 170, 400, 20)
GUICtrlSetTip(-1, "Max ping retrys")
GUICtrlCreateLabel("Ping timeout", 10, 200, 88, 20)
Global $InputPingTimeOut = GUICtrlCreateInput("Ping Time Out", 120, 200, 400, 20)
GUICtrlSetTip(-1, "Ping time out")
GUICtrlCreateLabel("After power off", 10, 230, 109, 20)
Global $InputAfterPowerOffDelay = GUICtrlCreateInput("After Power Off Delay", 120, 230, 400, 20)
GUICtrlSetTip(-1, "After Power Off Delay")
GUICtrlCreateLabel("After power on", 10, 260, 102, 20)
Global $InputAfterPowerOnDelay = GUICtrlCreateInput("After Power On Delay", 120, 260, 400, 20)
GUICtrlSetTip(-1, "After Power On Delay")
GUICtrlCreateLabel("Log file path", 10, 290, 95, 20)
Global $InputLogFilePath = GUICtrlCreateInput("Log path", 120, 290, 400, 20)
GUICtrlSetTip(-1, "Log path")
GUICtrlCreateLabel("Options path", 10, 320, 95, 20)
Global $InputOptionFilePath = GUICtrlCreateInput("Option path", 120, 320, 400, 20)
GUICtrlSetTip(-1, "Option path")
GUICtrlCreateLabel("Firefox path", 10, 350, 88, 20)
Global $InputFirefoxPath = GUICtrlCreateInput("Firefox path", 120, 350, 400, 20)
GUICtrlSetTip(-1, "Firefox path")
Global $CheckBoxStopOnFail = GUICtrlCreateCheckbox("Stop On Fail", 120, 375, 100, 20)
GUICtrlSetTip(-1, "Stop on fail")
Global $CheckBoxCleanShutDown = GUICtrlCreateCheckbox("Clean shut down", 230, 375, 100, 20)
GUICtrlSetTip(-1, "Clean shut down")

Global $ButtonSaveOptions = GUICtrlCreateButton("Save options", 10, 10, 80, 25)
GUICtrlSetTip(-1, "Save current options to an opt file")
Global $ButtonLoadOptions = GUICtrlCreateButton("Load options", 100, 10, 80, 25)
GUICtrlSetTip(-1, "Load options from an opt file")
Global $ButtonDefaultOptions = GUICtrlCreateButton("Default options", 190, 10, 80, 25)
GUICtrlSetTip(-1, "Load options from an opt file")

Global $ButtonMainOptions = GUICtrlCreateButton("Main form", 330, 10, 70, 25)
GUICtrlSetTip(-1, "Display the main form")
Global $ButtonToolsOptions = GUICtrlCreateButton("Tools form", 400, 10, 70, 25)
GUICtrlSetTip(-1, "Display the tools form")

Global $ButtonOptionAbout = GUICtrlCreateButton("About", 470, 10, 60, 25)
GUICtrlSetTip(-1, "About the Option window")

; Tools form
Global $ToolsFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, _
        $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $ToolsForm = GUICreate("Tools", 450, 150, 10, 10, $ToolsFormOptions)

Global $ButtonMainTools = GUICtrlCreateButton("Main form", 10, 10, 80, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Display the main form")
Global $ButtonOptionsTools = GUICtrlCreateButton("Options form", 100, 10, 80, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Display the options form")
Global $ButtonToolsAbout = GUICtrlCreateButton("About", 190, 10, 60, 25, $WS_GROUP)
GUICtrlSetTip(-1, "About the tools window")
Global $ButtonViewLog = GUICtrlCreateButton("View log", 260, 10, 60, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Display the log")
Global $ButtonClear = GUICtrlCreateButton("Clear log", 330, 10, 60, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Clear results list and log")

Global $ButtonOn = GUICtrlCreateButton("SUT On", 10, 40, 50, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Power on SUT")
Global $ButtonOff = GUICtrlCreateButton("SUT Off", 70, 40, 50, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Power off SUT")
Global $ButtonPingSUT = GUICtrlCreateButton("Ping SUT", 130, 40, 60, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Ping SUT")
Global $ButtonStatus = GUICtrlCreateButton("PWR Stat", 200, 40, 60, 25, $WS_GROUP)
GUICtrlSetTip(-1, "SUT power status")

Global $ButtonTestForPowerSwitch = GUICtrlCreateButton("PS test", 270, 40, 60, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Test for power switch")
Global $ButtonTestForSUT = GUICtrlCreateButton("SUT test", 340, 40, 60, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Test for SUT")
Global $ButtonCleanShutDown = GUICtrlCreateButton("Clean Down", 10, 70, 70, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Clean shutdown")
Global $ButtonWOL = GUICtrlCreateButton("WOL", 100, 70, 70, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Send WOL")
Global $ButtonStartFF = GUICtrlCreateButton("Start FF", 200, 70, 60, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Start Firefox")
Global $CheckBoxDebug = GUICtrlCreateCheckbox("Debug", 300, 70, 60, 20)
GUICtrlSetTip(-1, "Debug")
Global $LabelSUTInfoTools = GUICtrlCreateLabel("SUTInfo", 10, 100, 430, 20, $SS_SUNKEN)
Global $LabelStatusTools = GUICtrlCreateLabel("status", 10, 130, 430, 20, $SS_SUNKEN)
;-----------------------------------------------
LoadDefaultOptions()
VerifyPaths()
TestForPowerSwitch('Start')
TestForSUT('')

GUISetState(@SW_SHOW, $MainForm)
GUISetState(@SW_HIDE, $OptionForm)
GUISetState(@SW_HIDE, $ToolsForm)

If $CmdLine[0] > 0 Then
    GUICtrlSetState($CheckBoxDebug, $GUI_CHECKED)
Else
    StartFirefox()
EndIf
;-----------------------------------------------
While 1
    Global $msg = GUIGetMsg(True)
    Switch $msg[0]
        Case $GUI_EVENT_CLOSE
            If $msg[1] = $MainForm Then
                _FFQuit()
                Exit
            EndIf
            If $msg[1] = $OptionForm Then
                GUISetState(@SW_HIDE, $OptionForm)
                GUISetState(@SW_SHOW, $MainForm)
                GUISetState(@SW_HIDE, $ToolsForm)
            EndIf
            If $msg[1] = $ToolsForm Then
                GUISetState(@SW_HIDE, $ToolsForm)
                GUISetState(@SW_SHOW, $MainForm)
                GUISetState(@SW_HIDE, $OptionForm)
            EndIf

        Case $ButtonMainOptions
            GUISetState(@SW_HIDE, $OptionForm)
            GUISetState(@SW_SHOW, $MainForm)
            GUISetState(@SW_HIDE, $ToolsForm)

        Case $ButtonMainTools
            GUISetState(@SW_HIDE, $OptionForm)
            GUISetState(@SW_SHOW, $MainForm)
            GUISetState(@SW_HIDE, $ToolsForm)

        Case $ButtonOptionsMain
            GUISetState(@SW_SHOW, $OptionForm)
            GUISetState(@SW_HIDE, $MainForm)
            GUISetState(@SW_HIDE, $ToolsForm)

        Case $ButtonOptionsTools
            GUISetState(@SW_SHOW, $OptionForm)
            GUISetState(@SW_HIDE, $MainForm)
            GUISetState(@SW_HIDE, $ToolsForm)

        Case $ButtonToolsMain
            GUISetState(@SW_HIDE, $OptionForm)
            GUISetState(@SW_SHOW, $ToolsForm)
            GUISetState(@SW_HIDE, $MainForm)

        Case $ButtonToolsOptions
            GUISetState(@SW_HIDE, $OptionForm)
            GUISetState(@SW_SHOW, $ToolsForm)
            GUISetState(@SW_HIDE, $MainForm)

        Case $ButtonExit
            ExitLoop
        Case $ButtonMainAbout
            About("Power")

        Case $ButtonPingSUT
            GUICtrlSetState($CheckBoxAbort, $GUI_UNCHECKED)
            PingLoop(True)
        Case $ButtonClear
            FileDelete(GUICtrlRead($InputLogFilePath))
            _GUICtrlListBox_ResetContent($ListResults)
        Case $ButtonRun
            LogOptions()
            VerifyPaths()
            $SUTMACAddress = GetMacAddress(GUICtrlRead($InputSUTIPAddress))
            GUICtrlSetState($CheckBoxAbort, $GUI_UNCHECKED)
            $Running = True
        Case $CheckBoxAbort
            Status(" Running: " & $Running)
        Case $CheckBoxStopOnFail
        Case $CheckBoxCleanShutDown
        Case $CheckBoxDebug
        Case $ButtonViewLog
            If FileExists(GUICtrlRead($InputLogFilePath)) Then
                ConsoleWrite(@ScriptLineNumber & " " & GUICtrlRead($InputLogFilePath) & @CRLF)
                ShellExecuteWait("notepad.exe", GUICtrlRead($InputLogFilePath), @SW_HIDE)
                GuiDisable($GUI_ENABLE)
            EndIf

        Case $ButtonCleanShutDown ; shut it down
            CleanShutDown()

        Case $ButtonWOL ; Bring it up
            SendWOL()
        Case $ButtonStartFF
            StartFirefox()

        Case $ButtonOptionAbout
            About("Options")

        Case $ButtonLoadOptions
            LoadOptions()
        Case $ButtonSaveOptions
            SaveOptions()
        Case $ButtonDefaultOptions
            LoadDefaultOptions()
        Case $ButtonTestForPowerSwitch
            TestForPowerSwitch('menu')
        Case $ButtonTestForSUT
            TestForSUT('menu')

        Case $ButtonOn
            PowerOnCommand()
            ConsoleWrite(@ScriptLineNumber & " " & $Running & @CRLF)
        Case $ButtonOff
            PowerOffCommand()
            ConsoleWrite(@ScriptLineNumber & " " & $Running & @CRLF)
        Case $ButtonStatus
            GetPowerStatus()
        Case $ButtonToolsAbout
            About("Tools")
    EndSwitch

    If $Running Then
        RunLoop()
    EndIf
WEnd
;-----------------------------------------------
Func RunLoop()
    LogFile("RunLoop")
    If TestForSUT(' ') = False Then
        $Running = False
        Return
    EndIf
    If TestForPowerSwitch(' ') = False Then
        $Running = False
        Return
    EndIf

    If GUICtrlRead($CheckBoxCleanShutDown) = $GUI_CHECKED Then
        CleanShutDown()
        Wait(GUICtrlRead($InputAfterPowerOffDelay))
    EndIf
    PowerOffCommand()
    Wait(10)
    PingLoop(False)

    PowerOnCommand()
    Wait(10)
    If GUICtrlRead($CheckBoxCleanShutDown) = $GUI_CHECKED Then
        SendWOL()
    EndIf

    Wait(GUICtrlRead($InputAfterPowerOnDelay))
    PingLoop(True)
    $IterationCount = $IterationCount + 1
    If $IterationCount >= GUICtrlRead($InputMaxIterations) Then $Running = False
    If GUICtrlRead($CheckBoxAbort) = $GUI_CHECKED Then $Running = False
    Local $StatusString = "Iteration " & $IterationCount & " of " & GUICtrlRead($InputMaxIterations) & _
            " Passed: " & $PassCount & " Power on fails: " & $FailCountOn & " Power off fails: " & $FailCountOff
    Status($StatusString)
    LogFile($StatusString)
EndFunc   ;==>RunLoop
;-----------------------------------------------
; $Result is true or false
; $result false means we expected a response
; $result true means we did not
; Returns true if the result of pingSUT is as expected
Func PingLoop($result)
    LogFile("PingLoop")
    Local $PingRetry = 0
    While $PingRetry < GUICtrlRead($InputMaxPingRetrys)
        If PingSUT() = $result Then
            LogFile("PingLoop returned expected results: " & $result)
            $PassCount = $PassCount + 1
            Return True
        Else
            LogFile("PingLoop failed to returned expected results: " & $result)
            If $result Then
                $FailCountOff = $FailCountOff + 1
            Else
                $FailCountOn = $FailCountOn + 1
            EndIf
            If GUICtrlRead($CheckBoxStopOnFail) = $GUI_CHECKED Then
                LogFile("Stop on error activated")
                GUICtrlSetState($CheckBoxAbort, $GUI_CHECKED)
                $Running = False
            EndIf
            Return False
        EndIf
        $PingRetry = $PingRetry + 1
        Wait(GUICtrlRead($InputPingDelay))
        If GUICtrlRead($CheckBoxAbort) = $GUI_CHECKED Then ExitLoop
    WEnd
EndFunc   ;==>PingLoop
;-----------------------------------------------
; If SUT answers then return true else return false
Func PingSUT()
    Local $result = Ping(GUICtrlRead($InputSUTIPAddress), GUICtrlRead($InputPingTimeOut))
    If $result = 0 Then
        LogFile("No ping return " & GUICtrlRead($InputSUTIPAddress))
        Return False
    Else
        LogFile("Ping return OK. timeout (ms):" & $result)
        Return True
    EndIf
EndFunc   ;==>PingSUT
;-----------------------------------------------
; wait some number of seconds
Func Wait($Delay)
    LogFile("Wait: " & $Delay)
    Status("Wait:" & $Delay)
    Local $SavedTime = TimerInit()
    ConsoleWrite(@ScriptLineNumber & " Started delay: " & $Delay & @CRLF)
    Local $TimeDiff = 0
    While $TimeDiff < $Delay
        $TimeDiff = TimerDiff($SavedTime) / 1000 ; seconds
        Sleep(500)
        Status("*", True)
        If GUICtrlRead($CheckBoxAbort) = $GUI_CHECKED Then ExitLoop
    WEnd
    ConsoleWrite(@ScriptLineNumber & " Stopped " & TimerDiff($SavedTime) & @CRLF)
EndFunc   ;==>Wait
;-----------------------------------------------
;Powers on the SUT. Returns true if success
Func PowerOnCommand()
    LogFile("Run power on command")
    If GUICtrlRead($CheckBoxDebug) = $GUI_CHECKED Then
        LogFile("Run On DEBUG")
        Return
    EndIf
    _FFTabAdd(GUICtrlRead($InputSwitcherIPAddress) & "/Set.cmd?CMD=SetPower&P60=1&P61=1&P62=1&P63=1")
    LogFile(_FFReadText())
    _FFTabClose()

    Return True
EndFunc   ;==>PowerOnCommand
;-----------------------------------------------
; Attempts to power up SUT
Func SendWOL();Send a WOL packet
    LogFile("SendWOL")
    If $SUTMACAddress = '' Then
        $SUTMACAddress = InputBox("SendWOL", "Please enter the MAC address", "00:1b:21:88:57:b4")
    EndIf

    $cmd = $AuxPath & "mc-wol.exe"
    Local $A = StringSplit(GUICtrlRead($InputSUTIPAddress), ".")
    Local $address = $SUTMACAddress & " -a " & $A[1] & "." & $A[2] & ".255.255"
    Local $ReturnValue = ShellExecute($cmd, $address)
    LogFile("Command:" & $cmd & "  address:" & $address & "  RV:" & $ReturnValue & "  error:" & @error)
EndFunc   ;==>SendWOL
;-----------------------------------------------
; Shut SUT down cleanly
Func CleanShutDown()
    LogFile("Run clean shutdown")
    $cmd = $AuxPath & "plink.exe"
    Local $params = "root@" & GUICtrlRead($InputSUTIPAddress) & " -pw s init 0"
    ConsoleWrite(@ScriptLineNumber & " " & $cmd & " " & $params & @CRLF)
    ShellExecute($cmd, $params) ;, "open", @SW_SHOW)
    Status("Going down")
EndFunc   ;==>CleanShutDown
;-----------------------------------------------
;Powers off the SUT. (not a clean shutdown)
;Returns true if success
Func PowerOffCommand()
    LogFile("Run power off command")
    If GUICtrlRead($CheckBoxDebug) = $GUI_CHECKED Then
        LogFile("Run Off DEBUG")
        Return
    EndIf
    _FFTabAdd(GUICtrlRead($InputSwitcherIPAddress) & "/Set.cmd?CMD=SetPower&P60=0&P61=0&P62=0&P63=0")
    LogFile(_FFReadText())
    _FFTabClose()
    Return True
EndFunc   ;==>PowerOffCommand
;-----------------------------------------------
Func GetPowerStatus()
    LogFile("Get power status")
    _FFTabAdd(GUICtrlRead($InputSwitcherIPAddress) & "/Set.cmd?CMD=GetPower")
    ; Status(_FFReadText())
    LogFile(_FFReadText())
    _FFTabClose()
    Return True
EndFunc   ;==>GetPowerStatus
;-----------------------------------------------
Func LogOptions()
    LogFile("LogOptions")
    LogFile("Path to firefox: " & GUICtrlRead($InputFirefoxPath))
    LogFile("Path to options:" & GUICtrlRead($InputOptionFilePath))
    LogFile("SUT IP address: " & GUICtrlRead($InputSUTIPAddress))
    LogFile("Switcher IP address: " & GUICtrlRead($InputSwitcherIPAddress))
    LogFile("Max test iterations: " & GUICtrlRead($InputMaxIterations))
    LogFile("Ping delay: " & GUICtrlRead($InputPingDelay))
    LogFile("Ping timeout: " & GUICtrlRead($InputPingTimeOut))
    LogFile("After Power Off Delay: " & GUICtrlRead($InputAfterPowerOffDelay))
    LogFile("After Power On Delay: " & GUICtrlRead($InputAfterPowerOnDelay))
    LogFile("Max Ping Retrys: " & GUICtrlRead($InputMaxPingRetrys))
    LogFile("Stop on fail: " & GUICtrlRead($CheckBoxStopOnFail))
    LogFile("Clean shut down: " & GUICtrlRead($CheckBoxCleanShutDown))
    LogFile("Debug: " & GUICtrlRead($CheckBoxDebug))
EndFunc   ;==>LogOptions
;-----------------------------------------------
Func LoadDefaultOptions()
    LogFile("LoaddefaultOptions")
    GUICtrlSetData($InputSUTIPAddress, '191.0.0.100')
    GUICtrlSetData($InputSwitcherIPAddress, '192.168.10.100')
    GUICtrlSetData($InputMaxIterations, 10)
    GUICtrlSetData($InputPingDelay, 5)
    GUICtrlSetData($InputPingTimeOut, 1000)
    GUICtrlSetData($InputMaxPingRetrys, 100)
    GUICtrlSetData($InputAfterPowerOffDelay, 60)
    GUICtrlSetData($InputAfterPowerOnDelay, 120)
    GUICtrlSetData($InputOptionFilePath, "C:\Program Files (x86)\AutoIt3\Dougs\AUXFiles\Power_" & @ComputerName & ".opt")
    GUICtrlSetData($InputFirefoxPath, "C:\Program Files (x86)\Mozilla Firefox\firefox.exe")
    GUICtrlSetData($InputLogFilePath, "C:\Program Files (x86)\AutoIt3\Dougs\power.log")
    GUICtrlSetState($CheckBoxStopOnFail, $GUI_CHECKED)
    GUICtrlSetState($CheckBoxCleanShutDown, $GUI_CHECKED)
    GUICtrlSetState($CheckBoxDebug, $GUI_UNCHECKED)
EndFunc   ;==>LoadDefaultOptions
;-----------------------------------------------
;This saves an options file
Func SaveOptions()
    LogFile("SaveOptions")
    Local $SavePath = GUICtrlRead($InputOptionFilePath)
    Local $NewPath = FileSaveDialog("Save options file", $AuxPath, _
            $ProgramName & "Options (*.opt)|Power (Power.*)|All files (*.*)", 18, "Power-" & @ComputerName & ".opt")

    If @error = 0 Then
        GUICtrlSetData($InputOptionFilePath, $NewPath)
    Else
        GUICtrlSetData($InputOptionFilePath, $SavePath)
        Return
    EndIf

    Local $File = FileOpen(GUICtrlRead($InputOptionFilePath), 2)
    ; Check if file opened for writing OK
    If $File = -1 Then
        MsgBox(16, @ScriptLineNumber & " SaveOptions", "Unable to open file for writing: " & @CRLF & GUICtrlRead($InputOptionFilePath))
        LogFile("SaveOptions: Unable to open file for writing: " & GUICtrlRead($InputOptionFilePath))
        Return
    EndIf
    FileWriteLine($File, "Valid for " & $ProgramName & " options")
    FileWriteLine($File, "Options file for " & $ProgramName & "  " & _DateTimeFormat(_NowCalc(), 0))
    FileWriteLine($File, "Help 1 is enabled, 4 is disabled for checkboxes")
    FileWriteLine($File, "SUTIPAddress:" & GUICtrlRead($InputSUTIPAddress))
    FileWriteLine($File, "SwitcherIPAddress:" & GUICtrlRead($InputSwitcherIPAddress))
    FileWriteLine($File, "MaxIterations:" & GUICtrlRead($InputMaxIterations))
    FileWriteLine($File, "PingDelay:" & GUICtrlRead($InputPingDelay))
    FileWriteLine($File, "PingTimeOut:" & GUICtrlRead($InputPingTimeOut))
    FileWriteLine($File, "MaxPingRetrys:" & GUICtrlRead($InputMaxPingRetrys))
    FileWriteLine($File, "AfterPowerOffDelay:" & GUICtrlRead($InputAfterPowerOffDelay))
    FileWriteLine($File, "AfterPowerOnDelay:" & GUICtrlRead($InputAfterPowerOnDelay))
    FileWriteLine($File, "FirefoxPath:" & GUICtrlRead($InputFirefoxPath))
    FileWriteLine($File, "LogFilePath:" & GUICtrlRead($InputLogFilePath))
    FileWriteLine($File, "OptionFilePath:" & GUICtrlRead($InputOptionFilePath))
    FileWriteLine($File, "CheckBoxStopOnFail:" & GUICtrlRead($CheckBoxStopOnFail))
    FileWriteLine($File, "CheckBoxCleanShutDown:" & GUICtrlRead($CheckBoxCleanShutDown))
    FileWriteLine($File, "CheckBoxDebug:" & GUICtrlRead($CheckBoxDebug))

    FileClose($File)

EndFunc   ;==>SaveOptions
;-----------------------------------------------
;This loads an options file
Func LoadOptions()
    LogFile("LoadOptions")
    Local $SavePath = GUICtrlRead($InputOptionFilePath)
    Local $NewPath = FileOpenDialog("Load options file", $AuxPath, _
            $ProgramName & "Options (*.opt)|Power (Power.*)|All files (*.*)", 18, "Power-" & @ComputerName & ".opt")

    If @error = 0 Then
        GUICtrlSetData($InputOptionFilePath, $NewPath)
    Else
        GUICtrlSetData($InputOptionFilePath, $SavePath)
        Return
    EndIf

    Local $File = FileOpen(GUICtrlRead($InputOptionFilePath), 0)
    ; Check if file opened for reading OK
    If $File = -1 Then
        MsgBox(16, @ScriptLineNumber & " LoadOptions", "Unable to open file for reading: " & @CRLF & GUICtrlRead($InputOptionFilePath))
        LogFile("LoadOptions: Unable to open file for reading: " & GUICtrlRead($InputOptionFilePath))
        Return
    EndIf

    ; Read in the first line to verify the file is of the correct type
    If StringCompare(FileReadLine($File, 1), "Valid for Power options") <> 0 Then
        MsgBox(16, @ScriptLineNumber, "Not an options file for Power")
        FileClose($File)
        Return
    EndIf

    ; Read in lines of text until the EOF is reached
    While 1
        Local $LineIn = FileReadLine($File)
        If @error = -1 Then ExitLoop
        If StringInStr($LineIn, ";") = 1 Then ContinueLoop
        If StringInStr($LineIn, "SUTIPAddress:") Then GUICtrlSetData($InputSUTIPAddress, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "SwitcherIPAddress:") Then GUICtrlSetData($InputSwitcherIPAddress, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "MaxIterations:") Then GUICtrlSetData($InputMaxIterations, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "PingDelay:") Then GUICtrlSetData($InputPingDelay, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "PingTimeOut:") Then GUICtrlSetData($InputPingTimeOut, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "MaxPingRetrys:") Then GUICtrlSetData($InputMaxPingRetrys, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "AfterPowerOffDelay:") Then GUICtrlSetData($InputAfterPowerOffDelay, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "AfterPowerOnDelay:") Then GUICtrlSetData($InputAfterPowerOnDelay, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "FirefoxPath:") Then GUICtrlSetData($InputFirefoxPath, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "LogFilePath:") Then GUICtrlSetData($InputLogFilePath, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "OptionFilePath:") Then GUICtrlSetData($InputOptionFilePath, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckBoxStopOnFail:") Then GUICtrlSetState($CheckBoxStopOnFail, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckBoxCleanShutDown:") Then GUICtrlSetState($CheckBoxCleanShutDown, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckBoxDebug:") Then GUICtrlSetState($CheckBoxDebug, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
    WEnd
    FileClose($File)
EndFunc   ;==>LoadOptions
;-----------------------------------------------
Func GuiDisable($choice) ;$GUI_ENABLE $GUI_DISABLE
    For $X = 1 To 200
        GUICtrlSetState($X, $choice)
    Next
EndFunc   ;==>GuiDisable
;-----------------------------------------------
;Displays a status message on the gui. If append is true then $string is appended to the current message
Func Status($String, $Append = False)
    Local $T
    If $Append Then
        $T = GUICtrlRead($LabelStatusMain)
        If StringLen($T) > 80 Then $T = StringLeft($T, 10)
    EndIf
    GUICtrlSetData($LabelStatusMain, $T & $String)
    GUICtrlSetData($LabelStatusTools, $T & $String)
EndFunc   ;==>Status
;-----------------------------------------------
Func LogFile($String, $ShowMSGBox = False)
    ConsoleWrite(@ScriptLineNumber & " " & $String & @CRLF)
    ;GUICtrlSetData($LabelStatus, String)
    ;$String = StringReplace($String, "-1", "", 1)
    _FileWriteLog(GUICtrlRead($InputLogFilePath), $String)

    Local $X = _GUICtrlListBox_AddString($ListResults, $String)
    _GUICtrlListBox_SetTopIndex($ListResults, $X)
    Status($String)
    _Debug($String, $ShowMSGBox)
EndFunc   ;==>LogFile
;-----------------------------------------------
Func TestForPowerSwitch($type)
    Local $result = Ping(GUICtrlRead($InputSwitcherIPAddress), 1000)
    If $result = 0 Then
        LogFile("Power switch not detected: " & GUICtrlRead($InputSwitcherIPAddress))
        MsgBox(48, "TestForPowerSwitch", "Power switch not detected" & @CRLF & GUICtrlRead($InputSwitcherIPAddress))
        Return False
    Else
        LogFile("TestForPowerSwitch Ping return OK. timeout (ms):" & $result)
        If StringCompare($type, 'menu') = 0 Then

            $PowerMACAddress = GetMacAddress(GUICtrlRead($InputSwitcherIPAddress))
            ConsoleWrite(@ScriptLineNumber & " " & $PowerMACAddress & @CRLF)
            If StringLen($PowerMACAddress) = 0 Then
                MsgBox(48, "TestForPowerSwitch", "Unable to get PowerSwitch MAC address")
                Return False
            Else
                MsgBox(48, "TestForPowerSwitch", _
                        "Power switch detected. " & @CRLF & _
                        "IP: " & GUICtrlRead($InputSwitcherIPAddress) & @CRLF & _
                        "MAC adress: " & $PowerMACAddress)
            EndIf
        EndIf
    EndIf
    Return True
EndFunc   ;==>TestForPowerSwitch
;-----------------------------------------------
Func TestForSUT($type)
    GUICtrlSetData($LabelSUTInfoTools, '')
    Local $result = Ping(GUICtrlRead($InputSUTIPAddress), 1000)
    If $result = 0 Then
        LogFile("SUT not detected: " & GUICtrlRead($InputSUTIPAddress))
        MsgBox(48, "TestForSUT", "SUT not detected" & @CRLF & GUICtrlRead($InputSUTIPAddress))
        Return False
    Else
        LogFile("TestForSUT Ping return OK. timeout (ms):" & $result)
        If StringCompare($type, 'menu') = 0 Then
            $SUTMACAddress = GetMacAddress(GUICtrlRead($InputSUTIPAddress))
            If StringLen($SUTMACAddress) = 0 Then
                GUICtrlSetData($LabelSUTInfoTools, $SUTMACAddress)
                MsgBox(48, "TestForSUT", "Unable to get SUT MAC address")
                Return False
            Else
                GUICtrlSetData($LabelSUTInfoTools, "  SUT IP: " & GUICtrlRead($InputSUTIPAddress) & "  SUT MAC: " & $SUTMACAddress)
                MsgBox(48, "TestForSUT", _
                        "SUT detected. " & @CRLF & _
                        "IP: " & GUICtrlRead($InputSUTIPAddress) & @CRLF & _
                        "MAC adress: " & $SUTMACAddress)
            EndIf
        EndIf
    EndIf
    Return True
EndFunc   ;==>TestForSUT
;---------------------------------
Func VerifyPaths()
    Local $szDrive, $szDir, $szFName, $szExt

    If Not FileExists(GUICtrlRead($InputFirefoxPath)) Then
        LogFile("Failed VerifyPaths: " & GUICtrlRead($InputFirefoxPath))
        MsgBox(48, "Failed VerifyPaths: ", GUICtrlRead($InputFirefoxPath))
    EndIf

    _PathSplit(GUICtrlRead($InputLogFilePath), $szDrive, $szDir, $szFName, $szExt)
    If Not FileExists($szDrive & $szDir) Then
        LogFile("Failed VerifyPaths: " & $szDrive & $szDir)
        MsgBox(48, "Failed VerifyPaths: ", $szDrive & $szDir)
    EndIf

    _PathSplit(GUICtrlRead($InputLogFilePath), $szDrive, $szDir, $szFName, $szExt)
    If Not FileExists($szDrive & $szDir) Then
        LogFile("Failed VerifyPaths: " & $szDrive & $szDir)
        MsgBox(48, "Failed VerifyPaths: ", $szDrive & $szDir)
    EndIf

EndFunc   ;==>VerifyPaths
;-----------------------------------------------
Func GetMACAddress($IPAddressToTest)
    Local $tmpfile = "arp.txt"
    RunWait(@ComSpec & ' /c arp -a | find " ' & $IPAddressToTest & ' " > ' & $tmpfile, "", @SW_HIDE)

    Local $TA = StringSplit(StringStripWS(FileRead($tmpfile), 7), ' ')
    If FileGetSize($tmpfile) > 0 Then
        $TA[2] = StringReplace($TA[2], '-', ':') ; replace - with :
        LogFile("GetMACAddress: " & $IPAddressToTest & "  " & $TA[2])
        Return $TA[2]
    Else
        LogFile("GetMACAddress failed to get address: " & $IPAddressToTest)
        Return 0
    EndIf
EndFunc   ;==>GetMACAddress
;-----------------------------------------------
Func About(Const $FormID)
    LogFile("About: " & $FormID)
    GuiDisable($GUI_DISABLE)
    ConsoleWrite(@ScriptLineNumber & "  " & $FormID & @CRLF)
    ;LogFile(@ScriptLineNumber & " About " & $FormID)
    Local $d = WinGetPos($FormID)
    Local $WinPos
    If IsArray($d) = True Then
        ConsoleWrite(@ScriptLineNumber & " " & $FormID & @CRLF)
        $WinPos = StringFormat("%s" & @CRLF & "WinPOS: %d  %d " & @CRLF & "WinSize: %d %d " & @CRLF & "Desktop: %d %d ", _
                $FormID, $d[0], $d[1], $d[2], $d[3], @DesktopWidth, @DesktopHeight)
    Else
        $WinPos = ">>>About ERROR, Check the window name<<<"
    EndIf
    MsgBox(48, "About " & $FormID, $FormID & @CRLF & $SystemS & @CRLF & $WinPos & @CRLF & "Written by Doug Kaynor!")
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>About
;-----------------------------------------------
;This function starts firefox and logs in to the power control web page
Func StartFirefox()
    LogFile("Starting Firefox")
    ShellExecute(GUICtrlRead($InputFirefoxPath), "-height 500 -width 800")
    If _FFConnect(Default, Default, 3000) Then
        _FFAction("Min")
        _FFOpenURL(GUICtrlRead($InputSwitcherIPAddress))
        _FFClick("Submitbtn", "name", 0)
    EndIf
EndFunc   ;==>StartFirefox
;-----------------------------------------------