#region
#AutoIt3Wrapper_icon=../icons/freebsd.ico
#AutoIt3Wrapper_outfile=ShutdownTest.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Comment=GUI wrapper for ShutdownTest
#AutoIt3Wrapper_Res_Description=Wrapper Set up ShutdownTest
#AutoIt3Wrapper_Res_Fileversion=1.0.0.25
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
    Reboot the system
    verify data ping works
    pause or abort test
    
    Verify shutdowntest.exe exists before looping
    Stop on fail not implemented
    More inputs to verify
    verify connection speed (max possible)
    specify test time of number of iterations
    log all relevant data
#ce

_Debug('DBGVIEWCLEAR')
Opt('MustDeclareVars', 1)
;Opt('TrayMenuMode', 3)
;Opt('TrayOnEventMode', 1)
;Opt('TrayIconDebug', 1)

#include <ButtonConstants.au3>
#include <Constants.au3>
#include <Date.au3>
#include <GUIConstantsEx.au3>
#include <GuiListBox.au3>
#include <Misc.au3>
#include <StaticConstants.au3>
#include <String.au3>
#include <WindowsConstants.au3>
#include <_DougFunctions.au3>
#include <_intel.au3>

DirCreate('AUXFiles')
Global $AuxPath = @ScriptDir & '\AUXFiles\'

Global Const $ProgramName = 'ShutdownTest'
Global $Project_filename = @ScriptDir & '\AUXFiles\' & $ProgramName & '.prj'
Const $Log_filename = @ScriptDir & '\AUXFiles\' & $ProgramName & '.log'
Const $Loop_filename = @ScriptDir & '\AUXFiles\' & $ProgramName & '.loop'

If _Singleton($ProgramName, 1) = 0 Then
    MsgBox(48, 'Already running', $ProgramName & ' is already running!', 10)
    Exit
EndIf

Global Const $FileVersion = '  Ver: ' & FileGetVersion(@AutoItExe, 'Fileversion')
Global $SystemS = $ProgramName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & @OSServicePack & @CRLF & @OSType & @CRLF & @OSArch
Global $loops
Global $TestIPAddress
Global $ShutdownDelay
Global $PingTimeOut
Global $PingRetries
Global $ExpectedSpeed

Global $StopOnFail
Global $AutoClearLog
Global $DebugFlag
;----------------------------------------------
; Mainform
Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $MainForm = GUICreate($ProgramName & '  ' & $FileVersion, 500, 300, 400, 10, $MainFormOptions)
GUISetFont(9, 400, 0, 'Courier New')

#region ### START Koda GUI section ### Form=

Global $FileMenu = GUICtrlCreateMenu('File')
Global $ClearLogItem = GUICtrlCreateMenuItem('Clear log', $FileMenu)
Global $ClearDisplayItem = GUICtrlCreateMenuItem('Clear display', $FileMenu)
GUICtrlCreateMenuItem('', $FileMenu, 4)

Global $ViewLogItem = GUICtrlCreateMenuItem('View log', $FileMenu)
Global $ViewProjectItem = GUICtrlCreateMenuItem('View project', $FileMenu)
Global $ViewEditAFile = GUICtrlCreateMenuItem('View\Edit a file', $FileMenu)

GUICtrlCreateMenuItem('', $FileMenu, 6)
Global $InstallItem = GUICtrlCreateMenuItem('Install', $FileMenu)
Global $UnInstallItem = GUICtrlCreateMenuItem('Un-Install', $FileMenu)
GUICtrlCreateMenuItem('', $FileMenu, 8)
Global $ExitItem = GUICtrlCreateMenuItem('Exit', $FileMenu)

Global $OtherMenu = GUICtrlCreateMenu('Other')
Global $Pingitem = GUICtrlCreateMenuItem('Ping', $OtherMenu)
Global $NICInfoItem = GUICtrlCreateMenuItem('NIC Info', $OtherMenu)

Global $StartDebuggerItem = GUICtrlCreateMenuItem('Start Debugger', $OtherMenu)
Global $StartEventViewerItem = GUICtrlCreateMenuItem('Start Eventviewer', $OtherMenu)
Global $StartResourceMonitorItem = GUICtrlCreateMenuItem('Start Resource monitor', $OtherMenu)
Global $StartWiresharkItem = GUICtrlCreateMenuItem('Start Wireshark', $OtherMenu)
Global $StartOIDSItem = GUICtrlCreateMenuItem('Start OIDS', $OtherMenu)

Global $SettingsMenu = GUICtrlCreateMenu('Settings')
Global $ShowSettingsItem = GUICtrlCreateMenuItem('Show settings', $SettingsMenu)
Global $SaveSettingsItem = GUICtrlCreateMenuItem('Save settings', $SettingsMenu)
Global $LoadSettingsItem = GUICtrlCreateMenuItem('load settings', $SettingsMenu)
Global $DefaultSettingsItem = GUICtrlCreateMenuItem('Defaults settings', $SettingsMenu)

GUICtrlCreateMenuItem('', $SettingsMenu, 4)

Global $LoopsItem = GUICtrlCreateMenuItem('Loops', $SettingsMenu)
Global $TestIPAddresssItem = GUICtrlCreateMenuItem('Test IP address', $SettingsMenu)
Global $ShutdownDelayItem = GUICtrlCreateMenuItem('Shutdown delay', $SettingsMenu)
Global $PingTimeOutItem = GUICtrlCreateMenuItem('Ping timeout', $SettingsMenu)
Global $PingRetriesItem = GUICtrlCreateMenuItem('Ping retries', $SettingsMenu)
Global $ExpectedSpeedItem = GUICtrlCreateMenuItem('Expected speed', $SettingsMenu)
Global $StopOnFailItem = GUICtrlCreateMenuItem('Stop On Fail', $SettingsMenu)
Global $AutoClearLogItem = GUICtrlCreateMenuItem('Auto clear log', $SettingsMenu)
Global $DebugFlagItem = GUICtrlCreateMenuItem('Debug flag', $SettingsMenu)

Global $InfoMenu = GUICtrlCreateMenu('Info')
Global $AboutItem = GUICtrlCreateMenuItem('About', $InfoMenu)
Global $HelpItem = GUICtrlCreateMenuItem('Help', $InfoMenu)

;-----------------------------------------------
Global $ButtonStart = GUICtrlCreateButton('Start', 10, 5, 50, 25)
GUICtrlSetTip(-1, 'Start the test')
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $CheckAbort = GUICtrlCreateCheckbox('Abort', 70, 5, 50, 25)
GUICtrlSetTip(-1, 'Abort the test')
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ListResults = GUICtrlCreateList('', 10, 40, 480, 240, BitOR($WS_BORDER, $WS_VSCROLL))
GUICtrlSetData(-1, '')
GUICtrlSetResizing(-1, $GUI_DOCKBORDERS)

GUISetState(@SW_SHOW)
#endregion ### END Koda GUI section ###
DefaultSettings()

For $x = 1 To $CmdLine[0]
    logfile('CmdLine  ' & $x & ' >> ' & $CmdLine[$x])
    Select
        Case StringInStr($CmdLine[$x], 'help') > 0 Or StringInStr($CmdLine[$x], '?') > 0
            Help()
            Exit
        Case StringInStr($CmdLine[$x], 'run') > 0
            LogFile('Auto start detected')
            LoadSettings('loop')
            Start('loop')
        Case Else
            LogFile(' Unknown cmdline option found: ' & $CmdLine[$x] & '  ' & $x)
            MsgBox(16, @ScriptLineNumber, ' Unknown cmdline option found:' & @CRLF & $CmdLine[$x])
            Exit
    EndSelect
Next

LogFile('ShutdownTest started running.')

GUISetState(@SW_SHOW)

; network  A programmer is just a tool which converts caffeine into code
;-----------------------------------------------
Global $TS ; a temp string used by inputbox functions
While 1
    Switch GUIGetMsg()
        Case $ButtonStart
            LogFile('Start button detected')
            Start('menu')
            GuiDisable($GUI_ENABLE)
        Case $CheckAbort
            GuiDisable($GUI_ENABLE)
            UnInstallAutostart()
            ShellExecuteWait('shutdown.exe', '/a')
            LogFile('Check abort detected')

        Case $ShowSettingsItem
            ShowSettings()
        Case $SaveSettingsItem
            SaveSettings('menu')
        Case $LoadSettingsItem
            LoadSettings('menu')
        Case $DefaultSettingsItem
            DefaultSettings()

        Case $ClearLogItem
            FileDelete($Log_filename)

        Case $ViewLogItem
            ShellExecute(ChoseTextEditor(), $Log_filename, @SW_HIDE)
        Case $ViewProjectItem
            ShellExecute(ChoseTextEditor(), $Project_filename, @SW_HIDE)
        Case $ViewEditAFile
            ShellExecute(ChoseTextEditor(), '', @SW_HIDE)
        Case $ExitItem
            LogFile($ProgramName & ' ButtonExit')
            Exit

        Case $InstallItem
            InstallAutostart()
        Case $UnInstallItem
            UnInstallAutostart()
        Case $ClearDisplayItem
            _GUICtrlListBox_ResetContent($ListResults)
            GuiDisable($GUI_ENABLE)
        Case $Pingitem
            PingAdapter()
            GuiDisable($GUI_ENABLE)
        Case $NICInfoItem
            GetAdapterInfo(False)
            GuiDisable($GUI_ENABLE)
        Case $StartDebuggerItem
            ShellExecute($AuxPath & '\Dbgview.exe')
        Case $StartEventViewerItem
            ShellExecute('eventvwr.msc', '/s')
        Case $StartResourceMonitorItem
            ShellExecute('resmon.exe')
        Case $StartWiresharkItem
            LaunchWireShark()
        Case $StartOIDSItem
            LaunchOIDs()

            GuiDisable($GUI_ENABLE)
        Case $LoopsItem
            $TS = InputBox('Loops', 'Loops:', $loops, '', -1, 130)
            If @error = 0 Then $loops = $TS
        Case $TestIPAddresssItem
            $TS = InputBox('Test IP address', 'Test IP address:', $TestIPAddress, '', -1, 130)
            If @error = 0 Then $TestIPAddress = $TS
        Case $ShutdownDelayItem
            $TS = InputBox('Shutdown delay', 'Shutdown delay:', $ShutdownDelay, '', -1, 130)
            If @error = 0 Then $ShutdownDelay = $TS
        Case $PingTimeOutItem
            $TS = InputBox('Ping Timeout', 'Ping timeout:', $PingTimeOut, '', -1, 130)
            If @error = 0 Then $PingTimeOut = $TS
        Case $PingRetriesItem
            $TS = InputBox('Ping retries', 'Ping retries:', $PingRetries, '', -1, 130)
            If @error = 0 Then $PingRetries = $TS
        Case $ExpectedSpeedItem
            $TS = InputBox('Ping retries', 'Ping retries:', $ExpectedSpeed, '', -1, 130)
            If @error = 0 Then $ExpectedSpeed = $TS

        Case $StopOnFailItem
            If BitAND(GUICtrlRead($StopOnFailItem), $GUI_CHECKED) = $GUI_CHECKED Then
                GUICtrlSetState($StopOnFailItem, $GUI_UNCHECKED)
                $StopOnFail = False
            Else
                GUICtrlSetState($StopOnFailItem, $GUI_CHECKED)
                $StopOnFail = True
            EndIf
        Case $AutoClearLogItem
            If BitAND(GUICtrlRead($AutoClearLogItem), $GUI_CHECKED) = $GUI_CHECKED Then
                GUICtrlSetState($AutoClearLogItem, $GUI_UNCHECKED)
                $AutoClearLog = False
            Else
                GUICtrlSetState($AutoClearLogItem, $GUI_CHECKED)
                $AutoClearLog = True
                ConsoleWrite(@ScriptLineNumber & " " & $AutoClearLogItem & @CRLF)
            EndIf
        Case $DebugFlagItem
            If BitAND(GUICtrlRead($DebugFlagItem), $GUI_CHECKED) = $GUI_CHECKED Then
                GUICtrlSetState($DebugFlagItem, $GUI_UNCHECKED)
                $DebugFlag = False
            Else
                GUICtrlSetState($DebugFlagItem, $GUI_CHECKED)
                $DebugFlag = True
            EndIf

        Case $AboutItem
            About($ProgramName)
            GuiDisable($GUI_ENABLE)
        Case $HelpItem
            Help()
            GuiDisable($GUI_ENABLE)

        Case $GUI_EVENT_CLOSE
            Exit
    EndSwitch
WEnd
;-----------------------------------------------
Func VerifyInputs()
    GuiDisable($GUI_DISABLE)
    logfile('VerifyInputs')
    GUICtrlSetState($CheckAbort, $GUI_UNCHECKED)

    Local $Result = True
    If _TestIP($TestIPAddress) <> 'ERROR0' Then
        MsgBox(16, 'Bad test IP address', 'Bad test IP address: ' & $TestIPAddress)
        LogFile('Verify inputs failed. Bad test IP address:' & $TestIPAddress, False, True)
        $Result = False
    EndIf

    If Int($loops) < 0 Then
        MsgBox(16, 'Bad loop value', 'Bad loop value: ' & $loops)
        LogFile('Verify inputs failed. Bad loop value: >>' & $loops & '<<<')
        $Result = False
    EndIf

    Return $Result
EndFunc   ;==>VerifyInputs
;-----------------------------------------------
Func Start($type)
    ;If $AutoClearLog And StringInStr($type, 'menu') > 0 Then FileDelete($Log_filename)
    LogFile('Start detected ' & $type)
    _GUICtrlListBox_ResetContent($ListResults)
    If Not VerifyInputs() Then Return

    If Not FileExists('ShutdownTest.exe') Then
        MsgBox(48, 'ShutdownTest.exe missing', 'ShutdownTest.exe missing' & @CRLF & 'Test can not loop')
        UnInstallAutostart()
        GUICtrlSetState($CheckAbort, $GUI_CHECKED)
        Return
    EndIf

    If StringInStr($type, 'menu') > 0 Then
        IniWrite($Loop_filename, 'All', 'Counter', 0)
        SaveSettings('loop')
        ShowSettings()
        GetAdapterInfo(False)
        InstallAutostart()
    EndIf
    ;LogFile('Start ' & )

    If PingAdapter() Then
        Local $CurrentLoop = IniRead($Loop_filename, 'All', 'Counter', '0')
        ;The following is the actual test loop
        If $CurrentLoop < $loops Then
            LogFile('Looping. Loop ' & $CurrentLoop & ' of ' & $loops)
            $CurrentLoop = $CurrentLoop + 1
            IniWrite($Loop_filename, 'All', 'Counter', $CurrentLoop)

            LogFile('Start current loop: ' & $CurrentLoop & '  Max loops: ' & $loops)
            If Not $DebugFlag Then
                ShellExecuteWait('shutdown.exe', '/r /t ' & $ShutdownDelay)
                MsgBox(16, 'Wait for shutdown', 'Wait for shutdown ' & $ShutdownDelay)
            Else
                LogFile('Debug mode set. Shutdown skipped')
            EndIf
            GetAdapterInfo(True)
            PingAdapter()
            If LogFile('Start detected ' & $type) = $GUI_CHECKED Then Return
        EndIf
        LogFile('Looping complete:' & $CurrentLoop & ' Max loops:' & $loops)
    Else
        _GUICtrlListBox_AddString($ListResults, 'Initial ping failed  ' & $TestIPAddress)
        LogFile('Initiall ping failed')
        MsgBox(16, 'Initial ping failed', 'Initial ping failed')
        GUICtrlSetState($CheckAbort, $GUI_CHECKED)
    EndIf

EndFunc   ;==>Start
;-----------------------------------------------

Func PingAdapter()
    If Not VerifyInputs() Then Return
    Const $SleepTime = 500

    Local $Retries = 0
    Local $Response = 99
    While $Retries < $PingRetries
        $Response = Ping($TestIPAddress, $PingTimeOut)
        Local $ResultString = 'Ping (' & $TestIPAddress & ') response (ms):' & $Response & ' retries: ' & $Retries
        _GUICtrlListBox_AddString($ListResults, $ResultString)
        LogFile($ResultString)
        If $Response > 0 Then ExitLoop
        $Retries = $Retries + 1
        If GUICtrlRead($CheckAbort) = $GUI_CHECKED Then ExitLoop
        Sleep($SleepTime)
    WEnd

    If $Retries >= $PingRetries Then
        GUICtrlSetState($CheckAbort, $GUI_CHECKED)
        LogFile('Ping failed. Retries: ' & $Retries)
        Return False
    Else
        LogFile('Ping passed. Retries: ' & $Retries)
        Return True
    EndIf
EndFunc   ;==>PingAdapter
;-----------------------------------------------
Func GetAdapterInfo($SpeedOnly = False)
    If Not VerifyInputs() Then Return

    Const $wbemFlagReturnImmediately = 0x10
    Const $wbemFlagForwardOnly = 0x20

    Local $objWMI = ObjGet('winmgmts:\\localhost\root\CIMV2')
    Local $objItems = $objWMI.ExecQuery('SELECT * FROM Win32_NetworkAdapter', 'WQL', $wbemFlagReturnImmediately + $wbemFlagForwardOnly)

    For $objItem In $objItems
        If GUICtrlRead($CheckAbort) = $GUI_CHECKED Then Return
        If $objItem.NetEnabled = True And StringInStr($objItem.AdapterType, 'Ethernet 802.3') > 0 Then
            If StringInStr($objItem.Name, 'Intel') > 0 And Not $SpeedOnly Then
                _GUICtrlListBox_AddString($ListResults, 'Name: ' & $objItem.Name)
                _GUICtrlListBox_AddString($ListResults, 'Manufacturer: ' & $objItem.Manufacturer)
                _GUICtrlListBox_AddString($ListResults, 'Description: ' & $objItem.Description)
                _GUICtrlListBox_AddString($ListResults, 'MACAddress: ' & $objItem.MACAddress)
                _GUICtrlListBox_AddString($ListResults, 'PowerManagementSupported: ' & $objItem.PowerManagementSupported)
                _GUICtrlListBox_AddString($ListResults, 'Speed: ' & $objItem.Speed / 1e6 & 'Mbs')
                _GUICtrlListBox_AddString($ListResults, 'NetConnectionStatus: ' & NetConnectionStatus($objItem.NetConnectionStatus))
                _GUICtrlListBox_AddString($ListResults, 'NetEnabled: ' & $objItem.NetEnabled)
                _GUICtrlListBox_AddString($ListResults, 'Status: ' & $objItem.Status)
                _GUICtrlListBox_AddString($ListResults, 'StatusInfo: ' & $objItem.StatusInfo)
                _GUICtrlListBox_AddString($ListResults, 'AdapterType: ' & $objItem.AdapterType)
                _GUICtrlListBox_AddString($ListResults, 'Availability: ' & NetConnectionAvailabilty($objItem.Availability))
                _GUICtrlListBox_AddString($ListResults, @CRLF)
            EndIf
        EndIf
    Next


    If $ExpectedSpeed <> $objItem.Speed Then
        Logfile('Expected speed mismatch. Expected:' & $ExpectedSpeed & '  Actual: ' & $objItem.Speed, False, True)
    EndIf

    LogFile($objItem.Name & '  ' & 'Speed: ' & $objItem.Speed / 1e6 & 'Mbs' & '  ' & NetConnectionStatus($objItem.NetConnectionStatus))

EndFunc   ;==>GetAdapterInfo

;-----------------------------------------------

Func About(Const $FormID)
    Local $D = GetWinPos($FormID)
    Local $WinPos
    If IsArray($D) = True Then
        ConsoleWrite(@ScriptLineNumber & $FormID & @CRLF)
        $WinPos = StringFormat('%s' & @CRLF & 'WinPOS: %d  %d ' & @CRLF & 'WinSize: %d %d ' & @CRLF & 'Desktop: %d %d ', _
                $FormID, $D[0], $D[1], $D[2], $D[3], @DesktopWidth, @DesktopHeight)
    Else
        $WinPos = '>>>About ERROR, Check the window name<<<'
    EndIf
    MsgBox(64, 'About', $SystemS & @CRLF & $WinPos & @CRLF & 'Written by Doug Kaynor!')
EndFunc   ;==>About

;-----------------------------------------------
Func Help()
    Local $helpstr = 'Startup options: ' & @CRLF & _
            'help or ?   Display this help file' & @CRLF & _
            'Run         Start checking for link on startup'

    MsgBox(16, @ScriptName & $FileVersion, $helpstr)
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
        _Debug(ConsoleWrite('Filewrite error: ' & $G & '  ' & _ArrayToString($F) & @CRLF))
    WEnd
    Return ($F)
EndFunc   ;==>GetWinPos
;-----------------------------------------------
Func NetConnectionStatus($value)
    Switch $value
        Case 0
            Return 'Disconnected'
        Case 1
            Return 'Connecting'
        Case 2
            Return 'Connected'
        Case 3
            Return 'Disconnecting'
        Case 4
            Return 'Hardware Not present'
        Case 5
            Return 'Hardware disabled'
        Case 6
            Return 'Hardware malfunction'
        Case 7
            Return 'Media disconnected'
        Case 8
            Return 'Authenticating'
        Case 9
            Return 'Authentication succeeded'
        Case 10
            Return ' Authentication failed'
        Case 11
            Return 'Invalid address'
        Case 12
            Return 'Credentials required'
    EndSwitch
EndFunc   ;==>NetConnectionStatus
;-----------------------------------------------
Func NetConnectionAvailabilty($value)
    Switch $value
        Case 1
            Return 'Other'
        Case 2
            Return 'Unknown'
        Case 3
            Return 'Running Or Full Power'
        Case 4
            Return 'Warning'
        Case 5
            Return 'In Test'
        Case 6
            Return 'Not Applicable'
        Case 7
            Return 'Power Off'
        Case 8
            Return 'Off Line'
        Case 9
            Return 'Off Duty'
        Case 10
            Return 'Degraded'
        Case 11
            Return 'Not Installed'
        Case 12
            Return 'Install Error'
        Case 13
            Return 'Power Save - Unknown'
            ; The device is known To be in a power save state, but its exact status is unknown.
        Case 14
            Return 'Power Save - Low Power Mode'
            ; The device is in a power save state, but still functioning, and may exhibit degraded performance.
        Case 15
            Return 'Power Save - Standby'
            ;The device is not functioning, but could be brought To full power quickly.
        Case 16
            Return 'Power Cycle'
        Case 17
            Return 'Power Save - Warning'
            ;The device is in a warning state, though also in a power save state.
    EndSwitch
EndFunc   ;==>NetConnectionAvailabilty
;-----------------------------------------------
Func GuiDisable($choice) ;$GUI_ENABLE $GUI_DISABLE
    For $x = 1 To 100
        GUICtrlSetState($x, $choice)
    Next
    GUICtrlSetState($CheckAbort, $GUI_ENABLE)

EndFunc   ;==>GuiDisable
;-----------------------------------------------
Func LogFile($string, $Add2List = False, $ShowMSGBox = False)
    $string = StringReplace($string, '-1', '', 1)
    _FileWriteLog($Log_filename, $string)
    _Debug($string, $ShowMSGBox)
    If $Add2List Then _GUICtrlListBox_AddString($ListResults, 'Log: ' & $Add2List)
EndFunc   ;==>LogFile
;-----------------------------------------------
Func InstallAutostart()
    FileCreateShortcut(@ScriptDir & '\' & $ProgramName & '.exe', _
            @StartupCommonDir & '\' & $ProgramName & '.lnk', _
            @ScriptDir, _
            'run', _
            $ProgramName, _
            @ScriptDir & '\..\icons\freebsd.ico')
EndFunc   ;==>InstallAutostart
;-----------------------------------------------
Func UnInstallAutostart()
    FileDelete(@StartupCommonDir & '\' & $ProgramName & '.lnk')
EndFunc   ;==>UnInstallAutostart

;-----------------------------------------------
Func ShowSettings()
    _GUICtrlListBox_ResetContent($ListResults)
    _GUICtrlListBox_AddString($ListResults, 'Test IP address: ' & $TestIPAddress)
    _GUICtrlListBox_AddString($ListResults, 'Loops: ' & $loops)
    _GUICtrlListBox_AddString($ListResults, 'Shutdown delay: ' & $ShutdownDelay)
    _GUICtrlListBox_AddString($ListResults, 'Ping timeout: ' & $PingTimeOut)
    _GUICtrlListBox_AddString($ListResults, 'Ping retries: ' & $PingRetries)
    _GUICtrlListBox_AddString($ListResults, 'Expected speed: ' & $ExpectedSpeed / 1e6 & 'Mbs')
    _GUICtrlListBox_AddString($ListResults, 'Stop on fail: ' & TrueFalse($StopOnFail))
    _GUICtrlListBox_AddString($ListResults, 'Auto clear log: ' & TrueFalse($AutoClearLog))
    _GUICtrlListBox_AddString($ListResults, 'Debug: ' & TrueFalse($DebugFlag))
    _GUICtrlListBox_AddString($ListResults, 'Abort: ' & ShowState(GUICtrlRead($CheckAbort)))
EndFunc   ;==>ShowSettings
;-----------------------------------------------
Func DefaultSettings()
    $TestIPAddress = '191.2.0.105'
    $loops = 5
    $ShutdownDelay = 10
    $PingTimeOut = 4000
    $PingRetries = 5
    $ExpectedSpeed = 1e9

    GUICtrlSetState($StopOnFailItem, $GUI_CHECKED)
    $StopOnFail = True
    GUICtrlSetState($AutoClearLogItem, $GUI_CHECKED)
    $AutoClearLog = True
    GUICtrlSetState($DebugFlagItem, $GUI_UNCHECKED)
    $DebugFlag = False

EndFunc   ;==>DefaultSettings
;-----------------------------------------------
Func SaveSettings($type = 'loop')
    If StringCompare($type, 'menu') = 0 Then
        $Project_filename = FileSaveDialog('Save options file', $AuxPath, _
                $ProgramName & 'Options (*.opt)|ShutdownTest (ShutdownTest.*)|All files (*.*)', 18, 'ShutdownTest-' & '.opt')
    EndIf

    Local $File = FileOpen($Project_filename, 2)
    ; Check if file opened for writing OK
    If $File = -1 Then
        LogFile(@ScriptLineNumber & ' SaveOptions: Unable to open file for writing: ' & $Project_filename)
        LogFile('SaveOptions: Unable to open file for writing: ' & $Project_filename, True)
        Return
    EndIf

    FileWriteLine($File, 'Valid for ' & $ProgramName & ' options')
    FileWriteLine($File, 'Options file for ' & $Project_filename & '  ' & _DateTimeFormat(_NowCalc(), 0))
    FileWriteLine($File, 'Help 1 is enabled, 4 is disabled for checkboxes')
    FileWriteLine($File, 'TestIPAddress:' & $TestIPAddress)
    FileWriteLine($File, 'Loops:' & $loops)
    FileWriteLine($File, 'ShutdownDelay:' & $ShutdownDelay)
    FileWriteLine($File, 'PingTimeOut:' & $PingTimeOut)
    FileWriteLine($File, 'Pingretries:' & $PingRetries)
    FileWriteLine($File, 'ExpectedSpeed:' & $ExpectedSpeed)

    FileWriteLine($File, 'StopOnFail:' & $StopOnFail)
    FileWriteLine($File, 'AutoClearLog:' & $AutoClearLog)
    FileWriteLine($File, 'DebugFlag:' & $DebugFlag)
    FileClose($File)
EndFunc   ;==>SaveSettings
;-----------------------------------------------
;This loads the options file
Func LoadSettings($type = 'loop')

    logfile('LoadSettings start ' & $type)
    If StringCompare($type, 'menu') = 0 Then
        $Project_filename = FileOpenDialog('Load settings file', $AuxPath, _
                $ProgramName & 'Options (*.opt)|ShutdownTest (ShutdownTest.*)|All files (*.*)', 18, 'ShutdownTest-' & '.opt')
    EndIf

    Local $File = FileOpen($Project_filename, 0)
    ; Check if file opened for reading OK
    If $File = -1 Then
        LogFile('LoadSettings: Unable to open file for reading: ' & $Project_filename)
        MsgBox(16, $Project_filename, 'LoadSettings: Unable to open file for reading: ' & @CRLF & $Project_filename)
        Return
    EndIf

    LogFile('LoadOptions ' & $Project_filename)
    ; Read in the first line to verify the file is of the correct type
    If StringCompare(FileReadLine($File, 1), 'Valid for ShutdownTest options') <> 0 Then
        LogFile('Not an options file for ShutdownTest' & @CRLF & $Project_filename)
        FileClose($File)
        GuiDisable($GUI_ENABLE)
        Return
    EndIf
    ; Read in lines of text until the EOF is reached
    While 1
        Local $LineIn = FileReadLine($File)
        If @error = -1 Then ExitLoop
        If StringInStr($LineIn, ';') = 1 Then ContinueLoop

        If StringInStr($LineIn, 'TestIPAddress:') Then $TestIPAddress = StringMid($LineIn, StringInStr($LineIn, ':') + 1)
        If StringInStr($LineIn, 'Loops:') Then $loops = StringMid($LineIn, StringInStr($LineIn, ':') + 1)
        If StringInStr($LineIn, 'ShutdownDelay:') Then $ShutdownDelay = StringMid($LineIn, StringInStr($LineIn, ':') + 1)
        If StringInStr($LineIn, 'PingTimeOut:') Then $PingTimeOut = StringMid($LineIn, StringInStr($LineIn, ':') + 1)
        If StringInStr($LineIn, 'Pingretries:') Then $PingRetries = StringMid($LineIn, StringInStr($LineIn, ':') + 1)
        If StringInStr($LineIn, 'ExpectedSpeed:') Then $ExpectedSpeed = StringMid($LineIn, StringInStr($LineIn, ':') + 1)
        #cs
            If StringInStr($LineIn, 'StopOnFail:') Then $StopOnFail = StringMid($LineIn, StringInStr($LineIn, ':') + 1)
            If $StopOnFail Then GUICtrlSetState($StopOnFailItem, $GUI_CHECKED)
            If StringInStr($LineIn, 'AutoCkearLog:') Then $AutoClearLog = StringMid($LineIn, StringInStr($LineIn, ':') + 1)
            If $AutoClearLog Then GUICtrlSetState($AutoClearLogItem, $GUI_CHECKED)
            If StringInStr($LineIn, 'DebugFlag:') Then $DebugFlag = StringMid($LineIn, StringInStr($LineIn, ':') + 1)
            If $DebugFlag Then GUICtrlSetState($DebugFlagItem, $GUI_CHECKED)
        #ce
    WEnd
    FileClose($File)
    logfile('LoadSettings complete')
EndFunc   ;==>LoadSettings
;-----------------------------------------------
