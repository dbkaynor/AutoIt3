#region
#AutoIt3Wrapper_icon=../icons/freebsd.ico
#AutoIt3Wrapper_outfile=NTTTPWin.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Comment=GUI wrapper for NTTTPWin
#AutoIt3Wrapper_Res_Description=Wrapper Set up NTTTPWin
#AutoIt3Wrapper_Res_Fileversion=1.0.0.15
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

#include <misc.au3>
#include <file.au3>
#include <array.au3>
#include <WindowsConstants.au3>
#include <StaticConstants.au3>
#include <Process.au3>
#include <GuiListBox.au3>
#include <GuiListView.au3>
#include <GUIConstantsEx.au3>
#include <GUIConstants.au3>
#include <EditConstants.au3>
#include <Constants.au3>
#include <ButtonConstants.au3>
#include <_DougFunctions.au3>
#include <_Intel.au3>

Opt("MustDeclareVars", 1)
Opt("TrayIconDebug", 1)
Opt("TrayAutoPause", 0)
Opt("WinTitleMatchMode", 2)

Global Const $ProgramName = "NTTTPWin"
Global $AUXPath = @ScriptDir & "\AUXFiles\"
Global $UtilPath = @ScriptDir & "\AUXFiles\Utils\"
Global $Project_filename = $AUXPath & $ProgramName & ".prj"
Global $LOG_filename = $AUXPath & $ProgramName & ".log"

If _Singleton($ProgramName, 1) = 0 Then
    MsgBox(48, "Already running", $ProgramName & " is already running!", 10)
    Exit
EndIf

Global Const $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")
Global $SystemS = $ProgramName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & @OSServicePack & @CRLF & @OSType & @CRLF & @OSArch

Global $logfileName = $ProgramName & ".log"
Global $datafileName = $ProgramName & ".dat"

Global $NTTTPWinR
Global $NTTTPWinS
Global $WireSharkFolder
If @OSArch = "X86" Then
    $NTTTPWinR = FileGetShortName($UtilPath & "NTTTCP 3.0\IA32\NTTTCPR.exe")
    $NTTTPWinS = FileGetShortName($UtilPath & "NTTTCP 3.0\IA32\NTTTCPS.exe")
    $WireSharkFolder = 'C:\Program Files (x86)\wireshark\'
ElseIf "X64" Then
    $NTTTPWinR = FileGetShortName($UtilPath & "NTTTCP 3.0\IA32e\NTTTCPR.exe")
    $NTTTPWinS = FileGetShortName($UtilPath & "NTTTCP 3.0\IA32e\NTTTCPS.exe")
    $WireSharkFolder = 'C:\Program Files\wireshark\'
ElseIf "IA64" Then
    $NTTTPWinR = FileGetShortName($UtilPath & "NTTTCP 3.0\IA64\NTTTCPR.exe")
    $NTTTPWinS = FileGetShortName($UtilPath & "NTTTCP 3.0\IA64\NTTTCPS.exe")
    $WireSharkFolder = 'C:\Program Files\wireshark\'
EndIf

If Not FileExists($NTTTPWinR) Then
    MsgBox(16, "Missing file", $NTTTPWinR & " not found." & @CRLF & "Must be located in the" & $UtilPath & "NTTTCP 3.0 folder")
    Exit
EndIf
_debug(@ScriptLineNumber & " " & $NTTTPWinR)

If Not FileExists($NTTTPWinS) Then
    MsgBox(16, "Missing file", $NTTTPWinS & " not found." & @CRLF & "Must be located in the " & $UtilPath & "NTTTCP 3.0 folder")
    Exit
EndIf
_debug(@ScriptLineNumber & " " & $NTTTPWinS)

#region ### START Koda GUI section ### Form=C:\Program Files (x86)\AutoIt3\Dougs\NTTTPWin.kxf
Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, _
        $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $MainForm = GUICreate($ProgramName & "  " & $FileVersion, 750, 325, 200, 200, $MainFormOptions)
GUISetFont(9, 400, 0, "Courier New")

Global $ToolsMenu = GUICtrlCreateMenu('Tools')
Global $OIDsItem = GUICtrlCreateMenuItem('OIDs', $ToolsMenu)
Global $WireSharkItem = GUICtrlCreateMenuItem('WireShark', $ToolsMenu)
Global $EventMonitorItem = GUICtrlCreateMenuItem('Event Monitor', $ToolsMenu)
Global $EventViewerItem = GUICtrlCreateMenuItem('Event Viewer', $ToolsMenu)
Global $PerformanceMonitorItem = GUICtrlCreateMenuItem('Performance Monitor', $ToolsMenu)
Global $ResourceMonitorItem = GUICtrlCreateMenuItem('Resource Monitor', $ToolsMenu)
Global $DeviceManagerItem = GUICtrlCreateMenuItem('Device Manager', $ToolsMenu)
Global $GroupPolicyEditorItem = GUICtrlCreateMenuItem('Group Policy Editor', $ToolsMenu)
Global $ServicesItem = GUICtrlCreateMenuItem('Services', $ToolsMenu)

;Global $PerformanceMonitorItem = GUICtrlCreateMenuItem('Performance Monitor', $ToolsMenu)

GUICtrlCreateLabel("Length of buffer:", 20, 10, 125, 20)
Global $InputLengthOfBuffer = GUICtrlCreateInput("", 150, 10, 50, 20)
GUICtrlCreateLabel("Number of buffers:", 10, 35, 130, 20)
Global $InputNumberOfBuffers = GUICtrlCreateInput("", 150, 35, 50, 20)
GUICtrlCreateLabel("Port base:", 65, 60, 75, 20)
Global $InputPortBase = GUICtrlCreateInput("", 150, 60, 50, 20)
GUICtrlCreateLabel("Outstanding I/O:", 25, 85, 115, 20)
Global $InputOutstandingIO = GUICtrlCreateInput("", 150, 85, 50, 20)
GUICtrlCreateLabel("Packet Array size:", 220, 10, 140, 20)
Global $InputPacketArraySize = GUICtrlCreateInput("", 355, 10, 50, 20)
GUICtrlCreateLabel("Receive buffer size:", 200, 35, 140, 20)
Global $InputRecieveBufferSize = GUICtrlCreateInput("", 355, 35, 50, 20)
GUICtrlCreateLabel("Send buffer size:", 225, 60, 125, 20)
Global $InputSendBufferSize = GUICtrlCreateInput("0", 355, 60, 50, 20)
GUICtrlCreateLabel("Run time:", 280, 85, 70, 20)
Global $InputRunTime = GUICtrlCreateInput("0", 360, 85, 50, 20)

GUICtrlCreateLabel("Output file name:", 10, 110, 140, 20)
Global $InputFileName = GUICtrlCreateInput("", 140, 110, 265, 20)

GUICtrlCreateLabel("IP address:", 10, 135, 85, 20)
Global $InputIPAddress = GUICtrlCreateInput('', 100, 135, 100, 20)

Global $CheckUseDefaults = GUICtrlCreateCheckbox("Use defaults", 220, 135, 100, 20)
Global $CheckTest = GUICtrlCreateCheckbox("Test", 340, 135, 50, 20)

GUICtrlCreateLabel("Processor:", 10, 160, 75, 20)
Global $InputProcessor = GUICtrlCreateInput('', 100, 160, 50, 20)
GUICtrlCreateLabel("Sessions:", 175, 160, 70, 20)
Global $InputSessions = GUICtrlCreateInput('', 245, 160, 50, 20)
Global $CheckInfiniteLoop = GUICtrlCreateCheckbox("Infinite loop", 420, 10, 120, 20)
Global $CheckUDPSendRecieve = GUICtrlCreateCheckbox("UDP send/recieve", 420, 30, 135, 20)
Global $CheckWSARecvWSASend = GUICtrlCreateCheckbox("WSARecv/WSASend", 420, 50, 120, 20)
Global $CheckVerifyFlag = GUICtrlCreateCheckbox("Verify flag", 420, 70, 120, 20)
Global $CheckVerboseMode = GUICtrlCreateCheckbox("Verbose mode", 420, 90, 120, 20)
Global $CheckEnableIPV6Mode = GUICtrlCreateCheckbox("Enable IPV6 mode", 420, 110, 145, 20)
Global $CheckFullBuffersPostedOnReads = GUICtrlCreateCheckbox("Full buffers posted on reads", 420, 130, 220, 20)
Global $CheckMultipleBufferPostMode = GUICtrlCreateCheckbox("Multiple buffer post mode", 420, 150, 215, 20)
Global $CheckSendTest = GUICtrlCreateCheckbox("Send test", 420, 170, 215, 20)

Global $ButtonStart = GUICtrlCreateButton("Start", 640, 10, 100, 20)
GUICtrlSetTip($ButtonStart, "Start the program")
Global $ButtonAbout = GUICtrlCreateButton("About", 640, 30, 100, 20)
GUICtrlSetTip($ButtonAbout, "About the program")
Global $ButtonExit = GUICtrlCreateButton("Exit", 640, 50, 100, 20)
GUICtrlSetTip($ButtonExit, "Exit the program")

Global $ListStatus = GUICtrlCreateList("Status", 10, 200, 730, 100, BitOR($WS_BORDER, $WS_VSCROLL))
GUICtrlSetTip($ListStatus, "Left click to clear status")

For $x = 0 To 100
    GUICtrlSetResizing($x, $GUI_DOCKALL)
Next
GUICtrlSetResizing($ListStatus, BitOR($GUI_DOCKTOP, $GUI_DOCKBOTTOM))

SetDefaults()
GUISetState(@SW_SHOW)
#endregion ### END Koda GUI section ###

;-----------------------------------------------
While 1
    Switch GUIGetMsg()
        ;--- Menus
        Case $OIDsItem
            OIDs()
        Case $WireSharkItem
            WireShark()
        Case $EventMonitorItem
            EventMonitor()
        Case $EventViewerItem
            Eventviewer()
        Case $ResourceMonitorItem
            ResourceMonitor()
        Case $PerformanceMonitorItem
            PerformanceMonitor()
        Case $DeviceManagerItem
            DeviceManager()
        Case $GroupPolicyEditorItem
            GroupPolicyEditor()
        Case $ServicesItem
            Services()
            ;--- Buttons
        Case $ButtonStart
            Doit()
        Case $ButtonAbout
            About($ProgramName)
            ;--- Other
        Case $ListStatus
            _GUICtrlListBox_ResetContent($ListStatus)
        Case $ButtonExit
            ExitLoop
        Case $GUI_EVENT_CLOSE
            Exit
    EndSwitch
WEnd
;-----------------------------------------------
Func SetDefaults()
    _Debug($ProgramName)
    Global $F = WinGetPos($ProgramName, "")
    If $F[0] > @DesktopWidth Or $F[1] > @DesktopHeight Then WinMove($ProgramName, "", 10, 10, 775, 325)
    If $F[0] < 0 Or $F[1] < 0 Then WinMove($ProgramName, "", 10, 10, 775, 325)
    GUICtrlSetState($CheckInfiniteLoop, $GUI_UNCHECKED)
    GUICtrlSetState($CheckUDPSendRecieve, $GUI_UNCHECKED)
    GUICtrlSetState($CheckWSARecvWSASend, $GUI_UNCHECKED)
    GUICtrlSetState($CheckVerifyFlag, $GUI_UNCHECKED)
    GUICtrlSetState($CheckVerboseMode, $GUI_UNCHECKED)
    GUICtrlSetState($CheckMultipleBufferPostMode, $GUI_UNCHECKED)
    GUICtrlSetState($CheckVerifyFlag, $GUI_UNCHECKED)
    GUICtrlSetState($CheckMultipleBufferPostMode, $GUI_UNCHECKED)
    GUICtrlSetState($CheckEnableIPV6Mode, $GUI_UNCHECKED)
    GUICtrlSetState($CheckFullBuffersPostedOnReads, $GUI_UNCHECKED)
    GUICtrlSetState($CheckUseDefaults, $GUI_CHECKED)

    ; GUICtrlSetState($CheckIPv6, $GUI_UNCHECKED)
    GUICtrlSetData($InputIPAddress, '1.0.0.1')
    GUICtrlSetData($InputLengthOfBuffer, 65536)
    GUICtrlSetData($InputNumberOfBuffers, 20480)
    GUICtrlSetData($InputPortBase, 5001)
    GUICtrlSetData($InputOutstandingIO, 2)
    GUICtrlSetData($InputPacketArraySize, 1)
    GUICtrlSetData($InputRecieveBufferSize, 65536)
    GUICtrlSetData($InputSendBufferSize, 0)
    GUICtrlSetData($InputRunTime, 100)
    GUICtrlSetData($InputProcessor, 0)
    GUICtrlSetData($InputSessions, 4)
    GUICtrlSetData($InputFileName, 'output.txt')
EndFunc   ;==>SetDefaults
;-----------------------------------------------

#cs
    NTTTCPR.exe: [-l|-n|-p|-a|-x|-rb|-sb|-i|-f|-u|-w|-d|-t|-v|-6|-fr|-mb] -m <mapping> [mapping]
    ]

    -l   <Length of buffer>     [default:  64K]
    -n   <Number of buffers>    [default:  20K]
    -p   <Port base>            [default: 5001]
    -a   [outstanding I/O]      [default:    2]
    -x   [PacketArray size]	    [default:    1]
    -rb  <Receive buffer size>  [default:  64K]
    -sb  <Send buffer size>     [default:    0]
    -i   Infinite Loop          [Only UDP mode]

    -f   <File Name>            [default: output.txt]
    -u   UDP send/recv
    -w   WSARecv/WSASend
    -d   Verify Flag
    -t   <Runtime> in seconds
    -v   enable verbose mode
    -6   enable IPv6 mode
    -fr  Full buffers posted on reads
    -mb  Multiple buffer post mode
    -m   <mapping> [mapping]
    where a mapping is a session(s),processor,receiver IP set
    e.g. -m 4,0,1.2.3.4 sets up:
    4 sessions on processor 0 to test a network on 1.2.3.4
#ce

;-----------------------------------------------
Func DoIt()
    GuiDisable($GUI_DISABLE)
    Local $CommandString = " -m " & Int(GUICtrlRead($InputSessions)) & ',' & Int(GUICtrlRead($InputProcessor)) & ',' & GUICtrlRead($InputIPAddress)
    If GUICtrlRead($CheckUDPSendRecieve) = $GUI_CHECKED Then $CommandString = $CommandString & " -u "
    If GUICtrlRead($CheckWSARecvWSASend) = $GUI_CHECKED Then $CommandString = $CommandString & " -w "
    If GUICtrlRead($CheckVerifyFlag) = $GUI_CHECKED Then $CommandString = $CommandString & " -d "
    If GUICtrlRead($CheckVerboseMode) = $GUI_CHECKED Then $CommandString = $CommandString & " -v "
    If GUICtrlRead($CheckEnableIPV6Mode) = $GUI_CHECKED Then $CommandString = $CommandString & " -6 "
    If GUICtrlRead($CheckFullBuffersPostedOnReads) = $GUI_CHECKED Then $CommandString = $CommandString & " -fr "
    If GUICtrlRead($CheckMultipleBufferPostMode) = $GUI_CHECKED Then $CommandString = $CommandString & " -mb "
    If GUICtrlRead($CheckInfiniteLoop) = $GUI_CHECKED Then $CommandString = $CommandString & " -i "

    $CommandString = $CommandString & " -t " & Int(GUICtrlRead($InputRunTime))
    If GUICtrlRead($CheckUseDefaults) = $GUI_UNCHECKED Then
        $CommandString = $CommandString & " -l " & Int(GUICtrlRead($InputLengthOfBuffer))
        $CommandString = $CommandString & " -n " & Int(GUICtrlRead($InputNumberOfBuffers))
        $CommandString = $CommandString & " -p " & Int(GUICtrlRead($InputPortBase))
        $CommandString = $CommandString & " -a " & Int(GUICtrlRead($InputOutstandingIO))
        $CommandString = $CommandString & " -x " & Int(GUICtrlRead($InputPacketArraySize))
        $CommandString = $CommandString & " -rb " & Int(GUICtrlRead($InputRecieveBufferSize))
        $CommandString = $CommandString & " -sb " & Int(GUICtrlRead($InputSendBufferSize))
        $CommandString = $CommandString & " -f " & GUICtrlRead($InputFileName)
    EndIf

    Local $Result = 0
    Local $NTTTPWinString
    If GUICtrlRead($CheckSendTest) = $GUI_CHECKED Then
        $NTTTPWinString = $NTTTPWinS
    Else
        $NTTTPWinString = $NTTTPWinR
    EndIf

    GUICtrlSetData($ListStatus, $NTTTPWinString & " " & $CommandString) ; & $Result
    If GUICtrlRead($CheckTest) = $GUI_UNCHECKED Then $Result = RunWait($NTTTPWinString & " " & $CommandString)
    GUICtrlSetData($ListStatus, "Result: " & $Result)

    _Debug(@ScriptLineNumber & " NTTTPWinString: " & $NTTTPWinString)
    _Debug(@ScriptLineNumber & " CommandString: " & $CommandString)
    _Debug(@ScriptLineNumber & " Result: " & $Result)
    GuiDisable($GUI_ENABLE)
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
    For $x = 1 To 100
        GUICtrlSetState($x, $setting)
    Next

EndFunc   ;==>GuiDisable
;-----------------------------------------------
Func About(Const $FormID)
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
EndFunc   ;==>About
;-----------------------------------------------
#cs

    NTTTCPR.exe: [-l|-n|-p|-a|-x|-rb|-sb|-i|-f|-u|-w|-d|-t|-v|-6|-fr|-mb] -m <mapping> [mapping]

    -l   <Length of buffer>     [default:  64K]
    -n   <Number of buffers>    [default:  20K]
    -p   <Port base>            [default: 5001]
    -a   [outstanding I/O]      [default:    2]
    -x   [PacketArray size]	    [default:    1]
    -rb  <Receive buffer size>  [default:  64K]
    -sb  <Send buffer size>     [default:    0]
    -i   Infinite Loop          [Only UDP mode]

    -f   <File Name>            [default: output.txt]
    -u   UDP send/recv
    -w   WSARecv/WSASend
    -d   Verify Flag
    -t   <Runtime> in seconds
    -v   enable verbose mode
    -6   enable IPv6 mode
    -fr  Full buffers posted on reads
    -mb  Multiple buffer post mode
    -m   <mapping> [mapping]

    where a mapping is a session(s),processor,receiver IP set
    e.g. -m 4,0,1.2.3.4 sets up:
    4 sessions on processor 0 to test a network on 1.2.3.4
#ce






