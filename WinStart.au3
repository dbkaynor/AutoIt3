#region
#AutoIt3Wrapper_icon=../icons/Sri_Aravan.ico
#AutoIt3Wrapper_outfile=WinStart.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Comment=GUI wrapper for WinStart
#AutoIt3Wrapper_Res_Description=Wrapper Set up WinStart
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


#ce

_Debug("DBGVIEWCLEAR")
Opt("MustDeclareVars", 1)
;Opt("TrayMenuMode", 3)
;Opt("TrayOnEventMode", 1)
;Opt("TrayIconDebug", 1)

#include <ButtonConstants.au3>
#include <Constants.au3>
#include <GUIConstantsEx.au3>
#include <Misc.au3>
#include <StaticConstants.au3>
#include <String.au3>
#include <WindowsConstants.au3>
#include <_DougFunctions.au3>
#include <_Intel.au3>


Global Const $ProgramName = "WinStart"
Global $Project_filename = $AuxPath & $ProgramName & ".prj"
Const $Log_filename = $AuxPath & $ProgramName & ".log"
Const $Loop_filename = $AuxPath & $ProgramName & ".lop"
If _Singleton($ProgramName, 1) = 0 Then
    MsgBox(48, "Already running", $ProgramName & " is already running!", 10)
    Exit
EndIf

Global Const $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")
Global $SystemS = $ProgramName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & @OSServicePack & @CRLF & @OSType & @CRLF & @OSArch

; Mainform
Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $MainForm = GUICreate($ProgramName & "  " & $FileVersion, 430, 200, 10, 10, $MainFormOptions)
GUISetFont(9, 400, 0, "Courier New")

;menu items -----------------------------------------------
Global $PostInstallMenu = GUICtrlCreateMenu('Post install')
Global $MostPostItem = GUICtrlCreateMenuItem('Run most post installs ', $PostInstallMenu)
Global $RegistryAddsItem = GUICtrlCreateMenuItem('Registry adds', $PostInstallMenu)
Global $MakeShortcutsItem = GUICtrlCreateMenuItem('Make Shortcuts', $PostInstallMenu)
Global $DisableFirewallItem = GUICtrlCreateMenuItem('Disable Firewall', $PostInstallMenu)
GUICtrlCreateMenuItem('', $PostInstallMenu)
Global $InstallBGInfoItem = GUICtrlCreateMenuItem('Install BG Info', $PostInstallMenu)

Global $AccessoriesMenu = GUICtrlCreateMenu('Accessories')
Global $EditorItem = GUICtrlCreateMenuItem('Editor', $AccessoriesMenu)
Global $DebugerItem = GUICtrlCreateMenuItem('Debugger', $AccessoriesMenu)

Global $TestsMenu = GUICtrlCreateMenu('Tests')
Global $URW4WinItem = GUICtrlCreateMenuItem("URW4Win", $TestsMenu)
Global $NTTTPWinItem = GUICtrlCreateMenuItem("NTTTPWin", $TestsMenu)

Global $ToolsMenu = GUICtrlCreateMenu('Tools')
Global $DeviceManagerItem = GUICtrlCreateMenuItem('Device Manager', $ToolsMenu)
Global $EventMonitorItem = GUICtrlCreateMenuItem('Event Monitor', $ToolsMenu)
Global $EventViewerItem = GUICtrlCreateMenuItem('Event Viewer', $ToolsMenu)
Global $FirewallItem = GUICtrlCreateMenuItem('Firewall', $ToolsMenu)
Global $GroupPolicyEditorItem = GUICtrlCreateMenuItem('Group Policy Editor', $ToolsMenu)
Global $NetworkCPLItem = GUICtrlCreateMenuItem('Network Control panel', $ToolsMenu)
Global $NetworkSharingItem = GUICtrlCreateMenuItem('Network sharing', $ToolsMenu)
Global $NICViewerItem = GUICtrlCreateMenuItem('NIC viewer', $ToolsMenu)
Global $OIDsItem = GUICtrlCreateMenuItem('OIDs', $ToolsMenu)
Global $PerformanceMonitorItem = GUICtrlCreateMenuItem('Performance Monitor', $ToolsMenu)
Global $ResourceMonitorItem = GUICtrlCreateMenuItem('Resource Monitor', $ToolsMenu)
Global $ServicesItem = GUICtrlCreateMenuItem('Services', $ToolsMenu)
Global $WiresharkItem = GUICtrlCreateMenuItem('WireShark', $ToolsMenu)

Global $AboutMenu = GUICtrlCreateMenu('About')
Global $AboutItem = GUICtrlCreateMenuItem("About", $AboutMenu)
Global $HelpItem = GUICtrlCreateMenuItem("Help", $AboutMenu)
;Buttons -----------------------------------------------

Global $ButtonEventMonitor = GUICtrlCreateButton("Event monitor", 15, 10, 120, 25)
Global $ButtonControlPanel = GUICtrlCreateButton("Control panel", 15, 35, 120, 25)

Global $ButtonResorceMonitor = GUICtrlCreateButton("Resource mon", 160, 10, 120, 25)
Global $ButtonPreformanceMonitor = GUICtrlCreateButton('Performance Mon', 160, 35, 120, 25)

Global $ButtonExit = GUICtrlCreateButton("Exit", 300, 10, 100, 25)
Global $ButtonGUIEnable = GUICtrlCreateButton("GUIEnable", 300, 35, 100, 25)

Global $EditStatusMain = GUICtrlCreateEdit("Status", 10, 70, 400, 100, $WS_HSCROLL)
;-----------------------------------------------

_CheckWindowLocation($MainForm, 'n')
GUISetState(@SW_SHOW)

; network  A programmer is just a tool which converts caffeine into code
;-----------------------------------------------
Global $TS ; a temp string used by inputbox functions
While 1
    Switch GUIGetMsg()
        ;--- Menus
        Case $OIDsItem
            OIDs()
        Case $WiresharkItem
            WireShark()
        Case $EventMonitorItem
            EventMonitor()
        Case $EventViewerItem
            Eventviewer()
        Case $ResourceMonitorItem
            ResourceMonitor()
        Case $NetworkCPLItem
            NetworkControlPanel()
        Case $NetworkSharingItem
            NetworkSharing()
        Case $PerformanceMonitorItem
            PerformanceMonitor()
        Case $DeviceManagerItem
            DeviceManager()
        Case $GroupPolicyEditorItem
            GroupPolicyEditor()
        Case $ServicesItem
            Services()
        Case $FirewallItem
            FirewallControlPanel()
        Case $MostPostItem
            DisableFirewall()
            MakeShortcuts()
            ; InstallBGInfo()
            ShellExecute($UtilPath & 'All.reg')

        Case $DisableFirewallItem
            DisableFirewall()
        Case $RegistryAddsItem
            GUICtrlSetData($EditStatusMain, 'Install some registry entries')
            ShellExecute($UtilPath & 'All.reg')
        Case $MakeShortcutsItem
            MakeShortcuts()
        Case $InstallBGInfoItem
            InstallBGInfo()
        Case $EditorItem
            GUICtrlSetData($EditStatusMain, 'Launch text file editor')
            $Editor = _ChoseTextEditor()
            ShellExecute($Editor, '', $AuxPath)
        Case $DebugerItem
            GUICtrlSetData($EditStatusMain, 'Start debugview')
            ShellExecute($UtilPath & 'Dbgview.exe')
        Case $URW4WinItem
            GUICtrlSetData($EditStatusMain, 'Start URW4WinW')
            ShellExecute('URW4WinW.exe')
        Case $NTTTPWinItem
            GUICtrlSetData($EditStatusMain, 'Start NTTTPWin')
            ShellExecute('NTTTPWin.exe')
        Case $NICViewerItem
            NICViewer()
        Case $AboutItem
            GUICtrlSetData($EditStatusMain, 'About')
            About($ProgramName)
        Case $HelpItem
            GUICtrlSetData($EditStatusMain, 'Help')
            ShellExecute($AuxPath & 'WinStart.htm')
            ;Help()

            ;--- Buttons
        Case $ButtonResorceMonitor
            GUICtrlSetData($EditStatusMain, 'Start resource monitor')
            ResourceMonitor()
        Case $ButtonPreformanceMonitor
            GUICtrlSetData($EditStatusMain, 'Start performance monitor')
            PerformanceMonitor()
        Case $ButtonEventMonitor
            GUICtrlSetData($EditStatusMain, 'Start event monitor')
            EventMonitor()
        Case $ButtonControlPanel
            GUICtrlSetData($EditStatusMain, 'Start control panel')
            ShellExecute('control.exe')
        Case $ButtonGUIEnable
            GUICtrlSetData($EditStatusMain, 'GUI enabled')
            GuiDisable($GUI_ENABLE)
        Case $ButtonExit
            Exit
        Case $GUI_EVENT_CLOSE
            Exit
    EndSwitch
WEnd
;-----------------------------------------------
Func InstallBGInfo()
    GUICtrlSetData($EditStatusMain, 'BGInfo create shortcut')
    If FileCreateShortcut($UtilPath & 'bginfo.exe', _
            @StartupCommonDir & '\bginfo.lnk', _
            $UtilPath, _
            'bginfo.bgi /timer=2') = 0 Then
        MsgBox(16, 'bginfo create shortcut error', _
                @StartupCommonDir & "\" & 'bginfo.lnk' & _
                $UtilPath & 'bginfo.exe' & @CRLF & _
                $UtilPath & 'bginfo.bgi /timer=1')
    EndIf
EndFunc   ;==>InstallBGInfo
;-----------------------------------------------
Func MakeShortcuts()
    GUICtrlSetData($EditStatusMain, 'Make shortcuts')
    FileCreateShortcut(@WindowsDir & '\system32\cmd.exe', @DesktopDir & '\cmd.lnk', 'c:\')
    FileCreateShortcut(@WindowsDir & '\system32\secpol.msc', @DesktopDir & '\Security Policy.lnk')
    FileCreateShortcut(@WindowsDir & '\system32\gpedit.msc', @DesktopDir & '\Group policy editor.lnk')
    FileCreateShortcut(@WindowsDir & '\system32\devmgmt.msc', @DesktopDir & '\DevMgmt.lnk')
    FileCreateShortcut(@WindowsDir & '\system32\perfmon.msc', @DesktopDir & '\Performance monitor.lnk')
    FileCreateShortcut(@WindowsDir & '\system32\resource.msc', @DesktopDir & '\Resource monitor.lnk')
    FileCreateShortcut('ncpa.cpl', @DesktopDir & '\Network CPL.lnk')
    FileCreateShortcut(@WindowsDir & '\explorer.exe', @DesktopDir & '\Network Sharing Center.lnk', @WindowsDir, _
            'shell:::{8E908FC9-BECC-40f6-915B-F4CA0E70D03D}')
    FileCreateShortcut('firewall.cpl', @DesktopDir & '\Firewall.lnk') ;firewall
EndFunc   ;==>MakeShortcuts
;-----------------------------------------------
Func DisableFirewall()
    ShellExecute("netsh", "advfirewall set currentprofile state off")
    ShellExecute("Netsh", "advfirewall set domainprofile state off")
    ShellExecute("netsh", "advfirewall set privateprofile state off")
EndFunc   ;==>DisableFirewall
;-----------------------------------------------
Func About(Const $FormID)
    GuiDisable($GUI_DISABLE)
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
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>About

;-----------------------------------------------
Func Help()
    GuiDisable($GUI_DISABLE)
    Local $helpstr = $ProgramName & ' options: ' & @CRLF & _
            'help or ?   Display this help file'

    MsgBox(16, @ScriptName & $FileVersion, $helpstr)
    GuiDisable($GUI_ENABLE)
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
Func GuiDisable($choice) ;$GUI_ENABLE $GUI_DISABLE
    For $x = 1 To 100
        GUICtrlSetState($x, $choice)
    Next
    GUICtrlSetState($ButtonGUIEnable, $GUI_ENABLE)
EndFunc   ;==>GuiDisable
;-----------------------------------------------

