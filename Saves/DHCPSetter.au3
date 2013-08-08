#region
#AutoIt3Wrapper_icon=../icons/Cockroach.ico
#AutoIt3Wrapper_outfile=DHCPSetter.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Comment=Set up DHCP for LOL
#AutoIt3Wrapper_Res_Description=Set up DHCP for LOL
#AutoIt3Wrapper_Res_Fileversion=1.0.0.63
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
;Opt("WinTitleMatchMode", 3)

Global $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")
Global $tmp = StringSplit(@ScriptName, ".")
Global Const $ProgramName = $tmp[1]
;Global $logfile = $ProgramName & ".txt"

Global $ProcessedLeaseArray[1]

If _Singleton($ProgramName, 1) = 0 Then
    MsgBox(32, "Already running", $ProgramName & " is already running!")
    Exit
EndIf

Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $MainForm = GUICreate($ProgramName & "  " & $FileVersion, 700, 400, 20, 20, $MainFormOptions)
GUISetFont(9, 400, 0, "Courier New")

Global $ButtonScopeStart = GUICtrlCreateButton("Scopes", 10, 10, 75, 25)
GUICtrlSetTip(-1, "Start setting the scopes")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ButtonLeaseStart = GUICtrlCreateButton("Leases", 10, 40, 75, 25)
GUICtrlSetTip(-1, "Fetch leases")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ButtonReserveRemove = GUICtrlCreateButton("Remove", 10, 70, 75, 25)
GUICtrlSetTip(-1, "Remove reserved addresses")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ButtonReserveStart = GUICtrlCreateButton("Reserves", 10, 100, 75, 25)
GUICtrlSetTip(-1, "Set reserved addresses")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

GUICtrlCreateLabel("Cabinet", 100, 10, 60, 25)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $InputCabinetNumber = GUICtrlCreateInput("2", 180, 10, 30, 25)
GUICtrlSetTip(-1, "Enter a cabinet Number")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

GUICtrlCreateLabel("Cabinet IP", 100, 40, 84, 17)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $InputCabinetIP = GUICtrlCreateInput("172.31.1.252", 180, 40, 120, 25)
GUICtrlSetTip(-1, "Enter a cabinet IP")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

GUICtrlCreateLabel("First scope", 320, 10, 84, 25)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $InputFirstScope = GUICtrlCreateInput("12", 400, 10, 30, 25)
GUICtrlSetTip(-1, "Enter first scope")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

GUICtrlCreateLabel("Last scope", 320, 40, 76, 17)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $InputLastScope = GUICtrlCreateInput("15", 400, 40, 30, 25)
GUICtrlSetTip(-1, "Enter last scope")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $LabelStatus = GUICtrlCreateLabel("? Status ?", 100, 70, 80, 17);, $SS_BLACKFRAME)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $CheckTest = GUICtrlCreateCheckbox("Check Test", 180, 70, 100, 20)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetTip(-1, "Don't actually write the results")

Global $CheckAbort = GUICtrlCreateCheckbox("Check Abort", 180, 100, 100, 20)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetTip(-1, "Abort the current operation")

Global $ButtonTest = GUICtrlCreateButton("Test", 500, 10, 75, 20)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetTip(-1, "Set selections for test mode")

Global $ButtonShowDHCPWindow = GUICtrlCreateButton("Show DHCP", 500, 35, 75, 20)
GUICtrlSetTip(-1, "Show DHCP Window")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ButtonDefaults = GUICtrlCreateButton("Defaults", 600, 10, 75, 20)
GUICtrlSetTip(-1, "Set selections to defaults")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ButtonDump = GUICtrlCreateButton("Dump", 600, 35, 75, 20)
GUICtrlSetTip(-1, "Save results to file")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ButtonExit = GUICtrlCreateButton("Exit", 600, 60, 75, 20)
GUICtrlSetTip(-1, "Exit the program")
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Const $Data1 = 0
Const $Data2 = 1
Const $Data3 = 2
Global $ListViewResults = GUICtrlCreateListView("Data1|Data2|Data3|Data4", 5, 130, 670, 250, $WS_VSCROLL, BitOR($WS_EX_CLIENTEDGE, $LVS_EX_GRIDLINES))
GUICtrlSetResizing(-1, $GUI_DOCKBORDERS)
GUICtrlSetTip(-1, "Display completed work")
_GUICtrlListView_SetBkColor(-1, $CLR_WHITE)
_GUICtrlListView_SetTextColor(-1, $CLR_BLACK)

GUISetState(@SW_SHOW)

Global $F = WinGetPos("DHCPSetter", "")
If $F[0] > @DesktopWidth Or $F[1] > @DesktopHeight Then WinMove("DHCPSetter", "", 10, 10, 630, 300)
If $F[0] < 0 Or $F[1] < 0 Then WinMove("DHCPSetter", "", 10, 10, 630, 300)

GUISetState()
;If FileExists($logfile) Then FileDelete($logfile)
;-----------------------------------------------
While 1
    Switch GUIGetMsg()
        Case $ButtonScopeStart
            ScopeStart()
        Case $ButtonLeaseStart
            LeaseStart()
        Case $ButtonReserveStart
            ReserveStart()
        Case $ButtonReserveRemove
            ReserveRemove()
        Case $ButtonTest
            GUICtrlSetData($InputCabinetNumber, "2")
            GUICtrlSetData($InputCabinetIP, "172.31.1.252")
            GUICtrlSetData($InputFirstScope, "9")
            GUICtrlSetData($InputLastScope, "11")
        Case $ButtonShowDHCPWindow
            ShowDHCPWindow()
        Case $ButtonDefaults
            GUICtrlSetData($InputCabinetNumber, "2")
            GUICtrlSetData($InputCabinetIP, "172.31.1.252")
            GUICtrlSetData($InputFirstScope, "1")
            GUICtrlSetData($InputLastScope, "99")
        Case $ButtonDump
            ConsoleWrite(@ScriptLineNumber & " " & GUICtrlRead(GUICtrlGetHandle($ListViewResults)) & @CRLF)
        Case $ButtonExit
            ExitLoop
        Case $GUI_EVENT_CLOSE
            Exit
    EndSwitch
WEnd
;-----------------------------------------------
; Launch the DHCP console window
Func ShowDHCPWindow()
    ConsoleWrite(@ScriptLineNumber & " " & ShellExecuteWait(@ScriptDir & "\test.bat") & "  " & @error & @CRLF)
    WinWaitActive("DHCP", "", 65)
    WinActivate("DHCP", "")
    WinMove("DHCP", "", 0, 0)
EndFunc   ;==>ShowDHCPWindow
;-----------------------------------------------
; verify that the various user input parameters make sense. Return 0 if all is well
Func VerifyParameters()
    GUICtrlSetState($CheckAbort, $GUI_UNCHECKED)
    GUICtrlSetData($LabelStatus, "VerifyParameters")
    Local $ip = StringSplit(GUICtrlRead($InputCabinetIP), ".")

    If $ip[0] <> 4 Then
        MsgBox(16, "IP address error", "IP address error" & @CRLF & GUICtrlRead($InputCabinetIP) & @CRLF & $ip[0])
        Return 1
    EndIf

    ; verify that the cabinet number is within a resonable range
    If GUICtrlRead($InputCabinetNumber) < 1 Or GUICtrlRead($InputCabinetNumber) > 8 Then
        MsgBox(16, "Cabinet value error", "Cabinet value error" & @CRLF & GUICtrlRead($InputCabinetNumber) & @CRLF & "Must be between 1 and 8 inclusive")
        Return 2
    EndIf

    Local $A = Number(GUICtrlRead($InputFirstScope))
    Local $B = Number(GUICtrlRead($InputLastScope))

    ConsoleWrite(@ScriptLineNumber & " >>" & $A & "<<>>" & $B & "<<" & @CRLF)

    If $A < 0 Or $B > 255 Or ($A > $B) Then
        MsgBox(16, "Scope range error", "Scope range error" & @CRLF & $A & "  " & $B)
        Return 3
    EndIf

    Return 0
EndFunc   ;==>VerifyParameters
;-----------------------------------------------
Func ReserveRemove()
    If VerifyParameters() <> 0 Then Return
    If UBound($ProcessedLeaseArray) < 2 Then
        MsgBox(16, "Nothing to reserve", "Has leases been run?")
        Return
    EndIf
    GUICtrlSetData($LabelStatus, "ReservesRemove")
    GuiDisable("Disable")

    _GUICtrlListView_DeleteAllItems(GUICtrlGetHandle($ListViewResults))

    Local $Command = "netsh.exe"
    Local $commandString = ''
    Local $Y = 0

    ;netsh dhcp server 172.31.0.0 scope 192.168.1.0 delete reservedip 192.168.1.189 080027C591F9
    _GUICtrlListView_SetColumn($ListViewResults, 0, "Command string", 650)
    ;_ArrayDisplay($ProcessedLeaseArray, @ScriptLineNumber)
    For $W In $ProcessedLeaseArray
        $Y = $Y + 1
        Local $SplitArray = StringSplit($W, "~~~", 3)

        ConsoleWrite(@ScriptLineNumber & ": " & _ArrayToString($SplitArray) & @CRLF)
        Local $ScopeArray = StringSplit($SplitArray[0], ".", 2)
        Local $Scope = $ScopeArray[0] & "." & $ScopeArray[1] & ".0.0"

        ;netsh dhcp server 172.31.0.0 scope 192.168.1.0 delete reservedip 192.168.1.189 00027C591F9

        $commandString = " dhcp server " & GUICtrlRead($InputCabinetIP) & " scope " & $Scope & _
                " delete reservedip " & $SplitArray[0] & " " & StringReplace($SplitArray[1], "-", "")
        _GUICtrlListView_AddItem($ListViewResults, $commandString) ; display data
        GUICtrlSetData($LabelStatus, $Y & " 0")

        Local $Result = 0
        If GUICtrlRead($CheckTest) = $GUI_UNCHECKED Then
            $Result = Run(@ComSpec & " /c " & $Command & " " & $commandString, " .", @SW_HIDE, $STDOUT_CHILD)
        EndIf
        If $Result = 0 Then _GUICtrlListView_AddItem($ListViewResults, "ERROR  add reservedip " & $Result & " " & @error)

        GUICtrlSetData($LabelStatus, $Y & " 1")
    Next


    ;netsh dhcp server 172.31.0.0 scope 192.168.1.0 add reservedip 192.168.1.181 080027C591F9
    GuiDisable("Enable")
EndFunc   ;==>ReserveRemove
;-----------------------------------------------
;Using the lease information, assign reserved addresses
Func ReserveStart()
    If VerifyParameters() <> 0 Then Return
    If UBound($ProcessedLeaseArray) < 2 Then
        MsgBox(16, "Nothing to reserve", "Has leases been run?")
        Return
    EndIf
    GUICtrlSetData($LabelStatus, "ReserveStart")
    GuiDisable("Disable")

    _GUICtrlListView_DeleteAllItems(GUICtrlGetHandle($ListViewResults))

    Local $Command = "netsh.exe"
    Local $commandString = ''
    Local $Y = 0

    _GUICtrlListView_SetColumn($ListViewResults, 0, "Command string", 650)
    ;_ArrayDisplay($ProcessedLeaseArray, @ScriptLineNumber)

    ;netsh dhcp server 172.31.0.0 scope 192.168.1.0 delete reservedip 192.168.1.189 080027C591F9
    ;netsh dhcp server 172.31.0.252 scope 2.15.0.0 add reservedip 2.15.0.7 080027C591F9 XVM1.WS2K8.DEV

    For $W In $ProcessedLeaseArray
        $Y = $Y + 1
        Local $SplitArray = StringSplit($W, "~~~", 3)
        Local $ScopeArray = StringSplit($SplitArray[0], ".", 2)

        ;ConsoleWrite(@ScriptLineNumber & ": " & _ArrayToString($SplitArray) & @CRLF)
        ;ConsoleWrite(@ScriptLineNumber & ": " & _ArrayToString($ScopeArray) & @CRLF)

        Local $Scope = $ScopeArray[0] & "." & $ScopeArray[1] & ".0.0"

        ;ConsoleWrite(@ScriptLineNumber & " " & $Scope & @CRLF)
        ;ConsoleWrite(@ScriptLineNumber & ": " & _ArrayToString($ScopeArray) & @CRLF)

        ;here the bits get moved around to accomadate the addressing scheme
        Local $K = StringSplit($SplitArray[0], ".")
        Local $L = StringSplit($SplitArray[2], "RC")

        ;$k[2] and $k[4] need to be tweaked
        ConsoleWrite(@ScriptLineNumber & " " & $K[2] & " " & $K[4] & @CRLF)

        ;$M is the final address
        Local $M = $K[1] & "." & $K[2] & "." & $L[3] & "." & StringLeft($K[2], 1)

        $commandString = " dhcp server " & GUICtrlRead($InputCabinetIP) & " scope " & $Scope & _
                " add reservedip " & $M & " " & StringReplace($SplitArray[1], "-", "") & " " & $SplitArray[2]

        _GUICtrlListView_AddItem($ListViewResults, $commandString) ; display data
        GUICtrlSetData($LabelStatus, $Y & " 0")
        Local $Result = 0
        If GUICtrlRead($CheckTest) = $GUI_UNCHECKED Then
            $Result = Run(@ComSpec & " /c " & $Command & " " & $commandString, " .", @SW_HIDE, $STDOUT_CHILD)
        EndIf
        If $Result = 0 Then _GUICtrlListView_AddItem($ListViewResults, "ERROR  add reservedip " & $Result & " " & @error)

        GUICtrlSetData($LabelStatus, $Y & " 1")
    Next

    ;netsh dhcp server 172.31.0.0 scope 192.168.1.0 delete reservedip 192.168.1.189 080027C591F9
    ;netsh dhcp server 172.31.0.0 scope 192.168.1.0 add reservedip 192.168.1.181 080027C591F9
    GuiDisable("Enable")
EndFunc   ;==>ReserveStart
;-----------------------------------------------
;Fetch the current leases
Func LeaseStart()
    If VerifyParameters() <> 0 Then Return
    GUICtrlSetData($LabelStatus, "LeaseStart")
    GuiDisable("Disable")

    Dim $ProcessedLeaseArray[1]
    _GUICtrlListView_SetColumn($ListViewResults, 0, "IP Address", 100)
    _GUICtrlListView_SetColumn($ListViewResults, 1, "MAC Address", 150)
    _GUICtrlListView_SetColumn($ListViewResults, 2, "Name", 150)
    _GUICtrlListView_SetColumn($ListViewResults, 3, "????", 150)
    Local $RawLeaseArray[1]
    _GUICtrlListView_DeleteAllItems(GUICtrlGetHandle($ListViewResults))
    Local $Command = "netsh.exe"

    Local $commandString = ''

    ; Loop through the scopes
    For $X = GUICtrlRead($InputFirstScope) To GUICtrlRead($InputLastScope)
        If GUICtrlRead($CheckAbort) = $GUI_CHECKED Then
            ConsoleWrite(@ScriptLineNumber & " CheckAbort" & @CRLF)
            GuiDisable("Enable")
            Return
        EndIf
        ; Read the scopes leased addresses, parse it out and write data to an $RawLeaseArray
        $commandString = " dhcp server " & GUICtrlRead($InputCabinetIP) & " scope " & GUICtrlRead($InputCabinetNumber) & "." & $X & ".0.0" & " show clients 1 "
        GUICtrlSetData($LabelStatus, $X & " 0")
        Local $Results = 666
        If GUICtrlRead($CheckTest) = $GUI_UNCHECKED Then
            $Results = Run(@ComSpec & " /c " & $Command & " " & $commandString, " .", @SW_HIDE, $STDOUT_CHILD)
        EndIf
        GUICtrlSetData($LabelStatus, $X & " 1 a")
        Local $A[1]
        Local $line
        While 1
            $line = StdoutRead($Results)
            If @error Then ExitLoop
            If GUICtrlRead($CheckAbort) = $GUI_CHECKED Then
                ConsoleWrite(@ScriptLineNumber & " CheckAbort  >>" & $line & "<<>>" & @error & "<<" & @CRLF)
                GuiDisable("Enable")
                Return
            EndIf
            GUICtrlSetData($LabelStatus, $X & " 1 b")
            If StringLen($line) > 0 Then
                _ArrayAdd($A, $line)
            EndIf
        WEnd

        GUICtrlSetData($LabelStatus, $X & " 2")
        ;Write the data in the array to $RawLeaseArray
        Local $C[1]
        For $B In $A
            $C = StringSplit($B, @CRLF, 2)
            For $D In $C
                If StringLen($D) > 0 Then
                    _ArrayAdd($RawLeaseArray, $D)
                EndIf
            Next
        Next
        GUICtrlSetData($LabelStatus, $X & " 3")
    Next

    GUICtrlSetData($LabelStatus, $X & " 4")

    ;_ArrayDisplay($RawLeaseArray, @ScriptLineNumber)

    ; Parse the lease data from $RawLeaseArray. The results go to $ProcessedLeaseArray
    ;If the string $X begins with a number and a . then it is a value we care about

    Local $Y
    Local $Z
    For $W In $RawLeaseArray
        If StringInStr($W, GUICtrlRead($InputCabinetNumber) & ".") = 1 Then
            $X = StringStripWS(StringLeft($W, 16), 8) ; IP Address
            $Y = StringStripWS(StringMid($W, 35, 18), 8) ; MAC Address
            $Z = StringStripWS(StringMid($W, 83, 25), 8) ; Name
            Local $Dx = _GUICtrlListView_AddItem($ListViewResults, $X) ; display data
            _GUICtrlListView_AddSubItem($ListViewResults, $Dx, $Y, $Data2)
            _GUICtrlListView_AddSubItem($ListViewResults, $Dx, $Z, $Data3)
            _ArrayAdd($ProcessedLeaseArray, $X & "~~~" & $Y & "~~~" & $Z)
        EndIf
    Next

    ;_ArrayDisplay($RawLeaseArray, @ScriptLineNumber)
    _ArrayDelete($ProcessedLeaseArray, 0)
    ;_ArrayDisplay($ProcessedLeaseArray, @ScriptLineNumber)

    GuiDisable("Enable")
EndFunc   ;==>LeaseStart
;-----------------------------------------------
;Create and process scopes
Func ScopeStart()
    If VerifyParameters() <> 0 Then Return
    GUICtrlSetData($LabelStatus, "ScopeStart")
    GuiDisable("Disable")
    _GUICtrlListView_DeleteAllItems(GUICtrlGetHandle($ListViewResults))

    Local $Command = "netsh.exe"
    Local $commandString = ''
    Local $Result = 666
    Local $TS
    Local $Dx
    _GUICtrlListView_SetColumn($ListViewResults, 0, "New Column 1", 150, 1)
    GUICtrlSendMsg($ListViewResults, $LVM_SETCOLUMNWIDTH, $Data1, 100)
    GUICtrlSendMsg($ListViewResults, $LVM_SETCOLUMNWIDTH, $Data2, 150)
    GUICtrlSendMsg($ListViewResults, $LVM_SETCOLUMNWIDTH, $Data3, 150)
    #CS
        Local $Dx = _GUICtrlListView_AddItem($ListViewResults, $X) ; display data
        _GUICtrlListView_AddSubItem($ListViewResults, $Dx, $Y, $Data2)
        _GUICtrlListView_AddSubItem($ListViewResults, $Dx, $Z, $Data3)
    #ce

    ; Loop through the scopes
    For $X = GUICtrlRead($InputFirstScope) To GUICtrlRead($InputLastScope)
        If GUICtrlRead($CheckAbort) = $GUI_CHECKED Then
            ConsoleWrite(@ScriptLineNumber & " CheckAbort" & @CRLF)
            GuiDisable("Enable")
            Return
        EndIf

        ; Add the scopes
        $commandString = " dhcp server " & GUICtrlRead($InputCabinetIP) & " add scope " & _
                GUICtrlRead($InputCabinetNumber) & "." & $X & ".0.0" & " 255.255.0.0" & " R" & GUICtrlRead($InputCabinetNumber) & "VLAN" & $X & " NewVLAN" & $X

        $Dx = _GUICtrlListView_AddItem($ListViewResults, "Add scope") ; display data
        _GUICtrlListView_AddSubItem($ListViewResults, $Dx, $Command, $Data2)
        _GUICtrlListView_AddSubItem($ListViewResults, $Dx, $commandString, $Data3)

        If GUICtrlRead($CheckTest) = $GUI_UNCHECKED Then $Result = ShellExecute($Command, $commandString)
        If $Result = 0 Then _GUICtrlListView_AddItem($ListViewResults, "ERROR  1")

        ; Add the address pool for each scope
        $commandString = " dhcp server " & GUICtrlRead($InputCabinetIP) & " scope " & GUICtrlRead($InputCabinetNumber) & "." & $X & ".0.0 add iprange " & _
                GUICtrlRead($InputCabinetNumber) & "." & $X & ".0.1 " & GUICtrlRead($InputCabinetNumber) & "." & $X & ".255.254"

        $Dx = _GUICtrlListView_AddItem($ListViewResults, "Add address pool") ; display data
        _GUICtrlListView_AddSubItem($ListViewResults, $Dx, $Command, $Data2)
        _GUICtrlListView_AddSubItem($ListViewResults, $Dx, $commandString, $Data3)

        If GUICtrlRead($CheckTest) = $GUI_UNCHECKED Then $Result = ShellExecute($Command, $commandString)

        If $Result = 0 Then _GUICtrlListView_AddItem($ListViewResults, "ERROR  2")

        ; Create and move scopes to a superscope
        ; netsh dhcp server scope 2.2.0.0 set superscope "Cabinet 2" 1

        #cs
            $commandString = " dhcp server scope " & GUICtrlRead($InputCabinetNumber) & "." & $X & ".0.0 set superscope Cabinet" & GUICtrlRead($InputCabinetNumber) & " 1"
            
            $Dx = _GUICtrlListView_AddItem($ListViewResults, "Add superscopes") ; display data
            _GUICtrlListView_AddSubItem($ListViewResults, $Dx, $Command, $Data2)
            _GUICtrlListView_AddSubItem($ListViewResults, $Dx, $commandString, $Data3)
            
            If GUICtrlRead($CheckTest) = $GUI_UNCHECKED Then $Result = ShellExecute($Command, $commandString)
            $TS = $TS + $Result & " " & $Command & " " & $commandString
            If $Result = 0 Then _GUICtrlListView_AddItem($ListViewResults, "ERROR  3")
        #ce
        _GUICtrlListView_AddItem($ListViewResults, "")

    Next
    GuiDisable("Enable")
EndFunc   ;==>ScopeStart
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
    GUICtrlSetState($CheckAbort, $GUI_ENABLE)
    GUICtrlSetState($LabelStatus, $GUI_ENABLE)
EndFunc   ;==>GuiDisable
;-----------------------------------------------







