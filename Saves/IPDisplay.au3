; Winpcap autoit3 UDF demo - V1.2c
; Copyleft GPL3 Nicolas Ricquemaque 2009-2011
#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_icon=../icons/Cryptkeeper.ico
#AutoIt3Wrapper_outfile=C:\Program Files (x86)\AutoIt3\Dougs\IPDisplay.exe
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Comment=Start various programs
#AutoIt3Wrapper_Res_Description=Start various programs
#AutoIt3Wrapper_Res_Fileversion=0.0.0.0
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=Y
#AutoIt3Wrapper_Res_ProductVersion=000
#AutoIt3Wrapper_Res_LegalCopyright=Copyright 2011 Douglas B Kaynor & Nicolas Ricquemaque
#AutoIt3Wrapper_Res_SaveSource=y
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Field=AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Res_Field=Compile date|%longdate% %time%
#AutoIt3Wrapper_Res_Field=Developer|Douglas Kaynor & Nicolas Ricquemaque
#AutoIt3Wrapper_Res_Field=Email|doug@kaynor.net
#AutoIt3Wrapper_Res_Field=Made By|Douglas Kaynor & Nicolas Ricquemaque
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=y
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Run_Tidy=y
#Tidy_Parameters=/tc 4 /kv 2
#AutoIt3Wrapper_Run_Debug_Mode=n
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <Array.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <Misc.au3>
#include <WindowsConstants.au3>
#include <GuiListView.au3>
#include <StaticConstants.au3>
#include <ComboConstants.au3>
#include <Winpcap.au3>

Opt("MustDeclareVars", 1) ; require pre-declared varibles
If _Singleton(@ScriptName, 1) = 0 Then
    MsgBox(16, @ScriptName, @ScriptName & @CRLF & " is already running!")
    Exit
EndIf

If _PcapSetup() = -1 Then
    MsgBox(16, "Pcap error !", "PcapSetup failed!")
    Exit
EndIf

Global $pcap_devices = _PcapGetDeviceList()
If ($pcap_devices = -1) Then
    MsgBox(16, "No devices found!", _PcapGetLastError())
    Exit
EndIf

Global Const $tmp = StringSplit(@ScriptName, ".")
Global Const $ProgramName = $tmp[1]
Global Const $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")

Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, _
        $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)

; Capture Main --------------------------
Global $MainForm = GUICreate($ProgramName & "  " & @ComputerName & "  " & $FileVersion, 500, 380, 10, 10, $MainFormOptions)
;GUISetFont(8, 400, -1, "Courier new")

GUICtrlSetStyle(GUICtrlCreateLabel("Interface :", 10, 20, 60), $SS_RIGHT)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ComboInterface = GUICtrlCreateCombo("", 80, 20, 350, Default, $CBS_DROPDOWNLIST)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetData(-1, "Pcap capture file")
For $i = 0 To UBound($pcap_devices) - 1
    GUICtrlSetData(-1, $i & " " & $pcap_devices[$i][1], "Pcap capture file")
Next

GUICtrlSetStyle(GUICtrlCreateLabel("IP Address: ", 10, 50, 60), $SS_RIGHT)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $InputIPAddress = GUICtrlCreateInput("", 80, 50, 300, Default, $ES_READONLY)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetStyle(GUICtrlCreateLabel("Filter: ", 10, 80, 60), $SS_RIGHT)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $InputFilter = GUICtrlCreateInput("", 80, 80, 300)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ButtonStart = GUICtrlCreateButton("Start", 10, 110, 60)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonStop = GUICtrlCreateButton("Stop", 80, 110, 60)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetState(-1, $GUI_DISABLE)
Global $ButtonClear = GUICtrlCreateButton("Clear", 150, 110, 60)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonStats = GUICtrlCreateButton("Stats", 220, 110, 60)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetState($ButtonStats, $GUI_DISABLE)
Global $ButtonDevices = GUICtrlCreateButton("Devices", 290, 110, 60)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonExit = GUICtrlCreateButton("Exit", 360, 110, 60)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $CheckPromiscuous = GUICtrlCreateCheckbox("Promiscuous", 10, 140)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $CheckSave = GUICtrlCreateCheckbox("Save packets", 100, 140)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

GUICtrlSetStyle(GUICtrlCreateLabel("Number of packets: ", 190, 145), $SS_RIGHT)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $InputNumberPackets = GUICtrlCreateInput("20", 300, 140, 30)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ListViewPacketWindow = GUICtrlCreateListView("No|Time|Len|Packet", 10, 180, 480, 180)
GUICtrlSetResizing(-1, 2 + 32 + 64)
_GUICtrlListView_SetColumn($ListViewPacketWindow, 0, "No", 40, 1)
_GUICtrlListView_SetColumnWidth($ListViewPacketWindow, 1, 80)
_GUICtrlListView_SetColumn($ListViewPacketWindow, 2, "Len", 40, 1)
_GUICtrlListView_SetColumnWidth($ListViewPacketWindow, 3, 300)

GUISetState()

Global $Interface
Global $pcap = 0
Global $Promiscuous
Global $packet = 0
Global $pcapfile = 0
Global $PacketCount = 0

While 1
    Global $msg = GUIGetMsg(True)
    Switch $msg[0]
        Case $GUI_EVENT_CLOSE
            If $msg[1] = $MainForm Then Exit

        Case $ComboInterface
            GUICtrlSetData($InputIPAddress, "")
            For $n = 0 To UBound($pcap_devices) - 1
                If $pcap_devices[$n][1] = StringStripWS(StringMid(GUICtrlRead($ComboInterface), 2, 100), 3) Then
                    GUICtrlSetData($InputIPAddress, $pcap_devices[$n][7])
                    $Interface = $pcap_devices[$n][0]
                    ExitLoop
                EndIf
            Next

            ;name ;desc ;mac ;ipv4_addr ;ipv4_netmask ;ipv4_broadaddr ;ipv6_addr ;ipv6_netmask ;ipv6_broadaddr ;flags) or -1 if error
        Case $ButtonDevices
            If Not IsArray($pcap_devices) Then MsgBox(16, "Devices", "Problem with devices")
            _ArrayDisplay($pcap_devices, "Devices list", -1, 1) ; display it

        Case $ButtonClear
            _PcapGetStats($pcap)
            _GUICtrlListView_DeleteAllItems($ListViewPacketWindow)
            $PacketCount = 0

        Case $ButtonStats
            Global $CaptureStats = _PcapGetStats($pcap)
            If Not IsArray($CaptureStats) Then MsgBox(16, "Stats", "Problem with stats")
            _ArrayDisplay($CaptureStats, "Capture statistics")

        Case $ButtonStop
            If IsPtr($pcapfile) Then
                _PcapStopCaptureFile($pcapfile)
                $pcapfile = 0
            EndIf
            If Not IsInt($pcap) Then _PcapStopCapture($pcap)
            $pcap = 0
            GUICtrlSetState($ComboInterface, $GUI_ENABLE)
            GUICtrlSetState($ButtonStart, $GUI_ENABLE)
            GUICtrlSetState($ButtonStop, $GUI_DISABLE)
            GUICtrlSetState($ButtonStats, $GUI_DISABLE)
            GUICtrlSetState($ButtonDevices, $GUI_ENABLE)
            GUICtrlSetState($CheckSave, $GUI_ENABLE)
            GUICtrlSetState($CheckPromiscuous, $GUI_ENABLE)
            GUICtrlSetState($InputFilter, $GUI_ENABLE)


        Case $ButtonExit
            If IsPtr($pcapfile) Then _PcapStopCaptureFile($pcapfile) ; A file is still open: close it
            If IsPtr($pcap) Then _PcapStopCapture($pcap) ; A capture is still running: close it
            _PcapFree()
            Exit


        Case $ButtonStart
            If GUICtrlRead($CheckPromiscuous) = $GUI_CHECKED Then
                $Promiscuous = 1
            Else
                $Promiscuous = 0
            EndIf
            If (GUICtrlRead($ComboInterface) = "Pcap capture file") Then OpenPCAPFile()

            $pcap = _PcapStartCapture($Interface, GUICtrlRead($InputFilter), $Promiscuous)
            If ($pcap = -1) Then
                MsgBox(16, "Pcap error !", _PcapGetLastError())
                ContinueLoop
            EndIf
            Global $linktype = _PcapGetLinkType($pcap)
            If ($linktype[1] <> "EN10MB") Then
                MsgBox(16, "Pcap error !", "This example only works for Ethernet captures")
                ContinueLoop
            EndIf
            If GUICtrlRead($CheckSave) = $GUI_CHECKED Then
                Global $File = FileSaveDialog("Pcap file to write to ?", ".", "Pcap (*.pcap)", 16)
                If ($File <> "") Then
                    If StringLower(StringRight($File, 5)) <> ".pcap" Then $File &= ".pcap"
                    $pcapfile = _PcapSaveToFile($pcap, $File)
                    If ($pcapfile = 0) Then MsgBox(16, "Pcap error !", _PcapGetLastError())
                EndIf
            EndIf
            GUICtrlSetState($ComboInterface, $GUI_DISABLE)
            GUICtrlSetState($ButtonStart, $GUI_DISABLE)
            GUICtrlSetState($ButtonStop, $GUI_ENABLE)
            GUICtrlSetState($ButtonStats, $GUI_ENABLE)
            GUICtrlSetState($ButtonDevices, $GUI_DISABLE)
            GUICtrlSetState($CheckSave, $GUI_DISABLE)
            GUICtrlSetState($CheckPromiscuous, $GUI_DISABLE)
            GUICtrlSetState($InputFilter, $GUI_DISABLE)

    EndSwitch

    If IsPtr($pcap) Then ; If $pcap is a Ptr, then the capture is running
        Global $time0 = TimerInit()
        While (TimerDiff($time0) < 500) ; Retrieve packets from queue for maximum 500ms before returning to main loop, not to "hang" the window for user
            $packet = _PcapGetPacket($pcap)
            If IsInt($packet) Then ExitLoop ; $InputNumberPackets
            GUICtrlCreateListViewItem($PacketCount & "|" & StringTrimRight($packet[0], 4) & "|" & $packet[2] & "|" & MyDissector($packet[3]), $ListViewPacketWindow)
            Global $data = $packet[3]
            _GUICtrlListView_EnsureVisible($ListViewPacketWindow, $PacketCount)
            $PacketCount += 1
            If IsPtr($pcapfile) Then _PcapWriteLastPacket($pcapfile)
        WEnd
    EndIf
WEnd
Exit
;-----------------------------------------------
Func OpenPCAPFile()
    Local $PCAPPath = FileOpenDialog("Pcap file to open ?", ".", "Pcap (*.pcap)|All files (*.*)", 1)
    If @error <> 0 Then Return
    If $PCAPPath = "" Then Return
    $Interface = "file://" & $PCAPPath
EndFunc   ;==>OpenPCAPFile
;-----------------------------------------------
Func MyDissector($data) ; Quick example packet dissector....
    Local $macdst = StringMid($data, 3, 2) & ":" & StringMid($data, 5, 2) & ":" & StringMid($data, 7, 2) & ":" & StringMid($data, 9, 2) & ":" & StringMid($data, 11, 2) & ":" & StringMid($data, 13, 2)
    Local $macsrc = StringMid($data, 15, 2) & ":" & StringMid($data, 17, 2) & ":" & StringMid($data, 19, 2) & ":" & StringMid($data, 21, 2) & ":" & StringMid($data, 23, 2) & ":" & StringMid($data, 25, 2)
    Local $ethertype = BinaryMid($data, 13, 2)
    Local $srcport
    Local $dstport

    If $ethertype = "0x0806" Then Return "ARP " & $macsrc & " -> " & $macdst

    If $ethertype = "0x0800" Then
        Local $src = Number(BinaryMid($data, 27, 1)) & "." & Number(BinaryMid($data, 28, 1)) & "." & Number(BinaryMid($data, 29, 1)) & "." & Number(BinaryMid($data, 30, 1))
        Local $dst = Number(BinaryMid($data, 31, 1)) & "." & Number(BinaryMid($data, 32, 1)) & "." & Number(BinaryMid($data, 33, 1)) & "." & Number(BinaryMid($data, 34, 1))
        Switch BinaryMid($data, 24, 1)
            Case "0x01"
                Return "ICMP " & $src & " -> " & $dst
            Case "0x02"
                Return "IGMP " & $src & " -> " & $dst
            Case "0x06"
                $srcport = Number(BinaryMid($data, 35, 1)) * 256 + Number(BinaryMid($data, 36, 1))
                $dstport = Number(BinaryMid($data, 37, 1)) * 256 + Number(BinaryMid($data, 38, 1))
                Local $flags = BinaryMid($data, 48, 1)
                Local $f = ""
                If BitAND($flags, 0x01) Then $f = "Fin "
                If BitAND($flags, 0x02) Then $f &= "Syn "
                If BitAND($flags, 0x04) Then $f &= "Rst "
                If BitAND($flags, 0x08) Then $f &= "Psh "
                If BitAND($flags, 0x10) Then $f &= "Ack "
                If BitAND($flags, 0x20) Then $f &= "Urg "
                If BitAND($flags, 0x40) Then $f &= "Ecn "
                If BitAND($flags, 0x80) Then $f &= "Cwr "
                $f = StringTrimRight(StringReplace($f, " ", ","), 1)
                Return "TCP(" & $f & ") " & $src & ":" & $srcport & " -> " & $dst & ":" & $dstport
            Case "0x11"
                $srcport = Number(BinaryMid($data, 35, 1)) * 256 + Number(BinaryMid($data, 36, 1))
                $dstport = Number(BinaryMid($data, 37, 1)) * 256 + Number(BinaryMid($data, 38, 1))
                Return "UDP " & $src & ":" & $srcport & " -> " & $dst & ":" & $dstport
            Case Else
                Return "IP " & BinaryMid($data, 24, 1) & " " & $src & " -> " & $dst
        EndSwitch
        Return BinaryMid($data, 13, 2) & " " & $src & " -> " & $dst
    EndIf

    If $ethertype = "0x8137" Or $ethertype = "0x8138" Or $ethertype = "0x0022" Or $ethertype = "0x0025" Or $ethertype = "0x002A" Or $ethertype = "0x00E0" Or $ethertype = "0x00FF" Then
        Return "IPX " & $macsrc & " -> " & $macdst
    EndIf

    Return "[" & $ethertype & "] " & $macsrc & " -> " & $macdst
EndFunc   ;==>MyDissector
