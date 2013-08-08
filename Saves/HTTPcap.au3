#NoTrayIcon
#AutoIt3Wrapper_icon=httpcap.ico

; httpcap file capture v1.0e
; Copyleft GPL3 Nicolas Ricquemaque 2009-2011

#include <Array.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GuiListView.au3>
#include <StaticConstants.au3>
#include <EditConstants.au3>
#include <WinAPI.au3>
#include <ComboConstants.au3>

#include <Winpcap.au3>

If Not FileExists(@ScriptDir & "\httpcap.ini") Then
	MsgBox(16, "Missing file", @ScriptDir & "\httpcap.ini" & @CRLF & "not found")
	Exit
EndIf

$winpcap = _PcapSetup()
If ($winpcap = -1) Then
	MsgBox(16, "Pcap error !", "WinPcap not found !")
	Exit
EndIf

$pcap_devices = _PcapGetDeviceList()
If ($pcap_devices = -1) Then
	MsgBox(16, "Pcap error !", _PcapGetLastError())
	Exit
EndIf

$int = SelectInterface($pcap_devices)

$pcap = _PcapStartCapture($pcap_devices[$int][0], "host " & $pcap_devices[$int][7] & " and tcp port (80 or 8080)", 0, 65536, 2 ^ 24, 0)
If IsInt($pcap) Then
	MsgBox(16, "Pcap error !", _PcapGetLastError())
	_PcapFree()
	Exit
EndIf

GUICreate("HTTP file capture", 600, 350)

$listenwindow = GUICtrlCreateListView("Time|Type|Len", 10, 90, 290, 200)
_GUICtrlListView_SetColumn($listenwindow, 0, "Time", 55, 0)
_GUICtrlListView_SetColumn($listenwindow, 1, "Type", 160, 0)
_GUICtrlListView_SetColumn($listenwindow, 2, "Len", 50, 0)

$typelistwindow = GUICtrlCreateListView("Type|Ext", 360, 90, 230, 200)
_GUICtrlListView_SetColumn($typelistwindow, 0, "Type", 160, 0)
_GUICtrlListView_SetColumn($typelistwindow, 1, "Ext", 55, 0)

$tr = GUICtrlCreateButton(" -> ", 315, 180, 30)
GUICtrlSetState(-1, $GUI_DISABLE)
$add = GUICtrlCreateButton(" + ", 540, 60, 20)
$sup = GUICtrlCreateButton(" - ", 565, 60, 20)
GUICtrlSetState(-1, $GUI_DISABLE)

GUICtrlCreateLabel("RealTime view :", 10, 70)
GUICtrlCreateLabel("MimeTypes to capture :", 360, 70)
GUICtrlCreateLabel("Listening interface :", 10, 30)
GUICtrlCreateInput($pcap_devices[$int][7] & " - " & _PcapCleanDeviceName($pcap_devices[$int][1]), 120, 27, 465, Default, $ES_READONLY)
GUICtrlCreateLabel("Captures in progress :", 30, 300)
GUICtrlCreateLabel("Captures terminated :", 30, 320)
$inprogress = GUICtrlCreateLabel("0", 170, 300, 80)
$terminated = GUICtrlCreateLabel("0", 170, 320, 80)
GUICtrlCreateLabel("Capture only files bigger than:", 360, 313)
Global $minlength = IniRead(@ScriptDir & "\httpcap.ini", "General", "MinLength", 0)
$_minlength = GUICtrlCreateInput($minlength, 520, 310, 70, 20, $ES_NUMBER)

; read ini file
Global $mimetocap[1][2] ; [0]: mime   [1]: ext
$i = 0
Do
	$mime = IniRead(@ScriptDir & "\httpcap.ini", "MimeTypes", "Mime" & $i, "")
	$mime2 = StringSplit($mime, "|")
	If $mime2[1] <> "" Then
		GUICtrlCreateListViewItem($mime, $typelistwindow)
		$mimetocap[$i][0] = $mime2[1]
		$mimetocap[$i][1] = $mime2[2]
		$i += 1
		ReDim $mimetocap[$i + 1][2]
	EndIf
Until $mime = ""

; active recordings
Global $recordings[1][6] ; [0] port  [1]: handler [2]: start sequence number [3]:filepos [4]:expected length [5]:filename
Global $filenumber = IniRead(@ScriptDir & "\httpcap.ini", "General", "FileNumber", 1)
Global $terminatedfiles = 0

GUISetState()

$i = 0
Do
	EnableControlIf($tr, StringInStr(GUICtrlRead(GUICtrlRead($listenwindow)), "/") And GUICtrlRead($listenwindow) > 0)
	EnableControlIf($sup, GUICtrlRead($typelistwindow) > 0)
	If GUICtrlRead($inprogress) <> (UBound($recordings) - 1) Then GUICtrlSetData($inprogress, UBound($recordings) - 1)
	If GUICtrlRead($terminated) < $terminatedfiles Then GUICtrlSetData($terminated, $terminatedfiles)
	If $minlength <> GUICtrlRead($_minlength) Then $minlength = GUICtrlRead($_minlength)

	$msg = GUIGetMsg()
	If $msg = $tr Then
		$mime = StringSplit(GUICtrlRead(GUICtrlRead($listenwindow)), "|")
		$ext = InputBox("New MimeType to capture", 'Please enter the extension the files of type "' & $mime[2] & '" will be saved with (starting with a "dot", for an example ".txt"); leave blank for a prompt for each file...')
		GUICtrlCreateListViewItem($mime[2] & "|" & $ext, $typelistwindow)
		$mimetocap[UBound($mimetocap) - 1][0] = $mime[2]
		$mimetocap[UBound($mimetocap) - 1][1] = $ext
		ReDim $mimetocap[UBound($mimetocap) + 1][2]
	EndIf

	If $msg = $add Then
		$mime = InputBox("New MimeType to capture", 'Please enter the mime/type to capture (ex: text/html)' & @CRLF & @CRLF & "Autoit3 regular expressions starting with '^' are accepted (example: '^video/(x-)?flv' will match either video/flv or video/x-flv)")
		If $mime = "" Then ContinueLoop
		$ext = InputBox("New MimeType to capture", 'Please enter the extension the files of type "' & $mime & '" will be saved with (starting with a "dot", for an example ".txt"); leave blank for a prompt for each file...')
		GUICtrlCreateListViewItem($mime & "|" & $ext, $typelistwindow)
		$mimetocap[UBound($mimetocap) - 1][0] = $mime
		$mimetocap[UBound($mimetocap) - 1][1] = $ext
		ReDim $mimetocap[UBound($mimetocap) + 1][2]
	EndIf

	If $msg = $sup Then
		$tmp = StringSplit(GUICtrlRead(GUICtrlRead($typelistwindow)), "|")
		$index = _ArraySearch($mimetocap, $tmp[1])
		If $index > -1 Then _ArrayDelete($mimetocap, $index)
		_GUICtrlListView_DeleteItemsSelected($typelistwindow)
	EndIf

	If IsPtr($pcap) Then ; If $pcap is a Ptr, then the capture is running
		$time0 = TimerInit()
		While (TimerDiff($time0) < 500) ; Retrieve packets from queue for maximum 500ms before returning to main loop, not to "hang" the window for user
			$packet = _PcapGetPacket($pcap)
			If IsInt($packet) Then ExitLoop
			$http = HttpCapture($packet[3])
			If Not IsArray($http) Then ContinueLoop
			GUICtrlCreateListViewItem(StringTrimRight($packet[0], 7) & "|" & $http[0] & "|" & $http[1], $listenwindow)
			_GUICtrlListView_EnsureVisible($listenwindow, $i)
			$i += 1
		WEnd
	EndIf

	If $msg = $GUI_EVENT_CLOSE Then
		If (UBound($recordings) - 1) > 0 Then
			$warning = "The following files are still beeing captured:" & @CRLF & @CRLF
			For $j = 0 To UBound($recordings) - 2
				$warning &= $recordings[$j][5] & @CRLF
			Next
			$warning &= @CRLF & "Confirm quitting anyway ?"
			If MsgBox(36, "Quitting now ?", $warning) = 6 Then ExitLoop
		Else
			ExitLoop
		EndIf
	EndIf

Until False

; close all remaining open captures
For $j = 0 To UBound($recordings) - 2
	_WinAPI_CloseHandle($recordings[$j][1])
Next

; close winpcap wrapper
_PcapStopCapture($pcap)
_PcapFree()

; update ini file
FileDelete(@ScriptDir & "\httpcap.ini")
For $i = 0 To UBound($mimetocap) - 2
	If $mimetocap[$i][0] <> "" Then IniWrite(@ScriptDir & "\httpcap.ini", "MimeTypes", "Mime" & $i, $mimetocap[$i][0] & "|" & $mimetocap[$i][1])
Next
IniWrite(@ScriptDir & "\httpcap.ini", "General", "FileNumber", $filenumber)
IniWrite(@ScriptDir & "\httpcap.ini", "General", "MinLength", $minlength)
Exit


Func EnableControlIf($control, $condition)
	If $condition Then
		If BitAND(GUICtrlGetState($control), $GUI_DISABLE) Then GUICtrlSetState($control, $GUI_ENABLE)
	ElseIf BitAND(GUICtrlGetState($control), $GUI_ENABLE) Then
		GUICtrlSetState($control, $GUI_DISABLE)
	EndIf
EndFunc   ;==>EnableControlIf



Func HttpCapture($data)
	ConsoleWrite(@ScriptLineNumber & " Got here" & @CRLF)
	Local $ipheaderlen = BitAND(_PcapBinaryGetVal($data, 15, 1), 0xF) * 4
	Local $tcpoffset = $ipheaderlen + 14
	Local $tcplen = _PcapBinaryGetVal($data, 17, 2) - $ipheaderlen ; ip total len - ip header len
	Local $tcpheaderlen = BitShift(_PcapBinaryGetVal($data, $tcpoffset + 13, 1), 4) * 4
	Local $tcpsrcport = _PcapBinaryGetVal($data, $tcpoffset + 1, 2)
	Local $tcpdstport = _PcapBinaryGetVal($data, $tcpoffset + 3, 2)
	Local $tcpsequence = _PcapBinaryGetVal($data, $tcpoffset + 5, 4)
	Local $tcpflags = _PcapBinaryGetVal($data, $tcpoffset + 14, 1)
	Local $r[2] = ["", ""]


	; Received RST or FIN on a recording => Close it
	If BitAND($tcpflags, 0x05) Then
		Local $i = _ArraySearch($recordings, $tcpsrcport) ; are we already recording this port ?
		If $i > -1 Then
			_WinAPI_CloseHandle($recordings[$i][1])
			$r[0] = "Stopped " & $recordings[$i][5]
			$r[1] = Size($recordings[$i][4])
			_ArrayDelete($recordings, $i)
			$terminatedfiles += 1
			Return $r
		EndIf
	EndIf

	; From here, we are watching http payload
	Local $httpoffset = $tcpoffset + $tcpheaderlen + 1
	Local $httplen = $tcplen - $tcpheaderlen
	If $httplen = 0 Then Return False

	; are we already recording this port ?
	Local $i = _ArraySearch($recordings, $tcpdstport)
	If $i > -1 Then
		Local $http = BinaryToString(BinaryMid($data, $httpoffset, $httplen))
		Local $tBuffer = DllStructCreate("byte[" & StringLen($http) & "]")
		Local $nBytes
		DllStructSetData($tBuffer, 1, $http)

		Local $seq = GetTcpRelativeSequence($recordings[$i][2], $tcpsequence) + 1

		; Is packet in sequence ?
		If $seq = $recordings[$i][3] Then
			_WinAPI_WriteFile($recordings[$i][1], DllStructGetPtr($tBuffer), StringLen($http), $nBytes)
			$recordings[$i][3] += StringLen($http)
		ElseIf $seq < $recordings[$i][3] Then
			_WinAPI_SetFilePointer($recordings[$i][1], $seq)
			_WinAPI_WriteFile($recordings[$i][1], DllStructGetPtr($tBuffer), StringLen($http), $nBytes)
			$recordings[$i][3] = _WinAPI_SetFilePointer($recordings[$i][1], 0, 2)
		ElseIf $seq > $recordings[$i][3] Then
			_WinAPI_SetFilePointer($recordings[$i][1], $seq)
			_WinAPI_WriteFile($recordings[$i][1], DllStructGetPtr($tBuffer), StringLen($http), $nBytes)
			$recordings[$i][3] = $seq + StringLen($http)
		EndIf
		If $recordings[$i][3] = $recordings[$i][4] Then
			_WinAPI_CloseHandle($recordings[$i][1])
			$r[0] = "Completed " & $recordings[$i][5]
			$r[1] = Size($recordings[$i][4])
			_ArrayDelete($recordings, $i)
			$terminatedfiles += 1
			Return $r
		EndIf
		; Received RST or FIN on a recording => Close it
		If BitAND($tcpflags, 0x05) Then
			_WinAPI_CloseHandle($recordings[$i][1])
			$r[0] = "Stopped " & $recordings[$i][5]
			$r[1] = Size($recordings[$i][4])
			_ArrayDelete($recordings, $i)
			$terminatedfiles += 1
			Return $r
		EndIf
		Return False
	EndIf

	Local $http = BinaryToString(BinaryMid($data, $httpoffset, $httplen))
	If StringMid($http, 1, 4) <> "HTTP" Then Return False ; not a new download !
	Local $contenttype = StringRegExp($http, "(?i)Content-Type: ([A-Za-z0-9-/.]*)", 1)
	If @error <> 0 Then Return False
	Local $contentlength = StringRegExp($http, "(?i)Content-Length: ([0-9]*)", 1)
	Local $filelength = 0
	If @error = 0 Then $filelength = $contentlength[0]

	; New recording ?
	For $i = 0 To UBound($mimetocap) - 2
		If StringLeft($mimetocap[$i][0], 1) = '^' Then ; mimetocap is a regexp
			If StringRegExp($contenttype[0], $mimetocap[$i][0]) = 1 Then ExitLoop
		Else
			If $contenttype[0] = $mimetocap[$i][0] Then ExitLoop
		EndIf
	Next
	If Number($filelength) >= Number($minlength) And $i < (UBound($mimetocap) - 1) Then
		Local $j
		If $mimetocap[$i][1] = "" Then
			Local $name = FileSaveDialog("Save " & $contenttype[0] & " file to", @ScriptDir, "All (*.*)", 18)
		Else
			Local $name = $filenumber & $mimetocap[$i][1]
			$filenumber += 1
		EndIf
		Local $h = _WinAPI_CreateFile($name, 1)
		If $h <> 0 Then
			Local $start = StringInStr($http, @CRLF & @CRLF) + 4
			Local $filedata = StringMid($http, $start)
			Local $tBuffer = DllStructCreate("byte[" & StringLen($filedata) & "]")
			Local $nBytes
			DllStructSetData($tBuffer, 1, $filedata)
			_WinAPI_WriteFile($h, DllStructGetPtr($tBuffer), StringLen($filedata), $nBytes)
			Local $k = UBound($recordings) - 1
			$recordings[$k][0] = $tcpdstport
			$recordings[$k][1] = $h
			$recordings[$k][2] = $tcpsequence + $start - 1
			If $recordings[$k][2] > 0xFFFFFFFF Then $recordings[$k][2] -= 0xFFFFFFFF
			$recordings[$k][3] = StringLen($filedata)
			$recordings[$k][4] = $filelength
			$recordings[$k][5] = StringTrimLeft($name, StringInStr($name, "\", 0, -1))
			ReDim $recordings[$k + 2][6]
			If $recordings[$k][3] = $recordings[$k][4] Then
				_WinAPI_CloseHandle($h)
				$r[0] = "Completed " & $recordings[$k][5]
				$r[1] = Size($recordings[$k][4])
				_ArrayDelete($recordings, $k)
				$terminatedfiles += 1
				Return $r
			ElseIf BitAND($tcpflags, 0x05) Then ; Received RST or FIN on a recording => Close it
				_WinAPI_CloseHandle($h)
				$r[0] = "Stopped " & $recordings[$k][5]
				$r[1] = Size($recordings[$k][4])
				_ArrayDelete($recordings, $k)
				$terminatedfiles += 1
				Return $r
			Else
				$r[0] = "Recording " & $recordings[$k][5]
				$r[1] = Size($filelength)
				Return $r
			EndIf
		EndIf
		$r[0] = "Error creating " & $name
		Return $r
	EndIf

	Local $r[2]
	$r[0] = $contenttype[0]
	$r[1] = Size($filelength)
	Return $r
EndFunc   ;==>HttpCapture


Func Size($filelength)
	If $filelength > 1000000 Then
		Return Int($filelength / 1000000) & "M"
	ElseIf $filelength > 1000 Then
		Return Int($filelength / 1000) & "K"
	ElseIf $filelength = 0 Then
		Return ""
	EndIf
	Return $filelength
EndFunc   ;==>Size


Func SelectInterface($devices) ; auto selects an ethernet pcap interface or prompt user for choice
	Local $ipv4 = 0, $int = 0, $i, $win0, $first, $interface, $ok, $which, $msg
	For $i = 0 To UBound($devices) - 1
		If $devices[$i][3] = "EN10MB" And StringLen($devices[$i][7]) > 6 Then ; for ethernet devices with valid ip address only !
			$ipv4 += 1
			$int = $i
		EndIf
	Next

	If $ipv4 = 0 Then
		MsgBox(16, "Error", "No network interface found with a valid IPv4 address !")
		_PcapFree()
		Exit
	EndIf

	If $ipv4 > 1 Then
		$win0 = GUICreate("Interface choice", 500, 50)
		$interface = GUICtrlCreateCombo("", 10, 15, 400, Default, $CBS_DROPDOWNLIST)
		$first = True
		For $i = 0 To UBound($devices) - 1
			If $devices[$i][3] = "EN10MB" And StringLen($devices[$i][7]) > 6 Then
				If $first Then
					GUICtrlSetData(-1, $devices[$i][7] & " - " & _PcapCleanDeviceName($devices[$i][1]), $devices[$i][7] & " - " & _PcapCleanDeviceName($devices[$i][1]))
					$first = False
				Else
					GUICtrlSetData(-1, $devices[$i][7] & " - " & _PcapCleanDeviceName($devices[$i][1]))
				EndIf
			EndIf
		Next
		$ok = GUICtrlCreateButton(" Ok ", 430, 15, 60)
		GUISetState()
		While True
			$msg = GUIGetMsg()
			If $msg = $ok Then
				$which = GUICtrlRead($interface)
				For $i = 0 To UBound($devices) - 1
					If StringLen($devices[$i][7]) > 6 And StringInStr($which, $devices[$i][7]) Then
						$int = $i
						ExitLoop
					EndIf
				Next
				GUIDelete($win0)
				ExitLoop
			EndIf
			If $msg = $GUI_EVENT_CLOSE Then Exit
		WEnd
	EndIf
	Return $int
EndFunc   ;==>SelectInterface

Func GetTcpRelativeSequence($initial, $actual)
	If $actual >= $initial Then Return $actual - $initial
	Return 0xFFFFFFFF - $initial + $actual + 1
EndFunc   ;==>GetTcpRelativeSequence