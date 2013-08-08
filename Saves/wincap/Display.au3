#include <Array.au3>
#include <Winpcap.au3>

Opt("MustDeclareVars", 1)

Global $pcap
Global $handle

Global $winpcap = _PcapSetup() ; initialize winpcap
Global $pcap_devices = _PcapGetDeviceList() ; get devices list
;_ArrayDisplay($pcap_devices, "Devices list", -1, 1) ; display it

Global $AA = 2
ConsoleWrite(@ScriptLineNumber & " Name: " & $pcap_devices[$AA][1] & @CRLF)
ConsoleWrite(@ScriptLineNumber & " IP: " & $pcap_devices[$AA][7] & @CRLF)
_PcapStartCapture($pcap_devices[0][1])
Sleep(10000)

global $filename = @ScriptDir & "\pcap.txt"
$handle = FileOpen($filename, 2)
FileWriteLine($handle,"hello")
ConsoleWrite(@ScriptLineNumber & " File handle: " & $handle & @CRLF)
ConsoleWrite(@ScriptLineNumber & " Is pointer: " & IsPtr($handle) & @CRLF)

_PcapSaveToFile($pcap, $handle)
_PcapWriteLastPacket($handle)

_PcapStopCaptureFile($handle)

_PcapStopCapture($pcap)
_PcapFree() ; close winpcap

ConsoleWrite(@ScriptLineNumber & " File exists: " & FileExists($filename) & @CRLF)
ConsoleWrite(@ScriptLineNumber & " File size: " & FileGetSize($filename) & @CRLF)
#cs
	;start a capture in non-blocking mode on device $DeviceName with
	_PcapStartCapture($DeviceName, $filter = "",
	$promiscuous = 0,
	$PacketLen = 65536,
	$buffersize = 0,
	$realtime = 1)
#ce