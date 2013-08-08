; Winpcap autoit3 UDF demo - V1.2c
; Copyleft GPL3 Nicolas Ricquemaque 2009-2011

#include <Array.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#Include <GuiListView.au3>
#include <StaticConstants.au3>
#include <ComboConstants.au3>

#include <Winpcap.au3>

$winpcap=_PcapSetup() ; initialise the Library

; Open pcap file for reading
$pcap=_PcapStartCapture(@ScriptDir & "\mycapture.pcap")

; Read whatever is in the file until its end.
Do
    $packet=_PcapGetPacket($pcap)
    If IsArray($packet) Then
        _arraydisplay($packet,@ScriptLineNumber)
    EndIf
Until $packet=-2  ; EOF

_PcapStopCapture($pcap) ; Stop capture
_PcapFree() ; release ressources