#include <Array.au3>
#include <Winpcap.au3>

; initialise the Library
$winpcap=_PcapSetup()
If ($winpcap=-1) Then
    MsgBox(16,"Pcap error !","WinPcap not found !")
    exit
EndIf

; Get the interfaces list for which a capture is possible
$pcap_devices=_PcapGetDeviceList()
If ($pcap_devices=-1) Then
    MsgBox(16,"Pcap error !",_PcapGetLastError())
    exit
EndIf

; Start a capture on interface #0, for ICMP packets only
$pcap=_PcapStartCapture($pcap_devices[0][0],"icmp")
If ($pcap=-1) Then
    MsgBox(16,"Pcap error !",_PcapGetLastError())
EndIf

; Detect of what type is the opened interface (ethernet, ATM, X25...)
$linktype=_PcapGetLinkType($pcap)
If ($linktype[1]<>"EN10MB") Then
    MsgBox(16,"Pcap error !","This example only accepts Ethernet devices...")
Endif

; Capture anything that matches our filter "ICMP" for 10 seconds...
$time0=TimerInit()
While (TimerDiff($time0)<10000) ; capture the packets for 10 seconds...
    $packet=_PcapGetPacket($pcap)
    If IsArray($packet) Then
        _arraydisplay($packet,@ScriptLineNumber)
    EndIf
Wend

; Stop capture
_PcapStopCapture($pcap)

; release ressources
_PcapFree()