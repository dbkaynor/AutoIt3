#include <Winpcap.au3>

$winpcap=_PcapSetup()   ; initialize winpcap
$pcap_devices=_PcapGetDeviceList()  ; get devices list

$pcap=_PcapStartCapture($pcap_devices[1][0]) ; my interface

$broadcastmac="FFFFFFFFFFFF" ; broacast
$src_mac=StringReplace($pcap_devices[1][6],":","") ; my mac address in hex
$dest_mac=""
$ethertype="3366"   ; fake ethertype, means nothing, just for example...
$mydata="0123456789"   ; dumb padding...

$mypacket="0x"&$broadcastmac&$mymac&$ethertype&$mydata ; stick together to a binary string !

ConsoleWrite(@ScriptLineNumber & " " & $mypacket & " " & @CRLF)
_PcapSendPacket($pcap,$mypacket) ; sends a valid ethernet broadcast !

_PcapFree() ; close winpcap
