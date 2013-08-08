TCPStartup()
HotKeySet("{esc}", "_close")

$ip = @IPAddress1
$ip = "191.0.0.100"
$port = 666

$connect = TCPConnect($ip, $port)

If @error Then
	ConsoleWrite(@ScriptLineNumber & " Could not connect to " & $ip & @CRLF)
	_Close()
EndIf

;$data = "0c 09 59 66 65 47 62 61 FF 0C AC 77 40 "
$data = " 77 "
For $x = 0 To 5
	TCPSend($connect, $data)
Next
Sleep(100)
_Close()

Func _Close()
	TCPShutdown()
	Exit
EndFunc   ;==>_Close