_connect()

Func _psftp_Connect()
	$exec = "C:\Program Files (x86)\PuTTY\psftp.exe"
	If Not FileExists($exec) Then
		_Err($exec & " not found!", 0)
		Exit
	EndIf

	$pid = Run($exec & ' ceg@192.168.1.108 -pw ceg -b get.txt')
	If Not $pid Then
		_Err("Failed to connect " & $pid, 0)
		Exit
	EndIf

	ConsoleWrite(@ScriptLineNumber & " " & @CRLF)
	Sleep(10)
EndFunc   ;==>_Connect


Func _Err($data, $pid)
	If $data And $data <> -1 Then MsgBox(0, "An Error has Occured", $data, 5)
	_Exit($pid)
EndFunc   ;==>_Err

Func _Exit($pid)
	ProcessClose($pid)
EndFunc   ;==>_Exit