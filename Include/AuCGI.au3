;----------------------------------------------------------------
;                AuCGI CGI Handler for AutoIt
;      by Erik Pilsits (wraithdu), and Josh Rowe (JRowe)
; original author Matt Roth (theguy0000) <theguy0000@gmail.com>
;
;                        Version 2.0
;                       May 11, 2011
;----------------------------------------------------------------
#NoTrayIcon
#AutoIt3Wrapper_Au3Check_Parameters=-q -d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#AutoIt3Wrapper_Res_Fileversion=2.0.0.3
#AutoIt3Wrapper_Run_Obfuscator=y
#Obfuscator_Parameters=/sf /sv /om /cs=0 /cn=0
#AutoIt3Wrapper_Run_After=del "%scriptdir%\%scriptfile%_Obfuscated.au3"

Opt("MustDeclareVars", 1)

Global $AuCGI_version = "2.0"

_Main()
Exit

Func _Main()
	;initialize environment variables, return size of any POST data
	Local $ContentLength = _envInit($AuCGI_version)
	;A sourcescript is sent to the CGI handler for processing as an au3 web object.
;~ 	Local $sourcescript = $CmdLine[1]
	Local $sourcescript = StringReplace(StringRegExpReplace(EnvGet("DOCUMENT_ROOT"), "^(.*?)\\*$", "${1}\\") & EnvGet("SCRIPT_NAME"), "/", "\")
	;Check if the $sourcescript exists, otherwise exit
	If Not FileExists($sourcescript) Then _error("There is no spoon.")
	;Else set the $source
	Local $source = FileRead($sourcescript)
	;Create new script path
	Local $newscriptpath = _generatePath()

	;formatting the source script about to be executed
	$source = _parseSource($source)
	$source = _prependSource($source)

	;create source file
	Local $hFile = FileOpen($newscriptpath, 2 + 8) ; overwrite + create dir structure
	FileWrite($hFile, $source)
	FileClose($hFile)

	;find and run AutoIt
	Local $RegKey = "HKLM\SOFTWARE\AutoIt v3\AutoIt"
	If @AutoItX64 Then $RegKey = "HKLM\SOFTWARE\Wow6432Node\AutoIt v3\AutoIt"
	Local $AutoItDir = RegRead($RegKey, "InstallDir")
	Local $AutoItExe = "AutoIt3.exe"
	; uncomment the next line if you want to run the webapp as native x64
;~  If @OSArch <> "X86" Then $AutoItExe = "AutoIt3_x64.exe"
	$AutoItExe = $AutoItDir & "\" & $AutoItExe
	If @error Or Not FileExists($AutoItExe) Then _error("Could not locate AutoIt executable.")
	Local $run = Run('"' & $AutoItExe & '" "' & $newscriptpath & '"', "", Default, 3) ; STDIN/OUT

	; give post data to AutoIt3.exe
	If $ContentLength Then
		Local $post = "", $timer = TimerInit()
		While StringLen($post) < $ContentLength
			$post &= ConsoleRead()
			If TimerDiff($timer) > 5000 Then ExitLoop ; 5 second timeout
		WEnd
		StdinWrite($run, $post)
	EndIf

	; give script output to server
	; run loop at least once, even if process has already exited
	; seems to work better this way on very fast servers
	Local $out = ""
	Do
		While 1
			$out = StdoutRead($run)
			If @extended Or @error Then ExitLoop
			Sleep(10)
		WEnd
		If $out Then ConsoleWrite($out)
		$out = ""
	Until (Not ProcessExists($run))
	; one final buffer check
	$out = StdoutRead($run)
	If @extended Then ConsoleWrite($out)
	StdioClose($run)
EndFunc   ;==>_Main

;Generate randomized temporary script paths to prevent server hang
;Randomly generates names, selects one that isn't already existing in the temp directory
Func _generatePath()
	Local $newscriptpath
	Do
		$newscriptpath = ""
		For $i = 1 To 7
			$newscriptpath &= Chr(Random(97, 122, 1))
		Next
		$newscriptpath = @ScriptDir & "\AUCGI\~" & $newscriptpath & ".tmp"
	Until Not FileExists($newscriptpath)
	Return $newscriptpath
EndFunc   ;==>_generatePath

;initialize environment variables and return POST content length
Func _envInit($version)
	EnvSet("AUCGI_VERS", $version)
	Local $ContentLength = Number(EnvGet("CONTENT_LENGTH"))
	EnvSet("CONTENT_LENGTH", String($ContentLength))
	EnvSet("QUERY_STRING", EnvGet("QUERY_STRING"))
	EnvSet("HTTP_COOKIE", EnvGet("HTTP_COOKIE"))
	EnvSet("REMOTE_ADDR", EnvGet("REMOTE_ADDR"))
	EnvSet("HTTP_ACCEPT_LANGUAGE", EnvGet("HTTP_ACCEPT_LANGUAGE"))
	EnvSet("HTTP_HOST", EnvGet("HTTP_HOST"))
	EnvSet("HTTP_ACCEPT_CHARSET", EnvGet("HTTP_ACCEPT_CHARSET"))
	EnvSet("HTTP_USER_AGENT", EnvGet("HTTP_USER_AGENT"))
	EnvSet("SERVER_SOFTWARE", EnvGet("SERVER_SOFTWARE"))
	EnvSet("SERVER_NAME", EnvGet("SERVER_NAME"))
	EnvSet("SERVER_PROTOCOL", EnvGet("SERVER_PROTOCOL"))
	EnvSet("SERVER_PORT", EnvGet("SERVER_PORT"))
	EnvSet("SCRIPT_NAME", EnvGet("SCRIPT_NAME"))
	EnvSet("HTTPS", EnvGet("HTTPS"))
	EnvSet("DOCUMENT_ROOT", EnvGet("DOCUMENT_ROOT"))
	EnvSet("HTTP_REFERER", EnvGet("HTTP_REFERER"))
	EnvSet("PATH", EnvGet("PATH"))
	EnvSet("REMOTE_HOST", EnvGet("REMOTE_HOST"))
	EnvSet("REMOTE_PORT", EnvGet("REMOTE_PORT"))
	EnvSet("REMOTE_USER", EnvGet("REMOTE_USER"))
	EnvSet("REQUEST_METHOD", EnvGet("REQUEST_METHOD"))
	EnvSet("REQUEST_URI", EnvGet("REQUEST_URI"))
	EnvSet("SCRIPT_FILENAME", EnvGet("SCRIPT_FILENAME"))
	EnvSet("SERVER_ADMIN", EnvGet("SERVER_ADMIN"))
	EnvSet("PATH_INFO", EnvGet("PATH_INFO"))
	EnvSet("PATH_TRANSLATED", EnvGet("PATH_TRANSLATED"))
	Return $ContentLength
EndFunc   ;==>_envInit

; parse source for <?au3 ?> code
Func _parseSource($source)
	Local $idx = 1, $idx2, $lastidx = 1, $parsed = "", $chunk = ""
	Do
		; get first code snippet
		$idx = StringInStr($source, "<?au3", 0, 1, $idx)
		If $idx Then
			If $idx > $lastidx Then
				; we have html
				$chunk = StringMid($source, $lastidx, $idx - $lastidx) ; get it
				$parsed &= _splitHTML($chunk) ; write it
			EndIf
			$idx += 5 ; start of code
			; get end of code tag
			$idx2 = StringInStr($source, "?>", 0, 1, $idx)
			If $idx2 Then
				; found end of code
				$chunk = StringMid($source, $idx, $idx2 - $idx) ; get it
				$parsed &= $chunk & @CRLF ; write it
				$lastidx = $idx2 + 2 ; new $lastidx value, set to position after end-code tag
				$idx = $lastidx ; next search start location
			Else
				; parse error, get out
				_error("Error parsing source.")
			EndIf
		Else
			; no code sections or last section of html
			$chunk = StringMid($source, $lastidx) ; get it
			If $chunk Then $parsed &= _splitHTML($chunk) ; check we actually have something this time, write it
		EndIf
	Until Not $idx
	Return $parsed
EndFunc   ;==>_parseSource

; split HTML chunks into lines to avoid ConsoleWrite buffer limit
Func _splitHTML($chunk)
	Local $sReturn = ""
	Local $aChunks = StringSplit($chunk, @CRLF, 1) ; split lines
	For $i = 1 To $aChunks[0] - 1
		$sReturn &= 'ConsoleWrite("' & StringReplace($aChunks[$i], '"', '""') & '" & @CRLF)' & @CRLF ; write it
	Next
	$sReturn &= 'ConsoleWrite("' & StringReplace($aChunks[$aChunks[0]], '"', '""') & '")' & @CRLF ; special case last line
	Return $sReturn
EndFunc   ;==>_splitHTML

; prepend core functions to prepared source
Func _prependSource($source)
	Local $prepend = _
			'FileDelete(@ScriptFullPath)' & @CRLF & _
			'Global $_POST_raw = _getPOSTData(Number(EnvGet("CONTENT_LENGTH")))' & @CRLF & _
			'Global $AuCGI_version = EnvGet("AUCGI_VERS")' & @CRLF & _
			'Global $_GET_raw = EnvGet("QUERY_STRING")' & @CRLF & _
			'Global $_Cookie_Raw = EnvGet("HTTP_COOKIE")' & @CRLF & _
			'Global $_REMOTE_ADDR = EnvGet("REMOTE_ADDR")' & @CRLF & _
			'Global $_ACCEPT_LANGUAGE = EnvGet("HTTP_ACCEPT_LANGUAGE")' & @CRLF & _
			'Global $_HOST = EnvGet("HTTP_HOST")' & @CRLF & _
			'Global $_ACCEPT_CHARSET = EnvGet("HTTP_ACCEPT_CHARSET")' & @CRLF & _
			'Global $_USER_AGENT = EnvGet("HTTP_USER_AGENT")' & @CRLF & _
			'Global $_SERVER_SOFTWARE = EnvGet("SERVER_SOFTWARE")' & @CRLF & _
			'Global $_SERVER_NAME = EnvGet("SERVER_NAME")' & @CRLF & _
			'Global $_SERVER_PROTOCOL = EnvGet("SERVER_PROTOCOL")' & @CRLF & _
			'Global $_SERVER_PORT = EnvGet("SERVER_PORT")' & @CRLF & _
			'Global $_SCRIPT_NAME = EnvGet("SCRIPT_NAME")' & @CRLF & _
			'Global $_HTTPS = EnvGet("HTTPS")' & @CRLF & _
			'Global $_DOCUMENT_ROOT = EnvGet("DOCUMENT_ROOT")' & @CRLF & _
			'Global $_HTTP_REFERER = EnvGet("HTTP_REFERER")' & @CRLF & _
			'Global $_PATH = EnvGet("PATH")' & @CRLF & _
			'Global $_REMOTE_HOST = EnvGet("REMOTE_HOST")' & @CRLF & _
			'Global $_REMOTE_PORT = EnvGet("REMOTE_PORT")' & @CRLF & _
			'Global $_REMOTE_USER = EnvGet("REMOTE_USER")' & @CRLF & _
			'Global $_REQUEST_METHOD = EnvGet("REQUEST_METHOD")' & @CRLF & _
			'Global $_REQUEST_URI = EnvGet("REQUEST_URI")' & @CRLF & _
			'Global $_SCRIPT_FILENAME = StringReplace(StringRegExpReplace($_DOCUMENT_ROOT, "^(.*?)\\*$", "${1}\\") & $_SCRIPT_NAME, "/", "\")' & @CRLF & _
			'Global $_SERVER_ADMIN = EnvGet("SERVER_ADMIN")' & @CRLF & _
			'Func _standardHeader()' & @CRLF & _
			'ConsoleWrite("Content-Type: text/html" & @CRLF & @CRLF)' & @CRLF & _
			'EndFunc' & @CRLF & _
			'Func _getPOSTData($length)' & @CRLF & _
			'Local $POST = ""' & @CRLF
	$prepend &= _
			'If $length > 0 Then' & @CRLF & _
			'Local $timer = TimerInit()' & @CRLF & _
			'While StringLen($POST) < $length' & @CRLF & _
			'$POST &= ConsoleRead()' & @CRLF & _
			'If TimerDiff($timer) > 5000 Then ExitLoop' & @CRLF & _
			'WEnd' & @CRLF & _
			'EndIf' & @CRLF & _
			'Return $POST' & @CRLF & _
			'EndFunc' & @CRLF & _
			'Func _Post($sVar)' & @CRLF & _
			'Local $varstring = $_POST_raw' & @CRLF & _
			'If Not StringInStr($varstring, $sVar & "=") Then Return ""' & @CRLF & _
			'Local $vars = StringSplit($varstring, "&")' & @CRLF & _
			'Local $var_array' & @CRLF & _
			'For $i = 1 To $vars[0]' & @CRLF & _
			'$var_array = StringSplit($vars[$i], "=")' & @CRLF & _
			'If $var_array[0] < 2 Then Return ""' & @CRLF & _
			'If $var_array[1] = $sVar Then Return _URLDecode($var_array[2])' & @CRLF & _
			'Next' & @CRLF & _
			'Return ""' & @CRLF & _
			'EndFunc' & @CRLF & _
			'Func _Get($sVar, $iVarType = 0)' & @CRLF & _
			'Local $varstring = $_GET_raw' & @CRLF & _
			'If Not StringInStr($varstring, $sVar & "=") Then Return ""' & @CRLF & _
			'Local $vars = StringSplit($varstring, "&")' & @CRLF & _
			'Local $var_array' & @CRLF & _
			'For $i = 1 To $vars[0]' & @CRLF & _
			'If $iVarType Then' & @CRLF & _
			'If $vars[$i] = $sVar Then Return True' & @CRLF & _
			'Else' & @CRLF & _
			'$var_array = StringSplit($vars[$i], "=")' & @CRLF & _
			'If $var_array[0] < 2 Then Return ""' & @CRLF & _
			'If $var_array[1] = $sVar Then Return _URLDecode($var_array[2])' & @CRLF & _
			'EndIf' & @CRLF & _
			'Next' & @CRLF & _
			'If $iVarType Then Return False' & @CRLF & _
			'Return ""' & @CRLF & _
			'EndFunc' & @CRLF & _
			'Func _PostMultipart($sVar)' & @CRLF & _
			'Local $ret[3], $temp_split, $name_pos, $name, $quote_pos, $pos' & @CRLF & _
			'Local $split = StringSplit($_POST_raw, "-----------------------------", 1)' & @CRLF & _
			'For $i = 1 To $split[0]' & @CRLF & _
			'$temp_split = StringSplit($split[$i], @CRLF, 1)' & @CRLF
	$prepend &= _
			'If $temp_split[0] < 4 Then ContinueLoop' & @CRLF & _
			'__ArrayDelete($temp_split, 0)' & @CRLF & _
			'If StringInStr($temp_split[1], "; filename=""") Then' & @CRLF & _
			'$name_pos = StringInStr($temp_split[1], "; name=""")' & @CRLF & _
			'$name = StringRight($temp_split[1], StringLen($temp_split[1]) -($name_pos + 7))' & @CRLF & _
			'$quote_pos = StringInStr($name, """")' & @CRLF & _
			'$name = StringLeft($name,(StringLen($name) -(StringLen($name) - $quote_pos)) - 1)' & @CRLF & _
			'If $name <> $sVar Then ContinueLoop' & @CRLF & _
			'$pos = StringInStr($temp_split[1], "; filename=""")' & @CRLF & _
			'$ret[0] = StringRight($temp_split[1], StringLen($temp_split[1]) -($pos + 11))' & @CRLF & _
			'$ret[0] = StringTrimRight($ret[0], 1)' & @CRLF & _
			'__ArrayDelete($temp_split, 1)' & @CRLF & _
			'$ret[1] = StringRight($temp_split[1], StringLen($temp_split[1]) - 14)' & @CRLF & _
			'__ArrayDelete($temp_split, 1)' & @CRLF & _
			'__ArrayDelete($temp_split, 1)' & @CRLF & _
			'__ArrayDelete($temp_split, 0)' & @CRLF & _
			'__ArrayDelete($temp_split, UBound($temp_split) - 1)' & @CRLF & _
			'$ret[2] = __ArrayToString($temp_split, @CRLF)' & @CRLF & _
			'Return $ret' & @CRLF & _
			'Else' & @CRLF & _
			'$name_pos = StringInStr($temp_split[1], "; name=""")' & @CRLF & _
			'$name = StringRight($temp_split[1], StringLen($temp_split[1]) -($name_pos + 7))' & @CRLF & _
			'$quote_pos = StringInStr($name, """")' & @CRLF & _
			'$name = StringLeft($name,(StringLen($name) -(StringLen($name) - $quote_pos)) - 1)' & @CRLF & _
			'If $name <> $sVar Then ContinueLoop' & @CRLF & _
			'Return $temp_split[3]' & @CRLF & _
			'EndIf' & @CRLF & _
			'Next' & @CRLF & _
			'Return 0' & @CRLF & _
			'EndFunc' & @CRLF & _
			'Func _Cookie($sVar)' & @CRLF & _
			'Local $varstring = $_Cookie_Raw' & @CRLF & _
			'If Not StringInStr($varstring, $sVar & "=") Then Return ""' & @CRLF & _
			'Local $vars = StringSplit($varstring, "&")' & @CRLF & _
			'Local $var_array' & @CRLF
	$prepend &= _
			'For $i = 1 To $vars[0]' & @CRLF & _
			'$var_array = StringSplit($vars[$i], "=")' & @CRLF & _
			'If $var_array[0] < 2 Then Return ""' & @CRLF & _
			'If $var_array[1] = $sVar Then Return $var_array[2]' & @CRLF & _
			'Next' & @CRLF & _
			'Return ""' & @CRLF & _
			'EndFunc' & @CRLF & _
			'Func _URLDecode($toDecode)' & @CRLF & _
			'Local $strChar = "", $iOne, $iTwo' & @CRLF & _
			'Local $aryHex = StringSplit($toDecode, "")' & @CRLF & _
			'For $i = 1 To $aryHex[0]' & @CRLF & _
			'If $aryHex[$i] = "%" Then' & @CRLF & _
			'$strChar &= Chr(Dec($aryHex[$i + 1] & $aryHex[$i + 2]))' & @CRLF & _
			'$i += 2' & @CRLF & _
			'Else' & @CRLF & _
			'$strChar &= $aryHex[$i]' & @CRLF & _
			'EndIf' & @CRLF & _
			'Next' & @CRLF & _
			'Return StringReplace($strChar, "+", " ")' & @CRLF & _
			'EndFunc' & @CRLF & _
			'Func __ArrayDelete(ByRef $avArray, $iElement)' & @CRLF & _
			'If Not IsArray($avArray) Then Return SetError(1, 0, 0)' & @CRLF & _
			'Local $iUBound = UBound($avArray, 1) - 1' & @CRLF & _
			'If Not $iUBound Then' & @CRLF & _
			'$avArray = ""' & @CRLF & _
			'Return 0' & @CRLF & _
			'EndIf' & @CRLF & _
			'If $iElement < 0 Then $iElement = 0' & @CRLF & _
			'If $iElement > $iUBound Then $iElement = $iUBound' & @CRLF & _
			'Switch UBound($avArray, 0)' & @CRLF & _
			'Case 1' & @CRLF & _
			'For $i = $iElement To $iUBound - 1' & @CRLF & _
			'$avArray[$i] = $avArray[$i + 1]' & @CRLF & _
			'Next' & @CRLF & _
			'ReDim $avArray[$iUBound]' & @CRLF & _
			'Case 2' & @CRLF & _
			'Local $iSubMax = UBound($avArray, 2) - 1' & @CRLF & _
			'For $i = $iElement To $iUBound - 1' & @CRLF & _
			'For $j = 0 To $iSubMax' & @CRLF & _
			'$avArray[$i][$j] = $avArray[$i + 1][$j]' & @CRLF & _
			'Next' & @CRLF & _
			'Next' & @CRLF & _
			'ReDim $avArray[$iUBound][$iSubMax + 1]' & @CRLF & _
			'Case Else' & @CRLF & _
			'Return SetError(3, 0, 0)' & @CRLF & _
			'EndSwitch' & @CRLF & _
			'Return $iUBound' & @CRLF & _
			'EndFunc' & @CRLF
	$prepend &= _
			'Func __ArrayToString(Const ByRef $avArray, $sDelim = "|", $iStart = 0, $iEnd = 0)' & @CRLF & _
			'If Not IsArray($avArray) Then Return SetError(1, 0, "")' & @CRLF & _
			'If UBound($avArray, 0) <> 1 Then Return SetError(3, 0, "")' & @CRLF & _
			'Local $sResult, $iUBound = UBound($avArray) - 1' & @CRLF & _
			'If $iEnd < 1 Or $iEnd > $iUBound Then $iEnd = $iUBound' & @CRLF & _
			'If $iStart < 0 Then $iStart = 0' & @CRLF & _
			'If $iStart > $iEnd Then Return SetError(2, 0, "")' & @CRLF & _
			'For $i = $iStart To $iEnd' & @CRLF & _
			'$sResult &= $avArray[$i] & $sDelim' & @CRLF & _
			'Next' & @CRLF & _
			'Return StringTrimRight($sResult, StringLen($sDelim))' & @CRLF & _
			'EndFunc' & @CRLF & _
			'Func inline($sText)' & @CRLF & _
			'ConsoleWrite($sText)' & @CRLF & _
			'EndFunc' & @CRLF & _
			'Func echo($sText)' & @CRLF & _
			'ConsoleWrite($sText & @CRLF)' & @CRLF & _
			'EndFunc' & @CRLF & _
			'Func echol($sText)' & @CRLF & _
			'ConsoleWrite($sText & "<br />" & @CRLF)' & @CRLF & _
			'EndFunc' & @CRLF & _
			'' & @CRLF & @CRLF
	$prepend &= $source
	Return $prepend
EndFunc   ;==>_prependSource

;error and exit
Func _error($str)
	ConsoleWrite($str)
	Exit
EndFunc   ;==>_error
