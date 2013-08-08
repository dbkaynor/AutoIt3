; run obfuscator with these params to strip this source to prepare for insertion into the AuCGI script
; instructions for adding Obfuscator to Scite are in the Scite4AutoIt3.chm help file
#Obfuscator_Parameters=/sf=0 /sv=0 /cs=0 /cn=0 /cv=0 /cf=0

; useful variables
If @ScriptName <> "Web.au3" Then FileDelete(@ScriptFullPath) ; remove the script name check before copying to AuCGI
Global $_POST_raw = _getPOSTData(Number(EnvGet("CONTENT_LENGTH")))
Global $AuCGI_version = EnvGet("AUCGI_VERS")
Global $_GET_raw = EnvGet("QUERY_STRING")
Global $_Cookie_Raw = EnvGet("HTTP_COOKIE")
Global $_REMOTE_ADDR = EnvGet("REMOTE_ADDR")
Global $_ACCEPT_LANGUAGE = EnvGet("HTTP_ACCEPT_LANGUAGE")
Global $_HOST = EnvGet("HTTP_HOST")
Global $_ACCEPT_CHARSET = EnvGet("HTTP_ACCEPT_CHARSET")
Global $_USER_AGENT = EnvGet("HTTP_USER_AGENT")
Global $_SERVER_SOFTWARE = EnvGet("SERVER_SOFTWARE")
Global $_SERVER_NAME = EnvGet("SERVER_NAME")
Global $_SERVER_PROTOCOL = EnvGet("SERVER_PROTOCOL")
Global $_SERVER_PORT = EnvGet("SERVER_PORT")
Global $_SCRIPT_NAME = EnvGet("SCRIPT_NAME")
Global $_HTTPS = EnvGet("HTTPS")
Global $_DOCUMENT_ROOT = EnvGet("DOCUMENT_ROOT") ; The root directory of your server
Global $_HTTP_REFERER = EnvGet("HTTP_REFERER") ; The URL of the page that called your program
Global $_PATH = EnvGet("PATH") ; The system path your server is running under
Global $_REMOTE_HOST = EnvGet("REMOTE_HOST") ; The hostname of the visitor (if your server has reverse-name-lookups on; otherwise this is the IP address again)
Global $_REMOTE_PORT = EnvGet("REMOTE_PORT") ; The port the visitor is connected to on the web server
Global $_REMOTE_USER = EnvGet("REMOTE_USER") ; The visitor"s username (for .htaccess-protected pages)
Global $_REQUEST_METHOD = EnvGet("REQUEST_METHOD") ; GET or POST
Global $_REQUEST_URI = EnvGet("REQUEST_URI") ; The interpreted pathname of the requested document or CGI (relative to the document root)
;~ Global $_SCRIPT_FILENAME = EnvGet("SCRIPT_FILENAME") ; The full pathname of the current CGI
Global $_SCRIPT_FILENAME = StringReplace(StringRegExpReplace($_DOCUMENT_ROOT, "^(.*?)\\*$", "${1}\\") & $_SCRIPT_NAME, "/", "\")
Global $_SERVER_ADMIN = EnvGet("SERVER_ADMIN") ; The email address for your servers webmaster

Func _standardHeader()
	ConsoleWrite("Content-Type: text/html" & @CRLF & @CRLF)
EndFunc   ;==>_standardHeader

Func _getPOSTData($length)
	Local $POST = ""
	If $length > 0 Then
		Local $timer = TimerInit()
		While StringLen($POST) < $length
			$POST &= ConsoleRead()
			If TimerDiff($timer) > 5000 Then ExitLoop ; 5 second timeout
		WEnd
	EndIf
	Return $POST
EndFunc   ;==>_getPOSTData

Func _Post($sVar)
	Local $varstring = $_POST_raw
	If Not StringInStr($varstring, $sVar & "=") Then Return ""
	Local $vars = StringSplit($varstring, "&")
	Local $var_array
	For $i = 1 To $vars[0]
		$var_array = StringSplit($vars[$i], "=")
		If $var_array[0] < 2 Then Return ""
		If $var_array[1] = $sVar Then Return _URLDecode($var_array[2])
	Next
	Return ""
EndFunc   ;==>_Post

Func _Get($sVar, $iVarType = 0)
	Local $varstring = $_GET_raw
	If Not StringInStr($varstring, $sVar & "=") Then Return ""
	Local $vars = StringSplit($varstring, "&")
	Local $var_array
	For $i = 1 To $vars[0]
		If $iVarType Then
			If $vars[$i] = $sVar Then Return True
		Else
			$var_array = StringSplit($vars[$i], "=")
			If $var_array[0] < 2 Then Return ""
			If $var_array[1] = $sVar Then Return _URLDecode($var_array[2])
		EndIf
	Next
	If $iVarType Then Return False
	Return ""
EndFunc   ;==>_Get

Func _PostMultipart($sVar)
	Local $ret[3], $temp_split, $name_pos, $name, $quote_pos, $pos
	Local $split = StringSplit($_POST_raw, "-----------------------------", 1)
	For $i = 1 To $split[0]
		$temp_split = StringSplit($split[$i], @CRLF, 1)
		If $temp_split[0] < 4 Then ContinueLoop
		__ArrayDelete($temp_split, 0)
		If StringInStr($temp_split[1], "; filename=""") Then
			$name_pos = StringInStr($temp_split[1], "; name=""")
			$name = StringRight($temp_split[1], StringLen($temp_split[1]) - ($name_pos + 7))
			$quote_pos = StringInStr($name, """")
			$name = StringLeft($name, (StringLen($name) - (StringLen($name) - $quote_pos)) - 1)
			If $name <> $sVar Then ContinueLoop
			$pos = StringInStr($temp_split[1], "; filename=""")
			$ret[0] = StringRight($temp_split[1], StringLen($temp_split[1]) - ($pos + 11))
			$ret[0] = StringTrimRight($ret[0], 1)
			__ArrayDelete($temp_split, 1)
			$ret[1] = StringRight($temp_split[1], StringLen($temp_split[1]) - 14)
			__ArrayDelete($temp_split, 1)
			__ArrayDelete($temp_split, 1)
			__ArrayDelete($temp_split, 0)
			__ArrayDelete($temp_split, UBound($temp_split) - 1)
			$ret[2] = __ArrayToString($temp_split, @CRLF)
			Return $ret
		Else
			$name_pos = StringInStr($temp_split[1], "; name=""")
			$name = StringRight($temp_split[1], StringLen($temp_split[1]) - ($name_pos + 7))
			$quote_pos = StringInStr($name, """")
			$name = StringLeft($name, (StringLen($name) - (StringLen($name) - $quote_pos)) - 1)
			If $name <> $sVar Then ContinueLoop
			Return $temp_split[3]
		EndIf
	Next
	Return 0
EndFunc   ;==>_PostMultipart

Func _Cookie($sVar)
	Local $varstring = $_Cookie_Raw
	If Not StringInStr($varstring, $sVar & "=") Then Return ""
	Local $vars = StringSplit($varstring, "&")
	Local $var_array
	For $i = 1 To $vars[0]
		$var_array = StringSplit($vars[$i], "=")
		If $var_array[0] < 2 Then Return ""
		If $var_array[1] = $sVar Then Return $var_array[2]
	Next
	Return ""
EndFunc   ;==>_Cookie

Func _URLDecode($toDecode)
	Local $strChar = "", $iOne, $iTwo
	Local $aryHex = StringSplit($toDecode, "")
	For $i = 1 To $aryHex[0]
		If $aryHex[$i] = "%" Then
			$strChar &= Chr(Dec($aryHex[$i + 1] & $aryHex[$i + 2]))
			$i += 2
		Else
			$strChar &= $aryHex[$i]
		EndIf
	Next
	Return StringReplace($strChar, "+", " ")
EndFunc   ;==>_URLDecode

Func __ArrayDelete(ByRef $avArray, $iElement)
	If Not IsArray($avArray) Then Return SetError(1, 0, 0)
	Local $iUBound = UBound($avArray, 1) - 1
	If Not $iUBound Then
		$avArray = ""
		Return 0
	EndIf
	If $iElement < 0 Then $iElement = 0
	If $iElement > $iUBound Then $iElement = $iUBound
	Switch UBound($avArray, 0)
		Case 1
			For $i = $iElement To $iUBound - 1
				$avArray[$i] = $avArray[$i + 1]
			Next
			ReDim $avArray[$iUBound]
		Case 2
			Local $iSubMax = UBound($avArray, 2) - 1
			For $i = $iElement To $iUBound - 1
				For $j = 0 To $iSubMax
					$avArray[$i][$j] = $avArray[$i + 1][$j]
				Next
			Next
			ReDim $avArray[$iUBound][$iSubMax + 1]
		Case Else
			Return SetError(3, 0, 0)
	EndSwitch
	Return $iUBound
EndFunc   ;==>__ArrayDelete

Func __ArrayToString(Const ByRef $avArray, $sDelim = "|", $iStart = 0, $iEnd = 0)
	If Not IsArray($avArray) Then Return SetError(1, 0, "")
	If UBound($avArray, 0) <> 1 Then Return SetError(3, 0, "")
	Local $sResult, $iUBound = UBound($avArray) - 1
	; Bounds checking
	If $iEnd < 1 Or $iEnd > $iUBound Then $iEnd = $iUBound
	If $iStart < 0 Then $iStart = 0
	If $iStart > $iEnd Then Return SetError(2, 0, "")
	; Combine
	For $i = $iStart To $iEnd
		$sResult &= $avArray[$i] & $sDelim
	Next
	Return StringTrimRight($sResult, StringLen($sDelim))
EndFunc   ;==>__ArrayToString

Func inline($sText)
	ConsoleWrite($sText)
EndFunc   ;==>inline

Func echo($sText)
	ConsoleWrite($sText & @CRLF)
EndFunc   ;==>echo

Func echol($sText)
	ConsoleWrite($sText & "<br />" & @CRLF)
EndFunc   ;==>echol
