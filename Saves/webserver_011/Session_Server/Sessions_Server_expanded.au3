#include <Date.au3>
#include <Array.au3>
#include <File.au3>
#include <String.au3>

#cs
	Resources:
	Internet Assigned Number Authority - all Content-Types: http://www.iana.org/assignments/media-types/
	World Wide Web Consortium - An overview of the HTTP protocol: http://www.w3.org/Protocols/
	
	Credits:
	Manadar for starting on the webserver.
	Alek for adding POST and some fixes
	Creator for providing the "application/octet-stream" MIME type.
#ce

; // OPTIONS HERE //
Dim $sRootDir = IniRead(@ScriptDir & "\http.ini", "settings", "wwwdir", @ScriptDir & "\WWW") ; The absolute path to the root directory of the server.
Dim $sIP = IniRead(@ScriptDir & "\http.ini", "settings", "ip", @IPAddress1) ; ip address as defined by AutoIt
Dim $iPort = IniRead(@ScriptDir & "\http.ini", "settings", "port", 8080) ; the listening port
Dim $iMaxUsers = IniRead(@ScriptDir & "\http.ini", "settings", "maxusers", 15) ; Maximum number of users who can simultaneously get/post
Dim $bDirListing = IniRead(@ScriptDir & "\http.ini", "security", "dirlisting", 0)
Dim $bLogActivity = IniRead(@ScriptDir & "\http.ini", "security", "logactivity", 1)
Dim $sLogfile = IniRead(@ScriptDir & "\http.ini", "security", "logfile", @ScriptDir & "\LOG.txt")
Dim $sAuthINI = IniRead(@ScriptDir & "\http.ini", "security", "authfile", @ScriptDir & "\auth.ini")
Dim $sessionTimeout = IniRead(@ScriptDir & "\http.ini", "security", "sessiontimeout", 120000)
; // END OPTIONS //

; // Error files etc //
Dim $err_404 = @ScriptDir & "\extrafiles\" & IniRead(@ScriptDir & "\http.ini", "extrafiles", "404", "404.html")
Dim $err_auth = @ScriptDir & "\extrafiles\" & IniRead(@ScriptDir & "\http.ini", "extrafiles", "auth", "AccessDenied.html")
Dim $err_dirlist = @ScriptDir & "\extrafiles\" & IniRead(@ScriptDir & "\http.ini", "extrafiles", "dirlist", "DirList.html")
; // END error files //

; // GLOBAL VARIABLES //
Dim $sRequestType = ""
Dim $sHost = ""
Dim $sUserAgent = ""
Dim $sCookie = ""
Dim $aSocket[$iMaxUsers] ; Creates an array to store all the possible users
Dim $sBuffer[$iMaxUsers] ; All these users have buffers when sending/receiving, so we need a place to store those
Dim $aSessions[15][3]
; [n][0] = ID
; [n][1] = Authenticated
; [n][2] = Last load

; // END GLOBAL VARIABLES //

For $x = 0 To UBound($aSocket) - 1 ; Fills the entire socket array with -1 integers, so that the server knows they are empty.
	$aSocket[$x] = -1
	$aSessions[$x][0] = -1
Next

If $bLogActivity Then
	Dim $hFileHandle = FileOpen($sLogfile, 1)
EndIf

TCPStartup() ; AutoIt needs to initialize the TCP functions

$iMainSocket = TCPListen($sIP, $iPort) ;create main listening socket
If @error Then ; if you fail creating a socket, exit the application
	MsgBox(0x20, "AutoIt Webserver", "Unable to create a socket on port " & $iPort & ".") ; notifies the user that the HTTP server will not run
	Exit ; if your server is part of a GUI that has nothing to do with the server, you'll need to remove the Exit keyword and notify the user that the HTTP server will not work.
EndIf

If $bLogActivity Then _WriteLog("Server created on http://" & $sIP & ":" & $iPort &"/")

While 1
	$iNewSocket = TCPAccept($iMainSocket) ;; Tries to accept incoming connections

	If $iNewSocket >= 0 Then ; Verifies that there actually is an incoming connection
		For $x = 0 To UBound($aSocket) - 1 ;; Attempts to store the incoming connection
			If $aSocket[$x] = -1 Then
				$aSocket[$x] = $iNewSocket ;store the new socket
				ExitLoop
			EndIf
		Next
	EndIf

	For $x = 0 To UBound($aSocket) - 1 ; A big loop to receive data from everyone connected
		If $aSessions[$x][0] <> -1 Then ;Test to see if the session has timed out
			If $aSessions[$x][1] <> -2 Then
				If $aSessions[$x][2] <> -1 Then
					If TimerDiff($aSessions[$x][2]) > $sessionTimeout Then
						Local $Hour = "", $Mins = "", $Secs = ""
						_TicksToTime(Int(TimerDiff(StringTrimLeft($aSessions[$x][0], 5))), $Hour, $Mins, $Secs)
						If $bLogActivity Then _WriteLog("Session #" & $aSessions[$x][0] & " timed out after " & StringFormat("%02i:%02i:%02i", $Hour, $Mins, $Secs))
						$aSessions[$x][1] = -2
					EndIf
				EndIf
			EndIf
		EndIf
		If $aSocket[$x] = -1 Then ContinueLoop ;; if the socket is empty, it will continue to the next iteration, doing nothing
		$sNewData = TCPRecv($aSocket[$x], 1024) ; Receives a whole lot of data if possible
		If @error Then ;; Client has disconnected
			$aSocket[$x] = -1 ; Socket is freed so that a new user may join
			ContinueLoop ; Go to the next iteration of the loop, not really needed but looks oh so good
		ElseIf $sNewData Then ; data received
			$sBuffer[$x] &= $sNewData ;store it in the buffer
			If StringInStr(StringStripCR($sBuffer[$x]), @LF & @LF) Then ; if the request has ended ..
				;$aBuffer = StringSplit(StringStripCR($sBuffer[$x]), @LF)
				$aBuffer = StringSplit($sBuffer[$x], @LF)
				For $iR = 1 To $aBuffer[0]
					Select
						Case StringInStr($aBuffer[$iR], "HTTP/", 1)
							$sRequestType = StringStripWS(StringLeft($aBuffer[$iR], StringInStr($aBuffer[$iR], " ") - 1), 2) ;; gets the type of the request
							$sRequest = StringStripWS(StringTrimLeft(StringTrimRight($aBuffer[$iR], StringLen($aBuffer[$iR]) - StringInStr($aBuffer[$iR], " HTTP/", 1)), StringInStr($aBuffer[$iR], " /", 1)), 2) ;; let's see what file he actually wants
							If StringRight($sRequest, 1) <> "/" Then
								If Not StringInStr(StringRight($sRequest, 5), ".") Then
									$sRequest &= "/"
								EndIf
							EndIf
						Case StringInStr($aBuffer[$iR], "User-Agent", 1)
							$sUserAgent = StringStripWS(StringTrimLeft($aBuffer[$iR], StringInStr($aBuffer[$iR], ": ", 1)), 2)
						Case StringInStr($aBuffer[$iR], "Host: ", 1)
							$sHost = StringStripWS(StringTrimLeft($aBuffer[$iR], StringInStr($aBuffer[$iR], ": ", 1)), 2)
						Case StringInStr($aBuffer[$iR], "Cookie", 1)
							$sCookie = StringStripWS(StringTrimLeft($aBuffer[$iR], StringInStr($aBuffer[$iR], ": ", 1)), 3)
					EndSelect
				Next
				$iCookieTest = _Test($sCookie)
				If $iCookieTest = 0 Then ; User is new, we'll assign them a session ID.
					$aSeshID = Random(10000, 99999, 1) & TimerInit()
					$aSessions[$x][0] = $aSeshID
					$aSessions[$x][2] = TimerInit()
					;They're not logged in so we'll redirect them to the login page.
					_SendFile($err_auth, "text/html", $aSocket[$x], "auth=" & $aSeshID)
				Else
					If $aSessions[$x][1] = -2 Then
						$day = _DateDayOfWeek(@WDAY, 1) & ", " & @MDAY & "-" & _DateToMonth(@MON, 1) & "-" & @YEAR - 1 & " 00:00:00 GMT"
						_SendHTML("Session Timed Out", $aSocket[$x], "auth=" & $aSessions[$x][0] & "; expires=" & $day)
						$aSessions[$x][0] = -1
						$aSessions[$x][1] = -1
						$aSessions[$x][2] = -1
					ElseIf $aSessions[$x][1] = 1 Then
						$aSessions[$x][2] = TimerInit()
						If $sRequestType = "GET" Then ;; user wants to download a file or whatever ..
							_SendPage($sRequest, $sHost)
						ElseIf $sRequestType = "POST" Then ;; user has come to us with data, we need to parse that data and based on that do something special
							
							$aPOST = _Get_Post($sBuffer[$x]) ;; parses the post data
							
							$sName = _POST("Name", $aPOST) ; Like PHPs _POST, but it requires the second parameter to be the return value from _Get_Post
							$sComment = _POST("Comment", $aPOST) ;; Gets the comment
							
							_POST_ConvertString($sName) ;; Needs to convert the POST HTTP string into a normal string
							_POST_ConvertString($sComment) ;; same ..
							
							FileWrite($sRootDir & "\index.html", "<br />" & $sName & " made comment: " & $sComment) ;Ofcourse, in real situations you have to prevent people to use HTML/PHP/Javascript etc. in their comments.
							;; The last line adds whatever Name:Comment said in the root file .. this creates some sort of chatty effect
							
							_SendFile($sRootDir & "\index.html", "text/html", $aSocket[$x]) ; Sends back the new file we just created
						EndIf
					Else ; They are not logged in.
						If $sRequestType = "POST" Then
							$aPOST = _Get_Post($sBuffer[$x]) ;; parses the post data
							$sUser = _POST("user", $aPOST) ; Like PHPs _POST, but it requires the second parameter to be the return value from _Get_Post
							$sPass = _POST("pass", $aPOST) ;; Gets the comment
							_POST_ConvertString($sUser) ;; Needs to convert the POST HTTP string into a normal string
							_POST_ConvertString($sPass) ;; same ..
							If _AuthTest($sUser, $sPass) Then
								$aSessions[$x][1] = 1
								If $bLogActivity Then _WriteLog("User '" & $sUser & "' authenticated themselves with session ID#" & $aSessions[$x][0])
								_SendPage($sRequest, $sHost)
							Else
								_SendFile($err_auth, "text/html", $aSocket[$x], "auth=" & $aSeshID)
							EndIf
						Else
							_SendFile($err_auth, "text/html", $aSocket[$x], "auth=" & $aSeshID)
						EndIf
					EndIf
					$sRequestType = ""
					$sHost = ""
					$sUserAgent = ""
					$sCookie = ""
					$sBuffer[$x] = "" ;; clears the buffer because we just used to buffer and did some actions based on them
					TCPCloseSocket($aSocket[$x]) ;; we have defined connection: close, so we close the connection
					$aSocket[$x] = -1 ;; reset the socket so that we may accept new clients
				EndIf
			EndIf
		EndIf
	Next
	Sleep(10)
WEnd

Func _POST_ConvertString(ByRef $sString) ;; converts any characters like %20 into space 8)
	$sString = StringReplace($sString, '+', ' ')
	StringReplace($sString, '%', '')
	For $t = 0 To @extended
		$Find_Char = StringLeft(StringTrimLeft($sString, StringInStr($sString, '%')), 2)
		$sString = StringReplace($sString, '%' & $Find_Char, Chr(Dec($Find_Char)))
	Next
EndFunc   ;==>_POST_ConvertString

Func _SendFile($sAddress, $sType, $sSocket, $cookie = "") ;; Sends a file back to the client on X socket, with X mime-type
	Local $hFile, $sImgBuffer, $sPacket, $a
	If $cookie <> "" Then $cookie = "Set-Cookie: " & $cookie & @CRLF
	
	$hFile = FileOpen($sAddress, 16)
	$sImgBuffer = FileRead($hFile)
	FileClose($hFile)

	$sPacket = Binary("HTTP/1.1 200 OK" & @CRLF & _
			"Server: ManadarX/1.3.26 (" & @OSVersion & ") AutoIt " & @AutoItVersion & @CRLF & _
			"Connection: close" & @CRLF & _
			$cookie & _
			"Content-Type: " & $sType & @CRLF & _
			@CRLF)
	TCPSend($sSocket, $sPacket)

	While BinaryLen($sImgBuffer) ;LarryDaLooza's idea to send in chunks to reduce stress on the application
		$a = TCPSend($sSocket, $sImgBuffer)
		$sImgBuffer = BinaryMid($sImgBuffer, $a + 1, BinaryLen($sImgBuffer) - $a)
	WEnd

	$sPacket = Binary(@CRLF & _
			@CRLF)
	TCPSend($sSocket, $sPacket)
	TCPCloseSocket($sSocket)
EndFunc   ;==>_SendFile

Func _SendError($sSocket, $error, $file, $cookie = "") ;; Sends back a basic 404 error
	Switch $error
		Case 404
			$text = StringReplace(FileRead($err_404), "<<PAGE>>", $file)
			_SendHTML($text, $sSocket, $cookie)
	EndSwitch
EndFunc   ;==>_SendError

Func _SendHTML($sHTML, $sSocket, $cookie = "")
	Local $iLen, $sPacket, $sSplit
	If $cookie <> "" Then $cookie = "Set-Cookie: " & $cookie & @CRLF

	$iLen = StringLen($sHTML)
	$sPacket = Binary("HTTP/1.1 200 OK" & @CRLF & _
			"Server: ManadarX/1.0 (" & @OSVersion & ") AutoIt " & @AutoItVersion & @CRLF & _
			"Connection: close" & @CRLF & _
			"Content-Lenght: " & $iLen & @CRLF & _
			$cookie & _
			"Content-Type: text/html" & @CRLF & _
			@CRLF & _
			$sHTML)
	$sSplit = StringSplit($sPacket, "")
	$sPacket = ""
	For $i = 1 To $sSplit[0]
		If Asc($sSplit[$i]) <> 0 Then ; Just make sure we don't send any null bytes, because they show up as ???? in your browser.
			$sPacket = $sPacket & $sSplit[$i]
		EndIf
	Next
	TCPSend($sSocket, $sPacket)
EndFunc   ;==>_SendHTML

Func _Get_Post($s_Buffer) ;; parses incoming POST data
	Local $sTempPost, $sLen, $sPostData, $sTemp

	;Get the lenght of the data in the POST
	$sTempPost = StringTrimLeft($s_Buffer, StringInStr($s_Buffer, "Content-Length:"))
	$sLen = StringTrimLeft($sTempPost, StringInStr($sTempPost, ": "))

	;Create the base struck
	$sPostData = StringSplit(StringRight($s_Buffer, $sLen), "&")

	Local $sReturn[$sPostData[0] + 1][2]

	For $t = 1 To $sPostData[0]
		$sTemp = StringSplit($sPostData[$t], "=")
		If $sTemp[0] >= 2 Then
			$sReturn[$t][0] = $sTemp[1]
			$sReturn[$t][1] = $sTemp[2]
		EndIf
	Next

	Return $sReturn
EndFunc   ;==>_Get_Post

Func _POST($sName, $sArray) ;; Returns a POST variable based on their name and not their array index. This function basically makes up for the lack of associative arrays in Au3
	For $i = 1 To UBound($sArray) - 1
		If $sArray[$i][0] = $sName Then
			Return $sArray[$i][1]
		EndIf
	Next
	Return ""
EndFunc   ;==>_POST

Func _Test($sCookieString)
	For $i = 0 To UBound($aSessions) - 1
		If "auth=" & $aSessions[$i][0] = $sCookieString Then
			Return 1
		EndIf
	Next
	Return 0
EndFunc   ;==>_Test

Func _AuthTest($sUser, $sPass)
	If IniRead($sAuthINI, "auth", $sUser, -1) == $sPass Then
		Return 1
	Else
		Return 0
	EndIf
EndFunc   ;==>_AuthTest

Func _WriteLog($string)
	$str = _Now() & @TAB & $string
	FileWriteLine($hFileHandle, $str)
	ConsoleWrite ($str & @CRLF)
EndFunc   ;==>_WriteLog

Func _DirList($sRequest, $host)
	Local $sHTML, $repeat2, $dir
	$dir = $sRequest
	$sfile = FileRead($err_dirlist)
	$repeat = _StringBetween($sfile, "<<STARTREPEAT>>", "<<ENDREPEAT>>")
	$repeat = $repeat[0]
	$sRequest2 = StringReplace($sRequest, "/", "\")
	$_Frarra = _FileListToArray($sRootDir & $sRequest2)
	If @error = 4 Then
		$repeat2 = StringReplace($repeat, "<<NAME>>", "No files/folders found")
		$repeat2 = StringReplace($repeat2, "<<ICON>>", "")
		$repeat2 = StringReplace($repeat2, "<<MODIFIED>>", "")
		$repeat2 = StringReplace($repeat2, "<<SIZE>>", "")
	Else
		If (StringLeft($sRequest, 1) = "/") And (StringRight($sRequest, 1) = "/") And ($sRequest <> "/") Then
			$repeat2 = StringReplace($repeat, "<<NAME>>", "<a href='" & StringTrimRight($sRequest, StringLen($sRequest) - StringInStr($sRequest, "/", Default, -2)) & "'>Parent Directory</a>")
			$repeat2 = StringReplace($repeat2, "<<ICON>>", "")
			$repeat2 = StringReplace($repeat2, "<<MODIFIED>>", "")
			$repeat2 = StringReplace($repeat2, "<<SIZE>>", "")
		EndIf
		For $i = 1 To $_Frarra[0]
			If StringInStr(FileGetAttrib($sRootDir & "\" & $_Frarra[$i]), "D") Then
				$text = StringReplace($repeat, "<<NAME>>", "<a href='" & $dir & $_Frarra[$i] & "/'>" & $_Frarra[$i] & "</a>")
				$text = StringReplace($text, "<<ICON>>", "DIR")
				$text = StringReplace($text, "<<SIZE>>", DirGetSize($sRootDir & "\" & $_Frarra[$i]))
			Else
				$text = StringReplace($repeat, "<<NAME>>", "<a href='" & $dir & $_Frarra[$i] & "'>" & $_Frarra[$i] & "</a>")
				$text = StringReplace($text, "<<ICON>>", StringTrimLeft($_Frarra[$i], StringInStr($_Frarra[$i], ".")))
				$text = StringReplace($text, "<<SIZE>>", FileGetSize($sRootDir & "\" & $_Frarra[$i]))
			EndIf
			$t = FileGetTime($sRootDir & "\" & $_Frarra[$i])
			$text = StringReplace($text, "<<MODIFIED>>", $t[0] & "/" & $t[1] & "/" & $t[2] & " " & $t[3] & ":" & $t[4] & ":" & $t[5])
			$repeat2 &= $text
		Next
	EndIf
	$sHTML = StringReplace(StringReplace($sfile, "<<PATH>>", $sRequest), "<<STARTREPEAT>>" & $repeat & "<<ENDREPEAT>>", $repeat2)
	Return $sHTML
EndFunc   ;==>_DirList

Func _SendPage($sRequest, $sHost)
	If StringInStr(StringReplace($sRequest, "\", "/"), "/.") Then ;; Disallow any attempts to go back a folder
		_SendError($aSocket[$x], 404, $sRequest) ;; sends back an error
	Else
		If StringRight($sRequest, 1) = "/" Then ;; user has requested the root
			If $bDirListing == 1 Then
				MsgBox(0, "", $bDirListing)
				_SendHTML(_DirList($sRequest, $sHost), $aSocket[$x])
				Return
			Else
				$var = IniReadSection(@ScriptDir & "\http.ini", "indexfiles")
				If @error Then
					$sRequest = "/index.html" ;; instead of root we'll give him the index page
				Else
					For $i = 1 To $var[0][0]
						$a = 0
						$filename = StringLeft($var[$i][1], StringInStr($var[$i][1], ".", Default, -1) - 1)
						$path = StringTrimLeft($var[$i][1], StringInStr($var[$i][1], ".", Default, -1))
						$salip = StringSplit($path, "|")
						For $ab = 1 To $salip[0]
							$file = $filename & "." & $salip[$ab]
							If FileExists($sRootDir & "\" & StringReplace($sRequest & $file, "/", "\")) Then
								$sRequest &= $file
								$a = 1
								ExitLoop
							EndIf
						Next
						If $a Then ExitLoop
					Next
				EndIf
			EndIf
		EndIf
		$sRequest = StringReplace($sRequest, "/", "\") ; convert HTTP slashes to windows slashes, not really required because windows accepts both
		If FileExists($sRootDir & $sRequest) Then ;; makes sure the file that the user wants exists
			_SendFile($sRootDir & "\" & $sRequest, _GetMIME($sRequest), $aSocket[$x])
		Else
			_SendError($aSocket[$x], 404, $sRequest) ;; File does not exist, so we'll send back an error..
		EndIf
	EndIf
EndFunc   ;==>_SendPage

Func _GetMIME($sRequest)
	Local $mime
	$sFileType = StringTrimLeft($sRequest, StringInStr($sRequest, ".")) ;; determines the file type, so that we may choose what MIME type to use
	Switch $sFileType
		Case "html", "htm" ;; in case of normal HTML files
			$mime = "text/html"
		Case "css" ;; in case of style sheets
			$mime = "text/css"
		Case "jpg", "jpeg" ;; for common images
			$mime = "image/jpeg"
		Case "png" ;; another common image format
			$mime = "image/png"
		Case Else ; this is for .exe, .zip, or anything else that is not supported is downloaded to the client using a application/octet-stream
			$mime = "application/octet-stream"
	EndSwitch
	Return $mime
EndFunc   ;==>_GetMIME