Opt("MustDeclareVars", 1) ; 0=no, 1=require pre-declare
#include <Array.au3>

;------------
Func IPAddress($IPAddress)
	Dim $Results[1] ;The array to hold the final results
	Local $Array = StringSplit($IPAddress, "+")

	If $Array[0] <> 2 Then
		Debug("No count value found. Testing and returning the value. " & @ScriptLineNumber)
		Local $T = TestIP($Array[1])
		If StringInStr($T, "ERROR0") = 0 Then
			_ArrayAdd($Results, "TestIP failed " & $Array[1] & " Return " & $T)
			Return $Results
		EndIf
		_ArrayAdd($Results, $Array[1])
		_ArrayDelete($Results, 0) ; This returns the count entry
		Return $Results
	EndIf

	Local $T = TestIP($Array[1])
	If StringInStr($T, "ERROR0") = 0 Then
		_ArrayAdd($Results, "TestIP failed " & $Array[1] & " Return " & $T)
		Return $Results
	EndIf
	
	Local $IPArray = StringSplit($Array[1], ".")
	Local $Count = $Array[2]
	Local $IPAddressHEX = Hex($IPArray[1], 2) & Hex($IPArray[2], 2) & Hex($IPArray[3], 2) & Hex($IPArray[4], 2)
	Local $IPAddressDEC = Dec($IPAddressHEX)
	
	For $X = 0 To $Count
		Local $tmp3 = Hex($IPAddressDEC)
		Local $IPout = Dec(StringMid($tmp3, 1, 2)) & "." & Dec(StringMid($tmp3, 3, 2)) & "." & Dec(StringMid($tmp3, 5, 2)) & "." & Dec(StringMid($tmp3, 7, 2))
		_ArrayAdd($Results, $IPout)
		$IPAddressDEC += 1
	Next
	_ArrayDelete($Results, 0) ; deletes a blank entry at the begining

	Return $Results
	
EndFunc   ;==>IPAddress

;------------
Func TestIP($IPAddress)
	Local $IPArray = StringSplit($IPAddress, ".") ;This is the ipaddress octets split on .
	
	If $IPArray[0] <> 4 Then
		Return "ERROR1  Not enough octets. 4 Required, " & $IPArray[0] & " Found."
	EndIf
	
	_ArrayDelete($IPArray, 0) ; This returns the count entry
	
	For $T In $IPArray  ;verify that the octet values are within range
		If $T < 0 Or $T > 255 Then
			Return "ERROR2 octet out of range (0 to 255). " & $T
		EndIf
	Next
	
	Return "ERROR0" ;good address
EndFunc   ;==>TestIP
;------------
Func Debug($msg, $ShowMsgBox = -1, $timeout = 0)
	$msg = "Debug: " & $msg
	DllCall("kernel32.dll", "none", "OutputDebugString", "str", $msg)
	ConsoleWrite($msg & @CRLF)
	If $ShowMsgBox > -1 Then MsgBox($ShowMsgBox, "Debug", $msg, $timeout)
EndFunc   ;==>Debug
;------------