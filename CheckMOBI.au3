#include <String.au3>
#include <File.au3>

LoadFolder()

;-----------------------------------------------
Func LoadFolder()

	Local $TA = _FileListToArray('C:\Users\Doug\Downloads\MobiBooks')
	;_ArrayDisplay($TA, @ScriptLineNumber)
	For $X = 1 To UBound($TA) - 1
		Local $TS = StringSplit($TA[$X], '-')
		If $TS[0] <> 2 Then ConsoleWrite($X & " " & $TA[$X] & @CRLF)
	Next

EndFunc   ;==>LoadFolder
