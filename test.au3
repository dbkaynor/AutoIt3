; *******************************************************
; Example 1 - Create browser windows with each of the example pages displayed.
;				The object variable returned can be used just as the object
;				variables returned by _IECreate or _IEAttach
; *******************************************************
#include <Array.au3>
#include <ButtonConstants.au3>
#include <Constants.au3>
#include <GUIConstantsEx.au3>
#include <Misc.au3>
#include <StaticConstants.au3>
#include <String.au3>
#include <WindowsConstants.au3>
#include <_DougFunctions.au3>
#include <IE.au3>

; "Overview of Software Power Consumers" table is the 4th table

;Row values to use (X)
Const $Wakeups = 1
Const $Description = 6

Local $oIE = _IECreate("C:\Users\dbkaynox\Downloads\powertop-hwdwcode-full.4.html")
Local $oTable = _IETableGetCollection($oIE, 4)
Local $aTableData = _IETableWriteToArray($oTable)

;$aTableData[x][y]  x=row, y=column
ConsoleWrite(@ScriptLineNumber & " " & $aTableData[$Wakeups][0] & @CRLF)
ConsoleWrite(@ScriptLineNumber & " " & $aTableData[$Description][0] & @CRLF)

For $a = 0 To UBound($aTableData, 2) - 1
	If StringInStr($aTableData[$Description][$a], '[40] i915') Then
		ConsoleWrite(@ScriptLineNumber & " " & $aTableData[$Description][$a] & "   " & $aTableData[$Wakeups][$a] & @CRLF)
	EndIf
	If StringInStr($aTableData[$Description][$a], 'kwork') Then
		ConsoleWrite(@ScriptLineNumber & " " & $aTableData[$Description][$a] & "   " & $aTableData[$Wakeups][$a] & @CRLF)
	EndIf
	If StringInStr($aTableData[$Description][$a], 'weston') Then
		ConsoleWrite(@ScriptLineNumber & " " & $aTableData[$Description][$a] & "   " & $aTableData[$Wakeups][$a] & @CRLF)
	EndIf
	If StringInStr($aTableData[$Description][$a], 'gears') Then
		ConsoleWrite(@ScriptLineNumber & " " & $aTableData[$Description][$a] & "   " & $aTableData[$Wakeups][$a] & @CRLF)
	EndIf
Next

_ArrayDisplay($aTableData, @ScriptLineNumber)