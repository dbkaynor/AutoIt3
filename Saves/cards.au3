;"This function uses code developed by Paul Campbell (PaulIA) for the Auto3Lib project"
; Created by Kohr

#include <GUIConstants.au3>

Opt("GUIOnEventMode", 1)
Opt("MustDeclareVars", 1)

;Globals for Auto3Lib
Global $gaLibDlls[64][2] = [[0, 0]]
Global Const $RECT = "int;int;int;int"
Global Const $RECT_LEFT = 1
Global Const $RECT_TOP = 2
Global Const $RECT_RIGHT = 3
Global Const $RECT_BOTTOM = 4

Global $GUImain = GUICreate("Cards", 950, 450)

Local $a = GUICtrlCreateButton("Clear", 110, 400, 50, 20)
GUICtrlSetOnEvent($a, "ClearBoard")

Local $b = GUICtrlCreateButton("Cards", 10, 400, 50, 20)
GUICtrlSetOnEvent($a, "ClearBoard")
GUICtrlSetOnEvent($b, "Cards")

Local $c = GUICtrlCreateButton("Backs", 60, 400, 50, 20)
GUICtrlSetOnEvent($a, "ClearBoard")
GUICtrlSetOnEvent($c, "BackCard")

Local $d = GUICtrlCreateButton("Move", 160, 400, 50, 20)
GUICtrlSetOnEvent($a, "ClearBoard")
GUICtrlSetOnEvent($d, "MoveCard")

Local $e = GUICtrlCreateButton("Bigcard", 210, 400, 50, 20)
GUICtrlSetOnEvent($a, "ClearBoard")
GUICtrlSetOnEvent($e, "BigCard")

Local $f = GUICtrlCreateButton("Exit", 260, 400, 50, 20)
GUICtrlSetOnEvent($f, "CloseClicked")

GUISetState(@SW_SHOW, $GUImain)
GUISetOnEvent($GUI_EVENT_CLOSE, "CloseClicked")

While 1
	Sleep(200)
WEnd

Func CloseClicked()
	Exit
EndFunc   ;==>CloseClicked

Func Cards()
	Local $hDLL = DllOpen("cards.dll")
	Local $hdc = DllCall("user32.dll", "int", "GetDC", "hwnd", $GUImain)
	$hdc = $hdc[0]
	DllCall($hDLL, "int", "cdtInit", "int*", 0, "int*", 0)
	Local $card = 0
	Local $x = 0
	Local $y = 0
	For $i = 1 To 13
		For $j = 1 To 4
			DllCall($hDLL, "int", "cdtDraw", "int", $hdc, "int", $x, "int", $y, "int", $card, "int", 0, "int", 0)
			$card += 1
			$y += 95
		Next
		$x += 70
		$y = 0
	Next
	DllCall($hDLL, "none", "cdtTerm")
	DllClose($hDLL)
	DllCall("user32.dll", "int", "ReleaseDC", "hwnd", $GUImain, "int", $hdc)
EndFunc   ;==>Cards

Func BackCard()
	Local $hDLL = DllOpen("cards.dll")
	Local $hdc = DllCall("user32.dll", "int", "GetDC", "hwnd", $GUImain)
	$hdc = $hdc[0]
	DllCall($hDLL, "int", "cdtInit", "int*", 0, "int*", 0)
	Local $card = 53
	Local $x = 0
	Local $y = 0
	For $i = 1 To 4
		For $j = 1 To 4
			DllCall($hDLL, "int", "cdtDraw", "int", $hdc, "int", $x, "int", $y, "int", $card, "int", 1, "int", 0)
			$card += 1
			$y += 95
		Next
		$x += 70
		$y = 0
	Next
	DllCall($hDLL, "none", "cdtTerm")
	DllClose($hDLL)
	DllCall("user32.dll", "int", "ReleaseDC", "hwnd", $GUImain, "int", $hdc)
EndFunc   ;==>BackCard

Func ClearBoard()
	Local $x = 0
	Local $y = 0
	Local $rRect
	$rRect = _DllStructCreate($RECT)
	_DllStructSetData($rRect, $RECT_TOP, $y)
	_DllStructSetData($rRect, $RECT_LEFT, $x)
	_DllStructSetData($rRect, $RECT_BOTTOM, $y + 450)
	_DllStructSetData($rRect, $RECT_RIGHT, $x + 950)
	_InvalidateRect($GUImain, $rRect, True)
EndFunc   ;==>ClearBoard

Func MoveCard()
	Local $rRect
	$rRect = _DllStructCreate($RECT)
	Local $hDLL = DllOpen("cards.dll")
	Local $hdc = DllCall("user32.dll", "int", "GetDC", "hwnd", $GUImain)
	$hdc = $hdc[0]
	DllCall($hDLL, "int", "cdtInit", "int*", 0, "int*", 0)
	Local $card = 1
	Local $x = 0
	Local $y = 0
	For $i = 1 To 10
		For $j = 1 To 10
			DllCall($hDLL, "int", "cdtDraw", "int", $hdc, "int", $x, "int", $y, "int", $card, "int", 0, "int", 0)
			Sleep(20)
			_DllStructSetData($rRect, $RECT_TOP, $y)
			_DllStructSetData($rRect, $RECT_LEFT, $x)
			_DllStructSetData($rRect, $RECT_BOTTOM, $y + 96)
			_DllStructSetData($rRect, $RECT_RIGHT, $x + 71)
			_InvalidateRect($GUImain, $rRect, True)
			$y += 10
		Next
		$x += 40
		$y = 0
	Next
	DllCall($hDLL, "none", "cdtTerm")
	DllClose($hDLL)
	DllCall("user32.dll", "int", "ReleaseDC", "hwnd", $GUImain, "int", $hdc)
EndFunc   ;==>MoveCard

Func BigCard()
	Local $hDLL = DllOpen("cards.dll")
	Local $hdc = DllCall("user32.dll", "int", "GetDC", "hwnd", $GUImain)
	$hdc = $hdc[0]
	DllCall($hDLL, "int", "cdtInit", "int*", 0, "int*", 0)
	Local $card = 34
	Local $x = 0
	Local $y = 0
	DllCall($hDLL, "int", "cdtDrawExt", "int", $hdc, "int", $x, "int", $y, "int", 150, "int", 170, "int", $card, "int", 0, "int", 0)
	$card += 1
	$y += 95
	DllCall($hDLL, "none", "cdtTerm")
	DllClose($hDLL)
	DllCall("user32.dll", "int", "ReleaseDC", "hwnd", $GUImain, "int", $hdc)
EndFunc   ;==>BigCard

#region --- Auto3Lib START ---
Func _DllOpen($sFileName)
	Local $hDLL
	Local $iIndex

	$sFileName = StringUpper($sFileName)
	For $iIndex = 1 To $gaLibDlls[0][0]
		If $sFileName = $gaLibDlls[$iIndex][0] Then
			Return $gaLibDlls[$iIndex][1]
		EndIf
	Next

	$hDLL = DllOpen($sFileName)
	$iIndex = $gaLibDlls[0][0] + 1
	$gaLibDlls[0][0] = $iIndex
	$gaLibDlls[$iIndex][0] = $sFileName
	$gaLibDlls[$iIndex][1] = $hDLL
	Return $hDLL
EndFunc   ;==>_DllOpen

Func _DllStructCreate($sStruct, $pPointer = 0)
	Local $rResult

	If $pPointer = 0 Then
		$rResult = DllStructCreate($sStruct)
	Else
		$rResult = DllStructCreate($sStruct, $pPointer)
	EndIf
	Return $rResult
EndFunc   ;==>_DllStructCreate

Func _DllStructSetData($rStruct, $iElement, $vValue, $iIndex = -1)
	Local $rResult

	If $iIndex = -1 Then
		$rResult = DllStructSetData($rStruct, $iElement, $vValue)
	Else
		$rResult = DllStructSetData($rStruct, $iElement, $vValue, $iIndex)
	EndIf
	Return $rResult
EndFunc   ;==>_DllStructSetData

Func _DllStructGetPtr($rStruct, $iElement = 0)
	Local $rResult

	If $iElement = 0 Then
		$rResult = DllStructGetPtr($rStruct)
	Else
		$rResult = DllStructGetPtr($rStruct, $iElement)
	EndIf
	Return $rResult
EndFunc   ;==>_DllStructGetPtr

Func _InvalidateRect($hWnd, $rRect = 0, $bErase = True)
	Local $pRect
	Local $aResult
	Local $hUser32

	If $rRect <> 0 Then $pRect = _DllStructGetPtr($rRect)
	$hUser32 = _DllOpen("User32.dll")
	$aResult = DllCall($hUser32, "int", "InvalidateRect", "hwnd", $hWnd, "ptr", $pRect, "int", $bErase)
	Return ($aResult[0] <> 0)
EndFunc   ;==>_InvalidateRect