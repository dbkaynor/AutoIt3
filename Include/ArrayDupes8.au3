; ArrayDupes8.au3   

#AutoIt3Wrapper_AU3Check_Parameters= -d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
Func _ArrayCompare(ByRef $arrItems, ByRef $arrItems2, $iDetails = 0, $iType = 0, $istarti = 0, $istartj = 0) ;$iType=1 returns the Dupes instead of non-matching lines
	If @OSTYPE = "WIN32_WINDOWS"  Then Return 0
	Local $j = $istartj, $i = $istarti, $objDictionary = ObjCreate("Scripting.Dictionary"), $objDictionary2 = ObjCreate("Scripting.Dictionary"), $arrItemsExtra[1], $arrItemsDupes[1]
	For $i = $istarti To UBound($arrItems) - 1
		If Not $objDictionary.Exists($arrItems[$i]) Then $objDictionary.Add($arrItems[$i], $arrItems[$i])
	Next
	For $j = $istartj To UBound($arrItems2) - 1
		If Not $objDictionary2.Exists($arrItems2[$j]) Then $objDictionary2.Add($arrItems2[$j], $arrItems2[$j])
	Next
	ReDim $arrItemsExtra[$i + $j + 2]
	ReDim $arrItemsDupes[$i + $j + 2]
	;================================
	Local $k = 0, $m = 1, $n = 0
	If $iType Then ;only want dupes
		For $i = $istarti To UBound($arrItems) - 1
			If $objDictionary2.Exists($arrItems[$i]) Then
				$arrItemsDupes[$k] = $arrItems[$i]
				If $iDetails Then $arrItemsDupes[$k] &= "|Dupes|" & $i
				$k += 1
			EndIf
		Next
		If $k = 0 Then ReDim $arrItemsDupes[1 ]
		If $k > 0 Then ReDim $arrItemsDupes[$k ]
	Else ;only want non-dupe lines
		For $i = $istarti To UBound($arrItems) - 1
			If Not $objDictionary2.Exists($arrItems[$i]) Then
				$arrItemsExtra[$m] = $arrItems[$i]
				If $iDetails Then $arrItemsExtra[$m] &= "|Extra $arrItems1|" & $i
				$m += 1
			EndIf
		Next
		For $j = $istartj To UBound($arrItems2) - 1
			If Not $objDictionary.Exists($arrItems2[$j]) Then
				$arrItemsExtra[$m + $n] = $arrItems2[$j]
				If $iDetails Then $arrItemsExtra[$m + $n] &= "|Extra $arrItems2|" & $j
				$n += 1
			EndIf
		Next
		ReDim $arrItemsExtra[$n + $m ]
		$arrItemsExtra[0 ] = UBound($arrItemsExtra) - 1
	EndIf
	$objDictionary = ""
	$objDictionary2 = ""
	If Not $iType Then Return $arrItemsExtra
	If $iType Then Return $arrItemsDupes
EndFunc   ;==>_ArrayCompare
;===============================================================================
;
; Description:  _ArrayDupes; deletes duplicates in an Array 1D
; Syntax:           _ArrayDupes(ByRef $ar_Array)
; Parameter(s):    	$ar_Array = 1d Array
; Requirement(s):   None
; Return Value(s):  On Success 	- Returns  array of  duplicates; options for unique or details of all
;								- byref Returns a sorted array with no duplicates
;                        On Failure -
;						@Error=1 P
;						@Error=2
;
; Author(s):        randallc
;===============================================================================
Func _ArrayDupes(ByRef $arrItemsF, $iDelete = 0, $iDetails = 0)
	local $timerstamp1 = TimerInit()
	If @OSTYPE = "WIN32_WINDOWS"  Then Return 0
	If not IsArray($arrItemsF) then   Return 0
	Local $arrItems = $arrItemsF
	Local $i = 0, $k = 0, $objDictionary = ObjCreate("Scripting.Dictionary"), $arrItemsDupes[UBound($arrItems)]
	Local $objDictDupes = ObjCreate("Scripting.Dictionary")
	For $strItem In $arrItems
		If Not $objDictionary.Exists($strItem) Then
			$objDictionary.Add($strItem, $strItem)
		Else
			If $iDetails Then
				$arrItemsDupes[$k] = $strItem & "|Dupes|" & $i
			Else
				If Not $objDictDupes.Exists($strItem) Then $objDictDupes.Add($strItem, $strItem)
			EndIf
			$k += 1
		EndIf
		$i += 1
	Next
	ReDim $arrItems[$objDictionary.Count ]
	If $k = 0 Then ReDim $arrItemsDupes[1 ]
	If $k > 0 Then ReDim $arrItemsDupes[$k ]
	$i = 0
	If $iDelete Then
		For $strKey In $objDictionary.Keys
			$arrItems[$i] = $strKey
			$i += 1
		Next
		$arrItemsF = $arrItems ;array deleted dupes
	EndIf
	$i = 0
	If Not $iDetails Then
		For $strKey In $objDictDupes.Keys
			$arrItemsDupes[$i] = $strKey
			$i += 1
		Next
		If $i > 0 Then ReDim $arrItemsDupes[$i]
	EndIf
	ConsoleWrite("Total Time= " & Round(TimerDiff($timerstamp1)) & "" & @TAB & " msec" & @LF)
	Return $arrItemsDupes
EndFunc   ;==>_ArrayDupes
