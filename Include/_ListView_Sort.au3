#include-once

#include <GuiConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GuiListView.au3>
#include <GuiImageList.au3>
#include <Array.au3>

;================================================================================
;Function Name:     _ListView_Sort()
; Description:      Sorting ListView items when column click
; Parameter(s):     cIndex - Column index
; Return Value(s):  None
; Requirement(s):   AutoIt 3.2.12.0 and above
; Author(s):        R.Gilman (a.k.a rasim) ;
;================================================================================
Func _ListView_Sort($cIndex = 0)
	Local $iColumnsCount, $iDimension, $iItemsCount, $aItemsTemp, $aItemsText, $iCurPos, $iImgSummand, $i, $j
	$iColumnsCount = _GUICtrlListView_GetColumnCount($hListView)
	$iDimension = $iColumnsCount * 2
	$iItemsCount = _GUICtrlListView_GetItemCount($hListView)
	Local $aItemsTemp[1][$iDimension]

	For $i = 0 To $iItemsCount - 1
		$aItemsTemp[0][0] += 1
		ReDim $aItemsTemp[$aItemsTemp[0][0] + 1][$iDimension]
		$aItemsText = _GUICtrlListView_GetItemTextArray($hListView, $i)
		$iImgSummand = $aItemsText[0] - 1
		For $j = 1 To $aItemsText[0]
			$aItemsTemp[$aItemsTemp[0][0]][$j - 1] = $aItemsText[$j]
			$aItemsTemp[$aItemsTemp[0][0]][$j + $iImgSummand] = _GUICtrlListView_GetItemImage($hListView, $i, $j - 1)
		Next
	Next
	$iCurPos = $aItemsTemp[1][$cIndex]
	_ArraySort($aItemsTemp, 0, 1, 0, $cIndex)
	If StringInStr($iCurPos, $aItemsTemp[1][$cIndex]) Then _ArraySort($aItemsTemp, 1, 1, 0, $cIndex)
	For $i = 1 To $aItemsTemp[0][0]
		For $j = 1 To $iColumnsCount
			_GUICtrlListView_SetItemText($hListView, $i - 1, $aItemsTemp[$i][$j - 1], $j - 1)
			_GUICtrlListView_SetItemImage($hListView, $i - 1, $aItemsTemp[$i][$j + $iImgSummand], $j - 1)
		Next
	Next
EndFunc   ;==>_ListView_Sort
;================================================================================
Func WM_NOTIFY($hWnd, $iMsg, $iwParam, $ilParam)
	Local $hWndFrom, $iIDFrom, $iCode, $tNMHDR, $hWndListView
	$hWndListView = $hListView
	If Not IsHWnd($hListView) Then $hWndListView = GUICtrlGetHandle($hListView)
	$tNMHDR = DllStructCreate($tagNMHDR, $ilParam)
	$hWndFrom = HWnd(DllStructGetData($tNMHDR, "hWndFrom"))
	$iCode = DllStructGetData($tNMHDR, "Code")
	Switch $hWndFrom
		Case $hWndListView
			Switch $iCode
				Case $LVN_COLUMNCLICK
					Local $tInfo = DllStructCreate($tagNMLISTVIEW, $ilParam)
					Local $ColumnIndex = DllStructGetData($tInfo, "SubItem")
					_ListView_Sort($ColumnIndex)
			EndSwitch
	EndSwitch
	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_NOTIFY