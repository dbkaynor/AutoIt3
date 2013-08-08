#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <TreeViewConstants.au3>
#include <GuiTreeView.au3>
#include <_DougFunctions.au3>

Opt('MustDeclareVars', 1)

Global $cTreeView

_Main()

Func _Main()

	Local $myArray[10][2], $nMsg, $iIndex, $sText, $Char
	Local $iTVStyles = BitOR($TVS_HASBUTTONS, $TVS_HASLINES, $TVS_LINESATROOT, _
			$TVS_DISABLEDRAGDROP, $TVS_SHOWSELALWAYS, $TVS_Checkboxes)
	Local $mainform = GUICreate("TreeView", 400, 300)
	$cTreeView = GUICtrlCreateTreeView(10, 10, 380, 280, $iTVStyles, $WS_EX_CLIENTEDGE)

	Local $var = GUICtrlCreateContextMenu($cTreeView)
	Local $preview = GUICtrlCreateMenuItem("Preview", $var)

	For $i = 1 To UBound($myArray) - 1
		$Char = Chr(64 + $i)
		$myArray[$i][0] = "This element contains " & $Char
		$myArray[$i][1] = GUICtrlCreateTreeViewItem($myArray[$i][0], $cTreeView)
	Next

	GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")
	GUISetState(@SW_SHOW)

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				Exit
			Case $preview
				$iIndex = GUICtrlRead($cTreeView)
				$sText = GUICtrlRead($cTreeView, 1);_GUICtrlTreeView_GetText($hTreeView, $iIndex)
				MsgBox(0, "Preview", "Index: " & $iIndex & @CRLF & "ItemText: " & $sText)
		EndSwitch
	WEnd

EndFunc   ;==>_Main


Func WM_NOTIFY($hWnd, $iMsg, $wParam, $lParam)
	#forceref $hWnd, $iMsg, $wParam
	Local $tNMHDR, $hWndFrom, $iIDFrom, $iCode
	$tNMHDR = DllStructCreate($tagNMHDR, $lParam)
	$hWndFrom = DllStructGetData($tNMHDR, "hWndFrom")
	$iIDFrom = DllStructGetData($tNMHDR, "IDFrom")
	$iCode = DllStructGetData($tNMHDR, "Code")
	;ConsoleWrite(@ScriptLineNumber & " Got here"  & @CRLF)
	Switch $iIDFrom
		Case $cTreeView
			Switch $iCode
				Case $NM_CLICK
					Local $tPoint = _WinAPI_GetMousePos(True, $hWndFrom), $tHitTest
					$tHitTest = _GUICtrlTreeView_HitTestEx($hWndFrom, DllStructGetData($tPoint, 1), DllStructGetData($tPoint, 2))
					If BitAND(DllStructGetData($tHitTest, "Flags"), $TVHT_ONITEM) Then
						_GUICtrlTreeView_SelectItem($hWndFrom, DllStructGetData($tHitTest, 'Item'))
					EndIf
					MsgBox(4160, "Information", _GUICtrlTreeView_GetText($cTreeView, _GUICtrlTreeView_GetSelection($cTreeView)))
			EndSwitch
	EndSwitch
EndFunc   ;==>WM_NOTIFY
