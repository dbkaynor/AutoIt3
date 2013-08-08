; _ArrayDisplaySortV.au3 
#AutoIt3Wrapper_Au3Check_Parameters= -d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#include <GuiListView.au3> 
#include <GuiEdit.au3>
#include <_ArrayDisplaySort.au3> 
#include <StringSplit.au3>
;~ #include <_ArraySortClibsetV.au3>
;~ #include <_ArraySortClib.au3>
Global $tTextBufferV, $sSeparatorV, $iCaseV, $avArrayV, $s_LISTVIEWSortOrderV, $iUBoundV, $iColV;, $iTimerProgressV
Global $vDescendingV[1], $arArAscV[1], $i_LISTVIEWPrevcolumnV = -1, $ar_LISTVIEWArrayV[1], $hGUI;, $arArDescV2V[1]
;~ Global $tagSourceV, $pSourceV, $procStrCmpV, $tSourceV, $hMsvcrtV, $sStrCmpV; for ClibSet
Func _ArrayDisplaySortV(Const ByRef $avArrayVF, $sTitle = "Array: ListView Display", $iItemLimit = 0, $iOptions = 0, $sSeparatorVF = "", $sReplace = "")
	#Region const
	Local Const $_ARRAYUUCONSTANT_GUI_EVENT_CLOSE = -3
	Local Const $_ARRAYUUCONSTANT_LVM_GETITEMCOUNT = (0x1000 + 4)
	Local Const $_ARRAYUUCONSTANT_LVM_GETITEMSTATE = (0x1000 + 44)
	Local Const $LVM_SETCOLUMNWIDTHU = (0x1000 + 30)
	Local Const $LVSCWU_AUTOSIZE = -1
	Local Const $LVSCWU_AUTOSIZE_USEHEADER = -2
	;===========================================added rjc
	;;virtual lv style
	Local Const $LVS_OWNERDATA = 0x1000


	Local Const $_ARRAYCONSTANT_GUI_DOCKBORDERS = 0x66
	Local Const $_ARRAYCONSTANT_GUI_DOCKBOTTOM = 0x40
	Local Const $_ARRAYCONSTANT_GUI_DOCKHEIGHT = 0x0200
	Local Const $_ARRAYCONSTANT_GUI_DOCKLEFT = 0x2
	Local Const $_ARRAYCONSTANT_GUI_DOCKRIGHT = 0x4
	Local Const $_ARRAYCONSTANT_GUI_EVENT_CLOSE = -3
	Local Const $_ARRAYCONSTANT_LVIF_PARAM = 0x4
	Local Const $_ARRAYCONSTANT_LVIF_TEXT = 0x1
	Local Const $_ARRAYCONSTANT_LVM_GETITEMCOUNT = (0x1000 + 4)
	Local Const $_ARRAYCONSTANT_LVM_GETITEMSTATE = (0x1000 + 44)
	Local Const $_ARRAYCONSTANT_LVM_INSERTITEMA = (0x1000 + 7)
	Local Const $_ARRAYCONSTANT_LVM_SETEXTENDEDLISTVIEWSTYLE = (0x1000 + 54)
	Local Const $_ARRAYCONSTANT_LVM_SETITEMA = (0x1000 + 6)
	Local Const $_ARRAYCONSTANT_LVS_EX_FULLROWSELECT = 0x20
	Local Const $_ARRAYCONSTANT_LVS_EX_GRIDLINES = 0x1
	Local Const $_ARRAYCONSTANT_LVS_SHOWSELALWAYS = 0x8
	Local Const $_ARRAYCONSTANT_WS_EX_CLIENTEDGE = 0x0200
	Local Const $_ARRAYCONSTANT_WS_MAXIMIZEBOX = 0x00010000
	Local Const $_ARRAYCONSTANT_WS_MINIMIZEBOX = 0x00020000
	Local Const $_ARRAYCONSTANT_WS_SIZEBOX = 0x00040000
	Local Const $_ARRAYCONSTANT_LVN_GETDISPINFO = -177
	Local Const $_ARRAYCONSTANT_tagLVITEM = "int Mask;int Item;int SubItem;int State;int StateMask;ptr Text;int TextMax;int Image;int Param;int Indent;int GroupID;int Columns;ptr pColumns"
	#EndRegion const

	Local $sTextBuff = "char[261]"
	If Execute("@Unicode") Then $sTextBuff = "w" & $sTextBuff
	$tTextBufferV = DllStructCreate($sTextBuff)
	Local $sBlank = ""
	Local $sCol0 = "", $vTmp, $iNumItems, $iBuffer = 2048000;,64;;$sSeparatorV3, $sSeparatorV2 = $sSeparatorV,
	Local $timerstamp1 = TimerInit(), $aItem, $i_ColWidth, $i_OrderColumn;, $timerstampprep = TimerInit()
	Local $iTranspose = BitAND($iOptions, 1), $iSort = BitAND($iOptions, 2)
;~ 	ReDim $ar_LISTVIEWArrayV[4000]
	$iCaseV = BitAND($iOptions, 4)
	$sSeparatorV = $sSeparatorVF
	Local $iDelimited = BitAND($iOptions, 8)
	Local $iFast = BitAND($iOptions, 16)
	If Not IsArray($avArrayVF) And IsString($avArrayVF) Then
		$avArrayV = StringSplit(StringStripCR($avArrayVF), @LF)
;~ 		$iDelimited = 1
	Else
		$avArrayV = $avArrayVF
	EndIf
	If Not IsArray($avArrayV) Then Return SetError(1, 0, 0)
	; Dimension checking
	Local $iDimension = UBound($avArrayV, 0), $iSubItems = UBound($avArrayV, 2), $arTitle[1], $sTempHeader = 'Row' ;,
	$iUBoundV = UBound($avArrayV, 1)
	If $iDimension > 2 Then Return SetError(2, 0, 0)
	If $iDelimited And $iDimension > 1 Then $iDelimited = 0
	
	;===========================================added rjc
	$arTitle = StringSplit($sTitle, "|")
	If (UBound($arTitle) > 2 Or $iDelimited) And $iDimension = 1 Then
		$iSubItems = UBound($arTitle)
		$iDelimited = 1
		$ar_LISTVIEWArrayV = $avArrayV
	EndIf

	; Separator handling
	If $iDelimited Then
		$iTranspose = 0
		Local $arSeps = StringSplit(",|;", "")
		$iSort = 1
		If $sSeparatorVF = "" Then
			For $i = 1 To UBound($arSeps) - 1
				If StringInStr($avArrayV[0] & $avArrayV[1] & $avArrayV[UBound($avArrayV) - 1], $arSeps[$i]) Then
					$sSeparatorV = $arSeps[$i]
					ExitLoop
				EndIf
			Next
		EndIf
		If $iDelimited And $sSeparatorV <> "" Then $sReplace = $sSeparatorV
	EndIf
	If $iSort Then
		If $sSeparatorV = "" Then $sSeparatorV = ","
;~ 		$sBlank = " "
;~ 		$sSeparatorV2 = $sBlank & $sSeparatorV
		$sReplace = $sSeparatorV
	EndIf
	If $sSeparatorV = "" Then
		$sSeparatorV = Chr(1)
		If $sReplace == "" Then $sReplace = Chr(1); Then $chr = 1
	EndIf
	If $sReplace == "" Then $sReplace = "&"

	;===========================================added rjc
	If $sTitle = -1 Or $sTitle = "" Then $sTitle = "ListView array 1D and 2D Display"
	If StringInStr($sTitle, "|") Then; then check if there are header cols too;
		$arTitle = StringSplit($sTitle, "|")
		$sTitle = StringReplace($arTitle[1], $sSeparatorV, $sReplace, 0, 1);$arTitle[1]
		If $arTitle[2] = "" Then $sCol0 = "col 0"
		$sTempHeader &= $sSeparatorV & StringReplace($arTitle[2], $sSeparatorV, $sReplace, 0, 1) & $sCol0;$arTitle[2]
	Else
		$sTempHeader &= $sSeparatorV & "Col 0"
	EndIf

	; Separator handling
	If $sSeparatorV = "" Then
		$sSeparatorV = Chr(1)
		If $sReplace == "" Then $sReplace = Chr(1); Then $chr = 1
	EndIf
	If $sReplace == "" Then $sReplace = "&"

	;===========================================added rjc
	If $sSeparatorV <> $sReplace Or $iFast Then
		ConsoleWrite("==================================$sReplace=" & $sReplace & @LF)
		If $iDimension = 2 Then
			ReDim $avArrayV[$iUBoundV][$iSubItems]
			For $i = 0 To $iUBoundV - 1
				For $j = 0 To $iSubItems - 1
					$avArrayV[$i][$j] = StringReplace($avArrayV[$i][$j], $sSeparatorV, $sReplace, 0, 1);replace seaparator
				Next
			Next
		ElseIf $iDimension = 1 Then
			ReDim $avArrayV[$iUBoundV]
			For $i = 0 To $iUBoundV - 1
				$avArrayV[$i] = StringReplace($avArrayV[$i], $sSeparatorV, $sReplace, 0, 1) ;replace seaparator
			Next
		EndIf
	Else
		If $iDelimited Then
			$iSubItems = UBound($arTitle) - 2
			$iDelimited = 1
		EndIf
	EndIf
	ReDim $vDescendingV[$iSubItems + 2]
	ConsoleWrite("$iSubItems=" & $iSubItems & @LF)
;~ 	;===========================================added rjc

	If $iDelimited And $iTranspose = 1 Then
		$iTranspose = 0
		$sTitle = "** WARNING ** no transpose for delimited string array: " & $sTitle
	EndIf

	; Hard limits
	Local $iLVIAddUDFThreshold = 4000, $iColVLimit, $iWidth = @DesktopWidth * 2 / 3, $iHeight = @DesktopHeight * 2 / 3;, $sSeparatorV = $sBlank & $sSeparatorV, $sSeparatorV3
	If $iItemLimit > 40 Then $iLVIAddUDFThreshold = $iItemLimit
	$iColVLimit = $iLVIAddUDFThreshold
	If $iItemLimit = 1 Or $iItemLimit < 0 Or $iItemLimit = Default Then $iItemLimit = $iLVIAddUDFThreshold

	; Declare/define variables
	Local $sHeader ;= "Row" ;, $avArrayV[1];, $ar_LISTVIEWArrayV$i, $j,  $arSelItems,
	
	; Swap dimensions if transposing
	If $iTranspose Then
		Local $iTmp = $iUBoundV
		$iUBoundV = $iSubItems
		$iSubItems = $iTmp
	EndIf
	; Set limits for dimensions
	If Not $iUBoundV Then $iUBoundV = 1
	If Not $iSubItems Then $iSubItems = 1
	;===========================================added rjc
	If $iSubItems > $iColVLimit Then
		$iSubItems = $iColVLimit
		$sTitle = "** WARNING ** $iColVLimit maximum 4000: " & $sTitle
	EndIf

	;===========================================added rjc
	If $iItemLimit < 1 Then $iItemLimit = $iUBoundV
	If $iUBoundV > $iItemLimit Then $iUBoundV = $iItemLimit
	;===========================================added ult
	If $iLVIAddUDFThreshold > $iUBoundV - 1 Then $iLVIAddUDFThreshold = $iUBoundV - 1
	;===========================================added ult

	; Convert array into text for listview
	ReDim $ar_LISTVIEWArrayV[$iUBoundV]
	Local $iArrayType = $iDimension + 2 * $iTranspose - 1, $sTxt;, $iTemp=1;, $sExec1 = '"', $sExec2 = '"',$avArrayVText[$iUBoundV + 1]

	; Set header up;; make LV header
	;===========================================added rjc
	
	StringReplace($ar_LISTVIEWArrayV[UBound($ar_LISTVIEWArrayV) - 1], $sSeparatorV, $sSeparatorV)
	Local $FirstPipes = @extended, $uboundary = $iSubItems - 1, $iMaxPipes = $FirstPipes, $sStrRep
	ConsoleWrite("$iMaxPipes=" & $iMaxPipes & @LF)
	If $iDelimited Then
		$uboundary = $iMaxPipes
		$iSubItems = $uboundary + 1
		$avArrayV = $ar_LISTVIEWArrayV
		$avArrayV[0] = $ar_LISTVIEWArrayV[0]
		ReDim $arArAscV[$iSubItems + 1]
	Else
		ReDim $arArAscV[2] ;for 1D
		If $iDimension = 2 Then ReDim $arArAscV[UBound($avArrayV, 2) + 1]
	EndIf
	;===========================================added rjc
	If Not $iDelimited And $iDimension = 1 And Not $iTranspose And Not $iSort Then
		$uboundary = 0
	EndIf
	If $uboundary > $iColVLimit - 1 Then $uboundary = $iColVLimit - 1
	For $i = 1 To $uboundary
		If $i < UBound($arTitle) - 2 Then
			$sTempHeader &= $sSeparatorV & StringReplace($arTitle[$i + 2], $sSeparatorV, $sBlank & $sReplace, 0, 1);arTitle[$i + 2]
		Else
			$sTempHeader &= $sSeparatorV & "Col " & $i
		EndIf
	Next
	; Set header up
	If $iSort Then $sHeader = "Row" & $sSeparatorV & $sHeader

	$sHeader = $sTempHeader & $sBlank
	Local $iOnEventMode = Opt("GUIOnEventMode", 0), $sDataSeparatorChar = Opt("GUIDataSeparatorChar", $sSeparatorV)
	ConsoleWrite("StringLeft($sHeader,30) =" & StringLeft($sHeader, 30) & @LF)
	ConsoleWrite("StringRight($sHeader,30) =" & StringRight($sHeader, 30) & @LF)
	If $iSubItems > 5 Then ;10 Then
		_ArrayDisplaySort($avArrayVF, $sTitle & $sHeader, $iItemLimit, $iOptions, $sSeparatorV, $sReplace)
		Return
	EndIf
	;===========================================

	Local $iAddMask = BitOR($_ARRAYCONSTANT_LVIF_TEXT, $_ARRAYCONSTANT_LVIF_PARAM)
	Local $tBuffer = DllStructCreate("char Text[" & $iBuffer & "]"), $pBuffer = DllStructGetPtr($tBuffer)
	Local $tItem = DllStructCreate($_ARRAYCONSTANT_tagLVITEM), $pItem = DllStructGetPtr($tItem)
	DllStructSetData($tItem, "Param", 0)
	DllStructSetData($tItem, "Text", $pBuffer)
	DllStructSetData($tItem, "TextMax", $iBuffer)

	; Set interface up
	$hGUI = GUICreate($sTitle, $iWidth, $iHeight, Default, Default, BitOR($_ARRAYCONSTANT_WS_SIZEBOX, $_ARRAYCONSTANT_WS_MINIMIZEBOX, $_ARRAYCONSTANT_WS_MAXIMIZEBOX))
	Local $aiGUISize = WinGetClientSize($hGUI)
	Local $hListView = GUICtrlCreateListView($sHeader, 0, 0, $aiGUISize[0], $aiGUISize[1] - 26, $_ARRAYCONSTANT_LVS_SHOWSELALWAYS + $LVS_OWNERDATA)
	Local $hCopy = GUICtrlCreateButton("Copy Selected", 3, $aiGUISize[1] - 23, $aiGUISize[0] - 6, 20)

	GUICtrlSetResizing($hListView, $_ARRAYCONSTANT_GUI_DOCKBORDERS)
	GUICtrlSetResizing($hCopy, $_ARRAYCONSTANT_GUI_DOCKLEFT + $_ARRAYCONSTANT_GUI_DOCKRIGHT + $_ARRAYCONSTANT_GUI_DOCKBOTTOM + $_ARRAYCONSTANT_GUI_DOCKHEIGHT)
	GUICtrlSendMsg($hListView, $_ARRAYCONSTANT_LVM_SETEXTENDEDLISTVIEWSTYLE, $_ARRAYCONSTANT_LVS_EX_GRIDLINES, $_ARRAYCONSTANT_LVS_EX_GRIDLINES)
	GUICtrlSendMsg($hListView, $_ARRAYCONSTANT_LVM_SETEXTENDEDLISTVIEWSTYLE, $_ARRAYCONSTANT_LVS_EX_FULLROWSELECT, $_ARRAYCONSTANT_LVS_EX_FULLROWSELECT)
	GUICtrlSendMsg($hListView, $_ARRAYCONSTANT_LVM_SETEXTENDEDLISTVIEWSTYLE, $_ARRAYCONSTANT_WS_EX_CLIENTEDGE, $_ARRAYCONSTANT_WS_EX_CLIENTEDGE)
	GUICtrlSendMsg($hListView, $_ARRAYCONSTANT_LVM_SETEXTENDEDLISTVIEWSTYLE, $LVS_EX_DOUBLEBUFFER, $LVS_EX_DOUBLEBUFFER)
;~ 	_ArrayDisplay($avArrayV,"$avArrayV")
	; ================================================================================================================
	For $i = 0 To $iColVLimit ;;fix columns width
		GUICtrlSendMsg($hListView, 0x101E, $i, 80) ;;LVM_SETCOLUMNWIDTH = 0x101E
	Next
	If $iColVLimit > 5 Then $iColVLimit = 5
	;===========================================added rjc         ; ensure column width narrow enough to avoid crash with many columns
	GUICtrlSendMsg($hListView, $LVM_SETCOLUMNWIDTHU, $iSubItems, $LVSCWU_AUTOSIZE_USEHEADER)
	; ================================================================================================================
;~ 		ConsoleWrite("$iUBoundV=" &$iUBoundV& @LF)
	GUICtrlSendMsg($hListView, $LVM_SETITEMCOUNT, $iUBoundV, 0) ;;set total item count - necessary for virtual lv
	Local $wProcNew
	If $iDimension = 1 Then
		$wProcNew = DllCallbackRegister("_ArrayUDF_WM_NOTIFYSort2", "ptr", "hwnd;uint;long;ptr") ;;register new window proc
		ConsoleWrite("will sort with _ArrayUDF_WM_NOTIFYSort2 " & @LF)
	ElseIf $iDimension = 2 Then
		$wProcNew = DllCallbackRegister("_ArrayUDF_WM_NOTIFYVir2D", "ptr", "hwnd;uint;long;ptr") ;;register new window proc
		ConsoleWrite("will not sort with _ArrayUDF_WM_NOTIFYVir2D " & @LF)
	EndIf

	Local $wProcOld = _ArrayUDF_WinSubclassV($hGUI, DllCallbackGetPtr($wProcNew)) ;;sublass gui
	Assign('ArrayUDF_wProcOld', $wProcOld, 2) ;; $ArrayUDF_wProcOld = global keeper of old windowproc ptr
	
	;===========================================added rjc
	; Show dialog
	GUISetState(@SW_SHOW, $hGUI)
	ConsoleWrite($sTitle & " MineV Total Time=" & Round(TimerDiff($timerstamp1)) & "" & @TAB & " msec" & @LF)
	;===========================================avbs code for QuickSort
	If $iSort And $iDelimited Then
		Local $timerstamp11 = TimerInit()
		If $iDelimited Then
			ConsoleWrite("$iDelimited=" & $iDelimited & @LF)
			Local $timerstamp12 = TimerInit()
			For $i = 0 To UBound($ar_LISTVIEWArrayV) - 1
				StringReplace($ar_LISTVIEWArrayV[$i], $sSeparatorV, $sSeparatorV)
				If @extended > $iMaxPipes Then $iMaxPipes = @extended
			Next
			$uboundary = $iMaxPipes
			$iSubItems = $uboundary + 1
			If $iMaxPipes > $FirstPipes Then
				For $i = 1 To $iMaxPipes - $FirstPipes
					$sStrRep &= " " & $sSeparatorV & " "
				Next
			EndIf
			ConsoleWrite($sTitle & " $iMaxPipes Time=" & Round(TimerDiff($timerstamp12)) & "" & @TAB & " msec" & @LF)
		EndIf
	EndIf
	If $iSort Or $iDelimited Then
		Local $timerstamp14 = TimerInit(), $iGuessNumeric
;~ 		If $iDimension = 1 And Not $iDelimited Then $iGuessNumeric = IsNumber($avArrayV[0]) And IsNumber($avArrayV[1]) And IsNumber($avArrayV[UBound($avArrayV) - 1])
;~ 		If $iDimension = 2 Or $iDelimited Or $iGuessNumeric Then
		_SetupSortArrays($iDelimited, $sStrRep, $iSubItems, $iArrayType, $iSort)
		ConsoleWrite($sTitle & "_SetupSortArrays Time=" & Round(TimerDiff($timerstamp14)) & "" & @TAB & " msec" & @LF)
		_GUICtrlListView_BeginUpdate($hListView)
		$avArrayV = $ar_LISTVIEWArrayV ; we need this inside _GUICtrlListView_BeginUpdate
;~ 		Else
;~ 			Local $timerstamp14 = TimerInit(), $iGuessNumeric
;~ 			_ArraySortClibsetV($avArrayV)
;~ 			_GUICtrlListView_BeginUpdate($hListView)
;~ 		EndIf
;~ 		ConsoleWrite($sTitle & "_ArraySortClibsetV Time=" & Round(TimerDiff($timerstamp14)) & "" & @TAB & " msec" & @LF)
		DllCallbackFree($wProcNew)
		If $iDimension = 1 And Not $iDelimited Then
			$wProcNew = DllCallbackRegister("_ArrayUDF_WM_NOTIFYSort1D", "ptr", "hwnd;uint;long;ptr") ;;register new window proc
			ConsoleWrite("will NOW sort with _ArrayUDF_WM_NOTIFYSort1D " & @LF)
		Else
			$wProcNew = DllCallbackRegister("_ArrayUDF_WM_NOTIFYSort2", "ptr", "hwnd;uint;long;ptr") ;;register new window proc
			ConsoleWrite("will NOW sort with _ArrayUDF_WM_NOTIFYSort2 " & @LF)
		EndIf
		$s_LISTVIEWSortOrderV = "DESC"
		_GUICtrlListView_EndUpdate($hListView)
;~ 		Local $timerstamp14 = TimerInit(), $iGuessNumeric
;~ 		_ArraySortClibsetV($avArrayV)
;~ 		ConsoleWrite($sTitle & "_ArraySortClibsetV Time=" & Round(TimerDiff($timerstamp14)) & "" & @TAB & " msec" & @LF)
	EndIf

	While 1
		Switch GUIGetMsg()
			Case $_ARRAYUUCONSTANT_GUI_EVENT_CLOSE
				ExitLoop

			Case $hCopy
				Local $sClip = ""

				; Get selected indices [ _GUICtrlListView_GetSelectedIndices($hListView, True) ]
				Local $aiCurItems[1] = [0]
				For $i = 0 To GUICtrlSendMsg($hListView, $_ARRAYUUCONSTANT_LVM_GETITEMCOUNT, 0, 0)
					If GUICtrlSendMsg($hListView, $_ARRAYUUCONSTANT_LVM_GETITEMSTATE, $i, 0x2) Then
						$aiCurItems[0] += 1
						ReDim $aiCurItems[$aiCurItems[0] + 1]
						$aiCurItems[$aiCurItems[0]] = $i
					EndIf
				Next

				; Generate clipboard text
				If Not $aiCurItems[0] Then
					For $i = 0 To $iUBoundV - 1
						$sClip &= $ar_LISTVIEWArrayV[$i] & @CRLF
					Next
				Else
					For $i = 1 To UBound($aiCurItems) - 1
						$sClip &= $ar_LISTVIEWArrayV[$aiCurItems[$i]] & @CRLF
					Next
				EndIf
				ClipPut($sClip)
		EndSwitch
	WEnd
	GUIDelete($hGUI)
	ReDim $arArAscV[1]
	$s_LISTVIEWSortOrderV = ""
	$i_LISTVIEWPrevcolumnV = -1
;~ 	ReDim $arArDescV2V[1]
	If IsString($ar_LISTVIEWArrayV) Then $ar_LISTVIEWArrayV = StringSplit("", "|")
;~ 	_ArrayDisplaySortV($ar_LISTVIEWArrayV, "$ar_LISTVIEWArrayV at exit")
	ReDim $ar_LISTVIEWArrayV[1]
	ReDim $avArrayV[1]
;~ 	DllCall('kernel32.dll', 'hwnd', 'FreeLibrary', 'hwnd', $hMsvcrtV)
	Opt("GUIOnEventMode", $iOnEventMode)
	Opt("GUIDataSeparatorChar", $sDataSeparatorChar)

	Return 1
EndFunc   ;==>_ArrayDisplaySortV
Func _ArrayUDF_WM_NOTIFYSort2($hWnd, $Msg, $wParam, $lParam)

	Local $wProcOld = Eval('ArrayUDF_wProcOld'), $i, $j, $tNMLVDISPINFO, $text, $textlen, $maxlen, $a_Rowa2
	;;$WM_NOTIFY = 0x004E
	If $Msg <> 0x004E Then Return _ArrayUDF_CallWndProcV($wProcOld, $hWnd, $Msg, $wParam, $lParam)

	Local $tNMHDR = DllStructCreate("hwnd hwndfrom;int idfrom;int code", $lParam)
	Local $hWndFrom = DllStructGetData($tNMHDR, "hwndfrom"), $iCode = DllStructGetData($tNMHDR, "code")

	;; the following is not very good, better would be have listview handle available in global scope.
	If ControlGetHandle($hWnd, "", "[Class:SysListView32;Instance:1]") = $hWndFrom Then

		Switch $iCode
			Case $LVN_COLUMNCLICK
				Local $iNotSorted = ($s_LISTVIEWSortOrderV == ""); Or ($s_LISTVIEWSortOrderV == "Ready")
				If $iNotSorted Then
;~ 					$iTimerProgress=_Timer_SetTimer($hGUI, 250, "_ArrayUDF_WM_NOTIFYSort2")
					MsgBox(64, "Warning", "Not Ready to sort, please wait", 1)
					ContinueCase
;~ 				Else
;~ 					_Timer_KillTimer($hGUI, $iTimerProgress)
				EndIf
				$tNMLISTVIEW = DllStructCreate("hwnd hwndfrom;int idfrom;int code;" & _
						"int item;int subitem", $lParam)
				$iColV = DllStructGetData($tNMLISTVIEW, "subitem")
				
				;===========================================
;~ 				_GUICtrlListView_BeginUpdate($hWndFrom)
				$timer = TimerInit()
				_Sort($iColV)
				ConsoleWrite("_Sort " & Round(TimerDiff($timer), 2) & " msecs" & @LF)
				_GUICtrlListView_BeginUpdate($hWndFrom)
				$avArrayV = $ar_LISTVIEWArrayV
				_GUICtrlListView_EndUpdate($hWndFrom)
			Case - 150, -177 ;;$LVN_GETDISPINFOA = -150, $LVN_GETDISPINFOW = -177
				;;give requested items one by one (kinda slow if a lot of items are visible at the same time)
				$tNMLVDISPINFO = DllStructCreate("hwnd hwndfrom;int idfrom;int code;" & _
						"uint mask;int item;int subitem;uint state;uint statemask;ptr text;int textmax;int;dword lparam;int;uint[5]", $lParam)

				If BitAND(DllStructGetData($tNMLVDISPINFO, "mask"), $LVIF_TEXT) Then
					$i = Dec(Hex(DllStructGetData($tNMLVDISPINFO, "item")))
					$j = DllStructGetData($tNMLVDISPINFO, "subitem")
					;================================================================
					Local $stmp, $iNotSorted = ($s_LISTVIEWSortOrderV == ""); Or ($s_LISTVIEWSortOrderV == "Ready")
					Local $aTmp = StringSplit($avArrayV[UBound($avArrayV) - 1], $sSeparatorV, 1)

					If $j = 0 And $iNotSorted Then
						$text = "[" & $i & "]"
					Else
						If $i_LISTVIEWPrevcolumnV = $iColV And _
								$s_LISTVIEWSortOrderV == "DESC" Then
							$aTmp = StringSplit($avArrayV[$iUBoundV - $i - 1], $sSeparatorV, 1)
						Else
							$aTmp = StringSplit($avArrayV[$i], $sSeparatorV, 1)
						EndIf
						$text = ""
						$stmp = $j + 1 - $iNotSorted
						If IsArray($aTmp) And $stmp <= UBound($aTmp) - 1 Then _
								$text = $aTmp[$stmp]
					EndIf
					$textlen = StringLen($text)
					$maxlen = DllStructGetSize($tTextBufferV)
					If Execute("@Unicode") Then $maxlen = $maxlen / 2
					If $textlen > $maxlen - 1 Then $text = StringLeft($text, $maxlen - 1)

					DllStructSetData($tTextBufferV, 1, $text)
					DllStructSetData($tNMLVDISPINFO, "textmax", $textlen)
					DllStructSetData($tNMLVDISPINFO, "text", DllStructGetPtr($tTextBufferV))
;~ 					If $i > $iUBoundV - 3 then _
;~ 					ConsoleWrite("$i"&$i & @LF&"$j"&$j & @LF&"UBound($aTmp) - 1="&UBound($aTmp) - 1 & @LF&"Last="&(_GUICtrlListView_GetTopIndex ($hListView) + _GUICtrlListView_GetCounterPage ($hListView) - 1) & @LF)
;~ 					If $i >= ($iUBoundV - 1 Or $i >= (_GUICtrlListView_GetTopIndex ($hListView) + _GUICtrlListView_GetCounterPage ($hListView) - 1)) And _
;~ 							$j = UBound($aTmp) - 1 Then $s_LISTVIEWSortOrderV = "Ready"
				ElseIf BitAND(DllStructGetData($tNMLVDISPINFO, "mask"), $LVIF_STATE) Then
					;;meh. although the selected items array filling could be implemented here, so there wouldn't be no need
					;;to loop through entire thing again when "Copy Selected" button is clicked

				EndIf
			Case - 152, -179 ;;$LVN_ODFINDITEM = -152, $LVN_ODFINDITEMW = -179
				;;lv keyboard find functionality.
				If UBound($avArrayV, 0) <> 1 Then Return -1
				$tNMLVFINDITEM = DllStructCreate("hwnd hwndfrom;int idfrom;int code;" & _
						"int start;uint flags;ptr string;dword lparam;int pointx;int pointy;uint dir", $lParam)
				If BitAND(DllStructGetData($tNMLVFINDITEM, "flags"), $LVFI_STRING) = 0 Then ContinueCase
				Local $sStructSearch = "char[260]"
				If Execute("@Unicode") Then $sStructSearch = "w" & $sStructSearch
				Local $tSearchStr = DllStructCreate($sStructSearch, DllStructGetData($tNMLVFINDITEM, "string"))
				Local $sSearchString = DllStructGetData($tSearchStr, 1)
;~ 				ConsoleWrite($sSearchString & @CRLF)
				$i = _ArraySearch($avArrayV, $sSearchString)
				If Not @error Then
					Return $i
				EndIf
				Return -1
		EndSwitch
	EndIf

	;;pass the unhandled messages to default WindowProc
	Return _ArrayUDF_CallWndProcV($wProcOld, $hWnd, $Msg, $wParam, $lParam)
EndFunc   ;==>_ArrayUDF_WM_NOTIFYSort2
Func _ArrayUDF_WM_NOTIFYVir2D($hWnd, $Msg, $wParam, $lParam)

	Local $wProcOld = Eval('ArrayUDF_wProcOld'), $i, $j, $tNMLVDISPINFO, $text, $textlen, $maxlen, $a_Rowa2
	;;$WM_NOTIFY = 0x004E
	If $Msg <> 0x004E Then Return _ArrayUDF_CallWndProcV($wProcOld, $hWnd, $Msg, $wParam, $lParam)

	Local $tNMHDR = DllStructCreate("hwnd hwndfrom;int idfrom;int code", $lParam)
	Local $hWndFrom = DllStructGetData($tNMHDR, "hwndfrom"), $iCode = DllStructGetData($tNMHDR, "code")

	;; the following is not very good, better would be have listview handle available in global scope.
	If ControlGetHandle($hWnd, "", "[Class:SysListView32;Instance:1]") = $hWndFrom Then

		Switch $iCode
			Case $LVN_COLUMNCLICK
				Local $iNotSorted = ($s_LISTVIEWSortOrderV == ""); Or ($s_LISTVIEWSortOrderV == "Ready")
				If $iNotSorted Then
					MsgBox(64, "Warning", "Not Ready to sort, please wait", 1)
					ContinueCase
				EndIf
				$tNMLISTVIEW = DllStructCreate("hwnd hwndfrom;int idfrom;int code;" & _
						"int item;int subitem", $lParam)
				$iColV = DllStructGetData($tNMLISTVIEW, "subitem")
				
				;===========================================
				_GUICtrlListView_BeginUpdate($hWndFrom)
				$timer = TimerInit()
				_Sort($iColV)
				ConsoleWrite("_Sort " & Round(TimerDiff($timer), 2) & " msecs" & @LF)
				_GUICtrlListView_EndUpdate($hWndFrom)
			Case - 150, -177 ;;$LVN_GETDISPINFOA = -150, $LVN_GETDISPINFOW = -177
				;;give requested items one by one (kinda slow if a lot of items are visible at the same time)
				$tNMLVDISPINFO = DllStructCreate("hwnd hwndfrom;int idfrom;int code;" & _
						"uint mask;int item;int subitem;uint state;uint statemask;ptr text;int textmax;int;dword lparam;int;uint[5]", $lParam)

				If BitAND(DllStructGetData($tNMLVDISPINFO, "mask"), $LVIF_TEXT) Then
					$i = Dec(Hex(DllStructGetData($tNMLVDISPINFO, "item")))
					$j = DllStructGetData($tNMLVDISPINFO, "subitem")
					;================================================================
					Local $iNotSorted = ($s_LISTVIEWSortOrderV == ""); Or ($s_LISTVIEWSortOrderV == "Ready")
					If $j = 0 Then
						$text = "[" & $i & "]"
					Else
						$text = $avArrayV[$i][$j - 1]
					EndIf
					;================================================================
					$textlen = StringLen($text)
					$maxlen = DllStructGetSize($tTextBufferV)
					If Execute("@Unicode") Then $maxlen = $maxlen / 2
					If $textlen > $maxlen - 1 Then $text = StringLeft($text, $maxlen - 1)

					DllStructSetData($tTextBufferV, 1, $text)
					DllStructSetData($tNMLVDISPINFO, "textmax", $textlen)
					DllStructSetData($tNMLVDISPINFO, "text", DllStructGetPtr($tTextBufferV))
;~ 					If ($i = $iUBoundV - 1 Or $i = (_GUICtrlListView_GetTopIndex ($hListView) + _GUICtrlListView_GetCounterPage ($hListView) - 1)) And _
;~ 							$j = UBound($avArrayV, 2) - 1 Then $s_LISTVIEWSortOrderV = "Ready"
				ElseIf BitAND(DllStructGetData($tNMLVDISPINFO, "mask"), $LVIF_STATE) Then
					;;meh. although the selected items array filling could be implemented here, so there wouldn't be no need
					;;to loop through entire thing again when "Copy Selected" button is clicked

				EndIf
			Case - 152, -179 ;;$LVN_ODFINDITEM = -152, $LVN_ODFINDITEMW = -179
				;;lv keyboard find functionality.
				If UBound($avArrayV, 0) <> 1 Then Return -1
				$tNMLVFINDITEM = DllStructCreate("hwnd hwndfrom;int idfrom;int code;" & _
						"int start;uint flags;ptr string;dword lparam;int pointx;int pointy;uint dir", $lParam)
				If BitAND(DllStructGetData($tNMLVFINDITEM, "flags"), $LVFI_STRING) = 0 Then ContinueCase
				Local $sStructSearch = "char[260]"
				If Execute("@Unicode") Then $sStructSearch = "w" & $sStructSearch
				Local $tSearchStr = DllStructCreate($sStructSearch, DllStructGetData($tNMLVFINDITEM, "string"))
				Local $sSearchString = DllStructGetData($tSearchStr, 1)
				$i = _ArraySearch($avArrayV, $sSearchString)
				If Not @error Then
					Return $i
				EndIf
				Return -1
		EndSwitch
	EndIf

	;;pass the unhandled messages to default WindowProc
	Return _ArrayUDF_CallWndProcV($wProcOld, $hWnd, $Msg, $wParam, $lParam)
EndFunc   ;==>_ArrayUDF_WM_NOTIFYVir2D
Func _ArrayUDF_WM_NOTIFYSort1D($hWnd, $Msg, $wParam, $lParam)

	Local $wProcOld = Eval('ArrayUDF_wProcOld'), $i, $j, $tNMLVDISPINFO, $text, $textlen, $maxlen, $a_Rowa2
	;;$WM_NOTIFY = 0x004E
	If $Msg <> 0x004E Then Return _ArrayUDF_CallWndProcV($wProcOld, $hWnd, $Msg, $wParam, $lParam)

	Local $tNMHDR = DllStructCreate("hwnd hwndfrom;int idfrom;int code", $lParam)
	Local $hWndFrom = DllStructGetData($tNMHDR, "hwndfrom"), $iCode = DllStructGetData($tNMHDR, "code")

	;; the following is not very good, better would be have listview handle available in global scope.
	If ControlGetHandle($hWnd, "", "[Class:SysListView32;Instance:1]") = $hWndFrom Then

		Switch $iCode
			Case $LVN_COLUMNCLICK
				Local $iNotSorted = ($s_LISTVIEWSortOrderV == ""); Or ($s_LISTVIEWSortOrderV == "Ready")
				If $iNotSorted Then
					MsgBox(64, "Warning", "Not Ready to sort, please wait", 1)
					ContinueCase
				EndIf
				Local $tNMLISTVIEW = DllStructCreate("hwnd hwndfrom;int idfrom;int code;" & _
						"int item;int subitem", $lParam)
				$iColV = DllStructGetData($tNMLISTVIEW, "subitem")
				
				;===========================================
				_GUICtrlListView_BeginUpdate($hWndFrom)
				Local $timer = TimerInit()
;~ 				ConsoleWrite('$i_OrderColumn + 1= ' & $i_OrderColumn + 1 & @CRLF)
				If $i_LISTVIEWPrevcolumnV <> $iColV Then $s_LISTVIEWSortOrderV = "DESC"
				If $i_LISTVIEWPrevcolumnV = $iColV And $s_LISTVIEWSortOrderV == "DESC" Then ;
;~ 				_GUICtrlListView_BeginUpdate($hWndFrom)
					ConsoleWrite('_ArrayRevVerse $s_LISTVIEWSortOrderV=' & $s_LISTVIEWSortOrderV & @CRLF)
					$s_LISTVIEWSortOrderV = "ASC"
				Else
					If ($i_LISTVIEWPrevcolumnV <> $iColV Or Not $s_LISTVIEWSortOrderV) And _
							Not IsArray($arArAscV[$iColV]) And $iColV Then
						Local $iGuessNumeric = IsNumber($avArrayV[0]) And IsNumber($avArrayV[1]) And IsNumber($avArrayV[UBound($avArrayV) - 1])
						If $iGuessNumeric Then
							Local $timerstamp13 = TimerInit()
							_ArraySortDelim($ar_LISTVIEWArrayV, 1, $iCaseV, $sSeparatorV)
							ConsoleWrite("_ArraySortDelim Time=" & Round(TimerDiff($timerstamp13)) & "" & @TAB & " msec" & @LF)
;~ 							;============== Can't use 1D sort on numbers in col1 as 2 cols unless string sort
;~ 							Local $timerstamp13 = TimerInit()
;~ 							_Sort_vbs1DVStr($ar_LISTVIEWArrayV, 0, 0, 1)
;~ 							ConsoleWrite("_Sort_vbs1DVStr Time=" & Round(TimerDiff($timerstamp13)) & "" & @TAB & " msec" & @LF)
;~ 							Local $timerstamp13 = TimerInit()
;~ 							_Sort_vbs1DV($ar_LISTVIEWArrayV, 0, 0, 1)
;~ 							ConsoleWrite("_Sort_vbs1DVStr Time=" & Round(TimerDiff($timerstamp13)) & "" & @TAB & " msec" & @LF)
						Else
							Local $timerstamp13 = TimerInit()
							_JSort1DaV($ar_LISTVIEWArrayV, False, 0, 0, 0, Chr(2))
							ConsoleWrite("_JSort1DaV Time=" & Round(TimerDiff($timerstamp13)) & "" & @TAB & " msec" & @LF)
							Local $timerstamp13 = TimerInit()
							$ar_LISTVIEWArrayV = _StringSplit($ar_LISTVIEWArrayV, Chr(2))
							ConsoleWrite("_StringSplit Time=" & Round(TimerDiff($timerstamp13)) & "" & @TAB & " msec" & @LF)
							;						_ArrayDisplay($ar_LISTVIEWArrayV,"_JSort1DaV Time=")
;~ 							Local $timerstamp13 = TimerInit()
;~ 							;_ArraySortClib(ByRef $Array, $bCase=False, $bDescend=False, $iStart=0, $iEnd=0, $iColumn=0, $iWidth=256)
;~ 							_ArraySortClib($ar_LISTVIEWArrayV, False, False, 0, 0, 0);, $iWidth=256)
;~ 							ConsoleWrite("_ArraySortClib Time=" & Round(TimerDiff($timerstamp13)) & "" & @TAB & " msec" & @LF)
						EndIf
;~ 						_GUICtrlListView_BeginUpdate($hWndFrom)
						$arArAscV[$iColV] = $ar_LISTVIEWArrayV
					ElseIf ($i_LISTVIEWPrevcolumnV <> $iColV Or Not $s_LISTVIEWSortOrderV) Or $iColV = 0 Then
;~ 						_GUICtrlListView_BeginUpdate($hWndFrom)
						$ar_LISTVIEWArrayV = $arArAscV[$iColV]
					EndIf
;~ 					$avArrayV = $ar_LISTVIEWArrayV
					$i_LISTVIEWPrevcolumnV = $iColV
					$s_LISTVIEWSortOrderV = "DESC"
				EndIf
				ConsoleWrite("_Sort " & Round(TimerDiff($timer), 2) & " msecs" & @LF)
				_GUICtrlListView_BeginUpdate($hWndFrom)
				$avArrayV = $ar_LISTVIEWArrayV
				_GUICtrlListView_EndUpdate($hWndFrom)
			Case - 150, -177 ;;$LVN_GETDISPINFOA = -150, $LVN_GETDISPINFOW = -177
				;;give requested items one by one (kinda slow if a lot of items are visible at the same time)
				$tNMLVDISPINFO = DllStructCreate("hwnd hwndfrom;int idfrom;int code;" & _
						"uint mask;int item;int subitem;uint state;uint statemask;ptr text;int textmax;int;dword lparam;int;uint[5]", $lParam)

				If BitAND(DllStructGetData($tNMLVDISPINFO, "mask"), $LVIF_TEXT) Then
					$i = Dec(Hex(DllStructGetData($tNMLVDISPINFO, "item")))
					$j = DllStructGetData($tNMLVDISPINFO, "subitem")
					;================================================================
					Local $iNotSorted = ($s_LISTVIEWSortOrderV == ""); Or ($s_LISTVIEWSortOrderV == "Ready")
					Local $sTemp1, $aTmp; = StringSplit($ar_LISTVIEWArrayV[UBound($avArrayV) - 1], $sSeparatorV, 1)

					If $i_LISTVIEWPrevcolumnV = $iColV And _
							$s_LISTVIEWSortOrderV == "DESC" Then
						$sTemp1 = $ar_LISTVIEWArrayV[$iUBoundV - $i - 1]
					Else
						$sTemp1 = $ar_LISTVIEWArrayV[$i]
					EndIf
					$aTmp = StringSplit($sTemp1, $sSeparatorV, 1)

;~ 					If $i_LISTVIEWPrevcolumnV = $iColV And _
;~ 							$s_LISTVIEWSortOrderV == "DESC" Then
;~ ;						$aTmp = StringSplit($avArrayV[$iUBoundV - $i - 1], $sSeparatorV, 1)
;~ 						$sTemp1 = DllStructGetData($tSourceV, $iUBoundV - $i) ;$Array[$iEnd+$iStart-$i] = DllStructGetData($tSource, $i+1)
;~ 						If Not $sTemp1 Then $sTemp1 = DllStructGetData($tSourceV, $iUBoundV - $i - 1) ;$Array[$iEnd+$iStart-$i] = DllStructGetData($tSource, $i+1)
;~ 						If Not $sTemp1 Then $sTemp1 = DllStructGetData($tSourceV, $iUBoundV - $i - 2) ;$Array[$iEnd+$iStart-$i] = DllStructGetData($tSource, $i+1)
;~ 						$aTmp = StringSplit($sTemp1, $sSeparatorV, 1)
;~ 						If IsString($sTemp1) Then
;~ 							ConsoleWrite("$aTmp=" & $aTmp & @LF)
;~ 							Exit
;~ 						ElseIf IsArray($aTmp) Then
;~ 							_ArrayDisplay($aTmp,"if")
;~ 							Exit
;~ 						Else
;~ 							If $sTemp1 = "" Then
;~ 								ConsoleWrite("1$sTemp1=NIL" & @LF)
;~ 								Exit
;~ 							EndIf
;~ 						EndIf
;~ 					Else
;~ 						$aTmp = StringSplit($avArrayV[$i], $sSeparatorV, 1)
;~ 						$sTemp1 = DllStructGetData($tSourceV, $i + 1)
;~ 						if $sTemp1="" then ConsoleWrite("else $sTemp1=NIL"  & @LF)
;~ 						If Not $sTemp1 Then $sTemp1 = DllStructGetData($tSourceV, $i + 2) ;$Array[$iEnd+$iStart-$i] = DllStructGetData($tSource, $i+1)
;~ 						If Not $sTemp1 Then $sTemp1 = DllStructGetData($tSourceV, $i + 3) ;$Array[$iEnd+$iStart-$i] = DllStructGetData($tSource, $i+1)
;~ 						$aTmp = StringSplit($sTemp1, $sSeparatorV, 1)
;~ 						If IsString($sTemp1) Then
;~ 							ConsoleWrite("else $sTemp1=" & $sTemp1 & @LF)
;~ 							Exit
;~ 						ElseIf IsArray($aTmp) Then
;~ 							_ArrayDisplay($aTmp,"else ")
;~ 							Exit
;~ 						Else
;~ 							If $sTemp1 = "" Then
;~ 								ConsoleWrite("else  $sTemp1=NIL" & @LF)
;~ 								Exit
;~ 							EndIf
;~ 						EndIf
;~ 					EndIf
;~ 					$aTmp = StringSplit($sTemp1, $sSeparatorV, 1)
;~ 					If IsString($sTemp1) Then
;~ 						ConsoleWrite("$aTmp=" & $aTmp & @LF)
;~ 						Exit
;~ 					ElseIf IsArray($aTmp) Then
;~ 						_ArrayDisplay($aTmp,"2")
;~ 						Exit
;~ 					Else
;~ 						If $sTemp1 = "" Then
;~ 							ConsoleWrite("2$sTemp1=NIL" & @LF)
;~ 							Exit
;~ 						EndIf
;~ 					EndIf
;~ 					Exit
					$text = $aTmp[2 - $j]
					$textlen = StringLen($text)
					$maxlen = DllStructGetSize($tTextBufferV)
					If Execute("@Unicode") Then $maxlen = $maxlen / 2
					If $textlen > $maxlen - 1 Then $text = StringLeft($text, $maxlen - 1)

					DllStructSetData($tTextBufferV, 1, $text)
					DllStructSetData($tNMLVDISPINFO, "textmax", $textlen)
					DllStructSetData($tNMLVDISPINFO, "text", DllStructGetPtr($tTextBufferV))
;~ 					If ($i = $iUBoundV - 1 Or $i = (_GUICtrlListView_GetTopIndex ($hListView) + _GUICtrlListView_GetCounterPage ($hListView) - 1) ) And _
;~ 							$j = UBound($aTmp) - 1 Then $s_LISTVIEWSortOrderV = "Ready"
				ElseIf BitAND(DllStructGetData($tNMLVDISPINFO, "mask"), $LVIF_STATE) Then
					;;meh. although the selected items array filling could be implemented here, so there wouldn't be no need
					;;to loop through entire thing again when "Copy Selected" button is clicked

				EndIf
			Case - 152, -179 ;;$LVN_ODFINDITEM = -152, $LVN_ODFINDITEMW = -179
				;;lv keyboard find functionality.
				If UBound($avArrayV, 0) <> 1 Then Return -1
				$tNMLVFINDITEM = DllStructCreate("hwnd hwndfrom;int idfrom;int code;" & _
						"int start;uint flags;ptr string;dword lparam;int pointx;int pointy;uint dir", $lParam)
				If BitAND(DllStructGetData($tNMLVFINDITEM, "flags"), $LVFI_STRING) = 0 Then ContinueCase
				Local $sStructSearch = "char[260]"
				If Execute("@Unicode") Then $sStructSearch = "w" & $sStructSearch
				Local $tSearchStr = DllStructCreate($sStructSearch, DllStructGetData($tNMLVFINDITEM, "string"))
				Local $sSearchString = DllStructGetData($tSearchStr, 1)
;~ 				ConsoleWrite($sSearchString & @CRLF)
				$i = _ArraySearch($avArrayV, $sSearchString)
				If Not @error Then
					Return $i
				EndIf
				Return -1
		EndSwitch
	EndIf

	;;pass the unhandled messages to default WindowProc
	Return _ArrayUDF_CallWndProcV($wProcOld, $hWnd, $Msg, $wParam, $lParam)
EndFunc   ;==>_ArrayUDF_WM_NOTIFYSort1D
Func _JSort1DaV(ByRef $arEither2D1D, Const $bSkipFirst = False, $iDesc = 0, $iNumeric = 1, $icase = 1, $sSeparator = "|"); mod for numeric
	Local $code = '   function SortArray(arrArray,Case,Sep) {'
	$code &= @LF & '	    arrArray=arrArray.toArray().sort();' ;numerical asc
	$code &= @LF & '	    if (Case=0){'
	$code &= @LF & '	    	arrArray=bubbleSort(arrArray);'
	$code &= @LF & '	    }'
	$code &= @LF & '	    return arrArray.join(Sep);' ;
;~ 	$code &= @LF & '	    return arrArray.join("|");' ;
	$code &= @LF & '	  }'
	;===============================================================
	$code &= @LF & 'function bubbleSort(inputArray) {'
	$code &= @LF & '	for (var i = inputArray.length - 1; i >= 0;  i--) {'
	$code &= @LF & '		for (var j = 0; j <= i; j++) {'
	$code &= @LF & '			if (inputArray[j+1].toUpperCase() < inputArray[j].toUpperCase()) {'
	$code &= @LF & '				var tempValue = inputArray[j];'
	$code &= @LF & '				inputArray[j] = inputArray[j+1];'
	$code &= @LF & '				inputArray[j+1] = tempValue;'
	$code &= @LF & '      }'
	$code &= @LF & '   }'
	$code &= @LF & '}'
	$code &= @LF & 'return inputArray;'
	$code &= @LF & '}'
	;===============================================================
	Local $codeNumeric = '   function SortArray(arrArray,Case,Sep) {'
	$codeNumeric &= @LF & '	        return arrArray.toArray().sort(sNumAsc).join(Sep);' ;numerical asc
	$codeNumeric &= @LF & '	    }'
	$codeNumeric &= @LF & '   function sNumAsc(a, b) {'
	$codeNumeric &= @LF & '	        return ((+a > +b) ? 1 : ((+a < +b) ? -1 : 0));'
	$codeNumeric &= @LF & '	    }'
	;===============================================================
;~ 	If Not $icase And Not $iNumeric Then $code = $codeNoCase
	If $iNumeric Then $code = $codeNumeric
	;===============================================================
	Local $jvs = ObjCreate("ScriptControl")
	$jvs.language = "jscript"
	$jvs.Timeout = -1
	$jvs.addcode($code)
;~ 	FileWrite(@ScriptDir & "\subsortJSau3a.jvs", $code)
	Local $arEither2D1DSt = $jvs.Run("SortArray", $arEither2D1D, ($icase Or $iNumeric), $sSeparator);, $iIndex, $icase)
	$arEither2D1D = $arEither2D1DSt
;~ 	$arEither2D1D = StringSplit($arEither2D1DSt, "|")
;~ 	_ArrayDelete($arEither2D1D,0)
;~ 	If Not $icase And Not $iNumeric Then
;~ 		Local $arEither2D1DSt = $jvs.Run("quick_sort", $arEither2D1D);, $iIndex, $icase)
;~ 		$arEither2D1D = StringSplit($arEither2D1DSt, "|")
;~ 	EndIf
	
	;===============================================================
;~ 	If Not $icase And Not $iNumeric Then _Array1DIntroSort( $arEither2D1D);_ArraySort($arEither2D1D, $iDesc, 1)
;~ 	If Not $icase And Not $iNumeric Then _Array2D1DFieldSortSt($arEither2D1D);_ArraySort($arEither2D1D, $iDesc, 1)
	;===============================================================
	If $iDesc Then _ArrayRevV($arEither2D1D)
	$jvs = ""
	Return True
EndFunc   ;==>_JSort1DaV
Func _Sort_vbs1DV(ByRef $arEither2D1D, Const $bSkipFirst = False, $iDesc = 0, $icase = 0)
	Local $code = 'function ArrQSort(ByRef SortArray, iAsc,icase)'
	$code &= @LF & '	dim First, Last'
	$code &= @LF & '	First=0: Last=ubound(SortArray)'
;~ 	$code &= @LF & '	if not icase then  QuickSortColumn  SortArray,First, Last'
	$code &= @LF & '	QuickSortColumnCase  SortArray,First, Last'
;~ 	$code &= @LF & '	if not icase then  QuickSortColumn  SortArray,First, Last'
;~ 	$code &= @LF & '	if icase then QuickSortColumnCase  SortArray,First, Last'
	$code &= @LF & '	if iasc=1 then ReverseElements SortArray, 0, ubound(SortArray)'
	$code &= @LF & '	ArrQSort= SortArray'
	$code &= @LF & 'End function   '
;~ 	$code &= @LF & 'function QuickSortColumn(ByRef SortArray,  First, Last)'
;~ 	$code &= @LF & '	dim Low,High,collitem,arCol'
;~ 	$code &= @LF & '	dim Temp,ListSeparator'
;~ 	$code &= @LF & '	Low = First'
;~ 	$code &= @LF & '	High = Last'
;~ 	$code &= @LF & '	ListSeparator=SortArray((First + Last) / 2)'
;~ 	$code &= @LF & '	Do'
;~ 	$code &= @LF & '		While lcase(SortArray(Low)) < lcase(ListSeparator)'
;~ 	$code &= @LF & '			Low = Low + 1'
;~ 	$code &= @LF & '		WEnd'
;~ 	$code &= @LF & '		While lcase(SortArray(High)) > lcase(ListSeparator)'
;~ 	$code &= @LF & '			High = High - 1'
;~ 	$code &= @LF & '		WEnd'
;~ 	$code &= @LF & '		If (Low <= High) Then'
;~ 	$code &= @LF & '			Temp = SortArray(Low)'
;~ 	$code &= @LF & '			SortArray(Low) = SortArray(High)'
;~ 	$code &= @LF & '			SortArray(High) = Temp'
;~ 	$code &= @LF & '			Low = Low + 1'
;~ 	$code &= @LF & '			High = High - 1'
;~ 	$code &= @LF & '		End If'
;~ 	$code &= @LF & '	Loop While (Low <= High)'
;~ 	$code &= @LF & '	If (First < High) Then QuickSortColumn  SortArray,First, High '
;~ 	$code &= @LF & '	If (Low < Last) Then QuickSortColumn  SortArray,Low, Last '
;~ 	$code &= @LF & 'End function   '
	$code &= @LF & 'function QuickSortColumnCase(ByRef SortArray, First, Last)'
	$code &= @LF & '	dim Low,High,collitem,arCol'
	$code &= @LF & '	dim Temp,ListSeparator'
	$code &= @LF & '	Low = First'
	$code &= @LF & '	High = Last'
	$code &= @LF & '	ListSeparator=SortArray((First + Last) / 2)'
	$code &= @LF & '	Do'
	$code &= @LF & '		While (SortArray(Low) < ListSeparator)'
	$code &= @LF & '			Low = Low + 1'
	$code &= @LF & '		WEnd'
	$code &= @LF & '		While (SortArray(High) > ListSeparator)'
	$code &= @LF & '			High = High - 1'
	$code &= @LF & '		WEnd'
	$code &= @LF & '		If (Low <= High) Then'
	$code &= @LF & '			Temp = SortArray(Low)'
	$code &= @LF & '			SortArray(Low) = SortArray(High)'
	$code &= @LF & '			SortArray(High) = Temp'
	$code &= @LF & '			Low = Low + 1'
	$code &= @LF & '			High = High - 1'
	$code &= @LF & '		End If'
	$code &= @LF & '	Loop While (Low <= High)'
	$code &= @LF & '	If (First < High) Then QuickSortColumnCase  SortArray,First, High '
	$code &= @LF & '	If (Low < Last) Then QuickSortColumnCase  SortArray,Low, Last '
	$code &= @LF & 'End function   '
	$code &= @LF & 'Sub ReverseElements( arrToReverse, intAlphaRow, intOmegaRow )'
	$code &= @LF & '    Dim intPointer, intUpper, intLower, varHolder'
	$code &= @LF & '    For intPointer = 0 To Int( (intOmegaRow - intAlphaRow) / 2 )'
	$code &= @LF & '        intUpper = intAlphaRow + intPointer'
	$code &= @LF & '        intLower = intOmegaRow - intPointer'
	$code &= @LF & '        varHolder = arrToReverse(intLower)'
	$code &= @LF & '        arrToReverse(intLower) = arrToReverse(intUpper)'
	$code &= @LF & '        arrToReverse(intUpper) = varHolder'
	$code &= @LF & '    Next'
	$code &= @LF & 'End Sub'
;~ 	$code &= @LF & 'Function IIf( expr, truepart, falsepart )'
;~ 	$code &= @LF & '   IIf = falsepart'
;~ 	$code &= @LF & '   If expr Then IIf = truepart'''
;~ 	$code &= @LF & 'End Function'
	Local $vbs = ObjCreate("ScriptControl")
	$vbs.language = "vbscript"
	$vbs.Timeout = -1
	$vbs.addcode($code)
;~ 	FileWrite(@ScriptDir&"\subsort.vbs",$code)
	$arEither2D1D = $vbs.Run("ArrQSort", $arEither2D1D, $iDesc, $icase)
	$vbs = ""
	Return True
EndFunc   ;==>_Sort_vbs1DV
Func _Sort_vbs1DVStr(ByRef $arArray1D, Const $bSkipFirst = False, $iDesc = 0, $icase = 0)
;~ 	Local $code = 'function ArrQSort(ByRef SortArray, iAsc,icase)'
;~ 	$code &= @LF & '	dim First, Last'
;~ 	$code &= @LF & '	First=0: Last=ubound(SortArray)'
;~ 	$code &= @LF & '	if not icase then  QuickSortColumn  SortArray,First, Last'
;~ 	$code &= @LF & '	if icase then QuickSortColumnCase  SortArray,First, Last'
;~ 	$code &= @LF & '	if iasc=1 then ReverseElements SortArray, 0, ubound(SortArray)'
;~ 	$code &= @LF & '	ArrQSort= SortArray'
;~ 	$code &= @LF & 'End function   '
	Local $code = 'function ArraySort1(ByRef SortArray, iAsc,First, Last,icase)'
	$code &= @LF & '	dim strWrite,arCol,arColumn()'
	$code &= @LF & '	Dim intPointer, booIsNumeric: booIsNumeric = True'
	$code &= @LF & '	For intPointer = First To Last'
	$code &= @LF & '		If Not IsNumeric( SortArray(intPointer) ) Then'
	$code &= @LF & '			booIsNumeric = False'
	$code &= @LF & '			Exit For'
	$code &= @LF & '		End If'
	$code &= @LF & '	Next'
	$code &= @LF & '	For intPointer = First To Last'
	$code &= @LF & '		If booIsNumeric Then'
	$code &= @LF & '			SortArray(intPointer)  = CSng( SortArray(intPointer) )'
	$code &= @LF & '		End If'
	$code &= @LF & '	Next'
;~ 	$code &= @LF & '	if icase or booIsNumeric then QuickSortColumnCase  SortArray,First, Last'
	$code &= @LF & '	if icase or booIsNumeric then'; QuickSortColumnCase  SortArray,First, Last'
	$code &= @LF & '	   QuickSortColumnCase  SortArray,First, Last'
	$code &= @LF & '	else'
	$code &= @LF & '		QuickSortColumn  SortArray, First, Last'
;~ 	$code &= @LF & '	if not icase then  QuickSortColumn  SortArray, First, Last'
	$code &= @LF & '	End If'
;~ 	$code &= @LF & '	if iasc=1 then ReverseElements SortArray, 0, ubound(SortArray)'
	$code &= @LF & '	ArraySort1= SortArray'
	$code &= @LF & 'End function   '
	$code &= @LF & 'function QuickSortColumn(ByRef SortArray,  First, Last)'
	$code &= @LF & '	dim Low,High,collitem,arCol'
	$code &= @LF & '	dim Temp,ListSeparator'
	$code &= @LF & '	Low = First'
	$code &= @LF & '	High = Last'
	$code &= @LF & '	ListSeparator=SortArray((First + Last) / 2)'
	$code &= @LF & '	Do'
	$code &= @LF & '		While lcase(SortArray(Low)) < lcase(ListSeparator)'
	$code &= @LF & '			Low = Low + 1'
	$code &= @LF & '		WEnd'
	$code &= @LF & '		While lcase(SortArray(High)) > lcase(ListSeparator)'
	$code &= @LF & '			High = High - 1'
	$code &= @LF & '		WEnd'
	$code &= @LF & '		If (Low <= High) Then'
	$code &= @LF & '			Temp = SortArray(Low)'
	$code &= @LF & '			SortArray(Low) = SortArray(High)'
	$code &= @LF & '			SortArray(High) = Temp'
	$code &= @LF & '			Low = Low + 1'
	$code &= @LF & '			High = High - 1'
	$code &= @LF & '		End If'
	$code &= @LF & '	Loop While (Low <= High)'
	$code &= @LF & '	If (First < High) Then QuickSortColumn  SortArray,First, High '
	$code &= @LF & '	If (Low < Last) Then QuickSortColumn  SortArray,Low, Last '
	$code &= @LF & 'End function   '
	$code &= @LF & 'function QuickSortColumnCase(ByRef SortArray, First, Last)'
	$code &= @LF & '	dim Low,High,collitem,arCol'
	$code &= @LF & '	dim Temp,ListSeparator'
	$code &= @LF & '	Low = First'
	$code &= @LF & '	High = Last'
	$code &= @LF & '	ListSeparator=SortArray((First + Last) / 2)'
	$code &= @LF & '	Do'
	$code &= @LF & '		While (SortArray(Low) < ListSeparator)'
	$code &= @LF & '			Low = Low + 1'
	$code &= @LF & '		WEnd'
	$code &= @LF & '		While (SortArray(High) > ListSeparator)'
	$code &= @LF & '			High = High - 1'
	$code &= @LF & '		WEnd'
	$code &= @LF & '		If (Low <= High) Then'
	$code &= @LF & '			Temp = SortArray(Low)'
	$code &= @LF & '			SortArray(Low) = SortArray(High)'
	$code &= @LF & '			SortArray(High) = Temp'
	$code &= @LF & '			Low = Low + 1'
	$code &= @LF & '			High = High - 1'
	$code &= @LF & '		End If'
	$code &= @LF & '	Loop While (Low <= High)'
	$code &= @LF & '	If (First < High) Then QuickSortColumnCase  SortArray,First, High '
	$code &= @LF & '	If (Low < Last) Then QuickSortColumnCase  SortArray,Low, Last '
	$code &= @LF & 'End function   '
	$code &= @LF & 'Sub ReverseElements( arrToReverse, intAlphaRow, intOmegaRow )'
	$code &= @LF & '    Dim intPointer, intUpper, intLower, varHolder'
	$code &= @LF & '    For intPointer = 0 To Int( (intOmegaRow - intAlphaRow) / 2 )'
	$code &= @LF & '        intUpper = intAlphaRow + intPointer'
	$code &= @LF & '        intLower = intOmegaRow - intPointer'
	$code &= @LF & '        varHolder = arrToReverse(intLower)'
	$code &= @LF & '        arrToReverse(intLower) = arrToReverse(intUpper)'
	$code &= @LF & '        arrToReverse(intUpper) = varHolder'
	$code &= @LF & '    Next'
	$code &= @LF & 'End Sub'
;~ 	$code &= @LF & 'Function IIf( expr, truepart, falsepart )'
;~ 	$code &= @LF & '   IIf = falsepart'
;~ 	$code &= @LF & '   If expr Then IIf = truepart'''
;~ 	$code &= @LF & 'End Function'
	Local $vbs = ObjCreate("ScriptControl")
	$vbs.language = "vbscript"
	$vbs.Timeout = -1
	$vbs.addcode($code)
;~ 	FileWrite(@ScriptDir&"\subsort.vbs",$code)
;~ 	$arEither2D1D = $vbs.Run("ArrQSort", $arEither2D1D, $iDesc, $icase)
	$arArray1D = $vbs.Run("ArraySort1", $arArray1D, $iDesc, 0, $iUBoundV - 1, $icase)
	;ArraySort(ByRef SortArray, iAsc,First, Last,icase)
	$vbs = ""
	Return True
EndFunc   ;==>_Sort_vbs1DVStr
Func _SetupSortArrays($iDelimited, $sStrRep, $iSubItems, $iArrayType, $iSort)
	Local $timerstampS = TimerInit()
	If $iDelimited Then;And $iDimension = 2
		For $i = 0 To (UBound($ar_LISTVIEWArrayV) - 1)
			$ar_LISTVIEWArrayV[$i] = "[" & $i & "]" & $sSeparatorV & StringReplace($ar_LISTVIEWArrayV[$i], $sSeparatorV, " " & $sSeparatorV, 0, 1) & $sStrRep;$sBlank
		Next
		ConsoleWrite("$ar_LISTVIEWArrayV[" & $i - 1 & "]=" & $ar_LISTVIEWArrayV[$i - 1] & @LF)
	ElseIf $iSort Then
		Switch $iArrayType ;   $iDimension + 2 * $iTranspose - 1
			Case 0 ;1D
;~ 				If $iUBoundV <= 1000 Then
;~ 				Local $timerstampS2D = TimerInit()
				For $i = 0 To (UBound($avArrayV) - 1)
					$ar_LISTVIEWArrayV[$i] = $avArrayV[$i] & $sSeparatorV & "[" & $i & "]" ;$sSeparatorV &
				Next
;~ 				ConsoleWrite("Time 2D=" & Round(TimerDiff($timerstampS2D), 2) & @LF)
;~ 				Else
;~ 					Local $timerstampS2D = TimerInit()
;~ 					$ar_LISTVIEWArrayV = _ArraySetup1DSt($avArrayV, 0, 0, $sSeparatorV)
;~ 					ConsoleWrite("Time 2DVBS=" & Round(TimerDiff($timerstampS2D), 2) & @LF)
;~ 				EndIf
			Case 1 ;2D
				If $iUBoundV <= 1000 Then
					Local $timerstampS2D = TimerInit()
					For $i = 0 To (UBound($avArrayV) - 1)
						$ar_LISTVIEWArrayV[$i] = "[" & $i & "]" & $sSeparatorV
						For $j = 0 To $iSubItems - 1; Add to text array
							$ar_LISTVIEWArrayV[$i] &= $avArrayV[$i][$j] & $sSeparatorV;$sSeparatorV &
						Next
						$ar_LISTVIEWArrayV[$i] = StringTrimRight($ar_LISTVIEWArrayV[$i], StringLen($sSeparatorV))
					Next
					ConsoleWrite("Time 2D=" & Round(TimerDiff($timerstampS2D), 2) & @LF)
				Else
					Local $timerstampS2D = TimerInit()
					$ar_LISTVIEWArrayV = _ArraySetup2DSt($avArrayV, 0, 0, $sSeparatorV)
					ConsoleWrite("Time 2DVBS=" & Round(TimerDiff($timerstampS2D), 2) & @LF)
				EndIf
			Case 2 ;1DT
				For $i = 0 To (UBound($avArrayV) - 1)
					$ar_LISTVIEWArrayV[$i] = "[" & $i & "]" & $sSeparatorV
					For $j = 0 To $iSubItems - 1; Add to text array
						$ar_LISTVIEWArrayV[$i] &= $avArrayV[$j] & $sSeparatorV;$sSeparatorV &
					Next
;~ 					$ar_LISTVIEWArrayV[$i] &= $sSeparatorV
				Next
			Case 3 ;2DT
				If $iUBoundV <= 1000 Then
					For $i = 0 To (UBound($avArrayV) - 1)
						$ar_LISTVIEWArrayV[$i] = "[" & $i & "]" & $sSeparatorV
						For $j = 0 To $iSubItems - 1; Add to text array
							$ar_LISTVIEWArrayV[$i] &= $avArrayV[$j][$i] & $sSeparatorV;$sSeparatorV &
						Next
					Next
				Else
					Local $timerstampS2D = TimerInit()
					$ar_LISTVIEWArrayV = _ArraySetup2DTrSt($avArrayV, 0, 0, $sSeparatorV)
					ConsoleWrite("Time 2DVBS=" & Round(TimerDiff($timerstampS2D), 2) & @LF)
				EndIf
;~ 					$ar_LISTVIEWArrayV[$i] &= $sSeparatorV
		EndSwitch

	EndIf
	If $iSort Then $arArAscV[0] = $ar_LISTVIEWArrayV
EndFunc   ;==>_SetupSortArrays
Func _ArraySetup1DSt(ByRef Const $arArray, $i_Base = 0, $i_Transpose = 0, $s_Separator = "|")
	Local $ar_ArrayRet, $VBScode2 = 'function ArraySetup1DSt( byref Array,Base,Sep)'
	$VBScode2 &= @LF & '	dim r,ArraySetup21DStF()'
	$VBScode2 &= @LF & '	ReDim preserve ArraySetup21DStF(ubound(Array))'
	$VBScode2 &= @LF & '			For r = 0 To ubound(Array)'
	$VBScode2 &= @LF & '				ArraySetup21DStF(r) = Array(r)&Sep& "[" & r & "]" '
	$VBScode2 &= @LF & '			Next'
	$VBScode2 &= @LF & '	ArraySetup1DSt=ArraySetup21DStF'
	$VBScode2 &= @LF & 'End function   '
	Local $vbs = ObjCreate("ScriptControl")
	$vbs.language = "vbscript"
	$vbs.addcode($VBScode2)
	$ar_ArrayRet = $vbs.Run("ArraySetup1DSt", $arArray, $i_Base, $s_Separator)
	$vbs = ""
;~ 	If $i_Transpose Then _Array2DTranspose($ar6_Array)
	Return $ar_ArrayRet
EndFunc   ;==>_ArraySetup1DSt
Func _ArraySetup2DSt(ByRef Const $arArray, $i_Base = 0, $i_Transpose = 0, $s_Separator = "|")
	Local $ar_ArrayRet, $VBScode2 = 'function ArraySetup2DSt( byref Array,Base,Sep)'
	$VBScode2 &= @LF & '	dim r,c,ArraySetup2DStF()'
	$VBScode2 &= @LF & '	ReDim preserve ArraySetup2DStF(ubound(Array,2))'
	$VBScode2 &= @LF & '		For c = 0 To ubound(Array,2)'
	$VBScode2 &= @LF & '			ArraySetup2DStF(c) = "["&c&"]"&Sep'
	$VBScode2 &= @LF & '			For r = 0 To ubound(Array)' ;Base To ubound(Array)'
	$VBScode2 &= @LF & '				ArraySetup2DStF(c) = ArraySetup2DStF(c)&Array(r,c)&Sep'
	$VBScode2 &= @LF & '			Next'
	$VBScode2 &= @LF & '			ArraySetup2DStF(c)=left(ArraySetup2DStF(c),len(ArraySetup2DStF(c))-len(Sep))'
;~ 	$VBScode2 &= @LF & '			ArraySetup2DStF(r) = ArraySetup2DStF(r)&Sep'
	$VBScode2 &= @LF & '		Next'
	$VBScode2 &= @LF & '	ArraySetup2DSt=ArraySetup2DStF'
	$VBScode2 &= @LF & 'End function   '
	Local $vbs = ObjCreate("ScriptControl")
	$vbs.language = "vbscript"
	$vbs.addcode($VBScode2)
	$ar_ArrayRet = $vbs.Run("ArraySetup2DSt", $arArray, $i_Base, $s_Separator)
	$vbs = ""
;~ 	If $i_Transpose Then _Array2DTranspose($ar6_Array)
	Return $ar_ArrayRet
EndFunc   ;==>_ArraySetup2DSt
Func _ArraySetup2DTrSt(ByRef Const $arArray, $i_Base = 0, $i_Transpose = 0, $s_Separator = "|")
	Local $ar_ArrayRet, $VBScode2 = 'function ArraySetup2DSt( byref Array,Base,Sep)'
	$VBScode2 &= @LF & '	dim r,c,ArraySetup2DStF()'
	$VBScode2 &= @LF & '	ReDim preserve ArraySetup2DStF(ubound(Array))'
	$VBScode2 &= @LF & '		For r = 0 To ubound(Array)'
	$VBScode2 &= @LF & '			ArraySetup2DStF(r) = "["&r&"]"&Sep'
	$VBScode2 &= @LF & '			For c = 0 To ubound(Array,2)' ;Base To ubound(Array)'
	$VBScode2 &= @LF & '				ArraySetup2DStF(r) = ArraySetup2DStF(r)&Array(r,c)&Sep'
	$VBScode2 &= @LF & '			Next'
;~ 	$VBScode2 &= @LF & '			ArraySetup2DStF(r) = ArraySetup2DStF(r)&Sep'
	$VBScode2 &= @LF & '		Next'
	$VBScode2 &= @LF & '	ArraySetup2DSt=ArraySetup2DStF'
	$VBScode2 &= @LF & 'End function   '
	Local $vbs = ObjCreate("ScriptControl")
	$vbs.language = "vbscript"
	$vbs.addcode($VBScode2)
	$ar_ArrayRet = $vbs.Run("ArraySetup2DSt", $arArray, $i_Base, $s_Separator)
	$vbs = ""
;~ 	If $i_Transpose Then _Array2DTranspose($ar6_Array)
	Return $ar_ArrayRet
EndFunc   ;==>_ArraySetup2DTrSt
Func _Sort($i_OrderColumn)
	ConsoleWrite('$i_OrderColumn + 1= ' & $i_OrderColumn + 1 & @CRLF)
	If $i_LISTVIEWPrevcolumnV <> $i_OrderColumn Then $s_LISTVIEWSortOrderV = "DESC"
	If $i_LISTVIEWPrevcolumnV = $i_OrderColumn And $s_LISTVIEWSortOrderV == "DESC" Then ;
		ConsoleWrite('_ArrayRevVerse $s_LISTVIEWSortOrderV=' & $s_LISTVIEWSortOrderV & @CRLF)
		$s_LISTVIEWSortOrderV = "ASC"
	Else
		If ($i_LISTVIEWPrevcolumnV <> $i_OrderColumn Or Not $s_LISTVIEWSortOrderV) And _
				Not IsArray($arArAscV[$i_OrderColumn]) And $i_OrderColumn Then
;~ 			Local $timerstamp13 = TimerInit()
;~ 			Local $avArrayV2 = _Array2DCreateFromDelim($ar_LISTVIEWArrayV)
;~ 			ConsoleWrite("_Array2DCreate Time=" & Round(TimerDiff($timerstamp13)) & "" & @TAB & " msec" & @LF)
			Local $timerstamp13 = TimerInit()
			_ArraySortDelim($ar_LISTVIEWArrayV, $i_OrderColumn + 1, $iCaseV, $sSeparatorV)
			ConsoleWrite("_ArraySortDelim Time=" & Round(TimerDiff($timerstamp13)) & "" & @TAB & " msec" & @LF)
;~ 			Local $timerstamp13 = TimerInit()
			;_ArraySortClib(ByRef $Array, $bCase=False, $bDescend=False, $iStart=0, $iEnd=0, $iColumn=0, $iWidth=256)
;~ 			_ArraySortClib($avArrayV2, False, False, 0, 0, $i_OrderColumn);, $iWidth=256)
;~ 			ConsoleWrite("_ArraySortClib Time=" & Round(TimerDiff($timerstamp13)) & "" & @TAB & " msec" & @LF)
			$arArAscV[$i_OrderColumn] = $ar_LISTVIEWArrayV
		ElseIf ($i_LISTVIEWPrevcolumnV <> $i_OrderColumn Or Not $s_LISTVIEWSortOrderV) Or $i_OrderColumn = 0 Then
			$ar_LISTVIEWArrayV = $arArAscV[$i_OrderColumn]
		EndIf
		$i_LISTVIEWPrevcolumnV = $i_OrderColumn
		$s_LISTVIEWSortOrderV = "DESC"
	EndIf
EndFunc   ;==>_Sort
Func _ArrayUDF_WinSubclassV($hWnd, $lpNewWindowProc)
	;#define GWL_WNDPROC (-4)
	Local $aTmp, $sFunc = "SetWindowLong"
	If Execute("@Unicode") Then $sFunc &= "W"
	$aTmp = DllCall("user32.dll", "ptr", $sFunc, "hwnd", $hWnd, "int", -4, "ptr", $lpNewWindowProc)
	If @error Then Return SetError(1, 0, 0)
	If $aTmp[0] = 0 Then Return SetError(1, 0, 0)
	Return $aTmp[0]
EndFunc   ;==>_ArrayUDF_WinSubclassV
Func _ArrayUDF_CallWndProcV($lpPrevWndFunc, $hWnd, $Msg, $wParam, $lParam)
	Local $aRet = DllCall('user32.dll', 'uint', 'CallWindowProc', 'ptr', $lpPrevWndFunc, 'hwnd', $hWnd, 'uint', $Msg, 'wparam', $wParam, 'lparam', $lParam)
;~  If @error Then Return 0
	Return $aRet[0]
EndFunc   ;==>_ArrayUDF_CallWndProcV
Func _ArraySortDelim(ByRef $arEither2D1D, $iIndex = "1", $iCaseV = 0, $s_Separator = "|")
	If UBound($arEither2D1D, 0) = 2 Then
		Local $arSingle
		_Array2DToArStringsV($arEither2D1D, $arSingle)
		$arEither2D1D = $arSingle
	EndIf
	Local $VBScode = 'function SubSort( byref arSingle,iIndex,icase,Sep)'
	$VBScode &= @LF & '	arIndexN=split(iIndex,"|")'
	$VBScode &= @LF & '	dim iAsc:iAsc=0'
	$VBScode &= @LF & '	if arIndexN(0)<0 then iAsc=1'
	$VBScode &= @LF & '	dim arIndex()'
	$VBScode &= @LF & '	redim preserve arIndex(1)' ;:arIndex(1) =0
	$VBScode &= @LF & '	for a= 0 to ubound(arIndexN)'
	$VBScode &= @LF & '		arIndexN(a)=csng(arIndexN(a))'
	$VBScode &= @LF & '		if arIndexN(a)<>"" and arIndexN(a)<>0 then'
	$VBScode &= @LF & '			redim preserve arIndex(a)'
	$VBScode &= @LF & '			arIndex(a)=abs(arIndexN(a))-1'
	$VBScode &= @LF & '			arIndexN(a)=csng(Iif(arIndexN(a)<0,arIndexN(a)+1,arIndexN(a)-1))'
	$VBScode &= @LF & '		End If'
	$VBScode &= @LF & '	next'
	$VBScode &= @LF & '	ArraySort arSingle, iAsc,  0, UBound(arSingle), arIndex(0),icase,Sep  ' ;SortArray, iAsc,First, Last,iViewColNum
	$VBScode &= @LF & '	for indexcol=1 to ubound(arIndex)' ;col=2'; start with check on 2nd  index'; define first index as col 1 (in call, 1)
	$VBScode &= @LF & '		SubSort1 arSingle,arIndex,indexcol,iif(arIndexN(indexcol)<0,-1,1),icase,Sep  ' ;->, sort order
	$VBScode &= @LF & '	Next'
	$VBScode &= @LF & '	SubSort=arSingle'
	$VBScode &= @LF & 'end function   ' ;==>ArrayFieldSort
	$VBScode &= @LF & 'function SubSort1(byref arArray,arIndex,indexcol, iAsc,icase,Sep)'
	$VBScode &= @LF & '	dim Row,col:col=arIndex(indexcol)'
	$VBScode &= @LF & '	dim pcol:pcol=arIndex(indexcol-1)'
	$VBScode &= @LF & '	iAsc=Iif(iAsc<0,1,0)'
	$VBScode &= @LF & '	dim arTemp(),itemp,	sMarker:sMarker="Equal"' ;,sTemp(1)
	$VBScode &= @LF & '	redim preserve arTemp(1) ' ;redim extra row
	$VBScode &= @LF & '	for Row=1 to UBound(arArray) ' ; go through all rows of 2d array in that column to check for dupes
	$VBScode &= @LF & '			arRow=split(arArray(row),Sep)'
	$VBScode &= @LF & '			arRowB4=split(arArray(row-1),Sep)'
	$VBScode &= @LF & '		if indexcol>1 then ' ;check cols in each row first if more than 2 index cols
	$VBScode &= @LF & '			for c=0 to indexcol-1'
	$VBScode &= @LF & '				if lcase(arRow(arIndex(c)))<>lcase(arRowB4(arIndex(c))) Then'
	$VBScode &= @LF & '					sMarker="pColsNotEqual"'
	$VBScode &= @LF & '					c=indexcol'
	$VBScode &= @LF & '				End If'
	$VBScode &= @LF & '			Next'
	$VBScode &= @LF & '		End If'
	$VBScode &= @LF & '		if lcase(arRow(pcol))=lcase(arRowB4(pcol)) and sMarker="Equal" Then ' ;dupes in the prev col.
	$VBScode &= @LF & '			arTemp(itemp)=arArray(row-1)' ;Array2DToD( arArray2d,"",0,row-1,1) '; set first line of new potential sort array
	$VBScode &= @LF & '			redim preserve arTemp(itemp+1) ' ;redim extra row
	$VBScode &= @LF & '			arTemp(itemp+1)=arArray(row)' ;Array2DToD( arArray2d,"",0,row-1,1) '; set first line of new potential sort array
	$VBScode &= @LF & '			itemp=itemp+1'
	$VBScode &= @LF & '		Else'
	$VBScode &= @LF & '			sMarker="Equal"'
	$VBScode &= @LF & '			SubSortDo1 arArray,arTemp,itemp,iAsc,col,row,icase'
	$VBScode &= @LF & '		End If'
	$VBScode &= @LF & '	Next'
	$VBScode &= @LF & '	SubSortDo1 arArray,arTemp,itemp,iAsc,col,row,icase'
	$VBScode &= @LF & 'end function   '
	$VBScode &= @LF & 'function SubSortDo1(byref arArray,byref arTemp,byref itemp,iAsc,col,row,icase)'
	$VBScode &= @LF & '	dim sTemp()'
	$VBScode &= @LF & '	if itemp>0 then'
	$VBScode &= @LF & '		ArraySort arTemp,  iAsc,0, UBound(arTemp), col ,icase' ;sort on current col (pcol+1), asc?
	$VBScode &= @LF & '		for i= 0 to ubound(arTemp)'
	$VBScode &= @LF & '			arArray(row-ubound(arTemp)+i-1)=arTemp(i)'
	$VBScode &= @LF & '		Next'
	$VBScode &= @LF & '	End If'
	$VBScode &= @LF & '	redim preserve arTemp(1)'
	$VBScode &= @LF & '	itemp=0' ; change backto single, then get into 2d array properly, then go on to check next line'; row=row-1??
	$VBScode &= @LF & '	SubSortDo1= SortArray'
	$VBScode &= @LF & 'end function'
	$VBScode &= @LF & 'Function IIf( expr, truepart, falsepart )'
	$VBScode &= @LF & '   IIf = falsepart'
	$VBScode &= @LF & '   If expr Then IIf = truepart'''
	$VBScode &= @LF & 'End Function'
	$VBScode &= @LF & 'function ArraySort(ByRef SortArray, iAsc,First, Last,iViewColNum,icase,Sep)'
	$VBScode &= @LF & '	dim strWrite,arCol,arColumn()'
	$VBScode &= @LF & '	ReDim Preserve arColumn(ubound(SortArray))'
	$VBScode &= @LF & '	Dim intPointer, booIsNumeric: booIsNumeric = True'
	$VBScode &= @LF & '	For intPointer = First To Last'
	$VBScode &= @LF & '		arCol = Split( SortArray(intPointer), Sep, -1,0 )'
	$VBScode &= @LF & '		If Not IsNumeric( arCol(iViewColNum) ) Then'
	$VBScode &= @LF & '			booIsNumeric = False'
	$VBScode &= @LF & '			Exit For'
	$VBScode &= @LF & '		End If'
	$VBScode &= @LF & '	Next'
	$VBScode &= @LF & '	For intPointer = First To Last'
	$VBScode &= @LF & '		arCol = Split( SortArray(intPointer), Sep, -1,0 )'
	$VBScode &= @LF & '		If booIsNumeric Then'
	$VBScode &= @LF & '			arColumn(intPointer)  = CSng( arCol(iViewColNum) )'
	$VBScode &= @LF & '		else'
	$VBScode &= @LF & '			arColumn(intPointer)  =  arCol(iViewColNum)'
	$VBScode &= @LF & '		End If'
	$VBScode &= @LF & '	Next'
	$VBScode &= @LF & '	if not icase then  QuickSortColumn  SortArray, arColumn,First, Last,iViewColNum'
	$VBScode &= @LF & '	if icase or booIsNumeric then QuickSortColumnCase  SortArray, arColumn,First, Last,iViewColNum'
	$VBScode &= @LF & '	if iasc=1 then ReverseElements SortArray, 0, ubound(SortArray)'
	$VBScode &= @LF & '	ArraySort= SortArray'
	$VBScode &= @LF & 'End function   '
	$VBScode &= @LF & 'function QuickSortColumn(ByRef SortArray, ByRef arColumn, First, Last,iViewColNum)'
	$VBScode &= @LF & '	dim Low,High,collitem,arCol'
	$VBScode &= @LF & '	dim Temp,ListSeparator'
	$VBScode &= @LF & '	Low = First'
	$VBScode &= @LF & '	High = Last'
	$VBScode &= @LF & '	ListSeparator=arColumn((First + Last) / 2)'
	$VBScode &= @LF & '	Do'
	$VBScode &= @LF & '		While lcase(arColumn(Low)) < lcase(ListSeparator)'
	$VBScode &= @LF & '			Low = Low + 1'
	$VBScode &= @LF & '		WEnd'
	$VBScode &= @LF & '		While lcase(arColumn(High)) > lcase(ListSeparator)'
	$VBScode &= @LF & '			High = High - 1'
	$VBScode &= @LF & '		WEnd'
	$VBScode &= @LF & '		If (Low <= High) Then'
	$VBScode &= @LF & '			Temp = SortArray(Low)'
	$VBScode &= @LF & '			SortArray(Low) = SortArray(High)'
	$VBScode &= @LF & '			SortArray(High) = Temp'
	$VBScode &= @LF & '			Temp = arColumn(Low)'
	$VBScode &= @LF & '			arColumn(Low) = arColumn(High)'
	$VBScode &= @LF & '			arColumn(High) = Temp'
	$VBScode &= @LF & '			Low = Low + 1'
	$VBScode &= @LF & '			High = High - 1'
	$VBScode &= @LF & '		End If'
	$VBScode &= @LF & '	Loop While (Low <= High)'
	$VBScode &= @LF & '	If (First < High) Then QuickSortColumn  SortArray, arColumn,First, High,iViewColNum '
	$VBScode &= @LF & '	If (Low < Last) Then QuickSortColumn  SortArray, arColumn,Low, Last,iViewColNum '
	$VBScode &= @LF & 'End function   '
	$VBScode &= @LF & 'function QuickSortColumnCase(ByRef SortArray, ByRef arColumn, First, Last,iViewColNum)'
	$VBScode &= @LF & '	dim Low,High,collitem,arCol'
	$VBScode &= @LF & '	dim Temp,ListSeparator'
	$VBScode &= @LF & '	Low = First'
	$VBScode &= @LF & '	High = Last'
	$VBScode &= @LF & '	ListSeparator=arColumn((First + Last) / 2)'
	$VBScode &= @LF & '	Do'
	$VBScode &= @LF & '		While (arColumn(Low) < ListSeparator)'
	$VBScode &= @LF & '			Low = Low + 1'
	$VBScode &= @LF & '		WEnd'
	$VBScode &= @LF & '		While (arColumn(High) > ListSeparator)'
	$VBScode &= @LF & '			High = High - 1'
	$VBScode &= @LF & '		WEnd'
	$VBScode &= @LF & '		If (Low <= High) Then'
	$VBScode &= @LF & '			Temp = SortArray(Low)'
	$VBScode &= @LF & '			SortArray(Low) = SortArray(High)'
	$VBScode &= @LF & '			SortArray(High) = Temp'
	$VBScode &= @LF & '			Temp = arColumn(Low)'
	$VBScode &= @LF & '			arColumn(Low) = arColumn(High)'
	$VBScode &= @LF & '			arColumn(High) = Temp'
	$VBScode &= @LF & '			Low = Low + 1'
	$VBScode &= @LF & '			High = High - 1'
	$VBScode &= @LF & '		End If'
	$VBScode &= @LF & '	Loop While (Low <= High)'
	$VBScode &= @LF & '	If (First < High) Then QuickSortColumnCase  SortArray, arColumn,First, High,iViewColNum '
	$VBScode &= @LF & '	If (Low < Last) Then QuickSortColumnCase  SortArray, arColumn,Low, Last,iViewColNum '
	$VBScode &= @LF & 'End function   '
	$VBScode &= @LF & 'Sub ReverseElements( arrToReverse, intAlphaRow, intOmegaRow )'
	$VBScode &= @LF & '    Dim intPointer, intUpper, intLower, varHolder'
	$VBScode &= @LF & '    For intPointer = 0 To Int( (intOmegaRow - intAlphaRow) / 2 )'
	$VBScode &= @LF & '        intUpper = intAlphaRow + intPointer'
	$VBScode &= @LF & '        intLower = intOmegaRow - intPointer'
	$VBScode &= @LF & '        varHolder = arrToReverse(intLower)'
	$VBScode &= @LF & '        arrToReverse(intLower) = arrToReverse(intUpper)'
	$VBScode &= @LF & '        arrToReverse(intUpper) = varHolder'
	$VBScode &= @LF & '    Next'
	$VBScode &= @LF & 'End Sub'
	Local $vbs = ObjCreate("ScriptControl")
	$vbs.language = "vbscript"
	$vbs.Timeout = -1
	$vbs.addcode($VBScode)
;~ 	FileWrite(@ScriptDir&"\subsort.vbs",$code)
	$arEither2D1D = $vbs.Run("SubSort", $arEither2D1D, $iIndex, $iCaseV, $s_Separator)
	$vbs = ""
EndFunc   ;==>_ArraySortDelim
Func _Array2DToArStringsV(ByRef $ar5_Array, ByRef $a_Rowa2, $s_Separator = "|");_Array2DToArStringsV
	Local $code = 'function Array2DtoArrayStrings( byref ar2ArrayStrings)'
	$code &= @LF & '	dim r,c,intPointer,ar1ArrayStrings()'
	$code &= @LF & '	ReDim preserve ar1ArrayStrings(ubound(ar2ArrayStrings,2))'
	$code &= @LF & '		For r = 0 To ubound(ar2ArrayStrings,2) '
	$code &= @LF & '			For c = 0 To ubound(ar2ArrayStrings)'
	$code &= @LF & '						ar1ArrayStrings(r)=ar1ArrayStrings(r)&"' & $s_Separator & '"&ar2ArrayStrings(c,r) '
	$code &= @LF & '			Next'
	$code &= @LF & '			ar1ArrayStrings(r)=mid(ar1ArrayStrings(r),2,len(ar1ArrayStrings(r))-1)'
	$code &= @LF & '		Next'
	$code &= @LF & '	Array2DtoArrayStrings=ar1ArrayStrings'
	$code &= @LF & 'end function'
	Local $vbs = ObjCreate("ScriptControl")
	$vbs.language = "vbscript"
	$vbs.addcode($code)
	$vbs.Timeout = -1
	$a_Rowa2 = $vbs.Run("Array2DtoArrayStrings", $ar5_Array)
	$vbs = ""
	Return 1
EndFunc   ;==>_Array2DToArStringsV
Func _Array2DCreateFromDelim($ar1_Array_Strings, $i_Base = 0, $i_Transpose = 0, $s_Separator = "|")
	Local $ar6_Array, $VBScode2 = 'function Array2DCreate( byref ar1Strings,Base,Sep)'
	$VBScode2 &= @LF & '	dim r,c,NewUbound,ar2ArrayStrings()'
	$VBScode2 &= @LF & '	For r = Base To ubound(ar1Strings) '
	$VBScode2 &= @LF & '		aRowArray = split(ar1Strings(r),Sep)'
	$VBScode2 &= @LF & '		if ubound(aRowArray)>NewUbound then'
	$VBScode2 &= @LF & '			NewUbound=ubound(aRowArray)'
	$VBScode2 &= @LF & '		end if'
	$VBScode2 &= @LF & '	next'
	$VBScode2 &= @LF & '	ReDim preserve ar2ArrayStrings(NewUbound,ubound(ar1Strings)-base)'
	$VBScode2 &= @LF & '		For r = Base To ubound(ar1Strings)'
	$VBScode2 &= @LF & '		aRowArray = split(ar1Strings(r),Sep)'
	$VBScode2 &= @LF & '			For c = 0 To ubound(aRowArray)'
	$VBScode2 &= @LF & '				ar2ArrayStrings(c,r-base) = aRowArray(c)'
	$VBScode2 &= @LF & '			Next'
	$VBScode2 &= @LF & '		Next'
	$VBScode2 &= @LF & '	Array2DCreate=ar2ArrayStrings'
	$VBScode2 &= @LF & 'End function   '
	Local $vbs = ObjCreate("ScriptControl")
	$vbs.language = "vbscript"
	$vbs.addcode($VBScode2)
	$ar6_Array = $vbs.Run("Array2DCreate", $ar1_Array_Strings, $i_Base, $s_Separator)
	$vbs = ""
;~ 	If $i_Transpose Then _Array2DTranspose($ar6_Array)
	Return $ar6_Array
EndFunc   ;==>_Array2DCreateFromDelim
Func _ArrayRevV(ByRef $avArray, $i_Base = 0, $i_ubound = 0)
	If Not IsArray($avArray) Then
		SetError(1)
		Return 0
	EndIf
	Local $tmp, $last = UBound($avArray) - 1
	If $i_ubound < 1 Or $i_ubound > $last Then $i_ubound = $last
	For $i = $i_Base To $i_Base + Int(($i_ubound - $i_Base - 1) / 2)
		$tmp = $avArray[$i]
		$avArray[$i] = $avArray[$i_ubound]
		$avArray[$i_ubound] = $tmp
		$i_ubound = $i_ubound - 1
	Next
	Return 1
EndFunc   ;==>_ArrayRevV
Func _ArraySortClibsetV(ByRef $Array, $bCase = False, $bDescend = False, $iStart = 0, $iEnd = 0, $iColumn = 0, $iWidth = 256)

	Local $iArrayDims = UBound($avArrayV, 0)
	If @error Or $iArrayDims > 2 Then Return SetError(1, 0, 0) ;;if not array or more than 2D, abort

	Local $iArraySize = UBound($avArrayV, 1), $iColumnMax = 0, $i, $hMem, $iCount;, $hMsvcrtV, $tSourceV,, $sStrCmpV,$tagSourceV,$pSourceV,$procStrCmpV

	If $iArraySize < 2 Then Return 0 ;if no sorting necessary, abort

	If $iEnd < 1 Or $iEnd > $iArraySize - 1 Then $iEnd = $iArraySize - 1
	If ($iEnd - $iStart < 2) Then Return SetError(2, 0, 0) ;;invalid param, no sorting necessary

	If $iArrayDims = 2 Then
		$iColumnMax = UBound($avArrayV, 2)
		If ($iColumnMax - $iColumn < 0) Then Return SetError(2, 0, 0) ;;invalid param
	EndIf

	;; intitialize sorting proc
	;MSDN: The _strcmpi function is equivalent to _stricmp and is provided for backward compatibility only.
	$sStrCmpV = '_strcmpi' ;case insensitive
	If $bCase Then $sStrCmpV = 'strcmp' ;case sensitive
;~  $sStrCmpV = '_strnicmp' ;case insensitive
;~  If $bCase Then $sStrCmpV = 'strncmp' ;case sensitive
	$hMsvcrtV = DllCall('kernel32.dll', 'hwnd', 'LoadLibraryA', 'str', 'msvcrt.dll')
	$hMsvcrtV = $hMsvcrtV[0]
	$procStrCmpV = DllCall('kernel32.dll', 'ptr', 'GetProcAddress', 'ptr', $hMsvcrtV, 'str', $sStrCmpV)
	If $procStrCmpV[0] = 0 Then
		DllCall('kernel32.dll', 'hwnd', 'FreeLibrary', 'hwnd', $hMsvcrtV)
		Return -1
	Else
		$procStrCmpV = $procStrCmpV[0]
	EndIf

	;; initialize memory
	$tagSourceV = ""
	For $i = 1 To $iArraySize
		$tagSourceV &= "char[" & $iWidth & "];"
	Next
	$tagSourceV = StringTrimRight($tagSourceV, 1)
	$tSourceV = DllStructCreate($tagSourceV)

	;; fill memory
	ConsoleWrite("$iArrayDims=" & $iArrayDims & @LF)
	If $iArrayDims = 1 Then
		For $i = 0 To $iArraySize - 1
			ConsoleWrite('$avArrayV[$i] & $sSeparatorV & "[" & $i & "]"=' & $avArrayV[$i] & $sSeparatorV & "[" & $i & "]" & @LF)
			DllStructSetData($tSourceV, $i + 1, $avArrayV[$i] & $sSeparatorV & "[" & $i & "]")
;~ 			DllStructSetData($tSourceV, $i + 1, $avArrayV[$i])
		Next
	Else
		For $i = 0 To $iArraySize - 1
			DllStructSetData($tSourceV, $i + 1, $avArrayV[$i][$iColumn])
		Next
	EndIf
	Return 1
	;;pointer to search starting element
	$pSourceV = DllStructGetPtr($tSourceV, $iStart + 1)
	;;count of elements to search
	$iCount = $iEnd - $iStart + 1

	;;sort
	DllCall('msvcrt.dll', 'none:cdecl', 'qsort', 'ptr', $pSourceV, 'int', $iCount, 'int', $iWidth, 'ptr', $procStrCmpV)

;~ ;; read back the result
;~     If $iArrayDims = 1 Then
;~         ; 1D
;~         If $bDescend Then
;~             For $i = $iStart To $iEnd
;~                 $avArrayV[$iEnd+$iStart-$i] = DllStructGetData($tSourceV, $i+1)
;~             Next
;~         Else
;~             For $i = $iStart To $iEnd
;~                 $avArrayV[$i] = DllStructGetData($tSourceV, $i+1)
;~             Next
;~         EndIf
;~     Else
;~         ;2D
;~         Local $aTmp[$iArraySize][$iColumnMax], $aState[$iCount], $aRet, $iIndex = -1
;~         If $iStart > 0 Then
;~             For $i = 0 To $iStart-1
;~                 For $j = 0 To $iColumnMax-1
;~                     $aTmp[$i][$j] = $avArrayV[$i][$j]
;~                 Next
;~             Next
;~         EndIf
;~         If $iEnd < $iArraySize-1 Then
;~             For $i = $iEnd+1 To $iArraySize-1
;~                 For $j = 0 To $iColumnMax-1
;~                     $aTmp[$i][$j] = $avArrayV[$i][$j]
;~                 Next
;~             Next
;~         EndIf
;~         If $bDescend Then
;~             For $i = $iStart To $iEnd
;~                 $aRet = DllCall('msvcrt.dll', 'int:cdecl', 'bsearch', 'str', $avArrayV[$i][$iColumn], 'ptr', $pSourceV, 'int', $iCount, 'int', $iWidth, 'ptr', $procStrCmpV)
;~                 If Not @error Then $iIndex = ( Dec(Hex($aRet[0]))- Dec(Hex($pSourceV)) ) / $iWidth
;~                 While $iIndex > 0
;~                     If $avArrayV[$i][$iColumn] <> DllStructGetData($tSourceV, $iIndex+$iStart) Then ExitLoop
;~                     $iIndex -= 1
;~                 WEnd
;~                 While ($iIndex < $iCount) And ($aState[$iIndex] = 1)
;~                     $iIndex += 1
;~                 WEnd
;~                 $aState[$iIndex] = 1
;~                 $iIndex = $iEnd - $iIndex
;~                 For $j = 0 To $iColumnMax-1
;~                     $aTmp[$iIndex][$j] = $avArrayV[$i][$j]
;~                 Next
;~             Next
;~         Else
;~             ;ascending
;~             For $i = $iStart To $iEnd
;~                 $aRet = DllCall('msvcrt.dll', 'int:cdecl', 'bsearch', 'str', $avArrayV[$i][$iColumn], 'ptr', $pSourceV, 'int', $iCount, 'int', $iWidth, 'ptr', $procStrCmpV)
;~                 If Not @error Then $iIndex = ( Dec(Hex($aRet[0]))- Dec(Hex($pSourceV)) ) / $iWidth
;~                 While $iIndex > 0
;~                     If $avArrayV[$i][$iColumn] <> DllStructGetData($tSourceV, $iIndex+$iStart) Then ExitLoop
;~                     $iIndex -= 1
;~                 WEnd
;~                 While ($iIndex < $iCount) And ($aState[$iIndex] = 1)
;~                     $iIndex += 1
;~                 WEnd
;~                 $aState[$iIndex] = 1
;~                 $iIndex += $iStart
;~                 For $j = 0 To $iColumnMax-1
;~                     $aTmp[$iIndex][$j] = $avArrayV[$i][$j]
;~                 Next
;~             Next
;~         EndIf
;~         $avArrayV = $aTmp
;~     EndIf

;~     DllCall('kernel32.dll', 'hwnd', 'FreeLibrary', 'hwnd', $hMsvcrtV)
;~     Return 1
;~
EndFunc   ;==>_ArraySortClibsetV