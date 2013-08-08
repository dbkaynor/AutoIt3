;===============================================================================
; Function:         _ColorPicker
; Version:          AutoIt Version: 3.3.0.0
; Description:      Color Picker Tool for Excel Current Workbook Pallet
;                   Returns the selected color
; Syntax:           _ColorPicker($_CPleft, $_CPTop, $_DefaultColor = "", $_Title = "Color Picker")
; Parameter(s):     $_CPleft - Left position of the Color Picker Window
;                   $_CPTop - Top position of the Color Picker Window
;                   $_DefaultColor - Array (element)
;                   $_Title - Optional title of the Color Picker Window
;                   $_Font - Logical, if True, then it is a font color, important for the default return which should be black
; Requirement(s):   None
; Return Value(s):  On Success - Returns an Array Array with the following format:
;                       [0] AutoIt Hex color
;                       [1] Excel Color number
;                   When Default/Transparent button is clicked
;                       [0] AutoIt color 0xFFFFFF
;                       [1] Excel Color 0 (black/transparent)
;                   On Failure, more exactly on Escape or Enter - Returns:
;                       [0] -1 or the default Hex color
;                       [1] -1 or the default Excel Color number
;                        @error=1 - Specified object does not exist
; Author(s):        A. Greencan  - March 2009
; Note(s):          Excel Color information can be found on
;                   http://www.mvps.org/dmcritchie/excel/colors.htm
;
;                   Color table $ExcelColors
;                   The table is the standard palet of Excel 2003
;                   The color table can be expanded to contain all color elements, from 0 (transparent) to color 56
;                   The window size will grow dynamically
;                   Each color element contains:
;                       [0] Hex number of the Windows color
;                       [1] Number of the Excel color
;                       [2] Name of the color
;                       [3] Return value of the button
;
;===============================================================================
#include-once
#include <WindowsConstants.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <Array.au3>
Global $ExcelColors[40][4] = [ _
		[0x000000, 1, "Black", ""], _
		[0x993300, 53, "Brown", ""], _
		[0x333300, 52, "Olive Green", ""], _
		[0x003300, 51, "Dark Green", ""], _
		[0x003366, 49, "Dark Teal", ""], _
		[0x000080, 11, "Dark Blue", ""], _
		[0x333399, 55, "Indigo", ""], _
		[0x333333, 56, "Gray-80%", ""], _
		[0x800000, 9, "Dark Red", ""], _
		[0xFF6600, 46, "Orange", ""], _
		[0x808000, 12, "Dark Yellow", ""], _
		[0x008000, 10, "Green", ""], _
		[0x008080, 14, "Teal", ""], _
		[0x0000FF, 5, "Blue", ""], _
		[0x666699, 47, "Blue-Gray", ""], _
		[0x808080, 16, "Gray-50%", ""], _
		[0xFF0000, 3, "Red", ""], _
		[0xFF9900, 45, "Light Orange", ""], _
		[0x99CC00, 43, "Lime", ""], _
		[0x339966, 50, "Sea Green", ""], _
		[0x33CCCC, 42, "Aqua", ""], _
		[0x3366FF, 41, "Light Blue", ""], _
		[0x800080, 13, "Violet", ""], _
		[0x969696, 48, "Gray-40%", ""], _
		[0xFF00FF, 7, "Pink", ""], _
		[0xFFCC00, 44, "Gold", ""], _
		[0xFFFF00, 6, "Yellow", ""], _
		[0x00FF00, 4, "Bright Green", ""], _
		[0x00FFFF, 8, "Turquoise", ""], _
		[0x00CCFF, 33, "Sky Blue", ""], _
		[0x993366, 54, "Plum", ""], _
		[0xC0C0C0, 15, "Gray-25%", ""], _
		[0xFF99CC, 38, "Rose", ""], _
		[0xFFCC99, 40, "Tan", ""], _
		[0xFFFF99, 36, "Light Yellow", ""], _
		[0xCCFFCC, 35, "Light Green", ""], _
		[0xCCFFFF, 34, "Light Turquoise", ""], _
		[0x99CCFF, 37, "Pale Blue", ""], _
		[0xCC99FF, 39, "Lavender", ""], _
		[0xFFFFFF, 2, "White", ""]]

Func _RandomColor()
	;_ArrayDisplay($ExcelColors,UBound($ExcelColors))
	Local $_a = Random(0, UBound($ExcelColors)-1, 1)
	Local $res[1][2] = [[$ExcelColors[$_a][0], $ExcelColors[$_a][2]]]
	Return ($res)
EndFunc   ;==>RandomColor

Func _ColorPicker($_CPleft = -1, $_CPTop = -1, $_DefaultColor = "", $_Title = "Color picker", $_Font = False)
	Local $_CP_GUI, $_TransparentBtn, $_HorCounter, $_VertCounter, $_a, $_msg, $_result, $_DummytBtn

	If $_DefaultColor <> "" Then
		ConsoleWrite("default : " & "0x" & Hex($_DefaultColor[0][0], 6) & " - Excel: " & $_DefaultColor[0][1] & @CR)
		If $_DefaultColor[0][1] = 0 Then
			$_result = -1 ; set for Default/Transparent button
		Else
			$_result = _ArraySearch($ExcelColors, $_DefaultColor[0][0], -1, -1, -1, -1, 0)
		EndIf
	EndIf
	ConsoleWrite("result: " & $_result & @CR)
	If $_CPleft = -1 Then $_CPleft = ((@DesktopWidth - 215) / 2)
	If $_CPTop = -1 Then $_CPTop = ((@DesktopHeight - 165) / 2)

	$_CP_GUI = GUICreate($_Title, 215, 35 + (26 * Round((0.4 + (UBound($ExcelColors) / 8)), 0)), $_CPleft, $_CPTop, -1, $WS_EX_TOOLWINDOW)
	; dummy button, need to create this to catch the enter which means no change
	$_DummytBtn = GUICtrlCreateButton("", 0, 0, 0, 0, $BS_DEFPUSHBUTTON)
	ConsoleWrite( $_DummytBtn)
	If $_result = -1 Then
		GUICtrlCreateLabel("", 5, 5, 205, 26) ;,$SS_BLACKRECT )
		$_TransparentBtn = GUICtrlCreateButton("Default/Transparent", 7, 7, 201, 22)
	Else
		$_TransparentBtn = GUICtrlCreateButton("Default/Transparent", 5, 5, 205, 26)
	EndIf
	GUICtrlSetBkColor(-1, 0xFFFFFF)
	GUICtrlSetTip(-1, "Default")
	$_HorCounter = 0
	$_VertCounter = 0
	For $_a = 0 To UBound($ExcelColors) - 1
		If $_HorCounter = 8 Then
			$_HorCounter = 1
			$_VertCounter = $_VertCounter + 1
		Else
			$_HorCounter = $_HorCounter + 1
		EndIf
		If $_a = $_result Then ; default color
			GUICtrlCreateLabel("", 5 - 26 + ($_HorCounter * 26), 5 + 26 + ($_VertCounter * 26), 24, 24);,$SS_BLACKRECT )
			$ExcelColors[$_a][3] = GUICtrlCreateButton("", 7 - 26 + ($_HorCounter * 26), 7 + 26 + ($_VertCounter * 26), 20, 20)
		Else
			$ExcelColors[$_a][3] = GUICtrlCreateButton("", 5 - 26 + ($_HorCounter * 26), 5 + 26 + ($_VertCounter * 26), 24, 24)
		EndIf
		GUICtrlSetBkColor(-1, $ExcelColors[$_a][0])
		GUICtrlSetTip(-1, $ExcelColors[$_a][2])

	Next

	GUISetState()

	$_msg = 0
	While $_msg <> $GUI_EVENT_CLOSE
		$_msg = GUIGetMsg()
		Select
			Case $_msg = $_TransparentBtn ; Transparent / default button will return 0
				GUIDelete($_CP_GUI)
				If $_Font = True Then
					Local $res[1][2] = [[0x000000, 0]] ; foreground should be black
				Else
					Local $res[1][2] = [[0xFFFFFF, 0]] ; background should be white
				EndIf
				Return ($res)
			Case $_msg = 6 Or $_msg = $GUI_EVENT_CLOSE; enter or Escape - no change, will return -1 or the default color
				GUIDelete($_CP_GUI)
				If $_DefaultColor <> "" Then ; we will keep the default color if known, otherwise, return -1
					;local $res[1][2] = [[$ExcelColors[$_result][0],$ExcelColors[$_result][1]]]
					Return ($_DefaultColor)
				Else
					Local $res[1][2] = [[-1, -1]]
					Return ($res)
				EndIf
			Case Else
				; returns selected color numbers
				For $_a = 0 To UBound($ExcelColors) - 1
					If $_msg = $ExcelColors[$_a][3] Then
						GUIDelete($_CP_GUI)
						Local $res[1][2] = [[$ExcelColors[$_a][0], $ExcelColors[$_a][2]]]
						Return ($res)
					EndIf
				Next
		EndSelect
	WEnd
EndFunc   ;==>ColorPicker
