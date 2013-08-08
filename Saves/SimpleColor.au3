#include <ColorPicker.au3>
GUICreate('Color Picker', 179, 100)
$Label = GUICtrlCreateLabel('', 5, 5, 90, 90, 0x1000)
GUICtrlSetBkColor(-1, 0xFF6600)

; Create Picker
$Picker = _GUIColorPicker_Create('Color...', 102, 70, 70, 23, 0xFF6600, BitOR($CP_FLAG_DEFAULT, $CP_FLAG_TIP))
GUISetState()
While 1
	$Msg = GUIGetMsg()
	Switch $Msg
		Case - 3
			ExitLoop

		Case $Picker
			GUICtrlSetBkColor($Label, _GUIColorPicker_GetColor($Picker))
	EndSwitch
WEnd
