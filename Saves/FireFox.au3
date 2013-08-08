; *******************************************************
;
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <ff.au3>

GUICreate("Firefox Web Power control Test", 480, 100, 50, 50)
$GUI_Button_Start = GUICtrlCreateButton("Start", 10, 10, 50, 30)
$GUI_Button_Login = GUICtrlCreateButton("Login", 60, 10, 50, 30)
$GUI_Button_On = GUICtrlCreateButton("On", 110, 10, 50, 30)
$GUI_Button_Off = GUICtrlCreateButton("Off", 160, 10, 50, 30)
$GUI_Button_Read = GUICtrlCreateButton("Read", 210, 10, 50, 30)
$GUI_Button_Reset = GUICtrlCreateButton("Reset", 340, 10, 50, 30)
$GUI_Input = GUICtrlCreateInput("", 10, 50, 400, 30)

GUISetState() ;Show GUI

While 1
	$msg = GUIGetMsg()
	Select
		Case $msg = $GUI_EVENT_CLOSE
			ExitLoop
		Case $msg = $GUI_Button_Start
			ShellExecute("C:\Program Files (x86)\Mozilla Firefox\firefox.exe", "-height 500 -width 800")
			If _FFConnect(Default, Default, 3000) Then
				_FFAction("Min")
				_FFOpenURL("http://192.168.10.100/")
				_FFClick("Submitbtn", "name", 0)
			EndIf
		Case $msg = $GUI_Button_Login
		;	_FFClick("Submitbtn", "name", 0)
		Case $msg = $GUI_Button_Read
			_FFTabAdd("http://192.168.10.100/Set.cmd?CMD=GetPower")
			GUICtrlSetData($GUI_Input, _FFReadText())
			_FFTabClose()
		Case $msg = $GUI_Button_On
			_FFTabAdd("http://192.168.10.100/Set.cmd?CMD=SetPower&P60=1&P61=1&P62=1&P63=1")
			GUICtrlSetData($GUI_Input, _FFReadText())
			_FFTabClose()
		Case $msg = $GUI_Button_Off
			_FFTabAdd("http://192.168.10.100/Set.cmd?CMD=SetPower&P60=0&P61=0&P62=0&P63=0")
			GUICtrlSetData($GUI_Input, _FFReadText())
			_FFTabClose()
		Case $msg = $GUI_Button_Reset
			_FFDisconnect()
			MsgBox(64, @ScriptLineNumber, "Disconnected from Firefox")
	EndSelect
WEnd

Exit