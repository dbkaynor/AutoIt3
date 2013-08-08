Opt("MustDeclareVars", 1)
#include <GUIConstants.au3>
#include <GuiListView.au3>
#include <Misc.au3>
#include "_DougFunctions.au3"

If _Singleton(@ScriptName, 1) = 0 Then
	Debug(@ScriptName & " is already running!", 0x40)
	Exit
EndIf

Global $Main = GUICreate(@ScriptName & " " & @OSTYPE & "  " &  @OSArch, 400, 270, 20, 20) ;, $WS_SIZEBOX)

GUICtrlCreateGroup("Web browser", 10, 10, 80, 70)
Global $RadioFirefox = GUICtrlCreateRadio("Firefox", 20, 30, 50)
Global $RadioExplorer = GUICtrlCreateRadio("IE", 20, 50, 50)
GUICtrlCreateGroup("", -99, -99, 1, 1) ;close group

; "text", left, top [, width [, height

GUICtrlCreateGroup("Web sites", 100, 10, 200, 160)
Global $RadioWells = GUICtrlCreateRadio("Wells", 110, 30, 80)
Global $RadioBofA = GUICtrlCreateRadio("BofA", 110, 50, 80)
Global $RadioHorde = GUICtrlCreateRadio("Horde", 110, 70, 80)
Global $RadioVerizon = GUICtrlCreateRadio("Verizon", 110, 90, 80)
Global $RadioYahoo = GUICtrlCreateRadio("Yahoo", 110, 110, 80)
Global $RadioGoogle = GUICtrlCreateRadio("Google", 110, 130, 80)
Global $RadioWikipedia = GUICtrlCreateRadio("Wikipedia", 110, 150, 80)
Global $RadioKATU = GUICtrlCreateRadio("KATU", 200, 30, 80)
Global $RadioKGW = GUICtrlCreateRadio("KGW", 200, 50, 80)
Global $RadioKOIN = GUICtrlCreateRadio("KOIN", 200, 70, 80)
Global $RadioKPDX = GUICtrlCreateRadio("KPDX", 200, 90, 80)
Global $RadioWebMail = GUICtrlCreateRadio("GMail", 200, 110, 80)

GUICtrlCreateGroup("", -99, -99, 1, 1) ;close group

Global $label_string = GUICtrlCreateLabel("???", 10, 200, 500, 50)

Global $button_go = GUICtrlCreateButton("Go", 330, 10, 50)
Global $button_exit = GUICtrlCreateButton("Exit", 330, 50, 50)

Local $F1 = FileGetShortName("C:\Program Files (x86)\Mozilla Firefox\firefox.exe")
Local $F2 = FileGetShortName("C:\Program Files\Mozilla Firefox\firefox.exe")
Global $Firefox
If FileExists($F1) Then
	$Firefox = $F1
ElseIf FileExists($F2) Then
	$Firefox = $F2
Else
	Debug("FireFox not found.")
	Exit
EndIf

Local $F1 = FileGetShortName("C:\Program Files (x86)\Internet Explorer\iexplore.exe")
Local $F2 = FileGetShortName("C:\Program Files\Internet Explorer\iexplore.exe")
Global $Explorer
If FileExists($F1) Then
	$Explorer = $F1
ElseIf FileExists($F2) Then
	$Explorer = $F2
Else
	Debug("Explorer not found.")
	Exit
EndIf

GUICtrlSetState($radioFirefox, $GUI_CHECKED)
Global $Browser = $Firefox
GUICtrlSetState($RadioYahoo, $GUI_CHECKED)
Global $Web = "www.yahoo.com"


GUISetState()
; Run the GUI until the dialog is closed
While 1
	Local $msg = GUIGetMsg()
	Select
		Case $msg = $RadioFirefox And BitAND(GUICtrlRead($RadioFirefox), $GUI_CHECKED) = $GUI_CHECKED
			$Browser = $Firefox
		Case $msg = $RadioExplorer And BitAND(GUICtrlRead($RadioExplorer), $GUI_CHECKED) = $GUI_CHECKED
			$Browser = $Explorer
		Case $msg = $RadioVerizon And BitAND(GUICtrlRead($RadioVerizon), $GUI_CHECKED) = $GUI_CHECKED
			$Web = "www.verizon.com"
		Case $msg = $RadioHorde And BitAND(GUICtrlRead($RadioHorde), $GUI_CHECKED) = $GUI_CHECKED
			$Web = "webmail.kaynor.net/horde/imp/login.php"
		Case $msg = $RadioBofA And BitAND(GUICtrlRead($RadioBofA), $GUI_CHECKED) = $GUI_CHECKED
			$Web = "www.bofa.com"
		Case $msg = $RadioWells And BitAND(GUICtrlRead($RadioWells), $GUI_CHECKED) = $GUI_CHECKED
			$Web = "www.wellsfargo.com"
		Case $msg = $RadioYahoo And BitAND(GUICtrlRead($RadioYahoo), $GUI_CHECKED) = $GUI_CHECKED
			$Web = "www.yahoo.com"
		Case $msg = $RadioGoogle And BitAND(GUICtrlRead($RadioGoogle), $GUI_CHECKED) = $GUI_CHECKED
			$Web = "www.google.com"
		Case $msg = $RadioWikipedia And BitAND(GUICtrlRead($RadioWikipedia), $GUI_CHECKED) = $GUI_CHECKED
			$Web =  "en.wikipedia.org/wiki/Main_Page"
		Case $msg = $RadioKATU And BitAND(GUICtrlRead($RadioKATU), $GUI_CHECKED) = $GUI_CHECKED
			$Web = "www.katu.com"
		Case $msg = $RadioKGW And BitAND(GUICtrlRead($RadioKGW), $GUI_CHECKED) = $GUI_CHECKED
			$Web = "www.kgw.com"
		Case $msg = $RadioKOIN And BitAND(GUICtrlRead($RadioKOIN), $GUI_CHECKED) = $GUI_CHECKED
			$Web = "www.koin.com"
		Case $msg = $RadioKPDX And BitAND(GUICtrlRead($RadioKPDX), $GUI_CHECKED) = $GUI_CHECKED
			$Web = "www.kpdx.com"
		case $msg = $RadioWebMail And BitAND(GUICtrlRead($RadioWebMail), $GUI_CHECKED) = $GUI_CHECKED
			$Web = "mail.google.com/mail/?hl=en&tab=wm#inbox"
		Case $msg = $button_go
			Go();
		Case $msg = $button_exit Or $msg = $GUI_EVENT_CLOSE
			ExitLoop
		Case $msg = $GUI_EVENT_PRIMARYUP
			GUICtrlSetData($label_string, "Browser: " & $Browser & @CRLF & "Web site: " & $Web)
			ConsoleWrite("Browser: " & $Browser & "  Web site: " & $Web & @CRLF)
	EndSelect
WEnd

Func Go()
	If FileExists($Browser) = False Then
		MsgBox(0, "ERROR", $Browser & " does not exist", 5)
		Exit
	EndIf
	debug($Browser & " " & $Web)
	Run($Browser & " " & $Web)
	;RunWait($Browser & " " & $web)
EndFunc   ;==>Go

;$ie->gotoURL('https://login.verizonwireless.com/amserver/UI/Login?realm=vzw&goto=https%3A%2F%2Fwbillpay.verizonwireless.com%3A443%2Fvzw%2Faccountholder%2Foverview%2Funbilled-usage-minutes.do');