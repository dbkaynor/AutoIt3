#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Run_Tidy=y
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****

Opt("MustDeclareVars", 1) ; 0=no, 1=require pre-declare
Opt("GUICoordMode", 1) ; 0=relative, 1=absolute, 2=cell
Opt("GUIResizeMode", 1) ; 0=no resizing, <1024 special resizing

If _Singleton(@ScriptName, 1) = 0 Then
	Debug(@ScriptName & " is already running!", 0x40, 2)
	Exit
EndIf

#include <GUIConstants.au3>
#include <GuiListBox.au3>
#include <String.au3>
#include <Array.au3>
#include <Misc.au3>
#include "_DougFunctions.au3"

DirCreate("temp")
DirCreate("AUXFiles")

Local $tmp = StringSplit(@ScriptName, ".")
Global $Projectfilename = ".\AUXFiles\" & $tmp[1] & ".prj"

Global Const $IPConfig_filename = ".\temp\IPConfig.out"
Global Const $ARP_filename = ".\temp\arp.out"
Global Const $Initial_filename = ".\temp\ArpInit.out"
Global Const $Batch_filename = ".\temp\ARP_cmds.bat"
Local $SystemS = @OSVersion & "  " & @OSServicePack & "  " & @OSType & "  " & @CPUArch & "  " & @IPAddress1
Global $Main = GUICreate("Arp tool version: 0.0.0.2 running on " & $SystemS, 800, 630, 20, 20, $WS_SIZEBOX)

Local $font = "Courier new"
GUISetFont(10, 400, -1, $font)
GUISetState()

; GUICtrlCreateButton ( "text", left, top [, width [, height [, style [, exStyle]]]] )

Global $button_get = GUICtrlCreateButton("Get", 10, 15)
Global $button_Run = GUICtrlCreateButton("Run", 65, 15)
Global $button_Restore = GUICtrlCreateButton("Restore", 25, 50)

Global $button_loadPRJ = GUICtrlCreateButton("Load project", 130, 10)
Global $button_savePRJ = GUICtrlCreateButton("Save project", 130, 50)

GUICtrlCreateLabel("User", 300, 5)
Global $User_input = GUICtrlCreateInput("5", 340, 5, 40)
Global $User_updown = GUICtrlCreateUpdown($User_input)
GUICtrlSetLimit($User_updown, 10, 0)

GUICtrlCreateLabel("Slot", 300, 35)
Global $Slot_input = GUICtrlCreateInput("4", 340, 35, 40)
Global $Slot_updown = GUICtrlCreateUpdown($Slot_input)
GUICtrlSetLimit($Slot_updown, 10, 0)

GUICtrlCreateLabel("Speed", 400, 5)
Global $Speed_combo = GUICtrlCreateCombo("", 450, 5, 75) ; create first item
GUICtrlSetData(-1, "10 meg|100 meg|1 gig|10 gig", "1 gig") ; add other item snd set a new default

GUICtrlCreateGroup("OS select", 600, 1, 100, 80)
Global $radio_Auto = GUICtrlCreateRadio("Auto", 610, 20, 80, 20)
Global $radio_XP_2003 = GUICtrlCreateRadio("2003/XP", 610, 40, 80, 20)
Global $radio_Vista = GUICtrlCreateRadio("Vista", 610, 60, 80, 20)
GUICtrlCreateGroup("", -99, -99, 1, 1) ;close group
GUICtrlSetState($radio_Auto, $GUI_CHECKED)

Global $button_about = GUICtrlCreateButton("About", 730, 15)
Global $button_exit = GUICtrlCreateButton("Exit", 730, 50)

Local $MyOptions = $WS_BORDER + $WS_VSCROLL + $LBS_DISABLENOSCROLL
GUISetFont(8.5, 400, -1, $font) ;"text", left, top [, width [, height [, style [, exStyle]]]]
Global $IP_view = GUICtrlCreateList("", 10, 90, 775, 95, $LBS_SORT + $MyOptions, 0)
Global $ARP_view = GUICtrlCreateList("", 10, 180, 775, 200, $LBS_SORT + $MyOptions, 0)
Global $Command_view = GUICtrlCreateList("", 10, 370, 775, 245, $MyOptions, 0)

HotKeySet("{F1}", "GetIPInformation")
HotKeySet("{F2}", "RestoreDHCP")

LoadProject("start")
;-----------------------------------------------
; Run the GUI until the dialog is closed
While 1
	Local $msg = GUIGetMsg()
	Select
		Case $msg = $button_get
			GetIPInformation()
		Case $msg = $button_Run
			RunTheBatch()
		Case $msg = $button_Restore
			RestoreDHCP()
		Case $msg = $button_loadPRJ
			LoadProject("menu")
		Case $msg = $button_savePRJ
			SaveProject()
		Case $msg = $button_about
			Debug("Written by Doug Kaynor because I wanted to!", 0x40)
		Case $msg = $button_exit Or $msg = $GUI_EVENT_CLOSE
			Debug("Exit button detected")
			ExitLoop
	EndSelect
WEnd
;-----------------------------------------------
Func RunTheBatch()
	Debug("RunTheBatch")
	If FileExists($Batch_filename) Then
		Local $val = RunWait(@ComSpec & " /c " & $Batch_filename)
		Debug("Program returned with exit code: " & $val & " (0 is success)", 0, 10)
	Else
		Debug("Batch file not found : " & $Batch_filename, 0, 0)
	EndIf
EndFunc   ;==>RunTheBatch
;-----------------------------------------------
; get the current ip information
Func GetIPInformation()
	Debug("GetIPInformation")
	ClearLists()
	GetIPInfo()
	GetARPInfo()
	MakeBatch()
	SaveInit()
EndFunc   ;==>GetIPInformation
;-----------------------------------------------
Func SaveInit()
	Debug("SaveResults")
	;Save the initial results
	Local $file = FileOpen($Initial_filename, 2)
	; Check if file opened for writing OK
	If $file = -1 Then
		Debug("Unable to open file for writing: " & $Initial_filename, 0x10)
		Return
	EndIf

	For $x = 0 To _GUICtrlListBox_GetCount($IP_view) - 1
		FileWriteLine($file, _GUICtrlListBox_GetText($IP_view, $x))
	Next

	For $x = 0 To _GUICtrlListBox_GetCount($ARP_view) - 1
		FileWriteLine($file, _GUICtrlListBox_GetText($ARP_view, $x))
	Next

	FileClose($file)

EndFunc   ;==>SaveInit

;-----------------------------------------------
; This function will get the current system IP info and then parse and display it
; The data will be displayed in the top list box
; This runs ipconfig.exe and parses the output
Func GetIPInfo()
	Debug("GetIPInfo")
	RunWait(@ComSpec & " /c  ipconfig /all > " & $IPConfig_filename, ".", @SW_HIDE)
	RemoveBlankLines($IPConfig_filename)

	Local $file = FileOpen($IPConfig_filename, 0)
	; Check if file opened for reading OK
	If $file = -1 Then
		Debug("Unable to open file for reading: " & $IPConfig_filename, 0x10)
		Return
	EndIf

	; Read in lines of text until the EOF is reached
	; Look for Ethernet adapter Local Area Connection

	Local $ConnectionName = ""
	Local $IPAddress = ""
	Local $MACAddress = "*"
	Local $DHCP = ""
	While 1
		Dim $lineIn = FileReadLine($file)
		If @error = -1 Then ExitLoop
		If StringInStr($lineIn, "Ethernet adapter Local Area Connection") Then
			$ConnectionName = StringTrimRight(StringMid($lineIn, 18), 1)
		EndIf
		If StringInStr($lineIn, "IPv4 Address") Then
			$IPAddress = StringStripWS(StringMid($lineIn, StringInStr($lineIn, ":") + 1), 3)
		EndIf
		If StringInStr($lineIn, "IP Address") Then
			$IPAddress = StringStripWS(StringMid($lineIn, StringInStr($lineIn, ":") + 1), 3)
		EndIf
		If StringInStr($lineIn, "Physical Address") Then
			$MACAddress = StringStripWS(StringMid($lineIn, StringInStr($lineIn, ":") + 1), 3)
		EndIf
		If StringInStr($lineIn, "DHCP Enabled") Then
			$DHCP = StringStripWS(StringMid($lineIn, StringInStr($lineIn, ":") + 1), 3)
		EndIf
		; Local $lineD = StringRegExpReplace($IPAddress,"[^0-9.]","")

		If (StringLen($ConnectionName) > 0) And (StringLen($IPAddress) > 0) And (StringLen($MACAddress) > 0) And (StringLen($DHCP) > 0) Then
			GUICtrlSetData($IP_view, StringFormat("%-30s %17s %20s  DHCP: %s", $ConnectionName, StringRegExpReplace($IPAddress, "[^0-9.]", ""), $MACAddress, $DHCP))
			$ConnectionName = ""
			$IPAddress = ""
			$MACAddress = "*"
			$DHCP = ""
		EndIf
	WEnd
	FileClose($file)
EndFunc   ;==>GetIPInfo
;-----------------------------------------------
; This function will get the current system arp info and then parse and display it
; The data will be displayed in the middle list box
; It calls arp.exe
Func GetARPInfo()
	Debug("GetARPInfo")
	RunWait(@ComSpec & " /c arp.exe -a > " & $ARP_filename, ".", @SW_HIDE)
	RemoveBlankLines($ARP_filename)

	Local $file = FileOpen($ARP_filename, 0)
	; Check if file opened for reading OK
	If $file = -1 Then
		Debug("Unable to open file for reading: " & $ARP_filename, 0x10)
		Return
	EndIf

	Local $x = 0
	While 1
		Dim $lineIn = FileReadLine($file)
		If @error = -1 Then ExitLoop
		If StringLen($lineIn) > 0 And Not StringInStr($lineIn, "Internet address") Then
			$x = $x + 1
			GUICtrlSetData($ARP_view, StringFormat("%3d  %s", $x, $lineIn))
		EndIf
	WEnd

	FileClose($file)
EndFunc   ;==>GetARPInfo

;----------------------------------------------- a
; This function will create a batch file to set the IPAddress and ARP values
; The data will be displayed in the bottom list box
Func MakeBatch()
	Debug("MakeBatch")
	Debug("Build netsh commands to assign the IP addresss to each connection")
	_GUICtrlListBox_ResetContent($Command_view)
	If _GUICtrlListBox_GetCount($IP_view) = 0 Then
		Debug("No data to work with", 0x40, 5)
		Return
	EndIf

	Local $AddressList[1] ;array to hold the future address strings
	For $x = 0 To _GUICtrlListBox_GetCount($IP_view) - 1
		Local $string = _GUICtrlListBox_GetText($IP_view, $x)
		Local $Name = StringStripWS(StringMid($string, 1, 30), 3)
		Local $AddressL = "172.3" & $x & "." & GUICtrlRead($User_input) & "0.1"
		Local $T = "netsh interface ip set address """ & $Name & _
				""" static " & $AddressL & _
				" mask=255.255.0.0"
		_ArrayAdd($AddressList, StringFormat("%-30s %30s", $AddressL, $Name))
		;_ArrayAdd($AddressList, $AddressL)
		GUICtrlSetData($Command_view, $T)
	Next

	Debug("AddressList")
	_ArrayDelete($AddressList, 0)
	;_ArrayDisplay($AddressList)

	Local $ArpStrings[1] ;array to hold arp strings
	;create a list of all possible arp mac address values using the user inputs
	For $x = 0 To _GUICtrlListBox_GetCount($IP_view) - 1
		Local $string = _GUICtrlListBox_GetText($IP_view, $x)
		Local $Name = StringStripWS(StringMid($string, 1, 30), 3)

		Local $PCCard = 30;
		Local $PCCardIncrement = 0;

		For $SBCard = GUICtrlRead($Slot_input) To 12 Step +2
			For $SBPort = 0 To 3
				; Third octet of the IP address
				Local $IP3 = (100 + ($SBCard * 10) + $SBPort);
				; Build our ARP string
				Local $currentArpString = StringFormat("00-01-00-00-%0.2d-%0.2d", $SBCard, $SBPort)
				_ArrayAdd($ArpStrings, $currentArpString)
				$PCCardIncrement = 1;
				If $PCCardIncrement = 3 Then
					$PCCard += 1;
					$PCCardIncrement = 0;
				EndIf
			Next
		Next
	Next

	Debug("ArpStrings")
	_ArrayDelete($ArpStrings, 0)
	;	_ArrayDisplay($ArpStrings)

	Local $FinalStrings[1]
	_ArrayDelete($FinalStrings, 0)
	; Now build the actual commands

	Local $NumberOfARPAddress = 3
	If StringCompare(GUICtrlRead($Speed_combo), "10 gig") = 0 Then $NumberOfARPAddress = 6 ;figure out how many smartbits ports to use per card

	_ArrayAdd($FinalStrings, "netsh -c ""interface ipv4"" set neighbors ""Local Area Connection 3"" 172.30.50.142 00-01-00-00-04-02")
	_ArrayAdd($FinalStrings, "-----------------------")
	_ArrayReverse($ArpStrings) ;Do this so that _arraypop will work

	If GUICtrlRead($radio_Auto) = $GUI_CHECKED Then
		Global $TempOSVersion = @OSVersion
	ElseIf GUICtrlRead($radio_Vista) = $GUI_CHECKED Then
		Global $TempOSVersion = "WIN_VISTA"
	ElseIf GUICtrlRead($radio_XP_2003) = $GUI_CHECKED Then
		Global $TempOSVersion = "WIN_XP"
	EndIf

	GUICtrlSetData($Command_view, "rem OS version: " & $TempOSVersion)

	For $x In $AddressList ;This is the list of all addresses that we need to do
		For $Y = 1 To $NumberOfARPAddress ;This is a list of all possible arp mac addresses
			If StringCompare($TempOSVersion, "WIN_XP") = 0 Or StringCompare($TempOSVersion, "WIN_2003") = 0 Then
				Debug("Final version for WIN_XP or WIN_2003")
				; arp -s 172.33.50.183 00-01-00-00-08-03
				Local $T = StringFormat("arp -s %s %s", StringStripWS(StringLeft($x, 30), 8), _ArrayPop($ArpStrings))
				_ArrayAdd($FinalStrings, $T)
				GUICtrlSetData($Command_view, $T)

			ElseIf StringCompare($TempOSVersion, "WIN_VISTA") = 0 Then
				Debug("Final version for WIN_VISTA")
				; netsh -c "interface ipv4" set neighbors "Local Area Connection 3" 172.30.50.142 00-01-00-00-04-02
				Local $C = StringStripWS(StringMid($x, 1, 20), 2)
				Local $D = StringStripWS(StringMid($x, 20), 1)
				Local $E = _ArrayPop($ArpStrings)
				Local $T = StringFormat("netsh -c ""interface ipv4"" set neighbors ""%s"" %s %s", $D, $C, $E)
				;_ArrayAdd($FinalStrings, $T)
				GUICtrlSetData($Command_view, $T)
			Else
				Debug("This OS version is not supported " & $TempOSVersion, 0, 10)
			EndIf
		Next
	Next

	GUICtrlSetData($Command_view, "arp -a")
	GUICtrlSetData($Command_view, "Pause")

	Debug("FinalStrings")
	Local $file = FileOpen($Batch_filename, 2)
	; Check if file opened for writing OK
	If $file = -1 Then
		Debug("Unable to open file for write: " & $Batch_filename, 0x40, 5)
		Return
	EndIf

	For $x = 0 To _GUICtrlListBox_GetCount($Command_view) - 1
		FileWriteLine($file, _GUICtrlListBox_GetText($Command_view, $x))
	Next

	FileClose($file)
EndFunc   ;==>MakeBatch
;-----------------------------------------------
Func SaveProject()
	Debug("SaveProject")
	$Projectfilename = FileSaveDialog("Save project file", @ScriptDir & ".\AUXFiles\", _
			"ArpTool projects (A*.prj)|All projects (*.prj)|All files (*.*)", 18, @ScriptDir & ".\AUXFiles\ArpTool.prj")
	Local $file = FileOpen($Projectfilename, 2)
	; Check if file opened for writing OK
	If $file = -1 Then
		Debug("Unable to open project file for writing: " & $Projectfilename, 0x10)
		Return
	EndIf

	FileWriteLine($file, "Valid for ArpTool")
	FileWriteLine($file, "User:" & GUICtrlRead($User_input))
	FileWriteLine($file, "Slot:" & GUICtrlRead($Slot_input))
	FileWriteLine($file, "Speed:" & GUICtrlRead($Speed_combo))

	FileClose($file)
EndFunc   ;==>SaveProject
;-----------------------------------------------
Func LoadProject($type)
	Debug("LoadProject")
	If StringCompare($type, "menu") = 0 Then
		$Projectfilename = FileOpenDialog("Load project file", @ScriptDir & ".\AUXFiles\", _
				"ArpTool projects (A*.prj)|All projects (*.prj)|All files (*.*)", 18, @ScriptDir & ".\AUXFiles\ArpTool.prj")
	EndIf

	Local $file = FileOpen($Projectfilename, 0)
	; Check if file opened for reading OK
	If $file = -1 Then
		Debug("Unable to open project file for reading: " & $Projectfilename, 0x1, 5)
		Return
	EndIf

	; Read in the first line to verify the file is of the correct type
	If StringCompare(FileReadLine($file, 1), "Valid for ArpTool") <> 0 Then
		Debug("Not a valid project file for ArpTool", 0x20, 10)
		FileClose($file)
		Return
	EndIf

	; Read in lines of text until the EOF is reached
	While 1
		Local $lineIn = FileReadLine($file)
		If @error = -1 Then ExitLoop
		If StringInStr($lineIn, "User:") Then GUICtrlSetData($User_input, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
		If StringInStr($lineIn, "Slot:") Then GUICtrlSetData($Slot_input, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
		If StringInStr($lineIn, "Speed:") Then GUICtrlSetData($Speed_combo, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
	WEnd

	FileClose($file)
EndFunc   ;==>LoadProject
;-----------------------------------------------
Func RestoreDHCP()
	Debug("RestoreDHCP")
	GetIPInformation() ; get the current ip information
	For $x = 0 To _GUICtrlListBox_GetCount($IP_view) - 1
		Local $T = StringStripWS(StringLeft(_GUICtrlListBox_GetText($IP_view, $x), 30), 3)
		Debug($T)
		Local $U = "interface ip set address name=""" & $T & """ source=dhcp"
		Debug($U)
		ShellExecuteWait("netsh", $U)
	Next
	; Netsh interface ip set address name=”Local Area Connection” source=dhcp

EndFunc   ;==>RestoreDHCP
;-----------------------------------------------
Func ClearLists()
	Debug("ClearLists")
	_GUICtrlListBox_ResetContent($IP_view)
	_GUICtrlListBox_ResetContent($ARP_view)
	_GUICtrlListBox_ResetContent($Command_view)
EndFunc   ;==>ClearLists
;-----------------------------------------------
