#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Run_Tidy=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

Opt("MustDeclareVars", 1) ; 0=no, 1=require pre-declare
#RequireAdmin

Const $Version = "version: 0.0.0.1"

If _Singleton(@ScriptName, 1) = 0 Then
	Debug(@ScriptName & " is already running!", 0x40)
	Exit
EndIf

#include <GuiListBox.au3>
#include <GUIConstants.au3>
#include <String.au3>
#include <Misc.au3>
#include <_DougFunctions.au3>

Local $tmp = StringSplit(@ScriptName, ".")
Global $Project_filename = @ScriptDir & "\AUXFiles\" & $tmp[1] & ".prj"
Dim $EXE[1]
Global $SelectedLine = ''

;"text", left, top , width , height
Local $font = "Courier new"
GUISetFont(10, 400, -1, $font)
Local $SystemS = @OSVersion & "  " & @OSServicePack & "  " & @OSTYPE & "  " & @IPAddress1
Global $Main = GUICreate(@ScriptName & "  " & $Version & "  " & $SystemS, 530, 300, 20, 20)
GUISetState()

Global $button_launch = GUICtrlCreateButton("Launch program", 10, 10)
Global $button_editProject = GUICtrlCreateButton("Edit Project", 200, 10)
Global $button_loadProject = GUICtrlCreateButton("Load Project", 280, 10)
Global $button_about = GUICtrlCreateButton("About", 430, 10)
Global $button_exit = GUICtrlCreateButton("Exit", 480, 10)
Global $button_showlaunch = GUICtrlCreateButton("show", 380, 10)

Global $User_list1 = GUICtrlCreateList("", 10, 50, 500, 200, 0x00B01003)
;-----------------------------------------------
LoadProject("Start")
; Run the GUI until the dialog is closed
While 1
	Local $msg = GUIGetMsg()
	Switch $msg
		Case $button_launch
			Launch()
		Case $button_editProject
			EditProject()
		Case $button_loadProject
			LoadProject("menu")
		Case $button_showlaunch
			_ArrayDisplay($EXE)
		Case $GUI_EVENT_PRIMARYDOWN
			Local $A[1]
			$A = GUIGetCursorInfo()
			If $A[4] = $User_list1 Then ClickOnUser1()
		Case $button_about
			Debug(@CRLF & @ScriptName & "  " & $Version & @CRLF & "Written by Doug Kaynor because I wanted to!", 0x40, 5)
		Case $GUI_EVENT_CLOSE
			ExitLoop
		Case $button_exit
			ExitLoop
	EndSwitch
WEnd
;-----------------------------------------------
Func LoadProject($type)
	Debug("LoadProject  " & $type & "  " & @ScriptLineNumber)

	_GUICtrlListBox_ResetContent($User_list1)
	Dim $EXE[1]

	If StringCompare($type, "menu") = 0 Then
		$Project_filename = FileOpenDialog("Load project file", @ScriptDir & "\AUXFiles\", _
				"GLaunch projects (G*.prj)|All projects (*.prj)|All files (*.*)", 18, @ScriptDir & "\AUXFiles\GLaunch.prj")
	EndIf

	Local $file = FileOpen($Project_filename, 0)
	; Check if file opened for reading OK
	If $file = -1 Then
		Debug("LoadProject: Unable to open file for reading: " & $Project_filename, 0x10, 5)
		Return
	EndIf

	Debug("LoadProject 2   " & $Project_filename & "    " & $type)
	; Read in the first line to verify the file is of the correct type
	Local $LineIn = FileReadLine($file)
	If @error = -1 Then Return
	If StringInStr($LineIn, "Valid for CLaunch project") <> 1 Then
		Debug("Not a valid project file for CLaunch", 0x20, 5)
		FileClose($file)
		Return
	EndIf

	; Read in lines of text until the EOF is reached
	While 1
		Local $LineIn = FileReadLine($file)
		If @error = -1 Then ExitLoop
		$LineIn = StringStripWS($LineIn, 3)
		If StringLen($LineIn) > 5 And (StringInStr($LineIn, ';') <> 1) Then
			If StringInStr($LineIn, '_exe') > 0 Then
				_ArrayAdd($EXE, $LineIn)
			Else
				GUICtrlSetData($User_list1, $LineIn)
			EndIf
		EndIf
	WEnd
	FileClose($file)
	_ArrayDelete($EXE, 0);Get rid of the blank location
	TestProjectData()
EndFunc   ;==>LoadProject
;-----------------------------------------------
Func TestProjectData()
	Debug("TestProjectData " & @ScriptLineNumber)

	For $A = 0 To UBound($EXE) - 1 ;first test exe array
		Local $B = StringSplit($EXE[$A], '=')
		Local $C = StringStripWS($B[2], 3)
		$C = StringRegExpReplace($C, '\"', "")
		debug($C & "  " & @ScriptLineNumber)
		If FileExists($C) = 0 Then debug($C & " does not exists", 1, 10)
	Next

	For $A = 0 To _GUICtrlListBox_GetCount($User_list1) - 1
		Local $B = StringSplit(_GUICtrlListBox_GetText($User_list1, $A), '=')
		If StringInStr($B[1], "_script") > 0 Then
			Local $C = StringStripWS($B[2], 3)
			$C = StringRegExpReplace($C, '\"', "")
			debug($C & "  " & @ScriptLineNumber)
			If FileExists($C) = 0 Then debug($C & " does not exists", 1, 10)
		EndIf
	Next
EndFunc   ;==>TestProjectData
;-----------------------------------------------
Func EditProject()
	Debug("EditProject " & @ScriptLineNumber)
	RunWait("notepad.exe " & $Project_filename)
	LoadProject("Start")
EndFunc   ;==>EditProject
;-----------------------------------------------
;This lists the TIC files to $user_list2 when a TIC name is selected in $user_list1
Func ClickOnUser1()
	Debug("ClickOnUser1  " & @ScriptLineNumber)
	$SelectedLine = _GUICtrlListBox_GetText($User_list1, _GUICtrlListBox_GetCaretIndex($User_list1))
	debug($SelectedLine)
EndFunc   ;==>ClickOnUser1
;-----------------------------------------------
Func Launch()
	Debug("Launch  " & $SelectedLine & "   " & @ScriptLineNumber) ; perl_script = C:\Perl\Projects\music\TkPlaylist.pl
	Local $A = StringSplit($SelectedLine, '=') ; perl_exe = c:\perl\bin\perl.exe
	Local $B = _ArraySearch($EXE, StringStripWS($A[1], 3), 0, 0, 0, True)
	Local $C = StringSplit($EXE[$B], '=')
	Local $EXEProgram = StringStripWS($C[2], 3)
	Local $RunFile = StringStripWS($A[2], 3)
	debug($EXEProgram & "   " & $RunFile & "   " & @ScriptLineNumber)
	ShellExecute($EXEProgram, $RunFile)
EndFunc   ;==>Launch
;-----------------------------------------------