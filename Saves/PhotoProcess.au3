#Region
#AutoIt3Wrapper_Run_Au3check=y
#AutoIt3Wrapper_Au3Check_Stop_OnWarning=y
#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Tidy_Stop_OnError=y
;#Tidy_Parameters=/gd /sf
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Fileversion=1.0.0.28
#AutoIt3Wrapper_Res_FileVersion_AutoIncrement=Y
#AutoIt3Wrapper_Res_Description=PhotoProcess
#AutoIt3Wrapper_Res_LegalCopyright=GNU-PL
#AutoIt3Wrapper_Res_Comment=A program to process photos
#AutoIt3Wrapper_Res_Field=Developer|Douglas Kaynor
#AutoIt3Wrapper_Res_LegalCopyright=Copyright ? 2009 Douglas B Kaynor
#AutoIt3Wrapper_Res_Field= AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Res_Field=Compile date|%longdate% %time%
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Run_Before=
#AutoIt3Wrapper_Icon=eagle.ico
#EndRegion

Global $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")

Opt("MustDeclareVars", 1)

If _Singleton(@ScriptName, 1) = 0 Then
	Debug(@ScriptName & " is already running!", 0x40, 5)
	Exit
EndIf

#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <ListViewConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <TreeViewConstants.au3>

#include <GuiTreeView.au3>
#include <Date.au3>
#include <Misc.au3>
#include <String.au3>
#include "_DougFunctions.au3"

#Region ### START Koda GUI section ### Form=
Global $MainForm = GUICreate(@ScriptName & $FileVersion, 500, 350, 545, 280, BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS))
GUISetFont(10, 400, 0, "Courier New")

Global $ButtonChoseFolder = GUICtrlCreateButton("Folder", 16, 10, 60, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Chose a working folder")

Global $ButtonShow = GUICtrlCreateButton("Show", 16, 40, 60, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Show the selected selected picture")

Global $ButtonProc = GUICtrlCreateButton("Proc", 16, 70, 60, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Proceess all checked pictures")

Global $ButtonToggle = GUICtrlCreateButton("Toggle", 100, 10, 60, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Toggle all check boxes")

Global $ButtonExit = GUICtrlCreateButton("Exit", 100, 40, 60, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Exit the program")

;Global $hTreeView = GUICtrlCreateTreeView(16, 104, 201, 201, BitOR($TVS_HASBUTTONS, $TVS_HASLINES, $TVS_LINESATROOT, $TVS_DISABLEDRAGDROP, $TVS_SHOWSELALWAYS, $TVS_CHECKBOXES, $WS_GROUP, $WS_TABSTOP), $WS_EX_CLIENTEDGE)
;GUICtrlSetTip(-1, "List of files")

Global $hTreeView = GUICtrlCreateTreeView(16, 104, 201, 201, BitOR($TVS_HASBUTTONS, $TVS_HASLINES, $TVS_DISABLEDRAGDROP, $TVS_SHOWSELALWAYS, $WS_GROUP, $WS_TABSTOP), $WS_EX_CLIENTEDGE)
GUICtrlSetTip(-1, "List of files")

Global $Picture = GUICtrlCreatePic("-10-.jpg", 224, 8, 250, 250, BitOR($SS_NOTIFY, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS))
GUICtrlSetTip(-1, "This  shows the selected picture")

Global $InputStatus = GUICtrlCreateInput("Input  Status", 16, 312, 460, 25, 0x081)

GUICtrlSetTip(-1, "Status What's happening")

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

Global $Infraview
Global $DataArray[1]
Global $HandleArray[1]
Global $ToggleState = False
Global $WorkingFolder = @ScriptDir

TestForInfraView()
;-----------------------------------------------
While 1
	Global $nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $ButtonExit
			Exit
		Case $ButtonChoseFolder
			ChoseFolder()
		Case $ButtonShow
			Show()
		Case $ButtonToggle
			ToggleData()
		Case $ButtonProc
			Proceess()
	EndSwitch
WEnd
;-----------------------------------------------
Func ToggleData()
	Debug("ToggleData")

	;_ArrayDisplay($HandleArray)
	$ToggleState = Not $ToggleState

	For $X = 1 To UBound($HandleArray) - 1
		_GUICtrlTreeView_SetChecked($hTreeView, $HandleArray[$X], $ToggleState)
		;Debug(StringFormat("%2u   %6x   %d", $X, $HandleArray[$X], $ToggleState))
	Next

EndFunc   ;==>ToggleData

;-----------------------------------------------
Func Proceess()
	Debug("Prccess")
	Local $Tmp1 = _GUICtrlTreeView_GetSelection($hTreeView)
	Local $tmp2 = _GUICtrlTreeView_GetText($hTreeView, $Tmp1)
	Local $ShowName = $WorkingFolder & "\" & $tmp2
	debug(">>> " & $Tmp1 & "   " & $ShowName)



EndFunc   ;==>Proceess
;-----------------------------------------------
Func Show()
	Debug("Show")

	GUICtrlSetImage($Picture, @ScriptDir & "\one.gif")
	;_ArrayDisplay($DataArray, "DataArray")
	;_ArrayDisplay($HandleArray, "HandleArray")

	FileDelete("temp.gif")

	Local $Tmp1 = _GUICtrlTreeView_GetSelection($hTreeView)
	Local $tmp2 = _GUICtrlTreeView_GetText($hTreeView, $Tmp1)
	Local $ShowName = $WorkingFolder & "\" & $tmp2
	debug(">>> " & $Tmp1 & "   " & $ShowName)

	Local $ConvertCmd = $Infraview & "  " & $ShowName & " /resize=(250,250) /resample /aspectratio /convert=" & @ScriptDir & "\temp.gif";
	debug(RunWait($ConvertCmd))
	GUICtrlSetImage($Picture, @ScriptDir & "\temp.gif")

EndFunc   ;==>Show

;-----------------------------------------------
;"C:\Documents and Settings\dbkaynox\My Documents\My Pictures\"
; "d:\my pictures"
Func ChoseFolder()
	Debug("ChoseFolder")

	;Local $startinFolder = "C:\Documents and Settings\dbkaynox\My Documents\My Pictures\plants\"
	Local $StartingFolder = "D:\My Pictures\signs\"

	$WorkingFolder = FileSelectFolder("Chose a folder", "", 7, $StartingFolder)
	_GUICtrlTreeView_DeleteAll($hTreeView)

	FileChangeDir($WorkingFolder)
	GUICtrlSetData($InputStatus, $WorkingFolder)
	ReDim $HandleArray[1]

	Local $search = FileFindFirstFile("*.jpg")

	; Check if the search was successful
	If $search = -1 Then
		debug("No files/directories matched the search pattern   *.jpg")
		Return
	EndIf

	While 1
		Local $file = FileFindNextFile($search)
		If @error Then ExitLoop
		If StringInStr($file, ".jpg") Then _ArrayAdd($HandleArray, _GUICtrlTreeView_Add($hTreeView, 0, $file))
	WEnd

	FileClose($search) ; Close the search handle
	;_ArrayDisplay($HandleArray)
	FileChangeDir(@ScriptDir)

	Debug("ChoseFolder   " & $WorkingFolder)
EndFunc   ;==>ChoseFolder
;-----------------------------------------------
;This function gets all of the data from the treeview into an array
Func TreeViewToArray()
	Debug("TreeViewToArray")
	ReDim $DataArray[1]
	For $X = 1 To UBound($HandleArray) - 1
		_ArrayAdd($DataArray, _GUICtrlTreeView_GetText($hTreeView, $HandleArray[$X]))
	Next
	;_ArrayDisplay($HandleArray)
EndFunc   ;==>TreeViewToArray
;-----------------------------------------------
Func TestForInfraView()

	$Infraview = "C:\Program Files (x86)\IrfanView\i_view32.exe"
	If FileExists($Infraview) Then Return

	$Infraview = "C:\Program Files\IrfanView\i_view32.exe"
	If FileExists($Infraview) Then Return

	Debug("Unable to locate Infraview on this computer");
EndFunc   ;==>TestForInfraView
;-----------------------------------------------