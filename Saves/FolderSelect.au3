#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <GUIListBox.au3>
#include <GuiTreeView.au3>
#include <StaticConstants.au3>
#include <TreeViewConstants.au3>
#include <WindowsConstants.au3>
#include <Misc.au3>
#include "_DougFunctions.au3"

Global $WorkingFolder = EnvGet("USERPROFILE")
;Global $WorkingFolder = "C:\Program Files\AutoIt3\Dougs"
Global $HandleArray[1]
Global $ToggleState = False

Global $SelectFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $SelectForm = GUICreate("Select form", 450, 600, 10, 10, $SelectFormOptions)
GUISetFont(10, 400, 0, "Courier New")

Global $ButtonSelect = GUICtrlCreateButton("Select top folder", 15, 15, 140, 20)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKWIDTH + $GUI_DOCKTOP + $GUI_DOCKHEIGHT)
Global $InputSelectFolder = GUICtrlCreateInput($WorkingFolder, 160, 15, 250, 24)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKHEIGHT)

Global $ButtonGetFolders = GUICtrlCreateButton("Get folders", 15, 40, 120, 20)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKWIDTH + $GUI_DOCKTOP + $GUI_DOCKHEIGHT)

Global $ButtonProcess = GUICtrlCreateButton("Process", 150, 40, 120, 20)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKWIDTH + $GUI_DOCKTOP + $GUI_DOCKHEIGHT)

Global $ButtonToggle = GUICtrlCreateButton("Toggle", 285, 40, 120, 20)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKWIDTH + $GUI_DOCKTOP + $GUI_DOCKHEIGHT)

Global $LabelFolderList = GUICtrlCreateLabel("Folder list", 15, 75, 90, 20)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKHEIGHT)
Global $TreeViewFolders = GUICtrlCreateTreeView(10, 100, 400, 200, BitOR($TVS_HASBUTTONS, $TVS_HASLINES, $TVS_LINESATROOT, $TVS_DISABLEDRAGDROP, $TVS_SHOWSELALWAYS, $TVS_CHECKBOXES, $WS_GROUP, $WS_TABSTOP), $WS_EX_CLIENTEDGE)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKWIDTH + $GUI_DOCKTOP) ;+ $GUI_DOCKHEIGHT + $GUI_DOCKBOTTOM)

Global $LabelResults = GUICtrlCreateLabel("Results", 15, 300, 60, 20)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
Global $ListResults = GUICtrlCreateList("", 10, 320, 400, 246)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM)
GUICtrlSetData(-1, "ListResults")

GUISetState(@SW_SHOW)


While 1
	Global $nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $ButtonSelect
			ChoseFolders()
		Case $ButtonGetFolders
			$WorkingFolder = GUICtrlRead($InputSelectFolder)
			GetFolders()
		Case $ButtonProcess
			Process()
		Case $ButtonToggle
			ToggleChecked()
		Case $InputSelectFolder
			ConsoleWrite(@ScriptLineNumber & " InputSelectFolder" & @CRLF)
	EndSwitch
WEnd

;-----------------------------------------------
Func ChoseFolders()
	Local $tmp = FileSelectFolder("Chose a folder", "", 7, $WorkingFolder)
	If $tmp = "" Then Return
	$WorkingFolder = $tmp
	GUICtrlSetData($InputSelectFolder, $tmp)
EndFunc   ;==>ChoseFolders
;-----------------------------------------------
Func GetFolders()
	_GUICtrlTreeView_DeleteAll($TreeViewFolders)
	_GUICtrlListBox_ResetContent($ListResults)
	ReDim $HandleArray[1]

	Local $FolderArray = _FileListToArrayR($WorkingFolder, "*", 2, 1, 0, "", 1)
	_ArrayDelete($FolderArray, 0)
	;_ArrayDisplay($FolderArray, @ScriptLineNumber)
	_ArrayUnique($FolderArray)
	_ArraySort($FolderArray)
	;_ArrayDisplay($FolderArray, @ScriptLineNumber)

	If IsArray($FolderArray) Then
		For $X In $FolderArray
			;_GUICtrlTreeView_Add($TreeViewFolders, 0, $X)
			_ArrayAdd($HandleArray, _GUICtrlTreeView_Add($TreeViewFolders, 0, $X))
		Next
		_ArrayDelete($HandleArray, 0)
	Else
		MsgBox(16, "Path not found", "Path not found:" & @CRLF & $WorkingFolder)
	EndIf
EndFunc   ;==>GetFolders
;-----------------------------------------------
Func ToggleChecked()
	;_ArrayDisplay($HandleArray)
	$ToggleState = Not $ToggleState
	For $X = 0 To UBound($HandleArray) - 1
		_GUICtrlTreeView_SetChecked($TreeViewFolders, $HandleArray[$X], $ToggleState)
	Next
EndFunc   ;==>ToggleChecked
;-----------------------------------------------\
Func Process()
	Local $ArrayTmp[1]
	;_ArrayDisplay($HandleArray, @ScriptLineNumber)
	_GUICtrlListBox_ResetContent($ListResults)
	For $X = 0 To UBound($HandleArray) - 1
		Local $AA = _GUICtrlTreeView_GetChecked($TreeViewFolders, $HandleArray[$X])
		Local $BB = _GUICtrlTreeView_GetText($TreeViewFolders, $HandleArray[$X])
		If $AA Then _ArrayAdd($ArrayTmp, $BB)
	Next

	_ArrayDelete($ArrayTmp, 0)
	CleanPaths($ArrayTmp)

	For $X In $ArrayTmp
		_GUICtrlListBox_AddString($ListResults, $X)
	Next

EndFunc   ;==>Process
;-----------------------------------------------
Func CleanPaths(ByRef $Array)
	For $X = 0 To UBound($Array) - 1
		For $Y = 0 To UBound($Array) - 1
			If StringInStr($Array[$Y], $Array[$X]) <> 0 And _
					StringCompare($Array[$Y], $Array[$X]) <> 0 Then
				$Array[$Y] = ""
			EndIf
		Next
	Next
	_RemoveBlankLines($Array)
EndFunc   ;==>CleanPaths
;-----------------------------------------------