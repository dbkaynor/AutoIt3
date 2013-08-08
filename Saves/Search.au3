#Region
#AutoIt3Wrapper_Run_Au3check=y
#AutoIt3Wrapper_Au3Check_Stop_OnWarning=y
#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Tidy_Stop_OnError=y
;#Tidy_Parameters=/gd /sf
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Fileversion=1.0.0.2
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
#AutoIt3Wrapper_Icon=./icons/dbk.ico
#EndRegion

Opt("MustDeclareVars", 1)

If _Singleton(@ScriptName, 1) = 0 Then
	Debug(@ScriptName & " is already running!", 0x40, 5)
	Exit
EndIf

#include <Array.au3>
#include <Date.au3>
#include <iNet.au3>
#include <Misc.au3>
#include <String.au3>

#include <GUIConstants.au3>
#include <GuiListView.au3>
#include <GUIConstantsEx.au3>

#include <GuiTreeView.au3>
#include <GuiImageList.au3>

#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <ListViewConstants.au3>
#include <StaticConstants.au3>
#include <TreeViewConstants.au3>
#include <WindowsConstants.au3>
#include <FileListToArray3.au3>
#include "_DougFunctions.au3"

Global $RawDataArray[1]
Global $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")

Global $Tmp = StringSplit(@ScriptName, ".")
Global $Project_filename = @ScriptDir & "\AUXFiles\" & $Tmp[1] & ".prj"
Global $LOG_filename = @ScriptDir & "\AUXFiles\" & $Tmp[1] & ".log"
Global $WorkingFolder = @ScriptDir
Global $hItem[1]

Global $SystemS = @ScriptName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & @OSServicePack & @CRLF & @OSType & @CRLF & @OSArch

GUISetFont(10, 400, 0, "Courier New")

Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $MainForm = GUICreate(@ScriptName & " " & $FileVersion, 620, 420, 10, 10, $MainFormOptions)
Global $FileMenuItemRaw = GUICtrlCreateMenu("&Raw")
Global $MenuSaveChoseFolder = GUICtrlCreateMenuItem("Chose Folder", $FileMenuItemRaw)
Global $MenuGetRaw = GUICtrlCreateMenuItem("Get Raw", $FileMenuItemRaw)
Global $MenuSaveRaw = GUICtrlCreateMenuItem("Save raw", $FileMenuItemRaw)
Global $MenuLoadRaw = GUICtrlCreateMenuItem("Load raw", $FileMenuItemRaw)
Global $MenuViewRaw = GUICtrlCreateMenuItem("View raw", $FileMenuItemRaw)
Global $FileMenuItemResult = GUICtrlCreateMenu("R&esult")
Global $MenuSaveResults = GUICtrlCreateMenuItem("Save results", $FileMenuItemResult)
Global $MenuLoadResults = GUICtrlCreateMenuItem("Load results", $FileMenuItemResult)
Global $MenuViewResults = GUICtrlCreateMenuItem("View results", $FileMenuItemResult)
Global $FileMenuItemHelp = GUICtrlCreateMenu("&Help")
Global $MenuHelp = GUICtrlCreateMenuItem("Help", $FileMenuItemHelp)
Global $MenuAbout = GUICtrlCreateMenuItem("About", $FileMenuItemHelp)
Global $MenuExit = GUICtrlCreateMenuItem("Exit", $FileMenuItemHelp)
;------------------
Global $DBKResize = 802
Global $ButtonSearch = GUICtrlCreateButton("Search", 10, 20, 50, 20, $WS_GROUP)
GUICtrlSetResizing(-1, $DBKResize)
GUICtrlSetTip(-1, "Start searching")
Global $ButtonChoseFolder = GUICtrlCreateButton("Folder", 70, 20, 50, 20, $WS_GROUP)
GUICtrlSetResizing(-1, $DBKResize)
GUICtrlSetTip(-1, "Chose a Top level folder To begin serching In")

Global $ButtonSize = GUICtrlCreateButton("Size", 130, 20, 50, 20, $WS_GROUP)
GUICtrlSetResizing(-1, $DBKResize)
GUICtrlSetTip(-1, "Change the size")

Global $ButtonExit = GUICtrlCreateButton("Exit", 310, 20, 50, 20, $WS_GROUP)
GUICtrlSetResizing(-1, $DBKResize)
GUICtrlSetTip(-1, "Exit the program")

#include <DateTimeConstants.au3>
Global $style = "yyyy/MM/dd"
Global $DTM_FORMAT = 0x1032
GUICtrlCreateGroup("Date", 5, 50, 125, 200)
GUICtrlSetResizing(-1, $DBKResize)

Global $InputDate1 = GUICtrlCreateDate("1951/01/29", 10, 75, 105, 24)
GUICtrlSetResizing(-1, $DBKResize)
GUICtrlSendMsg($InputDate1, $DTM_FORMAT, 0, $style)

Global $InputDate2 = GUICtrlCreateDate("1951/01/29", 10, 100, 105, 24)
GUICtrlSetResizing(-1, $DBKResize)
GUICtrlSendMsg($InputDate2, $DTM_FORMAT, 0, $style)

Global $RadioDate1 = GUICtrlCreateRadio("Older", 10, 125, 110, 17)
GUICtrlSetResizing(-1, $DBKResize)
Global $RadioDate2 = GUICtrlCreateRadio("Between", 10, 150, 110, 17)
GUICtrlSetResizing(-1, $DBKResize)
Global $RadioDate3 = GUICtrlCreateRadio("Newer", 10, 175, 110, 17)
GUICtrlSetResizing(-1, $DBKResize)
Global $RadioDate4 = GUICtrlCreateRadio("Exact", 10, 200, 110, 17)
GUICtrlSetResizing(-1, $DBKResize)
Global $RadioDate5 = GUICtrlCreateRadio("None", 10, 225, 110, 17)
GUICtrlSetResizing(-1, $DBKResize)
GUICtrlSetState($RadioDate5, $GUI_CHECKED)
GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateGroup("Size", 135, 50, 125, 200)
GUICtrlSetResizing(-1, $DBKResize)
Global $InputSize1 = GUICtrlCreateInput("0", 140, 75, 105, 24, $ES_NUMBER)
GUICtrlSetResizing(-1, $DBKResize)
Global $InputSize2 = GUICtrlCreateInput("0", 140, 100, 105, 24, $ES_NUMBER)
GUICtrlSetResizing(-1, $DBKResize)
Global $RadioSize1 = GUICtrlCreateRadio("Smaller", 140, 125, 110, 17)
GUICtrlSetResizing(-1, $DBKResize)
Global $RadioSize2 = GUICtrlCreateRadio("Between", 140, 150, 110, 17)
GUICtrlSetResizing(-1, $DBKResize)
Global $RadioSize3 = GUICtrlCreateRadio("Bigger", 140, 175, 110, 17)
GUICtrlSetResizing(-1, $DBKResize)
Global $RadioSize4 = GUICtrlCreateRadio("Exact", 140, 200, 110, 17)
GUICtrlSetResizing(-1, $DBKResize)
Global $RadioSize5 = GUICtrlCreateRadio("None", 140, 225, 110, 17)
GUICtrlSetResizing(-1, $DBKResize)
GUICtrlSetState($RadioSize5, $GUI_CHECKED)
GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateGroup("Attributes", 265, 50, 100, 110)
GUICtrlSetResizing(-1, $DBKResize)
Global $CheckboxAttr1 = GUICtrlCreateCheckbox("System", 275, 72, 89, 17)
GUICtrlSetResizing(-1, $DBKResize)
Global $CheckboxAttr2 = GUICtrlCreateCheckbox("Hidden", 275, 92, 89, 17)
GUICtrlSetResizing(-1, $DBKResize)
Global $CheckboxAttr3 = GUICtrlCreateCheckbox("Archive", 275, 112, 89, 17)
GUICtrlSetResizing(-1, $DBKResize)
Global $CheckboxAttr4 = GUICtrlCreateCheckbox("Read only", 275, 132, 89, 17)
GUICtrlSetResizing(-1, $DBKResize)
GUICtrlCreateGroup("", -99, -99, 1, 1)

Global $TreeViewFileList = GUICtrlCreateTreeView(380, 10, 230, 180, BitOR($TVS_HASBUTTONS, $TVS_HASLINES, $TVS_LINESATROOT, $TVS_DISABLEDRAGDROP, $TVS_SHOWSELALWAYS, $WS_GROUP, $WS_TABSTOP, $WS_BORDER, $WS_CLIPSIBLINGS))
GUICtrlSetResizing(-1, 102)
GUICtrlSetTip(-1, "A list of all files")
Global $TreeViewResults = GUICtrlCreateTreeView(380, 200, 230, 180, BitOR($TVS_HASBUTTONS, $TVS_HASLINES, $TVS_LINESATROOT, $TVS_DISABLEDRAGDROP, $TVS_SHOWSELALWAYS, $WS_GROUP, $WS_TABSTOP, $WS_BORDER, $WS_CLIPSIBLINGS))
GUICtrlSetResizing(-1, 102)
GUICtrlSetTip(-1, "A list of files that match")

Global $LabelNameFilter = GUICtrlCreateLabel("Name filter", 16, 260, 92, 20)
GUICtrlSetResizing(-1, $DBKResize)
Global $InputNameFilter = GUICtrlCreateInput("", 16, 285, 97, 24)
GUICtrlSetResizing(-1, $DBKResize)
Global $LabelExtFilter = GUICtrlCreateLabel("Ext filter", 128, 260, 84, 20)
GUICtrlSetResizing(-1, $DBKResize)
Global $InputExtFilter = GUICtrlCreateInput("", 128, 285, 97, 24)
GUICtrlSetResizing(-1, $DBKResize)

; GUICtrlCreateGroup ( "text", left, top [, width [, height [, style [, exStyle]]]] )
GUICtrlCreateGroup("Options", 265, 210, 100, 100)
GUICtrlSetResizing(-1, $DBKResize)
Global $CheckboxCase = GUICtrlCreateCheckbox("Case", 275, 250, 90, 17)
GUICtrlSetResizing(-1, $DBKResize)
Global $RadioOption1 = GUICtrlCreateRadio("And", 275, 270, 90, 17)
GUICtrlSetResizing(-1, $DBKResize)
Global $RadioOption2 = GUICtrlCreateRadio("Or", 275, 290, 90, 17)
GUICtrlSetResizing(-1, $DBKResize)


;~ Global $CheckboxAnd = GUICtrlCreateCheckbox("And", 275, 280, 89, 17)
;~ GUICtrlSetResizing(-1, $DBKResize)
;~ Global $CheckboxOr = GUICtrlCreateCheckbox("Or", 275, 300, 89, 17)
;~ GUICtrlSetResizing(-1, $DBKResize)
;~ GUICtrlCreateGroup("", -99, -99, 1, 1)

Global $InputStatus = GUICtrlCreateInput("Status", 16, 345, 350, 24)
GUICtrlSetResizing(-1, 770)
ChangeSize();

GUISetState(@SW_SHOW)

While 1
	Global $nMsg = GUIGetMsg()
	Switch $nMsg
		Case $ButtonChoseFolder
			ChoseFolder()
		Case $MenuSaveChoseFolder
			ChoseFolder()
		Case $ButtonSearch
			StartSearching()
		Case $ButtonSize
			ChangeSize()
		Case $ButtonExit
			Exit
		Case $MenuExit
			Exit
		Case $GUI_EVENT_CLOSE
			Exit
		Case $MenuHelp
			Help()
		Case $MenuAbout
			About()
	EndSwitch
WEnd
;-----------------------------------------------
Func StartSearching()
	_GUICtrlTreeView_DeleteAll($TreeViewResults)
	;_GUICtrlTreeView_Add($TreeViewResults, 0, "I should be starting")
	; $InputNameFilter
	; $InputExtFilter
	; $InputDate1
	; $InputDate2
	; $InputSize1
	; $InputSize2

	; $RadioDate1-5
	; $RadioSize1-5
	; $CheckboxAttr1-4

	For $X = 1 To _GUICtrlTreeView_GetCount($TreeViewFileList)
		Local $A = _GUICtrlTreeView_GetText($TreeViewFileList, $hItem[$X])
		Local $B = StringSplit($A, "\", 0)
		Local $C = $B[UBound($B) - 1]
		Local $D = StringSplit($C, ".", 0)

		Local $R = 0
		;test for name
		$R += StringInStr($D[1], GUICtrlRead($InputNameFilter))
		;test for extension
		$R += StringInStr($D[2], GUICtrlRead($InputExtFilter))

		If $R <> 0 Then
			_GUICtrlTreeView_Add($TreeViewResults, 0, StringFormat("%s  1:%s   2:%s", $A, $D[1], $D[2]))
		EndIf

	Next

EndFunc   ;==>StartSearching
;-----------------------------------------------
Func ChoseFolder()
	Debug("ChoseFolder")

	Local $startingFolder = ''
	If FileExists("C:\Documents and Settings\dbkaynox\My Documents\My Pictures\") Then
		$startingFolder = "C:\Documents and Settings\dbkaynox\My Documents\My Pictures\"
	ElseIf FileExists("D:\My Pictures\signs\") Then
		$startingFolder = "D:\My Pictures\signs\"
	Else
		$startingFolder = "C:\"
	EndIf

	GUICtrlSetData($InputStatus, "")
	_GUICtrlTreeView_DeleteAll($TreeViewFileList)
	_GUICtrlTreeView_DeleteAll($TreeViewResults)
	$WorkingFolder = FileSelectFolder("Chose a folder", "", 7, $startingFolder)
	GUICtrlSetData($InputStatus, $startingFolder)

	$RawDataArray = _FileListToArray3($WorkingFolder, "*", 0, 1, 1, "", 0)
	GUICtrlSetData($InputStatus, StringFormat("%6d %s", $RawDataArray[0], $WorkingFolder))

	_ArrayDelete($RawDataArray, 0) ; remove the count value
	Global $hItem[1]
	For $X In $RawDataArray
		_ArrayAdd($hItem, _GUICtrlTreeView_Add($TreeViewFileList, 0, $X))
	Next
	;_ArrayDisplay($hItem, "$hItem")

	Debug("ChoseFolder   " & $WorkingFolder)
EndFunc   ;==>ChoseFolder
;-----------------------------------------------
Func ChangeSize()
	Global $TS
	If $TS = False Then
		WinMove("[active]", "", Default, Default, 620, 430) ; 620, 410,
		$TS = True
	Else
		WinMove("[active]", "", Default, Default, 1000, 430)
		$TS = False
	EndIf
EndFunc   ;==>ChangeSize
;-----------------------------------------------
Func Help()
	MsgBox(0, "Help informarion", "Someday this will be help info", 10)
EndFunc   ;==>Help
;-----------------------------------------------
;This function display About info and some Debug stuff
Func About(Const $FormID)
	Local $D = WinGetPos($FormID)
	Local $WinPos
	If IsArray($D) = True Then
		ConsoleWrite(@ScriptLineNumber & $FormID & @CRLF)
		$WinPos = StringFormat("%s" & @CRLF & "WinPOS: %d  %d " & @CRLF & "WinSize: %d %d " & @CRLF & "Desktop: %d %d ", _
				$FormID, $D[0], $D[1], $D[2], $D[3], @DesktopWidth, @DesktopHeight)
	Else
		$WinPos = ">>>About ERROR, Check the window name<<<"
	EndIf
	Debug(@CRLF & $SystemS & @CRLF & $WinPos & @CRLF & "Written by Doug Kaynor!", 0x40, 5)
EndFunc   ;==>About