#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=../icons/head_question.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Comment=Searchs the system path
#AutoIt3Wrapper_Res_Description=PathSearch
#AutoIt3Wrapper_Res_Fileversion=0.0.0.33
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=Y
#AutoIt3Wrapper_Res_LegalCopyright=Copyright 2011 Douglas B Kaynor
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Field=Developer|Douglas Kaynor
#AutoIt3Wrapper_Res_Field=AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Res_Field=Compile date|%longdate% %time%
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=y
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Run_Tidy=y
#Obfuscator_Parameters=/Convert_Strings=0 /Convert_Numerics=0 /showconsoleinfo=9
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****

#cs
	This area is used to store things todo, bugs, and other notes
	Fixed:
	Skip file button
	Search crashes with no data
	Tool tips
	
	Todo:
	Add a path refresh button
	
	Open the system properties/Enviroment Varibles window
	
	Path is saved back wards
	
	Duplicate not found if / on end
	Show full unprocessed path (scroll) ???
#CE


#include <Array.au3>
#include <Date.au3>
#include <file.au3>
#include <Misc.au3>
#include <String.au3>

#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstants.au3>
#include <GuiEdit.au3>
#include <GuiListView.au3>
#include <GUIConstantsEx.au3>
#include <GuiTreeView.au3>
#include <GuiImageList.au3>
#include <ListViewConstants.au3>
#include <ScrollBarConstants.au3>
#include <StaticConstants.au3>
#include <TreeViewConstants.au3>
#include <WindowsConstants.au3>
#include <_DougFunctions.au3>

Opt("MustDeclareVars", 1)
Global $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")
Global $SystemS = @ComputerName & @CRLF & @ScriptName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & @OSServicePack & @CRLF & @OSType & @CRLF & @OSArch
Global $FileListArray[1]
Global $EXETypes

If _Singleton(@ScriptName, 1) = 0 Then
	MsgBox(48, @ScriptName, @ScriptName & " is already running!")
	Exit
EndIf

#region ### START Koda GUI section ### Form
Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $MainForm = GUICreate(@ScriptName & " " & $FileVersion, 625, 625, 10, 10, $MainFormOptions)
GUISetFont(10, 400, 0, "Courier New")
GUISetHelp("notepad .\PathSearch.au3", $MainForm) ; Need a help file to call here

Global $ButtonSearch = GUICtrlCreateButton("Search", 10, 10, 85, 20, $WS_GROUP)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Start the regular expression search")

Global $ButtonFileInfo = GUICtrlCreateButton("File info", 10, 30, 85, 20, $WS_GROUP)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Display detailed information about selected file")

Global $ButtonEditPath = GUICtrlCreateButton("Edit path", 10, 50, 85, 20, $WS_GROUP)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Edit the current path (not implemented yet)")

Global $InputSearch = GUICtrlCreateInput("", 100, 10, 190, 24, BitOR($ES_CENTER, $ES_AUTOHSCROLL))
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "String to search for. (regular expression)")
Global $CheckNoStartFolder = GUICtrlCreateCheckbox("No start folder", 300, 10, 148, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Do not include the start folder in the path search")
GUICtrlSetState(-1, $GUI_CHECKED)
Global $CheckExecutableOnly = GUICtrlCreateCheckbox("Executable only", 300, 30, 148, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Show executable files only")
GUICtrlSetState(-1, $GUI_CHECKED)


Global $ButtonGetFiles = GUICtrlCreateButton("Get files", 460, 10, 85, 20, $WS_GROUP)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Get the data using current options (paths and files)")
Global $ButtonGetPath = GUICtrlCreateButton("Get path", 460, 30, 85, 20, $WS_GROUP)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Refresh paths")
Global $ButtonShowFiles = GUICtrlCreateButton("Show files", 460, 50, 85, 20, $WS_GROUP)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Show file list")

Global $ButtonHelp = GUICtrlCreateButton("Help", 550, 10, 65, 20, $WS_GROUP)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Display help about this program  (not implemented yet)")
Global $ButtonAbout = GUICtrlCreateButton("About", 550, 30, 65, 20, $WS_GROUP)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Display information about this program")
Global $ButtonExit = GUICtrlCreateButton("Exit", 550, 50, 65, 20, $WS_GROUP)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Exit this program")
Global $InputFolderCount = GUICtrlCreateInput("0", 100, 45, 60, 20, BitOR($ES_CENTER, $ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Number of folders in path")
Global $InputFileCount = GUICtrlCreateInput("0", 165, 45, 60, 20, BitOR($ES_CENTER, $ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Number of files in path")
Global $InputMatchCount = GUICtrlCreateInput("0", 230, 45, 60, 20, BitOR($ES_CENTER, $ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Number of matching files in path")


Global $InputEXETypes = GUICtrlCreateInput("", 10, 75, 600, 20, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Defined EXE types")
Global $InputPath = GUICtrlCreateEdit("", 10, 100, 600, 20, 0x0800);+ 0x00200000);, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetResizing(-1, 546)
GUICtrlSetTip(-1, "The raw path")
_GUICtrlEdit_Scroll(-1, $SB_BOTH)
Global $InputStatus = GUICtrlCreateInput("", 10, 125, 600, 20, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetResizing(-1, 546)
GUICtrlSetTip(-1, "Status bar")

Global $ListOptions = BitOR($TVS_HASBUTTONS, $TVS_HASLINES, $TVS_LINESATROOT, $TVS_DISABLEDRAGDROP, $TVS_SHOWSELALWAYS, $TVS_INFOTIP, $WS_GROUP, $WS_TABSTOP)
Global $TreeViewFolders = GUICtrlCreateTreeView(10, 150, 600, 150, $ListOptions, $WS_EX_CLIENTEDGE)
GUICtrlSetResizing(-1, $GUI_DOCKTOP)
GUICtrlSetTip(-1, "List of all folders in the path")
Global $TreeViewMatches = GUICtrlCreateTreeView(10, 305, 600, 150, $ListOptions, $WS_EX_CLIENTEDGE)
;GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP)
GUICtrlSetTip(-1, "Results of search results.")
Global $TreeViewProblems = GUICtrlCreateTreeView(10, 460, 600, 150, $ListOptions, $WS_EX_CLIENTEDGE)
;GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKBOTTOM)
GUICtrlSetTip(-1, "List of all problems.")

ConsoleWrite(@ScriptLineNumber & " " & $TreeViewProblems & @CRLF)

GUISetState(@SW_SHOW)
#endregion ### END Koda GUI section ###

GetPath()

While 1
	Global $t = GUIGetMsg()
	Switch $t
		Case $GUI_EVENT_CLOSE
			Exit
		Case $ButtonExit
			Exit
		Case $ButtonGetFiles
			GetPath()
			GetPathFiles()
		Case $ButtonShowFiles
			_GuiDisable("disable")
			_ArrayDisplay($FileListArray, @ScriptLineNumber & ' FIles in path directories')
			_GuiDisable("enable")
		Case $ButtonGetPath
			GetPath()
		Case $ButtonFileInfo
			ShowFileInfo()
		Case $ButtonSearch
			DoTheSearch()
		Case $InputSearch
			DoTheSearch()
		Case $ButtonAbout
			_About(@ScriptName, $SystemS)
		Case $ButtonHelp
			_Help('Help not ready yet')
		Case $CheckNoStartFolder
		Case $CheckExecutableOnly
		Case $ButtonEditPath
			EditPath()
		Case Else
			;ConsoleWrite($t & '  ')
	EndSwitch
WEnd
;-----------------------------------------------  Test for valid and duplicate folders in the path _arr
Func EditPath()
	;Local $var = RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\", "Path")
	Local $TmpFile = $AuxPath & "PathSearch.txt"

	Local $hItem = _GUICtrlTreeView_GetFirstItem($TreeViewFolders)
	Local $CurrentPath = _GUICtrlTreeView_GetText($TreeViewFolders, $hItem)
	FileWriteLine($TmpFile, $CurrentPath)

	While $hItem <> 0
		$hItem = _GUICtrlTreeView_GetNext($TreeViewFolders, $hItem)
		If $hItem <> 0 Then
			$CurrentPath = _GUICtrlTreeView_GetText($TreeViewFolders, $hItem)
			FileWriteLine($TmpFile, $CurrentPath)
		EndIf
	WEnd

	Local $editor = _ChoseTextEditor()
	ShellExecuteWait($editor, $TmpFile, "", "open")

	If MsgBox(36, "Edit path", "Do you want to save the changes?") = 6 Then
		Local $TmpArray
		Local $TS = ''
		_FileReadToArray($TmpFile, $TmpArray)
		_ArrayDelete($TmpArray, 0)

		For $x In $TmpArray
			$TS = $TS & ';' & $x
		Next
		ConsoleWrite(@ScriptLineNumber & " " & $TS & " " & @CRLF)
		ClipPut($TS)

		;If RegWrite("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\", "path", "REG_EXPAND_SZ", $TS) = 0 Then
		;	MsgBox(16, "Registry update error", '@error: ' & @error)
		;EndIf
		;EnvUpdate()
		;If @error = 1 Then
		;	MsgBox(16, "EnvUpdate error", '@error: ' & @error)
		;EndIf
		ShellExecute('systempropertiesadvanced.exe')
		Sleep(100)
		WinActivate('System Properties')
		Send("{TAB 3}{SPACE}{TAB 4}{DOWN 7}{TAB 2}{SPACE}{DELETE}{^V}")

		;ControlClick("[CLASS:Button]", 'Enviro&nment Variables...', '', 106)

		ConsoleWrite(@ScriptLineNumber & " " & RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\", "path") & @CRLF)

	EndIf
	If FileExists($TmpFile) Then FileDelete($TmpFile)

EndFunc   ;==>EditPath
;-----------------------------------------------
Func GetPath()
	_GuiDisable("disable")
	SplashImageOn("PathSearch is getting data. Please wait.", $AuxPath & "Working.jpg", -1, -1, -1, -1, 18)
	GUICtrlSetData($InputStatus, 'GetThePath starting')
	_GUICtrlTreeView_DeleteAll($TreeViewFolders)
	_GUICtrlTreeView_DeleteAll($TreeViewMatches)
	_GUICtrlTreeView_DeleteAll($TreeViewProblems)
	GUICtrlSetData($InputFileCount, 0)
	GUICtrlSetData($InputMatchCount, 0)

	ReDim $FileListArray[1]

	GUICtrlSetData($InputSearch, "")

	GUICtrlSetData($InputEXETypes, EnvGet("PathEXT"))
	$EXETypes = StringSplit(GUICtrlRead($InputEXETypes), ';', 2)

	If GUICtrlRead($CheckNoStartFolder) = $GUI_UNCHECKED Then
		_GUICtrlTreeView_Add($TreeViewFolders, 0, @WorkingDir)
	EndIf

	Local $B = StringStripWS(EnvGet("Path"), 3)
	GUICtrlSetData($InputPath, $B)

	If StringInStr($B, ";;") > 0 Then _GUICtrlTreeView_Add($TreeViewProblems, 0, "Double semi-colons found.")

	;This removes leading and trailing ;
	$B = StringRegExpReplace($B, "^;", "")
	$B = StringRegExpReplace($B, ";$", "")

	Local $A = StringSplit($B, ";")

	GUICtrlSetData($InputFolderCount, $A[0])
	Local $G[1]
	_ArrayDelete($A, 0)

	For $B In $A
		If StringInStr($B, ':\') = 2 Then
			_ArrayAdd($G, $B)
		Else
			_GUICtrlTreeView_Add($TreeViewProblems, 0, "Bad path entry found: " & $B)
		EndIf
	Next
	_ArrayDelete($G, 0)
	$A = $G

	CheckForDuplicates($A)

	For $B In $A
		If StringLen($B) < 3 Then ContinueLoop
		If FileExists($B) = 0 Then
			_GUICtrlTreeView_Add($TreeViewProblems, 0, "Path folder does not exist: " & $B)
		Else
			_GUICtrlTreeView_Add($TreeViewFolders, 0, $B)
		EndIf
	Next

	GUICtrlSetData($InputStatus, 'GetThePath complete')
	SplashOff()
	_GuiDisable("enable")
EndFunc   ;==>GetPath
;-----------------------------------------------
Func GetPathFiles()
	_GuiDisable("disable")
	GUICtrlSetData($InputStatus, 'GetPathFiles starting')
	SplashImageOn("PathSearch is getting data. Please wait.", $AuxPath & "Working.jpg", -1, -1, -1, -1, 18)
	GetTheFileList()
	SplashOff()
	GUICtrlSetData($InputStatus, 'GetPathFiles complete')
	_GuiDisable("enable")

EndFunc   ;==>GetPathFiles
;-----------------------------------------------
Func CheckForDuplicates(ByRef $A)
	;This builds a hash of the items. The name is the key and the count is the data.
	;The hash is a one dimensional array
	Global $TmpArray[2] ;Create 2 slots so that a crash with no data is avoided
	For $x = 0 To UBound($A) - 1
		Local $t = StringStripWS($A[$x], 3)
		ConsoleWrite(@ScriptLineNumber & " " & $t & @CRLF)
		Local $POS = _ArraySearch($TmpArray, $t)
		If $POS = -1 Then
			_ArrayAdd($TmpArray, $t)
			_ArrayAdd($TmpArray, 1)
		Else
			$TmpArray[$POS + 1] += 1
		EndIf
	Next
	;_ArrayDisplay($TmpArray, @ScriptLineNumber & " CheckForDups")
	;Convert the hash to a two dimensional array
	Local $Count = 0
	Global $ZArray[1][2]
	While UBound($TmpArray) > 0
		$ZArray[$Count][0] = _ArrayPop($TmpArray)
		$ZArray[$Count][1] = _ArrayPop($TmpArray)
		$Count += 1
		ReDim $ZArray[UBound($ZArray) + 1][2]
	WEnd

	For $x = 0 To UBound($ZArray) - 1
		If $ZArray[$x][0] <> 1 And StringLen($ZArray[$x][1]) Then
			_GUICtrlTreeView_Add($TreeViewProblems, 0, "Duplicate path: " & $ZArray[$x][1])
		EndIf
	Next
EndFunc   ;==>CheckForDuplicates
;-----------------------------------------------
Func GetTheFileList()
	GUICtrlSetData($InputStatus, 'GetTheFileList')
	GUICtrlSetData($InputFileCount, 0)
	GUICtrlSetData($InputMatchCount, 0)
	ReDim $FileListArray[1]

	Local $Data[1]
	Local $hItem = _GUICtrlTreeView_GetFirstItem($TreeViewFolders)
	_ArrayAdd($Data, $hItem)
	While $hItem <> 0
		$hItem = _GUICtrlTreeView_GetNext($TreeViewFolders, $hItem)
		If $hItem <> 0 Then _ArrayAdd($Data, $hItem)
	WEnd
	_ArrayDelete($Data, 0)

	For $item In $Data
		Local $CurrentPath = _GUICtrlTreeView_GetText($TreeViewFolders, $item)
		If FileChangeDir($CurrentPath) = 0 Then
			_GUICtrlTreeView_Add($TreeViewProblems, 0, "Unable to change to directory: " & $CurrentPath)
			ContinueLoop
		EndIf
		;_Debug(@ScriptLineNumber & ' Current path: ' & $CurrentPath)
		GUICtrlSetData($InputStatus, 'GetTheFileList: ' & $CurrentPath)
		Local $search = FileFindFirstFile("*.*")
		If $search = -1 Then ; Check if the search was successful
			_GUICtrlTreeView_Add($TreeViewProblems, 0, "Empty folder in path: " & $CurrentPath)
			ContinueLoop
		EndIf
		While 1
			Local $file = FileFindNextFile($search)
			If @error Then ExitLoop
			_ArrayAdd($FileListArray, $CurrentPath & '\' & $file)
		WEnd

		_ArrayDelete($FileListArray, 0)
		GUICtrlSetData($InputFileCount, UBound($FileListArray))
	Next

	GUICtrlSetData($InputStatus, 'GetTheFileList complete')

EndFunc   ;==>GetTheFileList
;-----------------------------------------------
Func DoTheSearch()
	If Not IsArray($EXETypes) Then Return
	If Not IsArray($FileListArray) Then Return
	ConsoleWrite(@ScriptLineNumber & " Files in list: " & UBound($FileListArray) & @CRLF)
	If UBound($FileListArray) < 2 Then
		MsgBox(48, "No files to search", "Run get files before searching")
		Return
	EndIf

	_GuiDisable("disable")
	GUICtrlSetData($InputMatchCount, 0)
	_GUICtrlTreeView_DeleteAll($TreeViewMatches)

	For $A In $FileListArray
		Local $B = StringSplit($A, "\", 2)
		Local $C = $B[UBound($B) - 1]

		If GUICtrlRead($CheckExecutableOnly) = $GUI_UNCHECKED Then
			If StringRegExp($C, GUICtrlRead($InputSearch), 0) = 1 Then _GUICtrlTreeView_Add($TreeViewMatches, 0, $A) ; extension does not matter
		Else
			For $D In $EXETypes
				If StringInStr($A, $D, 0, -1) = StringLen($A) - 3 Then ; Verify that it is an executable file
					If StringRegExp($C, GUICtrlRead($InputSearch), 0) = 1 Then _GUICtrlTreeView_Add($TreeViewMatches, 0, $A); Do the strings match?
				EndIf
			Next
		EndIf

	Next
	GUICtrlSetData($InputMatchCount, _GUICtrlTreeView_GetCount($TreeViewMatches))

	_GuiDisable("enable")
EndFunc   ;==>DoTheSearch
;-----------------------------------------------
Func ShowFileInfo()
	Local $Data[1]
	Local $hItem = _GUICtrlTreeView_GetFirstItem($TreeViewMatches)
	_ArrayAdd($Data, $hItem)
	While $hItem <> 0
		$hItem = _GUICtrlTreeView_GetNext($TreeViewMatches, $hItem)
		If $hItem <> 0 Then _ArrayAdd($Data, $hItem)
	WEnd
	_ArrayDelete($Data, 0)

	For $x = 0 To UBound($Data) - 1
		If _GUICtrlTreeView_GetSelected($TreeViewMatches, $Data[$x]) == True Then
			Local $file = _GUICtrlTreeView_GetText($TreeViewMatches, $Data[$x])
			_FileInfo($file)
		EndIf
	Next
EndFunc   ;==>ShowFileInfo
;-----------------------------------------------
