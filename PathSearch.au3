#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=../icons/head_question.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Comment=Searchs the system path
#AutoIt3Wrapper_Res_Description=PathSearch
#AutoIt3Wrapper_Res_Fileversion=0.0.0.35
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=Y
#AutoIt3Wrapper_Res_LegalCopyright=Copyright 2011 Douglas B Kaynor
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Field=Developer|Douglas Kaynor
#AutoIt3Wrapper_Res_Field=AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Res_Field=Compile date|%longdate% %time%
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=y
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Run_Au3Stripper=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#cs
	This area is used to store things todo, bugs, and other notes
	Fixed:
	Skip file button
	Search crashes with no data
	Added debug
	Tool tips

	Todo:
	Add a path refresh button
	Implement an edit path button
	Verify before save of editted path
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

Opt("MustDeclareVars", 1)
DirCreate("AUXFiles")
Global $AuxPath = @ScriptDir & "\AUXFiles\"
Global $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")
Global $SystemS = @ScriptName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & @OSServicePack & @CRLF & @OSType & @CRLF & @OSArch & @IPAddress1
Global $FileListArray[1]
Global $EXETypes
Global $Debug = False

_Debug("DBGVIEWCLEAR")

If _Singleton(@ScriptName, 1) = 0 Then
	MsgBox(48, @ScriptName, @ScriptName & " is already running!")
	Exit
EndIf

;#RequireAdmin
GUIRegisterMsg($WM_SIZE, 'onresize')

For $x = 1 To $CmdLine[0]
	_Debug($x & " >> " & $CmdLine[$x] & @CRLF)
	Select
		Case StringInStr($CmdLine[$x], "help") > 0 Or StringInStr($CmdLine[$x], "?") > 0
			Help()
			Exit
		Case StringInStr($CmdLine[$x], "debug") > 0
			$Debug = True
		Case Else
			MsgBox(32, @ScriptLineNumber & " Unknown cmdline option found", _
					"Unknown cmdline option found" & @CRLF & $CmdLine[$x])
			Exit
	EndSelect
Next



#Region ### START Koda GUI section ### Form=
Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $MainForm = GUICreate(@ScriptName & " " & $FileVersion, 625, 625, 10, 10, $MainFormOptions)
GUISetFont(10, 400, 0, "Courier New")
GUISetHelp("notepad .\PathSearch.au3", $MainForm) ; Need a help file to call here

Global $ButtonGetFiles = GUICtrlCreateButton("Get files", 10, 10, 85, 20, $WS_GROUP)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Get the data using current options (paths and files)")
Global $ButtonSearch = GUICtrlCreateButton("Search", 10, 30, 85, 20, $WS_GROUP)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Start the regular expression search")

Global $ButtonShowFiles = GUICtrlCreateButton("Show files", 10, 50, 85, 20, $WS_GROUP)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Show file list")

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
Global $CheckSkipFiles = GUICtrlCreateCheckbox("Skip files", 300, 50, 148, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Do not get the files in the folders")
If $Debug Then
	GUICtrlSetState(-1, $GUI_CHECKED)
Else
	GUICtrlSetState(-1, $GUI_UNCHECKED)
EndIf

Global $ButtonFileInfo = GUICtrlCreateButton("File info", 450, 5, 95, 20, $WS_GROUP)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Display detailed information about selected file")
Global $ButtonEditPath = GUICtrlCreateButton("Edit path", 450, 25, 95, 20, $WS_GROUP)
GUICtrlSetResizing(-1, 802)
Global $ButtonReloadPath = GUICtrlCreateButton("Reload path", 450, 45, 95, 20, $WS_GROUP)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Edit the current path (not implemented yet)")
Global $ButtonHelp = GUICtrlCreateButton("Help", 550, 5, 65, 20, $WS_GROUP)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Display help about this program  (not implemented yet)")
Global $ButtonAbout = GUICtrlCreateButton("About", 550, 25, 65, 20, $WS_GROUP)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Display information about this program")
Global $ButtonExit = GUICtrlCreateButton("Exit", 550, 45, 65, 20, $WS_GROUP)
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
;Global $InputPath = GUICtrlCreateEdit("", 10, 100, 600, 20, 0x0800);+ 0x00200000);, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
Global $InputPath = GUICtrlCreateEdit("", 10, 100, 600, 20, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetResizing(-1, 546)
GUICtrlSetTip(-1, "The raw path")
_GUICtrlEdit_Scroll(-1, $SB_BOTH)
Global $InputStatus = GUICtrlCreateInput("", 10, 125, 600, 20, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetResizing(-1, 546)
GUICtrlSetTip(-1, "Status bar")

Global $TreeViewFolders = GUICtrlCreateTreeView(10, 150, 600, 150, BitOR($TVS_HASBUTTONS, $TVS_HASLINES, $TVS_LINESATROOT, $TVS_DISABLEDRAGDROP, $TVS_SHOWSELALWAYS, $TVS_INFOTIP, $WS_GROUP, $WS_TABSTOP), $WS_EX_CLIENTEDGE)
GUICtrlSetResizing(-1, $GUI_DOCKtop)
GUICtrlSetTip(-1, "List of all folders in the path")

Global $TreeViewMatches = GUICtrlCreateTreeView(10, 305, 600, 150, BitOR($TVS_HASBUTTONS, $TVS_HASLINES, $TVS_LINESATROOT, $TVS_DISABLEDRAGDROP, $TVS_SHOWSELALWAYS, $TVS_INFOTIP, $WS_GROUP, $WS_TABSTOP), $WS_EX_CLIENTEDGE)
GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
GUICtrlSetTip(-1, "Results of search results")
Global $TreeViewProblems = GUICtrlCreateTreeView(10, 460, 600, 150, BitOR($TVS_HASBUTTONS, $TVS_HASLINES, $TVS_LINESATROOT, $TVS_DISABLEDRAGDROP, $TVS_SHOWSELALWAYS, $TVS_INFOTIP, $WS_GROUP, $WS_TABSTOP), $WS_EX_CLIENTEDGE)
GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
GUICtrlSetTip(-1, "List of all problems")

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

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
			GuiDisable("disable")
			_ArrayDisplay($FileListArray, @ScriptLineNumber & ' FIles in path directories')
			GuiDisable("enable")
		Case $ButtonFileInfo
			FileInfo()
		Case $ButtonSearch
			DoTheSearch()
		Case $InputSearch
			DoTheSearch()
		Case $ButtonAbout
			About()
		Case $ButtonHelp
			Help()
		Case $CheckNoStartFolder
		Case $CheckExecutableOnly
		Case $CheckSkipFiles
		Case $ButtonEditPath
			ShellExecuteWait("SystemPropertiesAdvanced.exe")
		Case $ButtonReloadPath
			MsgBox($MB_ICONINFORMATION, "Reload path information", _
					"The program may need to be restarted to ensure " & @CRLF & _
					" that the path is reloaded after editing")
			GetPath()
		Case Else
			;ConsoleWrite($t & '  ')
	EndSwitch
WEnd
;-----------------------------------------------
#cs
	ControlGetPos ( "title", "text", controlID )
	$aArray[0] = X position
	$aArray[1] = Y position
	$aArray[2] = Width
	$aArray[3] = Height

	( "title", "text", controlID, x, y [, width [, height]] )
#ce
Func onresize($hwnd, $iMsg, $iwParam, $ilParam)
	#forceref $hwnd, $iMsg, $iwParam, $ilParam

	Local $x1 = ControlGetPos("", "", $TreeViewFolders)
	If @error == 1 Then Return
	_Debug(@ScriptLineNumber & " TreeViewFolders  " & $x1[0] & " " & $x1[1] & " " & $x1[2] & " " & $x1[3])

	ControlMove("", "", $TreeViewMatches, $x1[0], $x1[1] + $x1[3], -1, -1)

	Local $x2 = ControlGetPos("", "", $TreeViewMatches)
	If @error == 1 Then Return
	_Debug(@ScriptLineNumber & " TreeViewMatches  " & $x2[0] & " " & $x2[1] & " " & $x2[2] & " " & $x2[3])
	Local $x3 = ControlGetPos("", "", $TreeViewProblems)
	If @error == 1 Then Return
	_Debug(@ScriptLineNumber & " TreeViewProblems " & $x3[0] & " " & $x3[1] & " " & $x3[2] & " " & $x3[3] & @CRLF)


	Return
EndFunc   ;==>onresize

;-----------------------------------------------
Func GetPath()
	GuiDisable("disable")
	SplashImageOn("PathSearch is getting data. Please wait.", $AuxPath & "Working.jpg", -1, -1, -1, -1, 18)
	EnvUpdate()
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
	;_ArrayDisplay($EXETypes, @ScriptLineNumber & " Exe types")

	If GUICtrlRead($CheckNoStartFolder) = $GUI_UNCHECKED Then
		_GUICtrlTreeView_Add($TreeViewFolders, 0, @WorkingDir)
	EndIf

	Local $B
	If $Debug Then
		$B = StringStripWS(";c:\xyz;C:\Program Files\Java\jdk1.6.0_23\bin;;C:\cygwin\bin;C:\cygwin\bin\;C:\windows", 3)
	Else
		$B = StringStripWS(EnvGet("Path"), 3)
	EndIf
	;MsgBox(16, @ScriptLineNumber & " Path " & $Debug, $B)
	GUICtrlSetData($InputPath, $B)

	If StringInStr($B, ";;") > 0 Then _GUICtrlTreeView_Add($TreeViewProblems, 0, "Double semi-colons found.")

	;This fixes leading and trailing ;
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
			$B = "Path does not exist: " & $B
			_GUICtrlTreeView_Add($TreeViewProblems, 0, $B)
		Else
			_GUICtrlTreeView_Add($TreeViewFolders, 0, $B)
		EndIf
	Next

	GUICtrlSetData($InputStatus, 'GetThePath complete')
	SplashOff()
	GuiDisable("enable")
EndFunc   ;==>GetPath
;-----------------------------------------------
Func GetPathFiles()
	GuiDisable("disable")
	GUICtrlSetData($InputStatus, 'GetPathFiles starting')
	SplashImageOn("PathSearch is getting data. Please wait.", $AuxPath & "Working.jpg", -1, -1, -1, -1, 18)
	GetTheFileList()
	SplashOff()
	GUICtrlSetData($InputStatus, 'GetPathFiles complete')
	GuiDisable("enable")

EndFunc   ;==>GetPathFiles
;-----------------------------------------------
Func CheckForDuplicates(ByRef $A)
	;This builds a hash of the items. The name is the key and the count is the data.
	;The hash is a one dimensional array
	Local $TmpArray[2] ;Create 2 slots so that a crash with no data is avoided
	For $x = 0 To UBound($A) - 1
		Local $t = StringStripWS($A[$x], 3)
		;ConsoleWrite(@ScriptLineNumber & " " & $t & @CRLF)
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
	Local $ZArray[1][2]
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
	If GUICtrlRead($CheckSkipFiles) == $GUI_CHECKED Then Return
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

	GuiDisable("disable")
	GUICtrlSetData($InputMatchCount, 0)
	_GUICtrlTreeView_DeleteAll($TreeViewMatches)

	For $A In $FileListArray
		Local $B = StringSplit($A, "\", 2)
		Local $C = $B[UBound($B) - 1]

		If GUICtrlRead($CheckExecutableOnly) == $GUI_UNCHECKED Then ; Check the extension if needed
			;If StringInStr($C, GUICtrlRead($InputSearch)) > 0 Then _GUICtrlTreeView_Add($TreeViewMatches, 0, $A)
			If StringRegExp($C, GUICtrlRead($InputSearch), 0) = 1 Then _GUICtrlTreeView_Add($TreeViewMatches, 0, $A)
		Else
			For $D In $EXETypes
				;If StringInStr($A, $D) Then
				;	If StringInStr($C, GUICtrlRead($InputSearch)) > 0 Then _GUICtrlTreeView_Add($TreeViewMatches, 0, $A)
				If StringRegExp($A, $D, 0) = 1 Then
					If StringRegExp($C, GUICtrlRead($InputSearch), 0) = 1 Then _GUICtrlTreeView_Add($TreeViewMatches, 0, $A)

				EndIf
			Next
		EndIf

	Next
	GUICtrlSetData($InputMatchCount, _GUICtrlTreeView_GetCount($TreeViewMatches))

	GuiDisable("enable")
EndFunc   ;==>DoTheSearch
;-----------------------------------------------
Func Help()
	MsgBox(64 + 8192, "Help " & @ScriptName, "Help is not available yet", 5)
EndFunc   ;==>Help
;-----------------------------------------------
Func About()
	Local $D = WinGetPos(@ScriptName)
	Local $WinPos
	If IsArray($D) = 1 Then
		$WinPos = StringFormat("%s" & @CRLF & "WinPOS: %d  %d " & @CRLF & "WinSize: %d %d " & @CRLF & "Desktop: %d %d ", _
				$MainForm, $D[0], $D[1], $D[2], $D[3], @DesktopWidth, @DesktopHeight)
	Else
		$WinPos = ">>>About ERROR. Check the window name<<<"
	EndIf
	_Debug(@CRLF & $SystemS & @CRLF & $WinPos & @CRLF & "Written by Doug Kaynor because I wanted to!", 0x40, 5)
EndFunc   ;==>About
;-----------------------------------------------;
Func FileInfo()
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

			MsgBox(64, 'File info', _
					StringFormat("%s %s %s %s %s %s %s %s %s", _
					FileGetLongName($file), _
					@CRLF & FileGetShortName($file), _
					@CRLF & "Attributes:    " & FileGetAttrib($file), _
					@CRLF & "Size:          " & FileGetSize($file), _
					@CRLF & "Version:       " & FileGetVersion($file), _
					@CRLF & "Modified Time: " & FormatedFileGetTime($file, 0), _
					@CRLF & "Create time:   " & FormatedFileGetTime($file, 1), _
					@CRLF & "Access Time:   " & FormatedFileGetTime($file, 2)))
		EndIf
	Next
EndFunc   ;==>FileInfo
;-----------------------------------------------
Func FormatedFileGetTime($file, $Type)
	Local $FT = FileGetTime($file, $Type)
	Return StringFormat("%s-%s-%s %s:%s:%s", $FT[0], $FT[1], $FT[2], $FT[3], $FT[4], $FT[5])
EndFunc   ;==>FormatedFileGetTime
;-----------------------------------------------
Func GuiDisable($choice)
	_Debug(@ScriptLineNumber & " GuiDisable  " & $choice)
	Static $LastState
	Local $setting

	If $choice = "Enable" Then
		$setting = $GUI_ENABLE
	ElseIf $choice = "Disable" Then
		$setting = $GUI_DISABLE
	ElseIf $choice = "Toggle" Then
		If $LastState = $GUI_DISABLE Then
			$setting = $GUI_ENABLE
		Else
			$setting = $GUI_DISABLE
		EndIf
	Else
		_Debug(@ScriptLineNumber & " Invalid choice at GuiDisable" & $choice, 0x40)
	EndIf

	GUICtrlSetState($ButtonSearch, $setting)
	GUICtrlSetState($ButtonGetFiles, $setting)
	GUICtrlSetState($ButtonShowFiles, $setting)
	GUICtrlSetState($ButtonFileInfo, $setting)
	GUICtrlSetState($ButtonEditPath, $setting)
	GUICtrlSetState($ButtonReloadPath, $setting)
	GUICtrlSetState($ButtonHelp, $setting)
	GUICtrlSetState($ButtonAbout, $setting)
	GUICtrlSetState($ButtonExit, $setting)
	GUICtrlSetState($InputSearch, $setting)
	GUICtrlSetState($CheckNoStartFolder, $setting)
	GUICtrlSetState($CheckExecutableOnly, $setting)
	GUICtrlSetState($CheckSkipFiles, $setting)

EndFunc   ;==>GuiDisable
;-----------------------------------------------
Func _Debug($msg, $ShowMsgBox = False, $Timeout = 0)
	Local $DebugMSG = ''
	$DebugMSG = "DEBUG " & @ScriptName & "  " & $msg & @CRLF
	DllCall("kernel32.dll", "none", "OutputDebugString", "str", $DebugMSG)
	ConsoleWrite($DebugMSG)
	If $ShowMsgBox = True Then MsgBox(48, @ScriptName & " Debug", $DebugMSG, $Timeout)
EndFunc   ;==>_Debug
