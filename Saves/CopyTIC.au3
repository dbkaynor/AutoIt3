#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=../icons/HotSun.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Comment=A program to get a clean copy of a TIC
#AutoIt3Wrapper_Res_Description=CopyTIC
#AutoIt3Wrapper_Res_Fileversion=3.0.0.10
#AutoIt3Wrapper_Res_FileVersion_AutoIncrement=Y
#AutoIt3Wrapper_Res_ProductVersion=666
#AutoIt3Wrapper_Res_LegalCopyright=Copyright © 2012 Douglas B Kaynor
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Field=Developer|Douglas Kaynor
#AutoIt3Wrapper_Res_Field=AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Res_Field=Compile date|%longdate% %time%
#AutoIt3Wrapper_Au3Check_Stop_OnWarning=y
#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Run_Tidy=y
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****

#CS
	
#CE

Opt("MustDeclareVars", 1) ; 0=no, 1=require pre-declare
Opt("GUICoordMode", 1) ; 0=relative, 1=absolute, 2=cell
Opt("GUIResizeMode", 1) ; 0=no resizing, <1024 special resizing

#include <ButtonConstants.au3>
#include <Constants.au3>
#include <Date.au3>
#include <GUIConstantsEx.au3>
#include <GuiComboBox.au3>

#include <ListBoxConstants.au3>
#include <Misc.au3>
#include <StaticConstants.au3>
#include <String.au3>
#include <WindowsConstants.au3>
#include <Constants.au3>
#include <GuiListBox.au3>
#include <Misc.au3>
#include "_DougFunctions.au3"

;Func _Debug($DebugMSG, $Log_filename = '', $ShowMsgBox = False, $Timeout = 0, $Verbose = False)

If _Singleton(@ScriptName, 1) = 0 Then
	_debug(@ScriptName & "is already running! ", '', True)
	Exit
EndIf
DirCreate("AUXFiles")
Global $Editor = ""

Global $tmp = StringSplit(@ScriptName, ".")
Global $ProgramName = $tmp[1]

Global $Projectfilename = @ScriptDir & ".\AUXFiles\" & $tmp[1] & ".prj"
Const $LOGfilename = @ScriptDir & ".\AUXFiles\" & $tmp[1] & ".log"
Global $Font = "Courier new"
GUISetFont(10, 400, -1, $Font)

Global $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")

Global $ResultLocation = ""
Global $SystemS = $ProgramName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & _
		@OSServicePack & @CRLF & @OSType & @CRLF & @OSArch

Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, _
		$WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $MainForm = GUICreate($ProgramName & $FileVersion, 600, 560, 10, 10, $MainFormOptions)

;menu items -----------------------------------------------
Global $FileMenu = GUICtrlCreateMenu('File')
Global $LoadProjectItem = GUICtrlCreateMenuItem('Load project', $FileMenu)
Global $SaveProjectItem = GUICtrlCreateMenuItem('Save project', $FileMenu)
Global $EditorItem = GUICtrlCreateMenuItem('Edit file', $FileMenu)

;;Buttons and such -----------------------------------------------
Global $ButtonSearch = GUICtrlCreateButton("Search", 10, 10, 70, 25)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonSelectBuild = GUICtrlCreateButton("Select Build", 80, 10, 70, 25)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonNetcopy = GUICtrlCreateButton("Net copy", 150, 10, 70, 25)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonClear = GUICtrlCreateButton("Clear", 310, 10, 70, 25)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonAbout = GUICtrlCreateButton("About", 390, 10, 70, 25)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonHelp = GUICtrlCreateButton("Help", 460, 10, 70, 25)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonExit = GUICtrlCreateButton("Exit", 530, 10, 70, 25)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ComboSearch = GUICtrlCreateCombo("VC*| ", 10, 40, 185, 20)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $CheckLinux32 = GUICtrlCreateCheckbox("Linux32", 10, 70)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $CheckLinux64 = GUICtrlCreateCheckbox("Linux64", 90, 70)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $CheckLinux_x64 = GUICtrlCreateCheckbox("Linux_x64", 180, 70)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $CheckWin32 = GUICtrlCreateCheckbox("Win32", 10, 90)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $CheckWin64 = GUICtrlCreateCheckbox("Win64", 90, 90)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $CheckWin64e = GUICtrlCreateCheckbox("Win64e", 180, 90)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $CheckPro100 = GUICtrlCreateCheckbox("Pro100", 10, 110)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $CheckPro1000 = GUICtrlCreateCheckbox("Pro1000", 90, 110)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $CheckProxgb = GUICtrlCreateCheckbox("Proxgb", 180, 110)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $CheckTools = GUICtrlCreateCheckbox("Tools", 260, 90)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $CheckAPPS = GUICtrlCreateCheckbox("Apps", 260, 110)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ComboSource = GUICtrlCreateCombo("", 10, 140, 400, 20)
GUICtrlSetTip(-1, 'ComboSource')
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetData($ComboSource, "\\elfin\MEDIA_TREE\LAD SW Reference|c:\temp|z:\| ")
Global $LabelSource = GUICtrlCreateLabel("TIC source", 420, 140, 300)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ComboDestination = GUICtrlCreateCombo("c:\temp| ", 10, 160, 400, 20)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetTip(-1, 'ComboDestination')
Global $LabelDestination = GUICtrlCreateLabel("TIC destination", 420, 160, 300)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $InputTICName = GUICtrlCreateInput("", 10, 180, 400, 20)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetTip(-1, 'ComboName')
Global $LabelName = GUICtrlCreateLabel("TIC name", 420, 180, 300)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ListBuilds = GUICtrlCreateList("", 10, 210, 500, 100, BitOR($LBS_DISABLENOSCROLL, $WS_BORDER, $WS_HSCROLL, $WS_VSCROLL))
GUICtrlSetResizing(-1, BitOR($GUI_DOCKTOP, $GUI_DOCKBOTTOM))
GUICtrlSetTip(-1, 'ListBuilds')
Global $LabelBuilds = GUICtrlCreateLabel("Projects", 520, 230, 300)

Global $ListTICInfo = GUICtrlCreateList("", 10, 320, 500, 100, BitOR($LBS_DISABLENOSCROLL, $WS_BORDER, $WS_HSCROLL, $WS_VSCROLL))
GUICtrlSetTip(-1, 'ListTICInfo')
Global $LabelTICInfo = GUICtrlCreateLabel("TIC info", 520, 340, 300)

Global $ListStatus = GUICtrlCreateList("", 10, 430, 500, 100, BitOR($LBS_DISABLENOSCROLL, $WS_BORDER, $WS_HSCROLL, $WS_VSCROLL))
GUICtrlSetTip(-1, 'Status')
Global $LabelStatus = GUICtrlCreateLabel("Status", 520, 450, 300)

HotKeySet("{F11}", "Enable")
Func Enable()
	_GuiDisable("Enable")
EndFunc   ;==>Enable

;-----------------------------------------------
GUISetState()
LoadProject("start")
_CheckWindowLocation($ProgramName)

;-----------------------------------------------
; Run the GUI until the dialog is closed
While 1
	Switch GUIGetMsg()
		Case $LoadProjectItem
			LoadProject("menu")
		Case $SaveProjectItem
			SaveProject()
		Case $EditorItem
			$Editor = _ChoseTextEditor()
			ShellExecute($Editor, $Projectfilename)

		Case $ButtonSearch
			GetTICList()
		Case $ButtonNetcopy
			CopyTIC()
		Case $ButtonClear
			_GUICtrlListBox_ResetContent($ListBuilds)
			_GUICtrlListBox_ResetContent($ListTICInfo)
			_GUICtrlListBox_ResetContent($ListStatus)
		Case $ButtonSelectBuild
			GetBuildName()
			GetTICInfo()
		Case $ButtonAbout
			_About($ProgramName, $SystemS)
		Case $ButtonHelp
			Help()
		Case $GUI_EVENT_CLOSE
			ExitLoop
		Case $ButtonExit
			ExitLoop
	EndSwitch
WEnd
;-----------------------------------------------
Func ToggleEnables()
	MsgBox(0, "Hot key detected", "Hot key detected")
EndFunc   ;==>ToggleEnables
;-----------------------------------------------
Func LoadProject($type)
	If StringCompare($type, "menu") = 0 Then
		$Projectfilename = FileOpenDialog("Loadproject file", @ScriptDir & ".\AUXFiles\", _
				$ProgramName & " projects (C*.prj)|All projects (*.prj)|All files (*.*)", 18, @ScriptDir & ".\AUXFiles\CopyTIC.prj")
	EndIf

	Local $file = FileOpen($Projectfilename, 0)
	; Check if file opened for reading OK
	If $file = -1 Then
		_Debug("Unable to open file for reading: " & $Projectfilename, '', True)
		Return
	EndIf

	_GUICtrlListBox_ResetContent($ListBuilds)
	_GUICtrlListBox_ResetContent($ListTICInfo)
	_GUICtrlListBox_ResetContent($ListStatus)
	_GUICtrlComboBox_ResetContent($ComboSearch)
	_GUICtrlComboBox_ResetContent($ComboSource)
	_GUICtrlComboBox_ResetContent($ComboDestination)
	GUICtrlSetData($InputTICName, "")

	; Read in the first line to verify the file is of the correct type
	If StringInStr(FileReadLine($file, 1), "Project file for " & $ProgramName) <> 1 Then
		ConsoleWrite(@ScriptLineNumber & " >" & FileReadLine($file, 1) & "<" & @CRLF)
		ConsoleWrite(@ScriptLineNumber & " >Project file for " & $ProgramName & "<" & @CRLF)
		ConsoleWrite(@ScriptLineNumber & " " & StringInStr(FileReadLine($file, 1), "Project file for " & $ProgramName) & @CRLF)
		_Debug("Not a Project file for " & $ProgramName, '', True)
		Return
	EndIf

	; Read in lines of text until the EOF is reached
	While 1
		Local $lineIn = FileReadLine($file)
		If @error = -1 Then ExitLoop
		If StringInStr($lineIn, "linux32:") Then GUICtrlSetState($CheckLinux32, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
		If StringInStr($lineIn, "linux64:") Then GUICtrlSetState($CheckLinux64, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
		If StringInStr($lineIn, "linux_x64:") Then GUICtrlSetState($CheckLinux_x64, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
		If StringInStr($lineIn, "win32:") Then GUICtrlSetState($CheckWin32, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
		If StringInStr($lineIn, "win64:") Then GUICtrlSetState($CheckWin64, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
		If StringInStr($lineIn, "win64e:") Then GUICtrlSetState($CheckWin64e, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
		If StringInStr($lineIn, "pro100:") Then GUICtrlSetState($CheckPro100, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
		If StringInStr($lineIn, "pro1000:") Then GUICtrlSetState($CheckPro1000, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
		If StringInStr($lineIn, "proxgb:") Then GUICtrlSetState($CheckProxgb, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
		If StringInStr($lineIn, "tools:") Then GUICtrlSetState($CheckTools, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
		If StringInStr($lineIn, "apps:") Then GUICtrlSetState($CheckAPPS, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
		If StringInStr($lineIn, "search:") Then GUICtrlSetData($ComboSearch, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
		If StringInStr($lineIn, "source:") Then GUICtrlSetData($ComboSource, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
		If StringInStr($lineIn, "destination:") Then GUICtrlSetData($ComboDestination, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
		If StringInStr($lineIn, "name:") Then GUICtrlSetData($InputTICName, StringMid($lineIn, StringInStr($lineIn, ":") + 1))
		_SetWindowPosition("MainWindow:", $ProgramName, $lineIn)
	WEnd
	FileClose($file)

	_GUICtrlComboBox_SetCurSel($ComboSearch, 0)
	_GUICtrlComboBox_SetCurSel($ComboSource, 0)
	_GUICtrlComboBox_SetCurSel($ComboDestination, 0)

	If Not FileExists(GUICtrlRead($ComboSource)) Then
		_Debug("TIC source directory not found " & @CRLF & GUICtrlRead($ComboSource), '', True)
	EndIf
	If Not FileExists(GUICtrlRead($ComboDestination)) Then
		_Debug("TIC destination directory not found " & GUICtrlRead($ComboDestination), '', True)
	EndIf



EndFunc   ;==>LoadProject
;-----------------------------------------------
Func SaveProject()
	$Projectfilename = FileSaveDialog("Save project file", @ScriptDir & ".\AUXFiles\", _
			$ProgramName & " projects (C*.prj)|All projects (*.prj)|All files (*.*)", 18, @ScriptDir & ".\AUXFiles\CopyTIC.prj")

	Local $file = FileOpen($Projectfilename, 2)
	; Check if file opened for writing OK
	If $file = -1 Then
		_Debug("Unable to open file for writing: " & $Projectfilename, '', True)
		Exit
	EndIf

	FileWriteLine($file, "Project file for " & $ProgramName & "  Saved on " & _DateTimeFormat(_NowCalc(), 0))
	; Write the lines of text to the file
	FileWriteLine($file, "Valid for " & $ProgramName)
	FileWriteLine($file, "linux32:" & GUICtrlRead($CheckLinux32))
	FileWriteLine($file, "linux64:" & GUICtrlRead($CheckLinux64))
	FileWriteLine($file, "linux_x64:" & GUICtrlRead($CheckLinux_x64))
	FileWriteLine($file, "win32:" & GUICtrlRead($CheckWin32))
	FileWriteLine($file, "win64:" & GUICtrlRead($CheckWin64))
	FileWriteLine($file, "win64e:" & GUICtrlRead($CheckWin64e))
	FileWriteLine($file, "pro100:" & GUICtrlRead($CheckPro100))
	FileWriteLine($file, "pro1000:" & GUICtrlRead($CheckPro1000))
	FileWriteLine($file, "proxgb:" & GUICtrlRead($CheckProxgb))
	FileWriteLine($file, "tools:" & GUICtrlRead($CheckTools))
	FileWriteLine($file, "apps:" & GUICtrlRead($CheckAPPS))
	FileWriteLine($file, "search:" & _GUICtrlComboBox_GetList($ComboSearch))
	FileWriteLine($file, "source:" & _GUICtrlComboBox_GetList($ComboSource))
	FileWriteLine($file, "destination:" & _GUICtrlComboBox_GetList($ComboDestination))
	FileWriteLine($file, "name:" & GUICtrlRead($InputTICName))
	_SaveWindowPosition("MainWindow:", $ProgramName, $file)
	FileClose($file)
EndFunc   ;==>SaveProject
;-----------------------------------------------
;This function searchs for and lists the matching TICs in User_list1 list box
Func GetTICList()
	_GUICtrlListBox_ResetContent($ListBuilds)
	_GUICtrlListBox_ResetContent($ListTICInfo)
	If Not FileExists(GUICtrlRead($ComboSource)) Then
		_Debug("TIC source directory not found " & GUICtrlRead($ComboSource), '', True)
		Return
	EndIf
	FileChangeDir(GUICtrlRead($ComboSource))

	Local $FileSearchString = GUICtrlRead($ComboSearch, 1)
	Local $search = FileFindFirstFile($FileSearchString)

	; Check if the search was successful
	If $search = -1 Then
		_Debug("No matches found: " & $FileSearchString)
		Return
	EndIf

	While 1
		Local $file = FileFindNextFile($search)
		If @error Then ExitLoop
		GUICtrlSetData($ListBuilds, $file)
	WEnd
	FileClose($search)
	ConsoleWrite(@ScriptLineNumber & " " & _GUICtrlListBox_GetCount($ListBuilds) & @CRLF)
	_GUICtrlListBox_SetTopIndex($ListBuilds, _GUICtrlListBox_GetCount($ListBuilds) - 1)
EndFunc   ;==>GetTICList
;-----------------------------------------------
;This lists the TIC files to $ListTICInfo when a TIC name is selected in $ListBuilds
Func GetTICInfo()
	_Debug("GetTICInfo")
	_GUICtrlListBox_ResetContent($ListTICInfo)
	Local $TmpS = GUICtrlRead($ComboSource) & "\" & _GUICtrlListBox_GetText($ListBuilds, _GUICtrlListBox_GetCurSel($ListBuilds))
	Local $XX = $TmpS & "\disk\verfile.tic"

	GUICtrlSetData($ListTICInfo, $XX)
	If FileExists($XX) Then
		Local $Time = FileGetTime($XX)
		GUICtrlSetData($ListTICInfo, "Date created: " & $Time[1] & '-' & $Time[2] & '-' & $Time[0] & ' ' & $Time[3] & ':' & $Time[4] & ':' & $Time[5])
		GUICtrlSetData($ListTICInfo, "TIC number:   " & GetTIC($XX))
	Else
		GUICtrlSetData($ListTICInfo, 'Unable to open: ' & $XX)
	EndIf
EndFunc   ;==>GetTICInfo
;-----------------------------------------------
Func GetBuildName()
	_Debug("GetBuildName")
	Local $Tmp1 = GUICtrlRead($ListBuilds)
	;Local $Tmp2 = StringMid(GUICtrlRead($ListTICInfo), 35)
	;$Tmp1 = StringRegExpReplace($Tmp1, "[\s,.]", "_")
	;$Tmp2 = StringRegExpReplace($Tmp2, "\s", "")
	GUICtrlSetData($InputTICName, $Tmp1)
EndFunc   ;==>GetBuildName
;-----------------------------------
;Reads the TIC file and returns the TIC value
; Sample: "\\elfin\media_tree\Kawela SW Reference\V1.0C0498\Disk\verfile.tic"
Func GetTIC($TICFile)
	Local $file = FileOpen($TICFile, 0)
	; Check if file opened for reading OK
	If $file = -1 Then
		_debug(@ScriptLineNumber & "  Unable to open file for reading: " & $TICFile, '', True)
		Return
	EndIf

	While 1
		Local $lineIn = FileReadLine($file)
		If @error = -1 Then ExitLoop
		If StringInStr($lineIn, ":  ") Then Local $tmp = $lineIn
	WEnd
	FileClose($file)
	_debug(@ScriptLineNumber & " " & $tmp)
	Return ($tmp)
EndFunc   ;==>GetTIC
;-----------------------------------------------
;This copies the appropriate files from source to destination
Func CopyTIC()
	Local $TopSourcePath
	If StringLen(GUICtrlRead($InputTICName)) < 3 Then ; verify that the TIC name is ok
		_Debug("TIC name is not valid: " & GUICtrlRead($InputTICName), '', True)
		Return
	EndIf

	_GuiDisable("Disable")

	Local $tmp = StringSplit(_GUICtrlListBox_GetText($ListTICInfo, _GUICtrlListBox_GetCaretIndex($ListTICInfo)), " ")

	$TopSourcePath = GUICtrlRead($ComboSource) & "\" & GUICtrlRead($InputTICName) & "\Disk\"
	_Debug(@ScriptLineNumber & "  " & $TopSourcePath)

	Local $TopDestinationPath = GUICtrlRead($ComboDestination) & "\" & GUICtrlRead($InputTICName) & '\'
	_Debug(@ScriptLineNumber & "  " & $TopDestinationPath)

	GUICtrlSetData($ListStatus, "TopSourcePath: " & $TopSourcePath)
	GUICtrlSetData($ListStatus, "TopDestinationPath: " & $TopDestinationPath)
	GUICtrlSetData($ListStatus, "Copying standard things")

	FileCopy($TopSourcePath & "verfile.tic", $TopDestinationPath, 9)

	If GUICtrlRead($CheckWin32) = $GUI_CHECKED Or GUICtrlRead($CheckWin64) = $GUI_CHECKED Or GUICtrlRead($CheckWin64e) = $GUI_CHECKED Then
		FileCopy($TopSourcePath & "autorun.*", $TopDestinationPath, 9)
		FileCopy($TopSourcePath & "AutoWire.bmp", $TopDestinationPath, 9)
		DirCopy($TopSourcePath & "APPS\SETUP", $TopDestinationPath & "APPS\SETUP", 1)
		DirCopy($TopSourcePath & "APPS\PROSETDX", $TopDestinationPath & "APPS\PROSETDX", 1)
	EndIf

	If GUICtrlRead($CheckTools) = $GUI_CHECKED Then
		If FileExists($TopSourcePath & "tools\") Then
			GUICtrlSetData($ListStatus, "Copying tools")
			DirCopy($TopSourcePath & "TOOLS\DOS", $TopDestinationPath & "TOOLS\DOS", 1)
			If GUICtrlRead($CheckWin32) = $GUI_CHECKED Then DirCopy($TopSourcePath & "TOOLS\WIN32", $TopDestinationPath & "TOOLS\WIN32", 1)
			If GUICtrlRead($CheckWin64) = $GUI_CHECKED Then DirCopy($TopSourcePath & "TOOLS\WIN64", $TopDestinationPath & "TOOLS\WIN64", 1)
			If GUICtrlRead($CheckWin64e) = $GUI_CHECKED Then DirCopy($TopSourcePath & "TOOLS\WINx64", $TopDestinationPath & "TOOLS\WINx64", 1)
			If GUICtrlRead($CheckLinux32) = $GUI_CHECKED Then DirCopy($TopSourcePath & "TOOLS\LINUX32", $TopDestinationPath & "TOOLS\LINUX32", 1)
			If GUICtrlRead($CheckLinux64) = $GUI_CHECKED Then DirCopy($TopSourcePath & "TOOLS\LINUX64", $TopDestinationPath & "TOOLS\LINUX64", 1)
			If GUICtrlRead($CheckLinux_x64) = $GUI_CHECKED Then DirCopy($TopSourcePath & "TOOLS\LINUX_x64", $TopDestinationPath & "TOOLS\LINUX_x64", 1)
			;EndIf
		Else
			GUICtrlSetData($ListStatus, $TopSourcePath & "Tools folder does not exist")
		EndIf
	EndIf

	If GUICtrlRead($CheckAPPS) = $GUI_CHECKED Then
		If FileExists($TopSourcePath & "APPS\") Then
			GUICtrlSetData($ListStatus, "Copying APPS")
			DirCopy($TopSourcePath & "APPS\DOS", $TopDestinationPath & "APPS\", 1)
		Else
			GUICtrlSetData($ListStatus, $TopSourcePath & "APPSfolder does not exist")
		EndIf
	EndIf


	If GUICtrlRead($CheckPro100) = $GUI_CHECKED Then
		If FileExists($TopSourcePath & "PRO100\") Then
			GUICtrlSetData($ListStatus, "Copying PRO100")
			If GUICtrlRead($CheckWin32) = $GUI_CHECKED Then DirCopy($TopSourcePath & "Pro100\Win32", $TopDestinationPath & "Pro100\Win32", 1)
			If GUICtrlRead($CheckWin64) = $GUI_CHECKED Then DirCopy($TopSourcePath & "Pro100\Win64", $TopDestinationPath & "Pro100\Win64", 1)
			If GUICtrlRead($CheckWin64e) = $GUI_CHECKED Then DirCopy($TopSourcePath & "Pro100\Winx64", $TopDestinationPath & "Pro100\Winx64", 1)
		Else
			GUICtrlSetData($ListStatus, $TopSourcePath & "PRO100 does not exist")
		EndIf
	EndIf

	If GUICtrlRead($CheckPro1000) = $GUI_CHECKED Then
		If FileExists($TopSourcePath & "PRO1000\") Then
			GUICtrlSetData($ListStatus, "Copying PRO1000")
			If GUICtrlRead($CheckWin32) = $GUI_CHECKED Then DirCopy($TopSourcePath & "Pro1000\Win32", $TopDestinationPath & "Pro1000\Win32", 1)
			If GUICtrlRead($CheckWin64) = $GUI_CHECKED Then DirCopy($TopSourcePath & "Pro1000\Win64", $TopDestinationPath & "Pro1000\Win64", 1)
			If GUICtrlRead($CheckWin64e) = $GUI_CHECKED Then DirCopy($TopSourcePath & "Pro1000\Winx64", $TopDestinationPath & "Pro1000\Winx64", 1)
			If GUICtrlRead($CheckLinux32) = $GUI_CHECKED Or GUICtrlRead($CheckLinux64) = $GUI_CHECKED Or GUICtrlRead($CheckLinux_x64) = $GUI_CHECKED Then
				DirCopy($TopSourcePath & "Pro1000\LINUX", $TopDestinationPath & "Pro1000\LINUX", 1)
			EndIf
		Else
			GUICtrlSetData($ListStatus, $TopSourcePath & "PRO1000 does not exist")
		EndIf
	EndIf

	If GUICtrlRead($CheckProxgb) = $GUI_CHECKED Then
		If FileExists($TopSourcePath & "PROXGB\") Then
			GUICtrlSetData($ListStatus, "Copying PROXGB")
			If GUICtrlRead($CheckWin32) = $GUI_CHECKED Then DirCopy($TopSourcePath & "PROXGB\Win32", $TopDestinationPath & "PROXGB\Win32", 1)
			If GUICtrlRead($CheckWin64) = $GUI_CHECKED Then DirCopy($TopSourcePath & "PROXGB\Win64", $TopDestinationPath & "PROXGB\Win64", 1)
			If GUICtrlRead($CheckWin64e) = $GUI_CHECKED Then DirCopy($TopSourcePath & "PROXGB\Winx64", $TopDestinationPath & "PROXGB\Winx64", 1)
			If GUICtrlRead($CheckLinux32) = $GUI_CHECKED Or GUICtrlRead($CheckLinux64) = $GUI_CHECKED Or GUICtrlRead($CheckLinux_x64) = $GUI_CHECKED Then
				DirCopy($TopSourcePath & "PROXGB\LINUX", $TopDestinationPath & "PROXGB\LINUX", 1)
			EndIf
		Else
			GUICtrlSetData($ListStatus, $TopSourcePath & "PROXGB does not exist")
		EndIf
	EndIf

	GUICtrlSetData($ListStatus, "Done copying")
	$tmp = DirGetSize($TopDestinationPath, 1)
	GUICtrlSetData($ListStatus, "Files: " & _commify($tmp[1]) & _
			" Dirs: " & _commify($tmp[2]) & _
			" Total size: " & _commify($tmp[0]))
	_GuiDisable("Enable")
EndFunc   ;==>CopyTIC
;-----------------------------------------------
Func Help()
	Local $helpstr = 'Command line startup options:'
	_Debug(@ScriptName & @CRLF & $FileVersion & @CRLF & $helpstr, '', True)
EndFunc   ;==>Help
;-----------------------------------------------

