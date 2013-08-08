#region
#AutoIt3Wrapper_Run_Au3check=y
#AutoIt3Wrapper_Au3Check_Stop_OnWarning=y
#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Tidy_Stop_OnError=y
#AutoIt3Wrapper_Tidy_Stop_OnError=y
;#Tidy_Parameters=/gd /sf
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Fileversion=0.0.0.21
#AutoIt3Wrapper_Res_FileVersion_AutoIncrement=Y
#AutoIt3Wrapper_Res_Description=Pingtool
#AutoIt3Wrapper_Res_LegalCopyright=GNU-PL
#AutoIt3Wrapper_Res_Comment=Web Grabber
#AutoIt3Wrapper_Res_Field=Developer|Douglas Kaynor
#AutoIt3Wrapper_Res_LegalCopyright=Copyright ? 2010 Douglas B Kaynor
#AutoIt3Wrapper_Res_Field= AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Res_Field=Compile date|%longdate% %time%
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Run_Before=
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Icon=./icons/freebsd.ico
#endregion

Opt("MustDeclareVars", 1)

If _Singleton(@ScriptName, 1) = 0 Then
	_Debug(@ScriptName & " is already running!", 0x40, 5)
	Exit
EndIf

#include <Array.au3>
#include <ButtonConstants.au3>
#include <Date.au3>
#include <EditConstants.au3>
#include <GUIConstants.au3>
#include <GuiListView.au3>
#include <GUIConstantsEx.au3>
#include <GuiTreeView.au3>
#include <GuiComboBox.au3>
#include <iNet.au3>
#include <IE.au3>

#include <ListViewConstants.au3>
#include <Misc.au3>
#include <String.au3>
#include <StaticConstants.au3>
#include <TreeViewConstants.au3>
#include <WindowsConstants.au3>
#include "_DougFunctions.au3"

TraySetIcon("./icons/freebsd.ico")

Global $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")

Global $tmp = StringSplit(@ScriptName, ".")

DirCreate(@ScriptDir & "\AUXFiles")
DirCreate(@ScriptDir & "\AUXFiles\WebGrabber\")
Global $AUXFiles = @ScriptDir & "\AUXFiles\"
Global $WebGrabber = @ScriptDir & "\AUXFiles\WebGrabber\"

Global $FireFoxPath
Const $FireFoxPath1 = "C:\Program Files (x86)\Mozilla Firefox 4.0 Beta 9\firefox.exe"
Const $FireFoxPath2 = "C:\Program Files\Mozilla Firefox 4.0 Beta 9\firefox.exe"
Const $FireFoxPath3 = "J:\PortableApps\FirefoxPortable\FirefoxPortable.exe"
If FileExists($FireFoxPath1) Then
	$FireFoxPath = $FireFoxPath1
ElseIf FileExists($FireFoxPath2) Then
	$FireFoxPath = $FireFoxPath2
Else
	$FireFoxPath = $FireFoxPath3
EndIf

If Not FileExists($FireFoxPath) Then
	If MsgBox(16 + 1 + 256, "Firefox error", "Unable to locate Firefox on this computer") = 2 Then Exit
EndIf

Global $Editor
Const $Edit1 = "c:\program files (x86)\notepad++\notepad++.exe"
Const $Edit2 = "c:\program files\notepad++\notepad++.exe"
Const $Edit3 = "notepad.exe"
If FileExists($Edit1) = 1 Then
	$Editor = $Edit1
ElseIf FileExists($Edit2) = 1 Then
	$Editor = $Edit2
Else
	$Editor = $Edit3
EndIf

If Not FileExists($Editor) Then
	If MsgBox(16 + 1 + 256, "Editor error", "Unable to locate Notepad or Notepad++ on this computer") = 2 Then Exit
EndIf


Global $WebGrabberPrj = $AUXFiles & $tmp[1] & ".prj"
ConsoleWrite(@ScriptLineNumber & " " & $WebGrabberPrj & @CRLF)
DirCreate($WebGrabber & "Results")
Global $ResultsLocation = $WebGrabber & "Results"
Global $SavedSourceFull = $WebGrabber & "SavedSourceFull.txt"
Global $SavedSourceParsed = $WebGrabber & "SavedSourceParsed.txt"
Global $PageSaveName
Global $ExternalLinks[1]
Global $InternalLinks[1]

Global $SystemS = @ScriptName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & @OSServicePack & @CRLF & @OSType & @CRLF & @OSArch & @IPAddress1
Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)

Global $MainForm = GUICreate(@ScriptName & " " & $FileVersion, 540, 401, 10, 10, $MainFormOptions)
GUISetFont(10, 400, 0, "Courier New")
Global $ButtonBegin = GUICtrlCreateButton("Begin", 10, 10, 70, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Begin getting the files")


;Global $ButtonFetch = GUICtrlCreateButton("Fetch", 10, 30, 70, 20)
;GUICtrlSetResizing(-1, 802)
;GUICtrlSetTip(-1, "Fetch")

Global $CheckStop = GUICtrlCreateCheckbox("Stop", 10, 85, 60, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Stop getting files")

Global $ButtonShowExternalLinks = GUICtrlCreateButton("ShowExt", 100, 10, 90, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Show External Links")
Global $LabelNumberExternallinks = GUICtrlCreateLabel("", 200, 10, 50, 20, $SS_SUNKEN)
GUICtrlSetResizing(-1, 802)

Global $ButtonShowInternalLinks = GUICtrlCreateButton("ShowInt", 100, 30, 90, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Show External Links")
Global $LabelNumberInternalLinks = GUICtrlCreateLabel("", 200, 30, 50, 20, $SS_SUNKEN)
GUICtrlSetResizing(-1, 802)


Global $LabelNumberOfLinks = GUICtrlCreateLabel("Number of Links", 100, 50, 150, 20, $SS_SUNKEN)
GUICtrlSetResizing(-1, 802)


Global $ButtonSaveProject = GUICtrlCreateButton("Save project", 260, 10, 120, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Save the current settings")

Global $ButtonLoadProject = GUICtrlCreateButton("Load project", 260, 30, 120, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Load saved settings")

Global $ButtonSetDefaults = GUICtrlCreateButton("Set defaults", 260, 50, 120, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Set default settings")

Global $ButtonEdit = GUICtrlCreateButton("Edit", 260, 70, 120, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Edit or view a file")

Global $ButtonAbout = GUICtrlCreateButton("About", 380, 10, 60, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "About button")
Global $ButtonHelp = GUICtrlCreateButton("Help", 380, 30, 60, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Help button")
Global $ButtonExit = GUICtrlCreateButton("Exit", 380, 50, 60, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Exit button")

Global $ButtonT1 = GUICtrlCreateButton("T1", 450, 10, 60, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Disable")
Global $ButtonT2 = GUICtrlCreateButton("T2", 450, 30, 60, 20)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip(-1, "Enable")
; www.zitscomics.com/home?destination=node%2F1
Global $ComboURL = GUICtrlCreateCombo("http://www.zitscomics.com/", 10, 126, 520, 24)
GUICtrlSetResizing(-1, 2 + 32 + 512)
GUICtrlSetTip(-1, "Input the top level URL to work on")

Global $InputResult = GUICtrlCreateInput("InputResult", 10, 156, 520, 24, $ES_READONLY)
GUICtrlSetResizing(-1, 2 + 32 + 512)
GUICtrlSetTip(-1, "The complete URL to fetch")


Global $MyList = GUICtrlCreateTreeView(10, 186, 520, 200, BitOR($TVS_HASBUTTONS, $TVS_HASLINES, $TVS_LINESATROOT, $TVS_DISABLEDRAGDROP, $TVS_SHOWSELALWAYS, $TVS_INFOTIP, $WS_GROUP, $WS_TABSTOP), $WS_EX_CLIENTEDGE)
GUICtrlSetResizing(-1, 2 + 32 + 64)
GUICtrlSetTip(-1, "This is the list of parsed data")

GUISetState(@SW_SHOW)

GUISetHelp("notepad", $MainForm) ; Need a help file to call here

AutoItSetOption("TrayIconDebug", 1)

_Debug("DBGVIEWCLEAR")

SetDefaults()

LoadProject("start")

GUISetState(@SW_SHOW)
;-----------------------------------------------
While 1
	Global $nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $ButtonExit
			Exit
		Case $ButtonHelp
			MsgBox(0, "Help", "Someday maybe", 5)
		Case $ButtonSaveProject
			SaveProject()
		Case $ButtonLoadProject
			LoadProject("Menu")
		Case $ButtonSetDefaults
			SetDefaults()
		Case $ButtonAbout
			About()
		Case $ButtonEdit
			Edit()
		Case $ButtonBegin
			Begin()
		Case $CheckStop
			_Debug(@ScriptLineNumber & "  CheckStop")
		Case $ButtonShowExternalLinks
			_ArrayDisplay($ExternalLinks, "External links")
		Case $ButtonShowInternalLinks
			_ArrayDisplay($InternalLinks, "Internal links")

		Case $ButtonT1
			Global $oIE = _IECreate("http://www.kaynor.net/music.php")
			Global $oForms = _IEFormGetCollection($oIE)
			Global $iNumForms = @extended
			MsgBox(0, "Forms Info", "There are " & $iNumForms & " forms on this page")
			For $i = 0 To $iNumForms - 1
				Global $oForm = _IEFormGetCollection($oIE, $i)
				MsgBox(0, "Form Info", $oForm.name)
			Next
		Case $ButtonT2
			GuiDisable("enable")
	EndSwitch
WEnd
;-----------------------------------------------
Func Begin()
	GuiDisable("disable")
	If IsArray($ExternalLinks) Then ReDim $ExternalLinks[1]
	If IsArray($InternalLinks) Then ReDim $InternalLinks[1]
	_GUICtrlTreeView_DeleteAll($MyList)
	GUICtrlSetData($LabelNumberOfLinks, "Number of Links")
	Local $URL = GUICtrlRead($ComboURL)
	GetURL($URL)
	GuiDisable("enable")
EndFunc   ;==>Begin
;-----------------------------------------------
Func Recurse()
	GuiDisable("disable")
	If IsArray($ExternalLinks) Then ReDim $ExternalLinks[1]
	If IsArray($InternalLinks) Then ReDim $InternalLinks[1]
	_GUICtrlTreeView_DeleteAll($MyList)
	GUICtrlSetData($LabelNumberOfLinks, "Number of Links")
	Local $URL = GUICtrlRead($ComboURL)

	GetURL($URL)


	GuiDisable("enable")
EndFunc   ;==>Recurse



;-----------------------------------------------
Func GetURL($URL)
	Local $TopURLa = StringSplit($URL, '/')
	Local $TopUrl = ''
	ConsoleWrite(@ScriptLineNumber & ": " & _ArrayToString($TopURLa) & @CRLF)
	If $TopURLa[0] >= 3 Then
		$TopUrl = $TopURLa[1] & "//" & $TopURLa[3]
		ConsoleWrite(@ScriptLineNumber & ": " & $TopUrl & @CRLF)
	EndIf

	Local $oIE = _IECreate($URL, 0, 0)
	If $oIE = 0 Then
		MsgBox(48, "_IECreate error", $URL & @CRLF & "Failed to create inet object." & @CRLF & @error)
		Return 1
	EndIf

	Local $oLinks = _IELinkGetCollection($oIE)
	Local $iNumLinks = @extended
	If $iNumLinks = 0 Then
		MsgBox(48, "_IELinkGetCollection", $URL & @CRLF & "Failed to get collection." & @CRLF & @error)
		Return 2
	EndIf

	GUICtrlSetData($LabelNumberOfLinks, "Number Links " & $iNumLinks)
	ConsoleWrite(@ScriptLineNumber & " " & $iNumLinks & @CRLF)
	ConsoleWrite(@ScriptLineNumber & " " & _IEPropertyGet($oIE, "locationname") & @CRLF)
	For $oLink In $oLinks
		ConsoleWrite(@ScriptLineNumber & " " & $oLink.href & " " & $TopUrl & " " & StringInStr($oLink.href, $TopUrl) & @CRLF)
		_GUICtrlTreeView_Add($MyList, 0, $oLink.href & @CRLF)
		If StringInStr($oLink.href, $TopUrl) > 0 Then
			_ArrayAdd($InternalLinks, $oLink.href)
			GUICtrlSetData($LabelNumberInternalLinks, UBound($InternalLinks) - 1)
		Else
			_ArrayAdd($ExternalLinks, $oLink.href)
			GUICtrlSetData($LabelNumberExternallinks, UBound($ExternalLinks) - 1)
		EndIf
	Next
	_IEQuit($oIE)
	Return 0
EndFunc   ;==>GetURL
;-----------------------------------------------
;This function displays about and debug information
Func About()
	Local $D = WinGetPos(@ScriptName)
	Local $WinPos
	If IsArray($D) = 1 Then
		$WinPos = StringFormat("%s" & @CRLF & "WinPOS: %d  %d " & @CRLF & "WinSize: %d %d " & @CRLF & "Desktop: %d %d ", _
				$MainForm, $D[0], $D[1], $D[2], $D[3], @DesktopWidth, @DesktopHeight)
	Else
		$WinPos = ">>>About ERROR, Check the window name<<<"
	EndIf
	_Debug(@CRLF & $SystemS & @CRLF & $WinPos & @CRLF & "Written by Doug Kaynor because I wanted too!", 0x40)
EndFunc   ;==>About
;-----------------------------------------------
;This function enables or disables the GUI componets
Func GuiDisable($choice) ;@SW_ENABLE @SW_disble
	_Debug(@ScriptLineNumber & "  GuiDisable  " & $choice)
	Global $LastState
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
		_Debug(@ScriptLineNumber & "  Invalid choice at GuiDisable" & $choice, 0x40)
	EndIf

	GUICtrlSetState($ButtonBegin, $setting)
	;GUICtrlSetState($ButtonFetch, $setting)
	GUICtrlSetState($ButtonLoadProject, $setting)
	GUICtrlSetState($ButtonSaveProject, $setting)
	GUICtrlSetState($ButtonSetDefaults, $setting)
	GUICtrlSetState($ButtonEdit, $setting)
	GUICtrlSetState($ButtonAbout, $setting)
	GUICtrlSetState($ButtonHelp, $setting)

	GUICtrlSetState($ButtonExit, $setting)

	GUICtrlSetState($ButtonShowExternalLinks, $setting)
	GUICtrlSetState($ButtonShowInternalLinks, $setting)
	GUICtrlSetState($ComboURL, $setting)

EndFunc   ;==>GuiDisable
;-----------------------------------------------
Func SetDefaults()
	GUICtrlSetState($CheckStop, $GUI_UNCHECKED)
	GUICtrlSetData($ComboURL, "http://localhost/KaynorNet/httpdocs/music.php")
EndFunc   ;==>SetDefaults
;-----------------------------------------------
;This function loads the project file
Func LoadProject($type)
	Local $Filename
	If StringCompare($type, "menu") = 0 Then
		_Debug(@ScriptLineNumber & "  LoadProject  " & $type & "  " & $WebGrabberPrj)
		$Filename = FileOpenDialog("Load project file", $WebGrabberPrj, _
				"WebGrabber projects (W*.prj)|All projects (*.prj)|All files (*.*)", 18, "WebGrabber.prj")
	Else
		$Filename = $WebGrabberPrj
	EndIf

	Local $file = FileOpen($Filename, 0)
	; Check if file opened for reading OK
	If $file = -1 Then
		_Debug(@ScriptLineNumber & "  LoadProject: Unable to open file for reading: " & $Filename, 0x10, 5)
		Return
	EndIf

	_Debug(@ScriptLineNumber & "  LoadProject  " & $type & "  " & $WebGrabberPrj)

	; Read in the first line to verify the file is of the correct type
	If StringCompare(FileReadLine($file, 1), "Valid for WebGrabber project") <> 0 Then
		_Debug(@ScriptLineNumber & "  Not a valid project file for WebGrabber", 0x20, 5)
		FileClose($file)
		Return
	EndIf

	; Read in lines of text until the EOF is reached
	While 1
		Local $LineIn = FileReadLine($file)
		If @error = -1 Then ExitLoop
		_Debug(@ScriptLineNumber & "  LoadProject   " & $LineIn)
		FileWriteLine($file, "Valid for WebGrabber project")

		If StringInStr($LineIn, "ComboURLAll:") Then
			_GUICtrlComboBox_ResetContent($ComboURL)
			GUICtrlSetData($ComboURL, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
		EndIf
		If StringInStr($LineIn, "ComboURL:") Then GUICtrlSetData($ComboURL, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
	WEnd

	FileClose($file)

EndFunc   ;==>LoadProject
;-----------------------------------------------
;This  function lsaves the project file
Func SaveProject()
	_Debug(@ScriptLineNumber & "   SaveProject")

	Local $Filename = FileSaveDialog("Save project file", $WebGrabberPrj, _
			"WebGrabber projects (W*.prj)|All projects (*.prj)|All files (*.*)", 18, "WebGrabber.prj")
	Local $file = FileOpen($Filename, 2)

	; Check if file opened for writing OK
	If $file = -1 Then
		_Debug(@ScriptLineNumber & "  SaveProject: Unable to open file for writing: " & $Filename, 0x10, 5)
		Return
	EndIf

	_Debug(@ScriptLineNumber & " SaveProject: " & $WebGrabber & " " & $Filename & @CRLF)

	FileWriteLine($file, "Valid for WebGrabber project")
	FileWriteLine($file, "ComboURLAll:" & _GUICtrlComboBox_GetList($ComboURL))
	FileWriteLine($file, "ComboURL:" & GUICtrlRead($ComboURL))

	FileClose($file)
EndFunc   ;==>SaveProject
;-----------------------------------------------
;This function allows the user to edit or view a text file
Func Edit()
	_Debug(@ScriptLineNumber & "  Edit")
	Local $Filename = FileOpenDialog("View a file", $WebGrabber, _
			"All (*.*)", 1, "WebGrabber.prj")
	ConsoleWrite(@ScriptLineNumber & $WebGrabber & @CRLF & $Filename & @CRLF)
	_Debug(@ScriptLineNumber & "  Edit  " & $Filename)
	ShellExecuteWait($Editor, $Filename)

EndFunc   ;==>Edit
;-----------------------------------------------






