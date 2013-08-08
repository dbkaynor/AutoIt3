#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_icon=../icons/Cockroach.ico
#AutoIt3Wrapper_outfile=\\chakotay\softval\iAMT\AT7.0\Build\BuildTool.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Comment=A Build tool for Intel MPG
#AutoIt3Wrapper_Res_Description=Build tool
#AutoIt3Wrapper_Res_Fileversion=0.0.0.82
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=Y
#AutoIt3Wrapper_Res_ProductVersion=666
#AutoIt3Wrapper_Res_LegalCopyright=Copyright 2010 Douglas B Kaynor
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Field=AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Res_Field=Compile date|%longdate% %time%
#AutoIt3Wrapper_Res_Field=Developer|Douglas Kaynor
#AutoIt3Wrapper_Res_Field=Email|doug@kaynor.net
#AutoIt3Wrapper_Res_Field=Made By|Douglas Kaynor
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=y
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Run_Tidy=y
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs
	Added clean up of unneed files and folders
	Corrected file naming to handle periods
	Enhanced logging
	Fixed file rename to overwrite the destination if needed and to create destination folder if needed
	
	Allow for the command line to be edited
#ce

Opt("MustDeclareVars", 1)

#include <Array.au3>
#include <Date.au3>
#include <Misc.au3>
#include <String.au3>
#include <GUIConstants.au3>
#include <GuiComboBox.au3>
#include <GuiComboBoxEx.au3>
#include <GuiListView.au3>
#include <GUIConstantsEx.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include "_DougFunctions.au3"

_Debug("DBGVIEWCLEAR")

TraySetIcon("../icons/Cockroach.ico")
HotKeySet("{F11}", "GUI_Enable")
Func GUI_Enable()
	GuiDisable($GUI_ENABLE)
EndFunc   ;==>GUI_Enable

Global $SourceLocation = @ScriptDir
Global $BuildLocation = "\\chakotay\softval\iAMT\AT7.0\Build"
Global Const $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")
Global Const $tmp = StringSplit(@ScriptName, ".")
Global Const $ProgramName = $tmp[1]

Global Const $SystemS = $ProgramName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & @OSServicePack & @CRLF & @OSType & @CRLF & @OSArch
Global $DEBUG = False

If _Singleton($ProgramName, 1) = 0 Then
	_Debug($ProgramName & " is already running!", 0x40, 5)
	Exit
EndIf

For $X = 1 To $CMDLine[0]
	ConsoleWrite($X & " >> " & $CMDLine[$X] & @CRLF)
	Select
		Case StringInStr($CMDLine[$X], "help") > 0 Or _
				StringInStr($CMDLine[$X], "?") > 0
			MsgBox(16, "Help", "No help yet")
			Exit
		Case StringInStr($CMDLine[$X], "sourcepath") > 0
			Global $Y = StringSplit($CMDLine[$X], "=")
			$SourceLocation = FileGetShortName(StringStripWS($Y[2], 3))
		Case StringInStr($CMDLine[$X], "debug") > 0
			$DEBUG = True
		Case Else
			_Debug("Unknown cmdline option found: >>" & $CMDLine[$X] & "<<", True)
			Exit
	EndSelect
Next

$SourceLocation = _AddSlash2PathString($SourceLocation)
Global $XMLLocation = $SourceLocation & "XML\"
Global $LOGfilename = $SourceLocation & $ProgramName & ".log"
Global $FITC = $SourceLocation & "fitc.exe"
Global $FITCLOG = $SourceLocation & "fitc.log"
Global $OutputFolderName

LogDebug("Command line arguments: " & $CmdLineRaw)

; Main form
Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $MainForm = GUICreate($ProgramName & $FileVersion, 650, 430, 10, 10, $MainFormOptions)
GUISetFont(10, 400, 0, "Courier New")

;left, top [, width [, height

Global $LabelXML = GUICtrlCreateLabel("XML", 10, 10, 40, 25, $SS_SUNKEN)
GUICtrlSetResizing(-1, 802)
Global $ComboXML = GUICtrlCreateCombo("", 55, 10, 395, 40)
GUICtrlSetTip(-1, "This is the list of available XML files")
GUICtrlSetResizing(-1, 546)
Global $LabelBIOSN = GUICtrlCreateLabel("BIOS", 10, 40, 40, 25, $SS_SUNKEN)
GUICtrlSetResizing(-1, 802)
GUICtrlSetTip($LabelBIOSN, GUICtrlRead($LabelBIOSN))
Global $LabelBIOS = GUICtrlCreateLabel("", 55, 40, 395, 40, $SS_SUNKEN)

GUICtrlSetResizing(-1, 546)
Global $LabelMEN = GUICtrlCreateLabel("ME", 10, 70, 40, 25, $SS_SUNKEN)
GUICtrlSetResizing(-1, 802)
Global $LabelME = GUICtrlCreateLabel("", 55, 70, 395, 40, $SS_SUNKEN)
GUICtrlSetTip(-1, "This is the ME file")
GUICtrlSetResizing(-1, 546)
Global $LabelGBEN = GUICtrlCreateLabel("GBE", 10, 100, 40, 25, $SS_SUNKEN)
GUICtrlSetResizing(-1, 802)
Global $LabelGBE = GUICtrlCreateLabel("", 55, 100, 395, 40, $SS_SUNKEN)
GUICtrlSetTip(-1, "This is the GBE file")
GUICtrlSetResizing(-1, 546)
Global $LabelOEMN = GUICtrlCreateLabel("OEM", 10, 130, 40, 25, $SS_SUNKEN)
GUICtrlSetResizing(-1, 802)
Global $LabelOEM = GUICtrlCreateLabel("", 55, 130, 395, 40, $SS_SUNKEN)
GUICtrlSetTip(-1, "This is the OEM file")
GUICtrlSetResizing(-1, 546)
Global $LabelPDRN = GUICtrlCreateLabel("PDR", 10, 160, 40, 25, $SS_SUNKEN)
GUICtrlSetResizing(-1, 802)
Global $LabelPDR = GUICtrlCreateLabel("", 55, 160, 395, 40, $SS_SUNKEN)
GUICtrlSetTip(-1, "This is the PDR file")
GUICtrlSetResizing(-1, 546)

Global $LabelOutName = GUICtrlCreateLabel("Out", 10, 190, 40, 25, $SS_SUNKEN)
GUICtrlSetResizing(-1, 802)
Global $InputOutName = GUICtrlCreateInput("", 55, 190, 395, 25, $SS_SUNKEN)
GUICtrlSetResizing(-1, 546)

Global $LabelCommand = GUICtrlCreateLabel("CMD", 10, 230, 40, 25, $SS_SUNKEN)
GUICtrlSetResizing(-1, 802)
Global $EditCommand = GUICtrlCreateEdit("", 55, 230, 400, 50, $WS_HSCROLL)
;Global $InputCommand = GUICtrlCreateInput("", 55, 220, 395, 25, $SS_SUNKEN)
GUICtrlSetResizing(-1, 546)

Global $ButtonCreateOutputName = GUICtrlCreateButton("Create output name", 455, 190, 190, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Create output name")
GUICtrlSetResizing(-1, 800)

Global $BIOSLabel = GUICtrlCreateLabel("BIOS Version", 10, 300, 100, 25, $SS_SUNKEN)
GUICtrlSetResizing(-1, 546)
Global $BIOSInput = GUICtrlCreateInput("????", 120, 300, 300, 25, $WS_GROUP)
GUICtrlSetTip(-1, "IN1")
GUICtrlSetResizing(-1, 800)

Global $MEKitLabel = GUICtrlCreateLabel("ME Kit", 10, 340, 100, 25, $SS_SUNKEN)
GUICtrlSetResizing(-1, 546)
Global $MEKitInput = GUICtrlCreateInput("????", 120, 340, 300, 25, $WS_GROUP)
GUICtrlSetTip(-1, "IN2")
GUICtrlSetResizing(-1, 800)

Global $LabelStatus1 = GUICtrlCreateLabel("Status", 10, 390, 65, 25, $SS_SUNKEN)
GUICtrlSetResizing(-1, 802)
Global $LabelStatus = GUICtrlCreateLabel("", 80, 390, 555, 25, $SS_SUNKEN)
GUICtrlSetTip(-1, "This displays the build result status")
GUICtrlSetResizing(-1, 546)
Global $ButtonBuild = GUICtrlCreateButton("Build", 455, 10, 90, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Build using the selected XML file")
GUICtrlSetResizing(-1, 800)
Global $ButtonBuildAll = GUICtrlCreateButton("Build all", 455, 40, 90, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Build using all XML files")
GUICtrlSetResizing(-1, 800)

Global $ButtonLoad = GUICtrlCreateButton("Load", 455, 70, 90, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Load combo boxes")
GUICtrlSetResizing(-1, 800)

Global $ButtonExploreSRC = GUICtrlCreateButton("BSD", 455, 300, 90, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Open explorer in the build source folder")
GUICtrlSetResizing(-1, 800)
Global $ButtonExploreBuild = GUICtrlCreateButton("BRS", 455, 340, 90, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Open explorer in the build run folder")
GUICtrlSetResizing(-1, 800)

Global $ButtonAbout = GUICtrlCreateButton("About", 555, 10, 90, 25, $WS_GROUP)
GUICtrlSetTip(-1, "About the program and some debug info")
GUICtrlSetResizing(-1, 800)
Global $ButtonHelp = GUICtrlCreateButton("Help", 555, 40, 90, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Help for the program")
GUICtrlSetResizing(-1, 800)
Global $ButtonExit = GUICtrlCreateButton("Exit", 555, 70, 90, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Exit the program")
GUICtrlSetResizing(-1, 800)
Global $ButtonViewEdit = GUICtrlCreateButton("View/Edit", 555, 100, 90, 25, $WS_GROUP)
GUICtrlSetTip(-1, "View or edit files in the build folder")
GUICtrlSetResizing(-1, 800)

GUISetState(@SW_SHOW)
DirCreate(@ScriptDir & "\AUXFiles")
LogDebug(_DateTimeFormat(_NowCalc(), 0) & " ------------ Start " & $ProgramName & " ------------")

GetXML()
SearchXML(GUICtrlRead($ComboXML))

GUICtrlSetData($EditCommand, " /b /o ")

FileDelete($FITCLOG)

If $DEBUG = True Then
	GUICtrlSetData($MEKitInput, "123.456")
	GUICtrlSetData($BIOSInput, "987.654")
EndIf

While 1
	Global $nMsg = GUIGetMsg(1)
	Switch $nMsg[0]
		Case $ComboXML
			SearchXML(GUICtrlRead($ComboXML))
		Case $ButtonBuild
			Build("")
		Case $ButtonBuildAll
			BuildAll()
		Case $ButtonLoad
			SearchXML(GUICtrlRead($ComboXML))
		Case $InputOutName

		Case $ButtonCreateOutputName
			CreateOutputName("")
		Case $LabelStatus

		Case $ButtonAbout
			GuiDisable($GUI_DISABLE)
			About($MainForm)
			GuiDisable($GUI_ENABLE)
		Case $ButtonViewEdit
			ViewEdit()
		Case $ButtonExploreSRC
			GotoSourceLocation()
		Case $ButtonExploreBuild
			GotoBuildLocation()
		Case $ButtonHelp
			GuiDisable($GUI_DISABLE)
			Help() ;MsgBox(0, "Help " & @ScriptLineNumber, "Someday maybe")
			GuiDisable($GUI_ENABLE)
		Case $ButtonExit
			Exit
		Case $GUI_EVENT_CLOSE
			Exit
	EndSwitch
WEnd
;-----------------------------------------------
Func GotoBuildLocation()
	LogDebug("GotoBuildLocation")
	Local $OpenCmd = "explorer " & $BuildLocation
	LogDebug($OpenCmd & " " & "RunWait results: " & RunWait($OpenCmd))
EndFunc   ;==>GotoBuildLocation
;-----------------------------------------------
Func GotoSourceLocation()
	LogDebug("GotoSourceLocation")
	Local $OpenCmd = "explorer " & $SourceLocation
	LogDebug($OpenCmd & " " & "RunWait results: " & RunWait($OpenCmd))
EndFunc   ;==>GotoSourceLocation
;-----------------------------------------------
Func GetXML()
	LogDebug("GetXML")
	_GUICtrlComboBoxEx_ResetContent($ComboXML)
	Local $search = FileFindFirstFile($XMLLocation & "*.xml")

	; Check if the search was successful
	If $search = -1 Then
		MsgBox(48, "Warning GetXML", "No files/directories matched the search pattern (XML's)")
		LogDebug("Warning GetXML: No files/directories matched the search pattern (XML's)")
		GuiDisable($GUI_ENABLE)
	EndIf

	While 1
		Local $file = FileFindNextFile($search)
		If @error Then ExitLoop
		GUICtrlSetData($ComboXML, $file, $file)
	WEnd

	FileClose($search)
EndFunc   ;==>GetXML
;-----------------------------------------------
Func CreateOutputName($XMLfile)

	If $XMLfile = "" Then $XMLfile = GUICtrlRead($ComboXML)
	LogDebug("CreateOutputName:  " & $XMLfile)
	Local $T = StringSplit($XMLfile, ".xml", 3)
	Local $U = GUICtrlRead($MEKitInput)
	Local $V = GUICtrlRead($BIOSInput)

	If StringInStr($V, "?") Or $V = "" Then
		MsgBox(16, "BIOS version error", "BIOS version " & $V & " is not valid")
		GUICtrlSetData($LabelStatus, "BIOS version " & $V & " is not valid")
		LogDebug("BIOS version is not valid")
		Return -2
	EndIf

	If StringInStr($U, "?") Or $U = "" Then
		MsgBox(16, "ME Kit name error", "ME Kit name " & $U & " is not valid")
		GUICtrlSetData($LabelStatus, "ME Kit name " & $U & " is not valid")
		LogDebug("ME Kit name " & $U & " is not valid")
		Return -1
	EndIf

	Local $name = $T[0] & "_BIOS-" & $V & "_ME-" & $U
	GUICtrlSetData($InputOutName, "output\" & $name & "\" & $name)
	$OutputFolderName = "output\" & $name
	GUICtrlSetData($LabelStatus, "CreateOutputName complete " & $XMLfile)
	Return 0
EndFunc   ;==>CreateOutputName
;-----------------------------------------------
Func Build($XMLfile)
	GuiDisable($GUI_DISABLE)
	FileChangeDir($SourceLocation)

	If CreateOutputName($XMLfile) <> 0 Then
		GuiDisable($GUI_ENABLE)
		Return
	EndIf

	Local $CMDLine = $FITC & " " & GUICtrlRead($EditCommand) & GUICtrlRead($InputOutName) & "  " & $XMLLocation & GUICtrlRead($ComboXML)
	GUICtrlSetData($LabelStatus, "Working: " & $CMDLine)
	LogDebug("Command: " & $CMDLine)
	Local $results = "RunWait results: " & RunWait($CMDLine)
	LogDebug("Results:  " & $results)

	;rename the resulting files here
	If StringCompare($results, "RunWait results: 0") = 0 Then
		Local $FilesIn[1]

		If Not DirRemove($OutputFolderName & "\Int", 1) Then
			MsgBox(16, "DirRemove error", $OutputFolderName & "\Int", 5)
			LogDebug("DirRemove error: " & $OutputFolderName & "\Int")
		Else
			LogDebug("DirRemove: " & $OutputFolderName & "\Int")
		EndIf

		$FilesIn = _FileListToArrayR($OutputFolderName, "*", 0, 0, 1, "*Int", 0)
		LogDebug("OutputFolderName: " & $OutputFolderName)

		If IsArray($FilesIn) Then
			_ArrayDelete($FilesIn, 0)
			Local $FilesOut = $FilesIn
			Local $V = GUICtrlRead($MEKitInput)

			For $F = 0 To UBound($FilesIn) - 1
				If StringInStr($FilesOut[$F], "(1)") > 0 Then
					$FilesOut[$F] = StringReplace($FilesOut[$F], "(1)", "") & "(SPI0).bin"
				EndIf
				If StringInStr($FilesOut[$F], "(2)") > 0 Then
					$FilesOut[$F] = StringReplace($FilesOut[$F], "(2)", "") & "(SPI2).bin"
				EndIf

				If StringLen($FilesOut[$F]) - StringLen($V) + 1 = StringInStr($FilesOut[$F], $V, 0, -1) Then
					$FilesOut[$F] = $FilesOut[$F] & "(SPI Full).bin"
				EndIf
				If StringInStr($FilesOut[$F], ".map") > 5 Then
					FileDelete($FilesOut[$F])
					LogDebug("FileDelete: " & $FilesOut[$F])
				EndIf
			Next

			For $F = 0 To UBound($FilesIn) - 1
				Local $Result = FileMove($FilesIn[$F], $FilesOut[$F], 9)
				If $Result = 1 Then
					LogDebug("Filemove: " & $FilesIn[$F] & "  FileMove passed")
				Else
					LogDebug("Filemove: " & $FilesIn[$F] & "  FileMove failed")
				EndIf
			Next
		EndIf

	EndIf
	GUICtrlSetData($LabelStatus, $results)
	GuiDisable($GUI_ENABLE)
EndFunc   ;==>Build
;-----------------------------------------------
Func BuildAll()
	If CreateOutputName("") <> 0 Then
		GuiDisable($GUI_ENABLE)
		Return
	EndIf

	Local $BuildAllArray = _GUICtrlComboBox_GetListArray($ComboXML)
	For $XMLfile In $BuildAllArray
		If StringInStr($XMLfile, ".xml") > 0 Then
			LogDebug("BuildAll: " & $XMLfile)
			Build($XMLfile)
		EndIf
	Next

EndFunc   ;==>BuildAll
;-----------------------------------------------
Func SearchXML($XMLfile)
	LogDebug("SearchXML: " & $XMLfile)
	GuiDisable($GUI_DISABLE)

	Local $file = FileOpen($XMLLocation & $XMLfile, 0)
	LogDebug("XMLLocation & XMLfile: " & $XMLLocation & $XMLfile)
	; Check if file opened for reading OK
	If $file = -1 Then
		LogDebug(" Unable to open file for reading: " & $XMLLocation & $XMLfile)
		GuiDisable($GUI_ENABLE)
		Return
	EndIf
	Local $ArraySTR = FileRead($file)
	FileClose($file)
	Local $FF = StringSplit($ArraySTR, "<")

	;_ArrayDisplay($FF, @ScriptLineNumber)
	_ArrayDelete($FF, 0)
	_ArrayDelete($FF, 0)
	For $i = 0 To UBound($FF) - 1
		If StringLen($FF[$i]) > 1 Then $FF[$i] = "<" & $FF[$i]
	Next

	;_ArrayDisplay($FF, @ScriptLineNumber)
	;_FileWriteFromArray("parsed.xml", $FF)

	For $T In $FF
		If StringInStr($T, "This is the BIOS image binary") <> 0 Then
			$T = StringSplit($T, 'InputFile value="', 1)
			$T = StringSplit($T[UBound($T) - 1], '" ', 1)
			If FileExists($T[1]) Then
				GUICtrlSetData($LabelBIOS, $T[1])
			Else
				GUICtrlSetData($LabelBIOS, "WARNING: " & $T[1])
				LogDebug("BIOS File not accessable: " & $T[1])
			EndIf
		EndIf

		If StringInStr($T, "This is the Gbe image binary") <> 0 Then
			$T = StringSplit($T, 'InputFile value="', 1)
			$T = StringSplit($T[UBound($T) - 1], '" ', 1)
			If FileExists($T[1]) Then
				GUICtrlSetData($LabelGBE, $T[1])
			Else
				GUICtrlSetData($LabelGBE, "WARNING: " & $T[1])
				LogDebug("Gbe File not accessable: " & $T[1])
			EndIf
		EndIf

		If StringInStr($T, "used for the ME region") <> 0 Then
			$T = StringSplit($T, 'InputFile value="', 1)
			$T = StringSplit($T[UBound($T) - 1], '" ', 1)
			If FileExists($T[1]) Then
				GUICtrlSetData($LabelME, $T[1])
			Else
				GUICtrlSetData($LabelME, "WARNING: " & $T[1])
				LogDebug("ME File not accessable: " & $T[1])
			EndIf
		EndIf

		If StringInStr($T, "are copied directly into the OEM section") <> 0 Then
			$T = StringSplit($T, 'InputFile value="', 1)
			$T = StringSplit($T[UBound($T) - 1], '" ', 1)
			If FileExists($T[1]) Then
				GUICtrlSetData($LabelOEM, $T[1])
			Else
				GUICtrlSetData($LabelOEM, "WARNING: " & $T[1])
				LogDebug("OEM File not accessable: " & $T[1])
			EndIf
		EndIf

		If StringInStr($T, "This is the PDR image binary") <> 0 Then
			$T = StringSplit($T, 'InputFile value="', 1)
			$T = StringSplit($T[UBound($T) - 1], '" ', 1)
			If FileExists($T[1]) Then
				GUICtrlSetData($LabelPDR, $T[1])
			Else
				GUICtrlSetData($LabelPDR, "WARNING: " & $T[1])
				LogDebug("PDR File not accessable: " & $T[1])
			EndIf
		EndIf

		If StringInStr($T, "BuildOutputFilename value=") <> 0 Then
			$T = StringSplit($T, 'BuildOutputFilename value="', 1)
			$T = StringSplit($T[UBound($T) - 1], '"/>', 1)

			GUICtrlSetData($InputOutName, $T[1])
		EndIf
	Next
	GuiDisable($GUI_ENABLE)
EndFunc   ;==>SearchXML
;-----------------------------------------------

Func About(Const $FormID)
	GuiDisable($GUI_DISABLE)
	Local $D = WinGetPos($FormID)
	Local $WinPos
	If IsArray($D) = True Then
		ConsoleWrite(@ScriptLineNumber & $FormID & @CRLF)
		$WinPos = StringFormat("WinPOS: %d  %d " & @CRLF & "WinSize: %d %d " & @CRLF & "Desktop: %d %d", _
				$D[0], $D[1], $D[2], $D[3], @DesktopWidth, @DesktopHeight)
	Else
		$WinPos = ">>>About ERROR, Check the window name<<<"
	EndIf
	_Debug(@CRLF & $SystemS & @CRLF & $WinPos & @CRLF & "Written by Doug Kaynor!", 0x40)
	GuiDisable($GUI_ENABLE)
EndFunc   ;==>About
;-----------------------------------------------
Func GuiDisable($choice) ;@SW_ENABLE @SW_disble
	GUICtrlSetState($LabelXML, $choice)
	GUICtrlSetState($LabelBIOS, $choice)
	GUICtrlSetState($LabelME, $choice)
	GUICtrlSetState($LabelGBE, $choice)
	GUICtrlSetState($LabelOEM, $choice)
	GUICtrlSetState($LabelPDR, $choice)
	GUICtrlSetState($InputOutName, $choice)
	GUICtrlSetState($EditCommand, $choice)
	GUICtrlSetState($LabelStatus, $choice)

	GUICtrlSetState($ButtonBuild, $choice)
	GUICtrlSetState($ButtonBuildAll, $choice)
	GUICtrlSetState($ButtonLoad, $choice)
	GUICtrlSetState($ButtonExploreSRC, $choice)
	GUICtrlSetState($ButtonExploreBuild, $choice)
	GUICtrlSetState($ButtonAbout, $choice)
	GUICtrlSetState($ButtonHelp, $choice)
	GUICtrlSetState($ButtonViewEdit, $choice)
	GUICtrlSetState($ButtonExit, $choice)
	GUICtrlSetState($ButtonCreateOutputName, $choice)
	GUICtrlSetState($InputOutName, $choice)
	GUICtrlSetState($MEKitInput, $choice)
	GUICtrlSetState($BIOSInput, $choice)
EndFunc   ;==>GuiDisable
;-----------------------------------------------
;This function allows the user to edit or view any file, useful for changing the config file
Func ViewEdit()
	LogDebug("ViewEdit")
	Local $Filename = FileOpenDialog("View or Edit a file", $SourceLocation, "Log (*.log)|All (*.*)", 1)
	ConsoleWrite(@ScriptLineNumber & " +++ " & $Filename & @CRLF)
	Const $edit1 = "c:\program files\notepad++\notepad++.exe"
	Const $edit2 = "c:\program files (x86)\notepad++\notepad++.exe"
	Const $edit3 = "notepad.exe"
	Local $editor = ""

	If FileExists($edit1) = 1 Then
		$editor = $edit1
	ElseIf FileExists($edit2) = 1 Then
		$editor = $edit2
	Else
		$editor = $edit3
	EndIf
	ShellExecute($editor, $Filename)
EndFunc   ;==>ViewEdit
;-----------------------------------------------
Func Help()
	_Debug("Help")
	Local $helpstr = 'Startup options: ' & @CRLF & _
			"help or ?     Display this help file" & @CRLF & _
			"sourcepath    Specify a source folder" & @CRLF & _
			"    The sourcepath must be wrapped in quotes" & @CRLF & _
			"Debug         Secret debug things"
	_Debug(@ScriptName & @CRLF & $FileVersion & @CRLF & @CRLF & $helpstr, 0x40)

EndFunc   ;==>Help
;-----------------------------------------------
Func LogDebug($STR)
	_Debug($STR)
	FileWriteLine($LOGfilename, $STR)
EndFunc   ;==>LogDebug
;-----------------------------------------------