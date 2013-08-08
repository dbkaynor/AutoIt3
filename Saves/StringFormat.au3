#Region
#AutoIt3Wrapper_Run_Au3check=y
#AutoIt3Wrapper_Au3Check_Stop_OnWarning=y
#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Tidy_Stop_OnError=y
;#Tidy_Parameters=/gd /sf
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Fileversion=0.0.0.0
#AutoIt3Wrapper_Res_FileVersion_AutoIncrement=Y
#AutoIt3Wrapper_Res_Description=StringFormat
#AutoIt3Wrapper_Res_LegalCopyright=GNU-PL
#AutoIt3Wrapper_Res_Comment=StringFormat
#AutoIt3Wrapper_Res_Field=Developer|Douglas Kaynor
#AutoIt3Wrapper_Res_LegalCopyright=Copyright 2010 Douglas B Kaynor
#AutoIt3Wrapper_Res_Field= AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Res_Field=Compile date|%longdate% %time%
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Run_Before=
#AutoIt3Wrapper_Icon=./icons/Cryptkeeper.ico
#EndRegion

TraySetIcon("./icons/Cryptkeeper.ico")

Opt("MustDeclareVars", 1)

#include <Array.au3>
#include <Date.au3>
#include <Misc.au3>
#include <String.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <SliderConstants.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>

#include "_DougFunctions.au3"

Global $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")

Global $tmp = StringSplit(@ScriptName, ".")
Global $ProgramName = $tmp[1]

If _Singleton($ProgramName, 1) = 0 Then
	Debug($ProgramName & " is already running!", 0x40, 5)
	Exit
EndIf

Global $DefaultWorkingFolder = "C:\Program Files (x86)\Apache Software Foundation\Apache2.2\htdocs\documents\"

Global $SystemS = @ScriptName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & @OSServicePack & @CRLF & @OSType & @CRLF & @OSArch & @IPAddress1
Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $MainForm = GUICreate($ProgramName & " " & $FileVersion, 500, 600, 10, 10, $MainFormOptions)

Global $ButtonOpenFile = GUICtrlCreateButton("Open file", 10, 10, 60, 30)
GUICtrlSetTip(-1, "Open a text file")
GUICtrlSetResizing(-1, 802)
Global $ButtonSaveFile = GUICtrlCreateButton("Save file", 84, 11, 60, 30)
GUICtrlSetTip(-1, "Save the text file")
GUICtrlSetResizing(-1, 802)
Global $ButtonProcess = GUICtrlCreateButton("Process", 157, 10, 60, 30)
GUICtrlSetTip(-1, "Process the file")
GUICtrlSetResizing(-1, 802)
Global $ButtonDefaults = GUICtrlCreateButton("Set defaults", 230, 10, 60, 30)
GUICtrlSetTip(-1, "Set any defaults")
GUICtrlSetResizing(-1, 802)
Global $ButtonAbout = GUICtrlCreateButton("About", 300, 10, 60, 30)
GUICtrlSetTip($ButtonAbout, "About button")
GUICtrlSetResizing(-1, 802)
Global $ButtonExit = GUICtrlCreateButton("Exit", 370, 10, 60, 30)
GUICtrlSetTip(-1, "Exit the program")
GUICtrlSetResizing(-1, 802)
Global $LabelStringLength = GUICtrlCreateLabel("String length", 10, 68, 60, 20)
GUICtrlSetResizing(-1, 802)
Global $SliderStringLength = GUICtrlCreateSlider(80, 60, 128, 36)
GUICtrlSetLimit(-1, 110, 10)
GUICtrlSetData(-1, 10)
GUICtrlSetTip(-1, "Set the target line length")
GUICtrlSetResizing(-1, 802)
Global $LabelStringLengthResult = GUICtrlCreateLabel("666", 226, 68, 52, 20)
GUICtrlSetResizing(-1, 802)
Global $LabelFile = GUICtrlCreateLabel("File", 10, 100, 400, 20)
GUICtrlSetResizing(-1, 802)
Global $LabelLongestLineOld = GUICtrlCreateLabel("0", 350, 100, 50, 20)
GUICtrlSetResizing(-1, 802)
Global $LabelLongestLineNew = GUICtrlCreateLabel("0", 420, 100, 50, 20)
GUICtrlSetResizing(-1, 802)

; "text", left, top [, width [, height [, style [, exStyle]]]]
Global $EditBoxIn = GUICtrlCreateEdit("", 16, 136, 460, 200)
GUICtrlSetResizing(-1, 33)
GUICtrlSetTip(-1, "View the input file")
GUICtrlSetState(-1, $GUI_DROPACCEPTED)

Global $EditBoxOut = GUICtrlCreateEdit("", 16, 350, 460, 200, Default)
GUICtrlSetResizing(-1, 65)
GUICtrlSetTip(-1, "View or edit the output file")
GUICtrlSetState(-1, $GUI_DROPACCEPTED)


GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

Defaults()

While 1
	Global $nMsg = GUIGetMsg(1)
	Switch $nMsg[0]
		Case $ButtonOpenFile
			LoadFile()
		Case $ButtonSaveFile
			SaveFile()
		Case $ButtonProcess
			Process()
		Case $ButtonDefaults
			Defaults()
		Case $SliderStringLength
			GUICtrlSetData($LabelStringLengthResult, GUICtrlRead($SliderStringLength))
		Case $ButtonAbout
			About()
		Case $GUI_EVENT_CLOSE
			Exit
		Case $ButtonExit
			Exit
	EndSwitch
WEnd
;-------------------------------------------------------------------------------------
Func Defaults()
	GUICtrlSetData($SliderStringLength, 80)
	GUICtrlSetData($LabelStringLengthResult, GUICtrlRead($SliderStringLength))
	$DefaultWorkingFolder = FileSelectFolder("Default Working Folder", $DefaultWorkingFolder, 7)
EndFunc   ;==>Defaults
;-------------------------------------------------------------------------------------
Func Process()
	GUICtrlSetData($LabelLongestLineOld, 0)
	GUICtrlSetData($LabelLongestLineNew, 0)

	Local $ArrayOfNewLines[1]

	Local $AllLines = GUICtrlRead($EditBoxIn)

	;$AllLines = StringRegExpReplace($AllLines, @LF, @CRLF, 0)

	Local $ArrayOfOldLines = StringSplit($AllLines, @CRLF, 3)

	For $Aline In $ArrayOfOldLines
		Local $LengthOfLine = StringLen($Aline)
		If GUICtrlRead($LabelLongestLineOld) < $LengthOfLine Then
			GUICtrlSetData($LabelLongestLineOld, $LengthOfLine)
		EndIf

		If $LengthOfLine < GUICtrlRead($LabelStringLengthResult) Then
			debug(StringFormat("%4d%4d%s", @ScriptLineNumber, StringLen($Aline), $Aline))
			_ArrayAdd($ArrayOfNewLines, $Aline)
		Else
			Local $ArrayOfWords = StringSplit($Aline, " ", 3)
			Local $NewString = '# '
			For $AWord In $ArrayOfWords
				$NewString = $NewString & " " & $AWord
				If StringLen($NewString) > GUICtrlRead($LabelStringLengthResult) Then
					debug(StringFormat("%4d%4d%s", @ScriptLineNumber, StringLen($NewString), $NewString))
					_ArrayAdd($ArrayOfNewLines, StringStripWS($NewString, 7))
					$NewString = ''
				EndIf
			Next
			$NewString = StringStripWS($NewString, 7)
			debug(StringFormat("%4d%4d%s", @ScriptLineNumber, StringLen($NewString), $NewString))
			_ArrayAdd($ArrayOfNewLines, StringStripWS($NewString, 7))
		EndIf
	Next

	_ArrayDelete($ArrayOfNewLines, 0)
	GUICtrlSetData($LabelLongestLineNew, 0)
	For $Aline In $ArrayOfNewLines
		$LengthOfLine = StringLen($Aline)
		If GUICtrlRead($LabelLongestLineNew) < $LengthOfLine Then
			GUICtrlSetData($LabelLongestLineNew, $LengthOfLine)
		EndIf
	Next

	Local $ts = _ArrayToString($ArrayOfNewLines, @CRLF)
	GUICtrlSetData($EditBoxOut, $ts)



EndFunc   ;==>Process
;-------------------------------------------------------------------------------------
Func LoadFile()
	Local $Filename = FileOpenDialog("Load file", $DefaultWorkingFolder, "Text file (*.txt)|All files (*.*)", 18, "")

	GUICtrlSetData($LabelFile, $Filename)

	Local $file = FileOpen($Filename, 0)
	; Check if file opened for reading OK
	If $file = -1 Then
		Debug("LoadFile unable to open file for reading: " & $Filename, 0x10, 5)
		Return
	EndIf

	Local $LineIn = FileRead($file)
	If @error <> 0 Then
		Debug("LoadFile read error " & $Filename, 0x10, 5)
		Return
	EndIf
	FileClose($file)


	$LineIn = StringRegExpReplace($LineIn, @LF, @CRLF, 0)

	GUICtrlSetData($EditBoxIn, $LineIn)
	GUICtrlSetData($EditBoxOut, '')
	GUICtrlSetData($LabelLongestLineOld, 0)
	Local $temp = StringSplit($LineIn, @CRLF, 1)
	For $X In $temp
		Local $Y = StringLen($X)
		If GUICtrlRead($LabelLongestLineOld) < $Y Then
			GUICtrlSetData($LabelLongestLineOld, $Y)
		EndIf
	Next
	GUICtrlSetData($LabelLongestLineNew, 0)


EndFunc   ;==>LoadFile
;-------------------------------------------------------------------------------------
Func SaveFile()

	Local $Filename = FileSaveDialog("Save file", $DefaultWorkingFolder, "Text file (*.txt)|All files (*.*)", 18, GUICtrlRead($LabelFile))
	Local $file = FileOpen($Filename, 2)
	; Check if file opened for writing OK
	If $file = -1 Then
		Debug("SaveFile: Unable to open file for writing: " & $Filename, 0x10, 5)
		Return
	EndIf

	Local $LineOut = GUICtrlRead($EditBoxOut)
	$LineOut = StringRegExpReplace($LineOut, @CRLF, @LF, 0)
	FileWrite($file, $LineOut)

	FileClose($file)

EndFunc   ;==>SaveFile
;-------------------------------------------------------------------------------------
Func About()
	Local $D = WinGetPos($ProgramName)
	Local $WinPos
	If IsArray($D) = 1 Then
		$WinPos = StringFormat("%s" & @CRLF & "WinPOS: %d  %d " & @CRLF & "WinSize: %d %d " & @CRLF & "Desktop: %d %d ", _
				$FormID, $D[0], $D[1], $D[2], $D[3], @DesktopWidth, @DesktopHeight)
	Else
		$WinPos = ">>>About ERROR, Check the window name<<<"
	EndIf
	Debug(@CRLF & $SystemS & @CRLF & $WinPos & @CRLF & "Written by Doug Kaynor because I wanted to!", 0x40, 5)
EndFunc   ;==>About