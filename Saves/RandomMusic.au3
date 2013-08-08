#region
#AutoIt3Wrapper_icon=../icons/Cryptkeeper.ico
#AutoIt3Wrapper_outfile=C:\Program Files\AutoIt3\Dougs\RandomMusic.exe
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Comment=RandomMusic
#AutoIt3Wrapper_Res_Description=RandomMusic
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=Y
#AutoIt3Wrapper_Res_Fileversion=0.0.0.8
#AutoIt3Wrapper_Res_LegalCopyright=Copyright 2011 Douglas B Kaynor
#AutoIt3Wrapper_Res_SaveSource=y
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Field=AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Res_Field=Compile date|%longdate% %time%
#AutoIt3Wrapper_Res_Field=Developer|Douglas Kaynor
#AutoIt3Wrapper_Res_Field=Email|doug@kaynor.net
#AutoIt3Wrapper_Res_Field=Made By|Douglas Kaynor
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=y
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Run_Tidy=y
#Tidy_Parameters=/tc 4 /kv 2
#endregion

Opt("MustDeclareVars", 1)

#include <Array.au3>
#include <Constants.au3>
#include <Date.au3>
#include <file.au3>
#include <Misc.au3>
#include <String.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstants.au3>
#include <GUIConstantsEx.au3>
#include <SliderConstants.au3>
#include <WindowsConstants.au3>
#include <StaticConstants.au3>
#include "_DougFunctions.au3"

Global Const $tmp = StringSplit(@ScriptName, ".")
Global Const $ProgramName = $tmp[1]

If _Singleton($ProgramName, 1) = 0 Then
    _debug(@ScriptLineNumber & " " & $ProgramName & " is already running!", True)
    Exit
EndIf

Global $Project_filename = $AuxPath & $ProgramName & ".prj"

Global $A_In

Global $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")
Global $SystemS = @ScriptName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & @OSServicePack & @CRLF & @OSType & @CRLF & @OSArch & @IPAddress1

Global $MainForm = GUICreate(@ScriptName & " " & $FileVersion, 550, 260, 1000, 500)
GUISetFont(10, 400, 0, "Courier New")
Global $ButtonBuildPlayList = GUICtrlCreateButton("Build List", 10, 10, 100)

Global $ButtonHelp = GUICtrlCreateButton("Help", 230, 10, 60)
Global $ButtonAbout = GUICtrlCreateButton("About", 290, 10, 60)
Global $ButtonExit = GUICtrlCreateButton("Exit", 350, 10, 60)
Global $CheckBoxRecursive = GUICtrlCreateCheckbox("Recursive", 450, 10, 140, 20)
GUICtrlSetState($CheckBoxRecursive, $GUI_CHECKED)
Global $ButtonLoadProject = GUICtrlCreateButton("Load Project", 10, 50, 120, 20)
Global $ButtonSaveProject = GUICtrlCreateButton("Save Project", 130, 50, 120, 20)

GUICtrlCreateLabel("Input path:", 10, 80, 110, 20)
ConsoleWrite(@ScriptLineNumber & " " & @UserProfileDir & @CRLF)
Global $InputMusicSourceFolder = GUICtrlCreateInput(@UserProfileDir & "\Music", 120, 80, 410, 20, $ES_AUTOHSCROLL)
GUICtrlCreateLabel("Output file:", 10, 110, 110, 20)
Global $InputMusicDestinationFolder = GUICtrlCreateInput(@UserProfileDir & "\Desktop\music.m3u", 120, 110, 410, 20, $ES_AUTOHSCROLL)
GUICtrlCreateLabel("Filters:", 10, 140, 110, 20)
Global $InputFilters = GUICtrlCreateInput("mp3", 120, 140, 410, 20, $ES_AUTOHSCROLL)
GUICtrlCreateLabel("Number of songs:", 10, 165, 140, 20)
Global $InputNumberOfSongs = GUICtrlCreateInput(10, 150, 165, 40, 20, $ES_AUTOHSCROLL)
Global $SliderNumberOfSongs = GUICtrlCreateSlider(10, 190, 520, 25, $TBS_AUTOTICKS)
GUICtrlSetLimit(-1, 200, 1)
GUICtrlSetData(-1, 1)
GUICtrlSetData($SliderNumberOfSongs, GUICtrlRead($InputNumberOfSongs))

GUICtrlCreateLabel("Total size:", 200, 165, 90, 20)
Global $InputTotalSize = GUICtrlCreateInput('', 300, 165, 100, 20)
Global $HelpString = "Command line startup options:" & @CRLF & @CRLF & _
        "Any positive number value" & @CRLF & _
        "F11 to unlock GUI" & @CRLF & _
        "help or ? Help information is displayed"
;Global $LabelPlayTime = GUICtrlCreateLabel("Play time:", 400, 160, 150, 20, $SS_SUNKEN)
Global $LabelStatus = GUICtrlCreateLabel("", 10, 230, 520, 20, $SS_SUNKEN)

For $X = 1 To $CmdLine[0]
    ConsoleWrite($X & " >> " & $CmdLine[$X] & @CRLF)
    Select
        Case StringInStr($CmdLine[$X], "help") > 0 Or _
                StringInStr($CmdLine[$X], "?") > 0
            _Help($HelpString)
            Exit
        Case Int($CmdLine[$X]) > 0
            GUICtrlSetData($InputNumberOfSongs, $CmdLine[$X])
        Case FileExists($CmdLine[$X])
            GUICtrlSetData($InputMusicSourceFolder, $CmdLine[$X])
        Case Else
            _Help("Unknown cmdline option found: " & $CmdLine[$X])
            Exit
    EndSelect
Next

LoadProject()
_CheckWindowLocation($MainForm)

GUISetState(@SW_SHOW)
GUICtrlSetData($LabelStatus, "Ready to begin")

While 1
    Global $t = GUIGetMsg()
    Switch $t
        Case $GUI_EVENT_CLOSE
            Exit
        Case $ButtonExit
            Exit
        Case $ButtonAbout
            _About($ProgramName, $SystemS)
        Case $ButtonHelp
            _Help($HelpString)
        Case $ButtonBuildPlayList
            BuildPlayList()
        Case $InputFilters
        Case $InputMusicSourceFolder
        Case $InputMusicDestinationFolder
        Case $SliderNumberOfSongs
            GUICtrlSetData($InputNumberOfSongs, GUICtrlRead($SliderNumberOfSongs))
        Case $InputNumberOfSongs
            If GUICtrlRead($InputNumberOfSongs) < 1 Then GUICtrlSetData($InputNumberOfSongs, 1)
            ;if $NumberOfSongs > 200 then $NumberOfSongs = 200
            GUICtrlSetData($SliderNumberOfSongs, GUICtrlRead($InputNumberOfSongs))
        Case $ButtonLoadProject
            LoadProject("menu")
        Case $ButtonSaveProject
            SaveProject()
    EndSwitch
WEnd

;-----------------------------------------------
Func Refresh()
    _GuiDisable('Disable')
    If GUICtrlRead($CheckBoxRecursive) = $GUI_CHECKED Then
        $A_In = _FileListToArrayR(GUICtrlRead($InputMusicSourceFolder), "*.mp3", 1, 1, 1, "", 1)
    Else
        $A_In = _FileListToArrayR(GUICtrlRead($InputMusicSourceFolder), "*.mp3", 1, 0, 1, "", 1)
    EndIf
    _ArrayDelete($A_In, 0)
    ;_ArrayDisplay($A_In, @ScriptLineNumber)
    GUICtrlSetData($LabelStatus, "Refresh completed")
    _GuiDisable('Enable')
EndFunc   ;==>Refresh
;-----------------------------------------------
Func BuildPlayList()
    GUICtrlSetData($LabelStatus, "Working")
    Refresh()
    If Not IsArray($A_In) Then Refresh()

    If UBound($A_In) < GUICtrlRead($InputNumberOfSongs) Then
        MsgBox(64, "Song count error", _
                "Unable to comply with your request" & @CRLF & _
                "Songs to chose from: " & UBound($A_In) & @CRLF & _
                "Songs to put in playlist: " & GUICtrlRead($InputNumberOfSongs))
        _GuiDisable('Enable')
        GUICtrlSetData($LabelStatus, "Unable to comply with your request")
        Return
    EndIf

    Local $tries = 0
    Local $count = GUICtrlRead($InputNumberOfSongs)
    Global $file = FileOpenDialog('Chose a file name', GUICtrlRead($InputMusicDestinationFolder), '*.m3u')
    If $file = 1 Then
        Return
    EndIf

    While $count > 0
        If $tries > 100 Then ExitLoop
        Global $RNumber = Random(0, UBound($A_In) - 1, 1)
        If StringLen(GUICtrlRead($InputFilters)) < 1 Then GUICtrlSetData($InputFilters, '.mp3')
        If StringInStr($A_In[$RNumber], GUICtrlRead($InputFilters)) > 0 Then
            If FileExists($A_In[$RNumber]) Then
                ConsoleWrite(@ScriptLineNumber & " " & $A_In[$RNumber] & " " & @CRLF)
                #cs
                    Local $aID3V2Tag[1]
                    _ID3ReadTag($A_In[$RNumber], 0, -1, 1)
                    _ArrayDisplay($aID3V2Tag, @ScriptLineNumber)

                    ;_ReadID3v2($A_In[$RNumber], $aID3V2Tag)

                    _ID3ReadTag(FileGetShortName($A_In[$RNumber]))
                    ConsoleWrite(@ScriptLineNumber & " " & _ID3GetTagField("TIT2") & @CRLF)
                    ConsoleWrite(@ScriptLineNumber & " " & _ID3GetTagField("TLEN") & @CRLF)
                    ConsoleWrite(@ScriptLineNumber & " " & _ID3GetTagField("TCON") & @CRLF)
                #ce

                GUICtrlSetData($InputTotalSize, FileGetSize($A_In[$RNumber]) + GUICtrlRead($InputTotalSize))
                FileWriteLine($file, $A_In[$RNumber])

                $A_In[$RNumber] = ""
                $count = $count - 1
            Else
                $tries = $tries + 1
            EndIf

        EndIf
        $tries = $tries + 1
    WEnd
    GUICtrlSetData($InputTotalSize, _commify(GUICtrlRead($InputTotalSize)))
    FileClose($file)
    GUICtrlSetData($LabelStatus, "Playlist build complete. Destination: " & GUICtrlRead($InputMusicDestinationFolder))

EndFunc   ;==>BuildPlayList
;-----------------------------------------------

Func SaveProject()
    _GuiDisable('Disable')
    $Project_filename = FileSaveDialog("Save Project file", $AuxPath, _
            "All Project (*.prj)|RandomMusic.prj (RandomMusic.prj)|All files (*.*)", 18, $AuxPath & $ProgramName & ".prj")
    Local $file = FileOpen($Project_filename, 2)
    ; Check if file opened for writing OK
    If $file = -1 Then
        GUICtrlSetData($LabelStatus, "Unable to open file for writing: " & $Project_filename)
        _GuiDisable('Enable')
        Return
    EndIf

    FileWriteLine($file, "Valid for " & $ProgramName & " Project")
    FileWriteLine($file, "Project file for " & $ProgramName & "  Saved on " & _DateTimeFormat(_NowCalc(), 0))
    FileWriteLine($file, "Help: 1 is enabled, 4 is disabled for checkboxes and radio buttons")
    FileWriteLine($file, "InPathString:" & GUICtrlRead($InputMusicSourceFolder))
    FileWriteLine($file, "OutFileString:" & GUICtrlRead($InputMusicDestinationFolder))
    FileWriteLine($file, "FilterString:" & GUICtrlRead($InputFilters))
    FileWriteLine($file, "NumberOfSongs:" & GUICtrlRead($InputNumberOfSongs))
    FileWriteLine($file, "CheckBoxRecursive:" & GUICtrlRead($CheckBoxRecursive))

    FileClose($file)
    GUICtrlSetData($LabelStatus, "Project saved: " & $Project_filename)
    _GuiDisable('Enable')
EndFunc   ;==>SaveProject

#cs
_ArrayAdd($Settings, "InPathString:")
_ArrayAdd($Settings, $InputMusicSourceFolder)
_ArrayAdd($Settings, "OutFileString:")
_ArrayAdd($Settings, $InputMusicDestinationFolder)
_ArrayAdd($Settings, "FilterString:")
_ArrayAdd($Settings, $InputFilters)
_ArrayAdd($Settings, "NumberOfSongs:")
_ArrayAdd($Settings, $InputNumberOfSongs)
_ArrayAdd($Settings, "CheckBoxRecursive:")
_ArrayAdd($Settings, $CheckBoxRecursive)
#ce

;-----------------------------------------------
;This opens and loads the Project file
Func LoadProject($type = "start")
    _GuiDisable('Disable')
    ConsoleWrite(@ScriptLineNumber & " LoadProject: " & $type & " " & @CRLF)

    If StringCompare($type, "menu") = 0 Then
        $Project_filename = FileOpenDialog("Load project file", $AuxPath, _
                "All project files (*.prj)|RandomMusic.prj (RandomMusi.prj)|All files (*.*)", 18, $AuxPath & $ProgramName & ".prj")
    EndIf

    Local $file = FileOpen($Project_filename, 0)
    ; Check if file opened for reading OK
    If $file = -1 Then
        GUICtrlSetData($LabelStatus, "Unable to open file for reading: " & $Project_filename)
        _GuiDisable('Enable')
        Return
    EndIf

    ; Read in the first line to verify the file is of the correct type
    If StringCompare(FileReadLine($file, 1), "Valid for " & $ProgramName & " Project") <> 0 Then
        MsgBox(64, "Invalid projet file", " Not an Project file for " & $ProgramName)
        FileClose($file)
        _GuiDisable('Enable')
        Return
    EndIf

    ; Read in lines of text until the EOF is reached
    While 1
        Local $LineIn = FileReadLine($file)
        If @error = -1 Then ExitLoop
        If StringInStr($LineIn, ";") = 1 Then ContinueLoop
        ConsoleWrite(@ScriptLineNumber & " " & $LineIn & @CRLF)
        If StringInStr($LineIn, "InPathString:") Then GUICtrlSetData($InputMusicSourceFolder, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "OutFileString:") Then GUICtrlSetData($InputMusicDestinationFolder, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "FilterString:") Then GUICtrlSetData($InputFilters, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "NumberOfSongs:") Then GUICtrlSetData($InputNumberOfSongs, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        GUICtrlSetData($SliderNumberOfSongs, GUICtrlRead($InputNumberOfSongs))
        If StringInStr($LineIn, "CheckBoxRecursive:") Then GUICtrlSetState($CheckBoxRecursive, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
    WEnd
    FileClose($file)
    GUICtrlSetData($LabelStatus, "Project loaded: " & $Project_filename)

    _GuiDisable('Enable')
EndFunc   ;==>LoadProject
;-----------------------------------------------


