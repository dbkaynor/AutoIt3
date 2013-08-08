#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_outfile_type=a3x
#AutoIt3Wrapper_icon=../icons/BrickWall.ico
#AutoIt3Wrapper_outfile=C:\Program Files\AutoIt3\Dougs\Renamer.a3x
#AutoIt3Wrapper_Res_Comment=File Renamer
#AutoIt3Wrapper_Res_Description=Renamer
#AutoIt3Wrapper_Res_Fileversion=0.0.0.0
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=Y
#AutoIt3Wrapper_Res_ProductVersion=666
#AutoIt3Wrapper_Res_LegalCopyright=Copyright 2010 Douglas B Kaynor
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
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****

Opt("MustDeclareVars", 1)
Opt("TrayIconDebug", 1)

#cs
    show errors
    show file renamed
    Fix the word case option
    make a better filter (wild cards)
    Make a search
#ce
#include <Array.au3>
#include <ButtonConstants.au3>
#include <Constants.au3>
#include <Date.au3>
#include <GuiEdit.au3>
#include <GuiComboBox.au3>
#include <GuiComboBoxEx.au3>
#include <GUIConstantsEx.au3>
#include <GUIListBox.au3>
#include <Misc.au3>
#include <String.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <_DougFunctions.au3>

If _Singleton(@ScriptName, 1) = 0 Then
    MsgBox(16, "Warning!!!  " & @ScriptLineNumber & " ", @ScriptName & " is already running!")
    Exit
EndIf
;Global $TopFolder = "C:\Users\Doug\Downloads\"
;Global $TopFolder = "J:\music\"
Global $TopFolder = "C:\Users\dbkaynox\Music\"
Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $MainForm = GUICreate("Renamer", 550, 620, 10, 10, $MainFormOptions)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ButtonChose = GUICtrlCreateButton("Folder", 20, 20, 50)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $InputFolder = GUICtrlCreateInput("", 80, 20, 250)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonRefresh = GUICtrlCreateButton("Refresh", 340, 20, 50)

Global $ButtonDoit = GUICtrlCreateButton("Doit", 395, 20, 50)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonRecycle = GUICtrlCreateButton("Recycle selected", 450, 20)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

GUICtrlCreateLabel("Find string", 20, 60, 150)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $InputFind = GUICtrlCreateInput("", 20, 80, 150)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

GUICtrlCreateLabel("Replacement string", 20, 100)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $InputReplace = GUICtrlCreateInput("", 20, 120, 150)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

GUICtrlCreateLabel("Filter string", 20, 140)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $InputFilter = GUICtrlCreateInput(".mp3", 20, 160, 150)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $GroupNameDisplayOptions = GUICtrlCreateGroup("Options", 200, 65, 320, 120)
Global $CheckboxSpaceFix = GUICtrlCreateCheckbox("Space fix", 210, 90, 100, 20)
Global $CheckDontAsk = GUICtrlCreateCheckbox("Don't ask", 210, 120, 100, 20)
Global $CheckboxCharacters = GUICtrlCreateCheckbox("Characters", 210, 150, 100, 20)
;Global $GroupPosition = GUICtrlCreateGroup("Position", 27, 100, 90, 65)
;Global $RadioTop = GUICtrlCreateRadio("Top", 37, 115, 65, 20)
;Global $RadioBottom = GUICtrlCreateRadio("Bottom", 37, 135, 65, 20)
GUICtrlCreateGroup("", -99, -99, 1, 1)


GUICtrlCreateLabel("Files", 20, 180)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ListFiles = GUICtrlCreateList("???", 20, 200, 500, 290)
GUICtrlSetData(-1, "")
GUICtrlSetTip(-1, "This is the List of matching files")
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKBOTTOM)

GUICtrlCreateLabel("Status", 20, 490)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ListStatus = GUICtrlCreateList("", 20, 510, 500, 100, BitOR($LBS_NOTIFY, $WS_BORDER, $WS_TABSTOP, $WS_GROUP, $LBS_DISABLENOSCROLL, $WS_VSCROLL, $WS_HSCROLL))
GUICtrlSetData(-1, "")
GUICtrlSetTip(-1, "This is the List of status infomation")
GUICtrlSetResizing(-1, $GUI_DOCKALL - $GUI_DOCKWIDTH)
GUISetState(@SW_SHOW, $MainForm)

While 1
    Global $nMsg = GUIGetMsg(1)
    Switch $nMsg[0]
        Case $GUI_EVENT_CLOSE ; The following handles the red X
            Exit
        Case $ButtonRecycle
            Run("explorer::{645FF040-5081-101B-9F08-00AA002F954E}")
        Case $ButtonChose
            Global $tmp = FileSelectFolder("Chose a folder", "", 7, $TopFolder)
            If StringInStr($tmp, ":") = 2 Then
                GUICtrlSetData($InputFolder, _AddSlash2PathString(_CleanUpPath($tmp)))
                _GUICtrlListBox_ResetContent($ListFiles)
                Global $TA = _FileListToArray(GUICtrlRead($InputFolder))
                If IsArray($TA) Then
                    For $X In $TA
                        If IsString($X) And StringInStr($X, GUICtrlRead($InputFilter)) <> 0 Then
                            _GUICtrlListBox_AddString($ListFiles, $X)
                        EndIf
                    Next
                EndIf
            EndIf
        Case $ButtonRefresh
            If StringInStr(GUICtrlRead($InputFolder), ":") = 2 Then
                _GUICtrlListBox_ResetContent($ListFiles)
                $TA = _FileListToArray(GUICtrlRead($InputFolder))
                If IsArray($TA) Then
                    For $X In $TA
                        If IsString($X) And StringInStr($X, GUICtrlRead($InputFilter)) <> 0 Then
                            _GUICtrlListBox_AddString($ListFiles, $X)
                        EndIf
                    Next
                EndIf
            EndIf
        Case $ListFiles
            Global $A = StringSplit(_GUICtrlListBox_GetText($ListFiles, _GUICtrlListBox_GetCurSel($ListFiles)), "-")
            Global $B = StringStripWS($A[1], 3)
            ClipPut($B)
            GUICtrlSetData($InputFind, $B)
            GUICtrlSetData($InputReplace, $B)
        Case $ListStatus
            If MsgBox(36, "Clear status?", "Are you sure that you want to clear status?") = 6 Then
                _GUICtrlListBox_ResetContent($ListStatus)
            EndIf
        Case $ButtonDoit
            If GUICtrlRead($CheckDontAsk) = $GUI_CHECKED Or _
                    MsgBox(49, "Rename", "Are you sure you want to replace occurances of" & @CRLF & _
                    GUICtrlRead($InputFind) & " with " & GUICtrlRead($InputReplace)) <> 1 Then
            EndIf
            For $X = 0 To _GUICtrlListBox_GetCount($ListFiles) - 1
                Global $InString = _GUICtrlListBox_GetText($ListFiles, $X)
                ConsoleWrite(@ScriptLineNumber & " " & $InString & @CRLF)
                If FileExists(GUICtrlRead($InputFolder) & $InString) Then
                    Global $OutString = StringReplace($InString, GUICtrlRead($InputFind), GUICtrlRead($InputReplace), 1)
                    _GUICtrlListBox_AddString($ListStatus, "In: " & $InString & " Out: " & $OutString)
                    FileMove(GUICtrlRead($InputFolder) & $InString, GUICtrlRead($InputFolder) & $OutString)
                    ;  ControlClick("Renamer", "", $ButtonRefresh, "left")
                EndIf
            Next
    EndSwitch
WEnd
;-----------------------------------------------
