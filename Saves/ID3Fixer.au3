#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_outfile_type=a3x
#AutoIt3Wrapper_icon=../icons/10.ico
#AutoIt3Wrapper_outfile=C:\Program Files\AutoIt3\Dougs\ID3Fixer.a3x
#AutoIt3Wrapper_Res_Comment=ID3Fixer
#AutoIt3Wrapper_Res_Description=ID3Fixer
#AutoIt3Wrapper_Res_Fileversion=0.0.0.0
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=Y
#AutoIt3Wrapper_Res_ProductVersion=000
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
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <Array.au3>
#include <ButtonConstants.au3>
#include <Constants.au3>
#include <Date.au3>
#include <EditConstants.au3>
#include <GuiEdit.au3>
#include <GuiComboBox.au3>
#include <GuiComboBoxEx.au3>
#include <GUIConstantsEx.au3>
#include <GUIListBox.au3>
#include <GuiTreeView.au3>
#include <Misc.au3>
#include <String.au3>
#include <StaticConstants.au3>
#include <TreeViewConstants.au3>
#include <WindowsConstants.au3>
#include <Process.au3>
#include <SliderConstants.au3>
#include <ListViewConstants.au3>
#include <GuiListView.au3>

#include "ID3.au3"
#include <_DougFunctions.au3>
Opt("MustDeclareVars", 1)
Opt("WinTitleMatchMode", 2)
Global $PathString = "f:\Music\"

Global Const $tmp = StringSplit(@ScriptName, ".")
Global Const $ProgramName = $tmp[1]
Global Const $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")
Global $SystemS = $ProgramName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & @OSServicePack & @CRLF & @OSType & @CRLF & @OSArch

If _Singleton($ProgramName, 1) = 0 Then
    ConsoleWrite(@ScriptLineNumber & " " & $ProgramName & " is already running" & @CRLF)
    Exit
EndIf

;-----------------------------------------------

Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $MainForm = GUICreate("ID Tags", 600, 400, 10, 10, $MainFormOptions)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ButtonClose = GUICtrlCreateButton("Close", 10, 10, 80, 20, $WS_GROUP)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonAbout = GUICtrlCreateButton("About", 100, 10, 80, 20, $WS_GROUP)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonRead = GUICtrlCreateButton("Read", 10, 40, 80, 20, $WS_GROUP)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonWrite = GUICtrlCreateButton("Write", 100, 40, 80, 20, $WS_GROUP)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $InputPath = GUICtrlCreateInput($PathString, 200, 10, 250, 24)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKHEIGHT)

Global $ListResults = GUICtrlCreateList("", 10, 70, 550, 300, BitOR($LBS_NOTIFY, $WS_BORDER, $WS_TABSTOP, $WS_GROUP, $LBS_DISABLENOSCROLL, $WS_VSCROLL, $WS_HSCROLL))
GUICtrlSetData(-1, "")
GUICtrlSetTip(-1, "This is the results")
GUICtrlSetResizing(-1, $GUI_DOCKALL - $GUI_DOCKBOTTOM)
;-----------------------------------------------
GUISetState(@SW_SHOW)

While 1
    Global $nMsg = GUIGetMsg(1)
    Switch $nMsg[0]
        Case $GUI_EVENT_CLOSE
            Switch $nMsg[1]
                Case $MainForm
                    Exit
            EndSwitch

        Case $ButtonClose
            Exit
        Case $ButtonAbout
            About("ID Tags")
        Case $ButtonWrite
            WriteID()
        Case $ButtonRead
            ReadId()
        Case $ListResults
            If MsgBox(36, "Clear results?", "Are you sure that you want to clear results?") = 6 Then
                _GUICtrlListBox_ResetContent($ListResults)
            EndIf
    EndSwitch
WEnd
;-----------------------------------------------
Func WriteID()
    $PathString = _AddSlash2PathString(GUICtrlRead($InputPath))
    Global $TA = _FileListToArray($PathString)
    If IsArray($TA) Then
        For $Filename In $TA
            $Filename = $PathString & $Filename
            If FileExists($Filename) And StringInStr($Filename, ".mp3") > 0 Then
                Global $AA = StringSplit($Filename, "\")
                Global $NameTitle = $AA[UBound($AA) - 1]
                Global $AB = StringSplit($NameTitle, "-.")
                _GUICtrlListBox_AddString($ListResults, $Filename)
                ; _ID3SetTagField("TCOM", StringStripWS($AB[1], 7))
                ; _ID3SetTagField("TPE2", StringStripWS($AB[1], 7))
                ; _ID3SetTagField("TIT2", StringStripWS($AB[2], 7))
                ; _ID3WriteTag($Filename)
            EndIf
        Next
    EndIf
EndFunc   ;==>WriteID
;-----------------------------------------------
Func ReadId()
    $PathString = _AddSlash2PathString(GUICtrlRead($InputPath))
    Global $TA = _FileListToArray($PathString)
    If IsArray($TA) Then
        ;_ArrayDisplay($TA, @ScriptLineNumber)
        For $Filename In $TA
            $Filename = $PathString & $Filename
            If FileExists($Filename) And StringInStr($Filename, ".mp3") > 0 Then

                Global $ID3Tag = _ID3ReadTag($Filename, 0, -1, 1)
                ConsoleWrite(@ScriptLineNumber & " " & $Filename & " " & $ID3Tag[0] & @CRLF)
                If IsArray($ID3Tag) Then
                    For $String In $ID3Tag
                        _GUICtrlListBox_AddString($ListResults, $String)
                    Next
                EndIf

            EndIf
        Next
    EndIf
EndFunc   ;==>ReadId
;-----------------------------------------------
Func About(Const $FormID)
    Local $D = WinGetPos($FormID)
    Local $WinPos
    If IsArray($D) = True Then
        $WinPos = StringFormat("%s" & @CRLF & "WinPOS: %d  %d " & @CRLF & "WinSize: %d %d " & @CRLF & "Desktop: %d %d ", _
                $FormID, $D[0], $D[1], $D[2], $D[3], @DesktopWidth, @DesktopHeight)
    Else
        $WinPos = ">>>About ERROR, check the window name<<<"
    EndIf
    MsgBox(48, "", $SystemS & @CRLF & $WinPos & @CRLF & "Written by Doug Kaynor!")

EndFunc   ;==>About
;-----------------------------------------------
