#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_icon=../icons/small_snake.ico
#AutoIt3Wrapper_outfile=C:\Program Files (x86)\AutoIt3\Dougs\StartUp.exe
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Comment=Start various programs
#AutoIt3Wrapper_Res_Description=Start various programs
#AutoIt3Wrapper_Res_Fileversion=0.0.0.12
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=Y
#AutoIt3Wrapper_Res_ProductVersion=000
#AutoIt3Wrapper_Res_LegalCopyright=Copyright 2012 Douglas B Kaynor
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
#AutoIt3Wrapper_Run_Debug_Mode=N
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****

Opt("MustDeclareVars", 1) ; require pre-declared varibles
If _Singleton(@ScriptName, 1) = 0 Then
    _Debug(@ScriptName & " is already running!", 0x40)
    Exit
EndIf

;#include <Array.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <Constants.au3>
#include <GUIMenu.au3>
#include <GUIConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiListBox.au3>
#include <file.au3>
#include <Misc.au3>
#include <String.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <_DougFunctions.au3>

Global Const $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")
Global $SystemS = @ScriptName & @LF & $FileVersion & @LF & @OSVersion & @LF & @OSServicePack & @LF & @OSType & @LF & @OSArch
Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, _
        $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
; Main form --------------------------
Global $MainForm = GUICreate(@ScriptName & $FileVersion, 720, 400, 10, 10, $MainFormOptions)
GUISetFont(10, 400, -1, "Courier new")

Global Const $StartFileName = "start.txt"

Global $ButtonGetCommands = GUICtrlCreateButton("Get commands", 10, 10, 110)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonRunACommand = GUICtrlCreateButton("Run command", 130, 10, 100)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonRunCommandGroup = GUICtrlCreateButton("Run group", 240, 10, 80)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonEditList = GUICtrlCreateButton("Edit list", 330, 10, 80)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ComboSelectPriority = GUICtrlCreateCombo("All", 420, 10, 110, -1, -1, -1)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetData(-1, "High|Medium|Low|AlmostNever", "High")
Global $ButtonAbout = GUICtrlCreateButton("About", 540, 10, 60)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonExit = GUICtrlCreateButton("Exit", 630, 10, 60)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ListBox = GUICtrlCreateList("", 10, 50, 700, 350, BitOR($WS_BORDER, $WS_VSCROLL)) ; $WS_TABSTOP, $LBS_NOTIFY)
GUICtrlSetResizing(-1, $GUI_DOCKTOP)

GetCommands()
GUISetState(@SW_SHOW, $MainForm)
;-----------------------------------------------
While 1
    Global $msg = GUIGetMsg(True)
    Switch $msg[0]
        Case $GUI_EVENT_CLOSE
            If $msg[1] = $MainForm Then Exit
        Case $ButtonExit
            ExitLoop
        Case $ButtonAbout
            _About('StartUp', $SystemS, 'Written by Doug Kaynor')
        Case $ButtonGetCommands
            GetCommands()
        Case $ButtonRunCommandGroup
            RunCommandGroup()
        Case $ButtonRunACommand
            RunACommand()
        Case $ButtonEditList
            EditList()
        Case $ComboSelectPriority

    EndSwitch
WEnd
;-----------------------------------------------
Func EditList()
    Local $TmpFileName = "tmpStart.txt"
    FileCopy($StartFileName, $TmpFileName, 1)

    ShellExecuteWait(_ChoseTextEditor(), $TmpFileName, "", "open")

    If MsgBox(4 + 32, "Edit start list", "Do you want to save the changes?") = 6 Then
        FileCopy($TmpFileName, $StartFileName, 1)
        GetCommands()
        MsgBox(32, "Save start file", " New Start file written")
    EndIf

EndFunc   ;==>EditList
;-----------------------------------------------
;Read in the start.txt file. Abort if not found
Func GetCommands()
    _GUICtrlListBox_ResetContent($ListBox)
    Local $TA

    If _FileReadToArray($StartFileName, $TA) <> 1 Then
        MsgBox(48, @ScriptLineNumber & " GetCommands", "Error reading command file:" & @LF & $StartFileName)
        Exit
    EndIf
    _ArrayDelete($TA, 0)

    For $String In $TA
        _GUICtrlListBox_AddString($ListBox, StringFormat("%s", $String))
    Next

EndFunc   ;==>GetCommands
;-----------------------------------------------
; ""High|Medium|Low|AlmostNever"
; GUICtrlRead($ComboSelectPriority)
Func RunCommandGroup()
    For $X = 0 To _GUICtrlListBox_GetCount($ListBox)
        Local $T = _GUICtrlListBox_GetText($ListBox, $X)
        ;ConsoleWrite(@ScriptLineNumber & " " & $T & @CRLF)
        ;ConsoleWrite(@ScriptLineNumber & " Delay: " & StringInStr($T, 'delay') & @CRLF)
        If StringInStr($T, '#') = 1 Or StringInStr($T, '"') = 0 Then
            ConsoleWrite(@ScriptLineNumber & " Blank or comment: " & $T & @CRLF)
        ElseIf StringInStr($T, 'delay') = 1 Then
            Local $Value = StringSplit(StringMid($T, StringInStr($T, 'delay'), 99), '" ')
            ConsoleWrite(@ScriptLineNumber & ": " & _ArrayToString($Value) & @CRLF)
            Sleep(Number($Value[2]))
        Else
            Local $A = StringMid($T, StringInStr($T, '"'), 99)
            ;ConsoleWrite(@ScriptLineNumber  & " " & GUICtrlRead($ComboSelectPriority) & "<<>>" & $T &  @LF)
            If StringInStr($T, GUICtrlRead($ComboSelectPriority)) = 1 Or _
                    StringInStr("All", GUICtrlRead($ComboSelectPriority)) = 1 Then
                ;   If Run($A) = 0 Then MsgBox(48, "RunCommandGroup", "Run command group error:" & @LF & $A & @LF & "Error: " & @error)
                ConsoleWrite(@ScriptLineNumber & " " & GUICtrlRead($ComboSelectPriority) & "  " & $T & "  " & $A & @LF)
            EndIf
        EndIf
    Next
EndFunc   ;==>RunCommandGroup
;-----------------------------------------------
Func RunACommand()
    Local $T = _GUICtrlListBox_GetText($ListBox, _GUICtrlListBox_GetCurSel($ListBox))
    ;ConsoleWrite(@ScriptLineNumber & " " & $T & @LF)
    If StringInStr($T, "#") = 1 Or StringInStr($T, '"') = 0 Then
        ConsoleWrite(@ScriptLineNumber & " RunACommand. Invalid command string:" & $T & @LF)
    Else
        Local $A = StringMid($T, StringInStr($T, '"'), 99)
        ConsoleWrite(@ScriptLineNumber & " " & $A & @LF)
        Local $Result = Run($A)
        If $Result = 0 Then
            MsgBox(48, "RunACommand", "Run a command error:" & @LF & $A & @LF & "Result: " & $Result & "  Error: " & @error)
        EndIf
        ConsoleWrite(@ScriptLineNumber & " " & GUICtrlRead($ComboSelectPriority) & "  " & $T & "  " & $A & @LF)
    EndIf
EndFunc   ;==>RunACommand
;-----------------------------------------------
