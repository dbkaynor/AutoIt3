#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_icon=../icons/FaithTones.ico
#AutoIt3Wrapper_outfile=MemGetStats.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Comment=Memory status
#AutoIt3Wrapper_Res_Description=A tool to watch memory statistics
#AutoIt3Wrapper_Res_Fileversion=1.0.0.14
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=Y
#AutoIt3Wrapper_Res_ProductVersion=666
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

#cs
    This area is used to store things todo, bugs, and other notes
    
    Fixed:
    
    Todo:
    Fix all help calls (F1)
    Add start value, min, and max value
    Logging would be nice
    
#CE
#include <array.au3>
#include <date.au3>
#include <GUIConstantsEx.au3>
#include <misc.au3>
#include <SliderConstants.au3>
#include <StaticConstants.au3>

Opt("MustDeclareVars", 1)
Opt("WinTitleMatchMode", 1)
Opt("TrayMenuMode", 3)

Global $array = MemGetStats()
Global $PercentStart = $array[0]
Global $PhStartU = $array[2]
Global $PhMinU = $array[2]
Global $PhMaxU = 0
Global $PaStartU = $array[4]
Global $PaMinU = $array[4]
Global $PaMaxU = 0
Global $VStartU = $array[6]
Global $VMinU = $array[6]
Global $VMaxU = 0

DirCreate(@ScriptDir & "\AUXFiles")
Global $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")
Global $tmp = StringSplit(@ScriptName, ".")
Global Const $ProgramName = $tmp[1]
Global $LOGFILE = FileGetShortName(@ScriptDir & "\" & $ProgramName & ".log")

If _Singleton($ProgramName, 1) = 0 Then
    MsgBox(32, "Already running", $ProgramName & " is already running!")
    Exit
EndIf

Global $MainForm = GUICreate("Memory Statistics (megabytes", 380, 270) ; will create a dialog box that when displayed is centered

Static $SavedTime = TimerInit()
GUISetFont(9, 400, 0, "Courier New")

Global $LabelCurrent = GUICtrlCreateLabel(GetTheStats(), 10, 10, 360, 100, $SS_SUNKEN)
GUICtrlSetTip(-1, "Cuurent values")

Global $LabelHistory = GUICtrlCreateLabel(GetHistory(), 10, 160, 360, 100, $SS_SUNKEN)
GUICtrlSetTip(-1, "Historical values")

Global $ButtonRefresh = GUICtrlCreateButton("Refresh now", 10, 120)
GUICtrlSetTip(-1, "Refresh the results")

Global $SliderSpeed = GUICtrlCreateSlider(120, 120, 150, 25, $TBS_AUTOTICKS);, BitOR($TBS_AUTOTICKS, $TBS_NOTICKS))
GUICtrlSetLimit(-1, 60, 1)
GUICtrlSetData(-1, 2)
GUICtrlSetTip(-1, "Auto refresh delay")

Global $LabelDelay = GUICtrlCreateLabel(GUICtrlRead($SliderSpeed), 100, 120, 20, 25, $SS_SUNKEN)
GUICtrlSetTip(-1, "Current refresh delay setting")

Global $ButtonLog = GUICtrlCreateButton("Log ", 280, 120)
GUICtrlSetTip(-1, "Add current data to logfile")

Global $ButtonExit = GUICtrlCreateButton("Exit ", 330, 120)
GUICtrlSetTip(-1, "Exit the program")

GUISetState()
GUICtrlSetData($LabelCurrent, GetTheStats())
GUICtrlSetData($LabelHistory, GetHistory())
DataLog('Startup')
;-----------------------------------------------
While 1
    Switch GUIGetMsg()
        Case $ButtonRefresh
            GUICtrlSetData($LabelCurrent, GetTheStats())
        Case $SliderSpeed
            GUICtrlSetData($LabelDelay, GUICtrlRead($SliderSpeed))
        Case $ButtonLog
            DataLog()
        Case $ButtonExit
            ExitLoop
        Case $GUI_EVENT_CLOSE
            DataLog('Close')
            Exit
    EndSwitch
    CheckChangeCounter()
WEnd
;-----------------------------------------------
Func CheckChangeCounter()
    Local $CurrentTime = TimerDiff($SavedTime) / 1000 ; seconds
    If Mod(Int($CurrentTime * 400), 10) <> 0 Then Return
    If $CurrentTime > GUICtrlRead($SliderSpeed) Then
        $SavedTime = TimerInit()
        GUICtrlSetData($LabelCurrent, GetTheStats())
        GUICtrlSetData($LabelHistory, GetHistory())
    EndIf
EndFunc   ;==>CheckChangeCounter
;-----------------------------------------------
Func DataLog(Const $String = '')
    Local $file = FileOpen($LOGFILE, 1)
    Local $tSystem = _Date_Time_GetLocalTime()

    FileWriteLine($file, _Date_Time_SystemTimeToDateTimeStr($tSystem, 1) & '  ' & $String)
    FileWriteLine($file, GUICtrlRead($LabelCurrent))
    FileWriteLine($file, GUICtrlRead($LabelHistory))

    FileWriteLine($file, "--------------------------------------------------")
    FileClose($file)
EndFunc   ;==>DataLog
;-----------------------------------------------
Func GetTheStats()
    Local $array = MemGetStats()
    TraySetToolTip($array[0] & "% of memory (current)")
    Return $array[0] & "% of memory in use (current)" & @CRLF & @CRLF & _
            "                Total     Available          Used" & @CRLF & _
            StringFormat("Physical:  %10.3f    %10.3f    %10.3f  ", $array[1] / 1000, $array[2] / 1000, ($array[1] - $array[2]) / 1000) & @CRLF & _
            StringFormat("Page file: %10.3f    %10.3f    %10.3f  ", $array[3] / 1000, $array[4] / 1000, ($array[3] - $array[4]) / 1000) & @CRLF & _
            StringFormat("Virtual:   %10.3f    %10.3f    %10.3f ", $array[5] / 1000, $array[6] / 1000, ($array[5] - $array[6]) / 1000 & @CRLF)
EndFunc   ;==>GetTheStats
;-----------------------------------------------
Func GetHistory()
    Local $array = MemGetStats()
    If $PhMinU > ($array[1] - $array[2]) Then $PhMinU = ($array[1] - $array[2])
    If $PhMaxU < ($array[1] - $array[2]) Then $PhMaxU = ($array[1] - $array[2])
    If $PaMinU > ($array[3] - $array[4]) Then $PaMinU = ($array[3] - $array[4])
    If $PaMaxU < ($array[3] - $array[4]) Then $PaMaxU = ($array[3] - $array[4])
    If $VMinU > ($array[5] - $array[6]) Then $VMinU = ($array[5] - $array[6])
    If $VMaxU < ($array[5] - $array[6]) Then $VMaxU = ($array[5] - $array[6])

    Return $PercentStart & "% of memory (at start)" & @CRLF & @CRLF & _
            "         Current used      Max used      Min used" & @CRLF & _
            StringFormat("Physical:  %10.3f    %10.3f    %10.3f  ", ($array[1] - $array[2]) / 1000, $PhMaxU / 1000, $PhMinU / 1000) & @CRLF & _
            StringFormat("Page file: %10.3f    %10.3f    %10.3f  ", ($array[3] - $array[4]) / 1000, $PaMaxU / 1000, $PaMinU / 1000) & @CRLF & _
            StringFormat("Virtual:   %10.3f    %10.3f    %10.3f ", ($array[5] - $array[6]) / 1000, $VMaxU / 1000, $VMinU / 1000)
EndFunc   ;==>GetHistory
;-----------------------------------------------
#cs
    $array[0] = Memory Load(Percentage of memory In use)
    $array[1] = Total physical RAM
    $array[2] = Available physical RAM
    $array[3] = Total Pagefile
    $array[4] = Available Pagefile
    $array[5] = Total virtual
    $array[6] = Available virtual
#ce
