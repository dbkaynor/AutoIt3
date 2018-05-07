#AutoIt3Wrapper_icon=../icons/Battery.ico
#AutoIt3Wrapper_outfile=BatteryStatus.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Comment=A battery status monitoring app
#AutoIt3Wrapper_Res_Description==Battery status
#AutoIt3Wrapper_Res_Fileversion=1.0.0.45
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=Y
#AutoIt3Wrapper_Res_ProductVersion=666
#AutoIt3Wrapper_Res_LegalCopyright=Copyright 2016 Douglas B Kaynor
#AutoIt3Wrapper_Res_SaveSource=y
#AutoIt3Wrapper_Res_Language=10334
#AutoIt3Wrapper_Res_Field=AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Res_Field=Compile date|%longdate% %time%
#AutoIt3Wrapper_Res_Field=Developer|Douglas Kaynor
#AutoIt3Wrapper_Res_Field=Email|doug@kaynor.net
#AutoIt3Wrapper_Res_Field=Made By|Douglas Kaynor
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=y
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Run_Tidy=y
#Tidy_Parameters=/tc 4 /kv 2
#AutoIt3Wrapper_Run_Au3Stripper=n

#include <Array.au3>
#include <GUIConstantsEx.au3>
#include <ColorConstantS.au3>
#include <FontConstants.au3>
#include <WinAPISys.au3>
#include <WindowsConstants.au3>
#include <TrayConstants.au3>
#include <Misc.au3>

Global Const $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")
Global Const $tmp = StringSplit(@ScriptName, ".")
Global Const $ProgramName = $tmp[1]

If @Compiled Then
    If _Singleton($ProgramName, 1) = 0 Then
        MsgBox(64, $ProgramName, $ProgramName & @CRLF & "is already running")
        Exit
    EndIf
EndIf

Global Const $TrayIconRunning = '../icons/Battery.ico'
TraySetIcon($TrayIconRunning)

Opt("TrayMenuMode", 1)
Opt("TrayOnEventMode", 0)
;Opt("TrayIconDebug", 1)

TraySetState($TRAY_ICONSTATE_SHOW)
TrayTip($ProgramName, $FileVersion, 3)

Global $hForm = GUICreate($ProgramName, 560, 30, -1, 5) ;, $WS_POPUP)

Global $ACLabel1 = GUICtrlCreateLabel('External power:', 10, 10, 90, 22)
Global $ACLabel2 = GUICtrlCreateLabel('Unknown', 90, 10, 50, 20)
Global $StatusLabel1 = GUICtrlCreateLabel('Status:', 150, 10, 35, 22)
Global $StatusLabel2 = GUICtrlCreateLabel('Unknown', 190, 10, 50, 20)
Global $ChargeLabel1 = GUICtrlCreateLabel('Charge:', 250, 10, 90, 22)
Global $ChargeLabel2 = GUICtrlCreateLabel('Unknown', 300, 10, 50, 20)
Global $TimeLabel1 = GUICtrlCreateLabel('Time:', 350, 10, 90, 22)
Global $TimeLabel2 = GUICtrlCreateLabel('Unknown', 380, 10, 50, 20)

Global $exitButton = GUICtrlCreateButton('Exit', 450, 5, 40, 20)
Global $SecondsButton = GUICtrlCreateButton("Secs", 500, 5, 40, 20)

Global $RawData
Global $DisplayData[4]

GUISetState(@SW_SHOW)

AdlibRegister('_BatteryStatus', 1000)

While 1
    Switch GUIGetMsg()
        Case $exitButton
            ExitLoop
        Case $GUI_EVENT_CLOSE
            ExitLoop
    EndSwitch
WEnd

Func _BatteryStatus()
    Local $RawData = _WinAPI_GetSystemPowerStatus()
    If @error Then Return

    If BitAND($RawData[1], 128) Then
        $RawData[0] = 'Not present'
        For $i = 1 To 3
            $RawData[$i] = 'Unknown'
        Next
    Else
        Switch $RawData[0]
            Case 0
                $RawData[0] = 'Offline'
            Case 1
                $RawData[0] = 'Online'
            Case Else
                $RawData[0] = 'Unknown'
        EndSwitch
        Switch $RawData[2]
            Case 0 To 100
                $RawData[2] &= '%'
            Case Else
                $RawData[2] = 'Unknown'
        EndSwitch
        Switch $RawData[3]
            Case -1
                $RawData[3] = 'Unknown'
            Case Else
                Local $H, $M
                $H = ($RawData[3] - Mod($RawData[3], 3600)) / 3600
                $M = ($RawData[3] - Mod($RawData[3], 60)) / 60 - $H * 60
                $RawData[3] = StringFormat($H & ':%02d', $M)
        EndSwitch
        If BitAND($RawData[1], 8) Then
            $RawData[1] = 'Charging'
        Else
            Switch BitAND($RawData[1], 0xF)
                Case 1
                    $RawData[1] = 'High'
                Case 2
                    $RawData[1] = 'Low'
                Case 4
                    $RawData[1] = 'Critical'
                Case Else
                    $RawData[1] = 'Unknown'
            EndSwitch
        EndIf
    EndIf

    GUICtrlSetData($ACLabel2, $RawData[0])
    GUICtrlSetData($StatusLabel2, $RawData[1])
    GUICtrlSetData($ChargeLabel2, $RawData[2])
    GUICtrlSetData($TimeLabel2, $RawData[3])
    GUICtrlSetData($SecondsButton, @SEC)

EndFunc   ;==>_BatteryStatus
