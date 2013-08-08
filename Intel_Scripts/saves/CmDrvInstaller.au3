#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_icon=../icons/canopus.ico
#AutoIt3Wrapper_outfile=T:\temp\CmDrvInstaller.exe
#AutoIt3Wrapper_Res_Description=Cross platform driver installer
#AutoIt3Wrapper_Res_Fileversion=0.0.0.18
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=Y
#AutoIt3Wrapper_Res_LegalCopyright=Copyright © 2010 Intel Corporation. All rights reserved.
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Field=Compiler|AutoIt Version: %AutoItVer%
#AutoIt3Wrapper_Res_Field=Author|Doug Kaynor
#AutoIt3Wrapper_Res_Field=Updated By|Doug Kaynor
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_UseX64=n
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Run_Debug_Mode=n

#comments-start PROGRAM HEADER
	;******************************************************************************************
	;** Intel Corporation, MPG MPAD
	;** Title			:  		CmDrvInstaller.au3
	;** Description	:
	;**		Driver installer
	;**
	;** Revision: 	Rev 2.0.0
	;******************************************************************************************
	;******************************************************************************************
	;** Revision History:
	;**
	;**
	;**	- Initial release - Doug Kaynor 04/15/2010
	;**
	;******************************************************************************************
#comments-end PROGRAM HEADER

; Script/File name
Const $SCRIPT_NAME = "Platform Driver Installer"
Const $SCRIPT_FILENAME = "CmDrvInstaller.au3"
Const $SCRIPT_VERSION = "V2.0.0"

; Include files
#include <Constants.au3>
#include <incAll.au3>
#include <Process.au3>

; Prepre for running
ScriptStarting($SCRIPT_NAME, $SCRIPT_FILENAME, $SCRIPT_VERSION)

If Not IsSupportedOS() Then
	Local $strError = "ERROR: Unsupported OS"
	ExitWithErrorMessage($SCRIPT_NAME, $strError)
	Exit
EndIf

; Allow user press ESC to abort the program
AbortProgram($SCRIPT_NAME)

Global $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")
Global $tmp = StringSplit(@ScriptName, ".")
Global $ProgramName = $tmp[1]
Global $ResultLocation = ""
Global $SystemProductName

Global $DEBUG = False

Global $OsName
Const $LOGFILE = "C:\LOGS\" & $ProgramName & ".LOG"
Const $InstallFolder = "C:\SCRIPTS\"
Const $Install_xml = $InstallFolder & "Install.xml"
Const $InstallTemp_xml = $InstallFolder & "InstallTEMP.xml"

Global $DEVCON

Global $ArrayConfigData[1]
Global $ArrayProblemsData[1]
Global $ArrayValidData[1]

FileDelete($LOGFILE)

WriteLog("Starting " & $ProgramName & ". Logging results to: " & $LOGFILE)
WriteMyLog(@ScriptLineNumber & " -----------------------------------------------------------")
WriteMyLog(@ScriptLineNumber & StringFormat(" Startup %s %s %s %s %s %s", $ProgramName, $FileVersion, @OSVersion, @OSServicePack, @OSType, @OSArch))

TestForStuff()
GetSystemInfo()
LoadConfig()
GetListOfProblemDevices()
BuildInstallTEMP_XML()

WriteMyLog(@ScriptLineNumber & " Completed " & $ProgramName & ". Logged results to: " & $LOGFILE)
WriteMyLog(@ScriptLineNumber & " -----------------------------------------------------------")

ScriptEndingRebootOthers($SCRIPT_NAME, "VManager -noui", "INSTALLED")
;-----------------------------------------------
;Compare the problem devces against the config file to get a list of drivers to install
Func BuildInstallTEMP_XML() ; dbk
	WriteMyLog(@ScriptLineNumber & " BuildInstallTEMP_XML")

	If Not FileExists($InstallTemp_xml) Then
		WriteMyLog(@ScriptLineNumber & " BuildInstallTEMP_XML " & $InstallTemp_xml & " not found")
		ExitWithErrorMessage($SCRIPT_NAME, "BuildInstallTEMP_XML " & $InstallTemp_xml & " not found")
	EndIf

	For $A In $ArrayProblemsData
		$B = StringSplit($A, "~", 2)
		Local $SaveRunCmd, $SaveName
		For $C In $ArrayValidData
			If StringInStr($C, "RunCmd>>") > 0 Then $SaveRunCmd = StringStripWS(StringReplace($C, "RunCmd>>", ""), 3)
			If StringInStr($C, "Name>>") > 0 Then $SaveName = StringStripWS(StringReplace($C, "Name>>", ""), 3)
			If StringInStr($B[0], $C) > 0 Then
				WriteMyLog(@ScriptLineNumber & " " & $C & " " & $SaveRunCmd & " " & $SaveName)
				Local $NewString = "<Driver>" & @CRLF & _
						"<App Name=" & Chr(34) & $SaveName & Chr(34) & _
						" RunCmd=" & Chr(34) & $SaveRunCmd & Chr(34) & _
						" Reboot=" & Chr(34) & "No" & Chr(34) & _
						" Order=" & Chr(34) & "20" & Chr(34) & _
						" DVD=" & Chr(34) & "Yes" & Chr(34) & _
						" Portable=" & Chr(34) & "Yes" & Chr(34) & _
						" Production=" & Chr(34) & "Yes" & Chr(34) & _
						" Tab=" & Chr(34) & "Install" & Chr(34) & "/>" & @CRLF & _
						"<SupportedOS>" & @CRLF & _
						"<OS>ALL</OS>" & @CRLF & _
						"</SupportedOS>" & @CRLF & _
						"<SupportedPlatform>" & @CRLF & _
						"<Platform>ALL</Platform>" & @CRLF & _
						"</SupportedPlatform>" & @CRLF & _
						"<OSDefault>" & @CRLF & _
						"<OS>ALL</OS>" & @CRLF & _
						"</OSDefault>" & @CRLF & _
						"<PlatformDefault>" & @CRLF & _
						"<Platform>ALL</Platform>" & @CRLF & _
						"</PlatformDefault>" & @CRLF & _
						"</Driver>" & @CRLF & _
						"</CRB>" & @CRLF
				_ReplaceStringInFile($InstallTemp_xml, "</CRB>", $NewString)
			EndIf
		Next
	Next
	WriteMyLog(@ScriptLineNumber & " BuildInstallTEMP_XML complete")
EndFunc   ;==>BuildInstallTEMP_XML
;-----------------------------------------------
;This function loads the config file in XML format and parses it
Func LoadConfig()
	WriteMyLog(@ScriptLineNumber & " LoadConfig")
	Global $ArrayConfigData[1]
	Local $ID, $Name, $InstallFile
	Local $Start = False

	WriteMyLog(@ScriptLineNumber & " LoadConfig " & $Install_xml)
	Local $Results = _FileReadToArray($Install_xml, $ArrayConfigData)
	; Check if file opened for reading OK
	If $Results = 0 Then
		WriteMyLog(@ScriptLineNumber & " LoadCFG: Unable to open file for reading: " & $Install_xml)
		Exit
	EndIf

	TrimArray($ArrayConfigData)

	Local $StartCRB = False
	For $T In $ArrayConfigData
		If StringInStr($T, "<CRB>") > 0 Then $StartCRB = True
		$T = StringReplace($T, "&amp;", "&")
		If $StartCRB Then
			If StringInStr($T, "<DRIVER") > 0 Then
				$Start = True
				;ConsoleWrite(@ScriptLineNumber & " true " & $T & @CRLF)
			EndIf
			If StringInStr($T, "</DRIVER>") > 0 Then
				$Start = False
				;ConsoleWrite(@ScriptLineNumber & " false " & $T & @CRLF)
			EndIf
			; <App Name="SOL - PCI Serial Port" RunCmd="ImMEI_SOL.exe" ......
			If $Start Then
				Local $Y, $Z
				If StringInStr($T, "<App Name=") <> 0 Or StringInStr($T, "<ID>") <> 0 Then
					If StringInStr($T, "<App Name=") <> 0 Then

						$C1 = StringInStr($T, 'App Name=') + StringLen("App Name=")
						$C2 = StringInStr($T, 'RunCmd=')
						$U = StringMid($T, $C1, $C2 - $C1)
						$U = StringReplace($U, '"', '')
						_ArrayAdd($ArrayValidData, "Name>>" & $U) ;this is the Device name

						$C1 = StringInStr($T, 'RunCmd=') + StringLen("RunCmd=")
						$C2 = StringInStr($T, 'Reboot=')
						$U = StringMid($T, $C1, $C2 - $C1)
						$U = StringReplace($U, '"', '')
						_ArrayAdd($ArrayValidData, "RunCmd>>" & $U) ;this is the RunCmd
					ElseIf StringInStr($T, "<ID>") <> 0 Then
						$T = StringReplace($T, "<ID>", "")
						$T = StringReplace($T, "</ID>", "")
						_ArrayAdd($ArrayValidData, $T) ; this is the devid
					EndIf
				EndIf
			EndIf
		EndIf
	Next
	_ArrayDelete($ArrayValidData, 0)
	WriteMyLog(@ScriptLineNumber & " Config file loaded: " & $Install_xml)
EndFunc   ;==>LoadConfig
;-----------------------------------------------
;Gets a list of problem devices using devcon
Func GetListOfProblemDevices()
	WriteMyLog(@ScriptLineNumber & " GetListOfProblemDevices")
	Const $TLOG = "C:\LOGS\ListOfProblemDevices.txt"
	FileDelete($TLOG)
	Global $TArray[1]
	$pid = Run(@ComSpec & " /c " & $DEVCON & " status * ", "", "", $STDOUT_CHILD)
	Sleep(500)
	Local $line
	While 1
		$line = StdoutRead($pid)
		If @error Then ExitLoop
		FileWrite($TLOG, $line)
	WEnd

	_FileReadToArray($TLOG, $TArray)

	Local $Save
	For $T In $TArray
		If StringInStr($T, "    ") = 0 Then $Save = $T
		If StringInStr($T, "Device has a problem: 28") <> 0 Then
			Local $R = StringSplit($Save, "&SUBSYS_", 3)
			_ArrayAdd($ArrayProblemsData, $R[0] & "~" & StringStripWS($T, 3))
		EndIf
	Next
	_ArrayDelete($ArrayProblemsData, 0)
	WriteMyLog(@ScriptLineNumber & " GetListOfProblemDevices complete")
EndFunc   ;==>GetListOfProblemDevices
;-----------------------------------------------
;Removes leading and trailing white spaces from every item in the array
Func TrimArray(ByRef $Array, $Mode = 7)
	Local $count = 0
	While True
		$count += 1
		If $count >= UBound($Array) Then ExitLoop
		$Array[$count] = StringStripWS($Array[$count], $Mode)
	WEnd
EndFunc   ;==>TrimArray
;-----------------------------------------------
;Writes a line to the log file
Func WriteMyLog($StrMsg)
	FileWriteLine($LOGFILE, $StrMsg)
EndFunc   ;==>WriteMyLog
;-----------------------------------------------
; This function verifies that Devcon and the config file exist
Func TestForStuff()
	WriteMyLog(@ScriptLineNumber & " TestForStuff")
	If @OSArch = "X86" Then
		$DEVCON = "C:\BIN\DEVCON.EXE"
	ElseIf @OSArch = "X64" Then
		$DEVCON = "C:\BIN\DEVCON_X64.EXE"
	Else
		WriteMyLog(@ScriptLineNumber & " Unsupported OSArch: " & @OSArch)
		Exit
	EndIf

	If FileExists($Install_xml) = False Then
		WriteMyLog(@ScriptLineNumber & " Test for config file failed. " & $Install_xml & " must exist.")
		Exit
	EndIf
	Return 0
EndFunc   ;==>TestForStuff
;-----------------------------------------------
;This function retrives some system info from the registry
Func GetSystemInfo()
	WriteMyLog(@ScriptLineNumber & " GetSystemInfo")
	$SystemProductName = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Intel\MPG", "PlatformName")
	$OsName = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Intel\MPG", "OsName")
	WriteMyLog(@ScriptLineNumber & " Get system info completed " & $SystemProductName & " " & $OsName)
EndFunc   ;==>GetSystemInfo
;-----------------------------------------------