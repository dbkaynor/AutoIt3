#region
#AutoIt3Wrapper_Run_Au3check= y
#AutoIt3Wrapper_Au3Check_Stop_OnWarning= y
#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Run_Tidy= y
#AutoIt3Wrapper_Tidy_Stop_OnError= y
;#Tidy_Parameters=/gd /sf
#AutoIt3Wrapper_Compression= 4
#AutoIt3Wrapper_Res_Fileversion=1.0.0.8
#AutoIt3Wrapper_Res_FileVersion_AutoIncrement= y
#AutoIt3Wrapper_Res_Description= USBLost
#AutoIt3Wrapper_Res_LegalCopyright= GNU-PL
#AutoIt3Wrapper_Res_Comment= Lost USB alert
#AutoIt3Wrapper_Res_Field= Developer|Douglas Kaynor
#AutoIt3Wrapper_Res_LegalCopyright= Copyright © 2011 Douglas B Kaynor
#AutoIt3Wrapper_Res_Field= AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Res_Field= Compile date|%longdate% %time%
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Run_Before=
#AutoIt3Wrapper_Icon=DBK.ico
#endregion

Opt("MustDeclareVars", 1)

#include <Misc.au3>
#include "_DougFunctions.au3"

If _Singleton(@ScriptName, 1) = 0 Then
	_Debug(@ScriptName & " is already running!", 0x40, 5)
	Exit
EndIf
#NoTrayIcon

Global $Message1
Global $Message2
Global $Message3
Global $Message4
Global $Title

$Title = "Help! I'm Lost!"
$Message1 = "I've been lost and my owner would love to get my data back."
$Message2 = "Please email me at: doug@kaynor.net"
$Message3 = "Thanks for your honesty in advance. - Doug K"
;_Debug($Title)
;_Debug($Message1)
;_Debug($Message2)
;_Debug($Message3)
;_Debug($Message4)
MsgBox(266304, $Title, $Message1 & @CRLF & $Message2 & @CRLF & $Message3, 5)
Run("explorer.exe " & @ScriptDir)