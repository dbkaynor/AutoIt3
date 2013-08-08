#region
#AutoIt3Wrapper_Run_Au3check=y
#AutoIt3Wrapper_Au3Check_Stop_OnWarning=y
#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Tidy_Stop_OnError=y
;#Tidy_Parameters=/gd /sf
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Fileversion=1.0.0.29
#AutoIt3Wrapper_Res_FileVersion_AutoIncrement=Y
#AutoIt3Wrapper_Res_Description=DMANG
#AutoIt3Wrapper_Res_LegalCopyright=GNU-PL
#AutoIt3Wrapper_Res_Comment=A program to map my favorite drives
#AutoIt3Wrapper_Res_Field=Developer|Douglas Kaynor
#AutoIt3Wrapper_Res_LegalCopyright=Copyright 2010 Douglas B Kaynor
#AutoIt3Wrapper_Res_Field= AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Res_Field=Compile date|%longdate% %time%
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Run_Before=
#AutoIt3Wrapper_Run_Obfuscator=n
#AutoIt3Wrapper_UseX64=n
#Obfuscator_Parameters= /Convert_Strings=0 /Convert_Numerics=0 /showconsoleinfo=9
#AutoIt3Wrapper_Icon=../icons/Sri_Aravan.ico
#endregion

Global $User = "amr\dbkaynox"
Global $Password = "Autoit3@"

Mapit("J:", "\\Lestat1\Automa", $User, $Password)
Mapit("K:", "\\Lestat1\Drivers", $User, $Password)
Mapit("L:", "\\Lestat1\Preload", $User, $Password)
Mapit("S:", "\\chakotay\softval", $User, $Password)
Mapit("T:", "\\chakotay\temp\dbkaynox", $User, $Password)
Exit

Func Mapit($Drive, $Path, $User, $PWD)
	Local $T
	DriveMapDel($Drive)
	If DriveMapAdd($Drive, $Path, 9, $User, $PWD) <> 1 Then
		Switch @error
			Case 1
				$T = "Undefined / Other error. " & @extended
			Case 2
				$T = "Access to the remote share was denied"
			Case 3
				$T = "The device is already assigned"
			Case 4
				$T = "Invalid device name"
			Case 5
				$T = "Invalid remote share"
			Case 6
				$T = "Invalid password"
			Case Else
				$T = "Unknown error returned " & @error
		EndSwitch
		MsgBox(0, "DriveMap failed", $Drive & @CRLF & $Path & @CRLF & $T, 6)
	Else
		MsgBox(0, "Drive map success", $Drive & " is mapped to" & DriveMapGet($Drive), 0.3)
	EndIf
EndFunc   ;==>Mapit