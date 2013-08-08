#comments-start HEADER
;******************************************************************************************
;** Intel Corporation, MPG MPAD
;** Title			:  		incNet.au3
;** Description	: 	
;**		Network related include file.
;**
;** Revision: 	Rev 2.0.8
;******************************************************************************************
;******************************************************************************************
;** Revision History:
;** 
;** Update for Rev 2.0.0		- Dick Lin 03/02/2006
;**	- Initial release
;**
;** Update for Rev 2.0.1		- Dick Lin 07/19/2006
;**	- Use M: for all of driver mapping functions.
;**
;** Update for Rev 2.0.2		- Dick Lin 08/21/2006
;**	- Return value for MapLicensedDriveMTN() function.
;**
;** Update for Rev 2.0.2		- Dick Lin 04/26/2007
;**	- RunNetUseCommand() check available drive letter before mapping.
;**
;** Update for Rev 2.0.3		- Dick Lin 06/26/2007
;**		- Use $BS_DEFPUSHBUTTON to replace _IsPressed() function.
;**
;** Update for Rev 2.0.4		- Dick Lin 07/12/2007
;**		- DriveMapAdd() replace net use command.
;**
;** Update for Rev 2.0.5		- Dick Lin 09/05/2007
;**		- Fixed CopyFromServer() issue with destDir == @HomeDrive.
;**
;** Update for Rev 2.0.6		- Dick Lin 10/17/2007
;**		- Added InstallMontevinaLAN() function.
;**		- IsLANDriverInstalled
;**
;** Update for Rev 2.0.7		- Jarek Szymanski 8/28/2008
;**		- Fixed some spelling errors
;**
;** Update for Rev 2.0.8		- Jarek Szymanski 4/2/2009
;**		- moved to AutoIt 3.3.0.0
;**		- removed Opt("RunErrorsFatal", 1) for InstallMontevinaLAN() - moving to AutoIt 3.3.0.0 - feature not supported in current version
;**		- added #include: <EditConstants.au3>; <ButtonConstants.au3>; <WindowsConstants.au3> - moving to AutoIt 3.3.0.0
;**
;******************************************************************************************
#comments-end HEADER

#include-once 
#include <GUIConstants.au3>
#include <EditConstants.au3> ;added for AutoIt 3.3.0.0
#include <ButtonConstants.au3> ;added for AutoIt 3.3.0.0
#include <WindowsConstants.au3> ;added for AutoIt 3.3.0.0
#Include <Misc.au3>

;***************************************************************************
;** Function: 		DirectoryCopy($scriptName, $sourceDir, $destDir)
;** Parameters:
;**		$strScriptName 	- script name call this function
;**		$strSourceDir 	- source directory to copy
;**		$strDestDir		- local destination directory to store files
;** Description: 				 
;**		This function is called to copy file from server.
;** Return:
;**		1 for no error occured, else 0.
;**
;***************************************************************************
Func DirectoryCopy($scriptName, $sourceDir, $destDir)
	; Concate source/destination dir
	SplashTextOn($scriptName, "Copying files from server, please be patient...", 450, 50, -1, -1, 2, "", 10)	
	$result = DirCopy($sourceDir, $destDir, 1)
	SplashOff()
	
	; Unmap all network drives
	UnmapNetworkDrives()
	
	If Not $result Then
		WriteLog("DirCopy() function failed")
		Return 0
	Else
		Return 1
	EndIf	
EndFunc

Func RoboDirCopy($scriptName, $sourceDir, $destDir)
	$strMsg = "Copying files, please be patient..."
	
	$strRcpy = @HomeDrive & "\bin\Robocopy.exe"
	$strLog  = @HomeDrive & "\logs\VMCopyLog.txt"
	$strFile = $strRcpy & " " & $sourceDir & " " & $destDir & " /E /XA:H /Log:" & $strLog & " /A-:R /TEE"
	SplashTextOn($scriptName, $strMsg, 450, 50, -1, -1, 2, "", 10)
	RunWait(@ComSpec & " /c " & $strFile)
	SplashOff()	

	Return 1
EndFunc

;***************************************************************************
;** Function: 		GetPreloadMediaVersion()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to get preload media version from registry.
;** Return:
;**		Media version string from registry.
;**
;***************************************************************************
Func GetPreloadMediaVersion()
	$strKey = "HKLM\SOFTWARE\Intel\MPG"
	$strSubKey = "PreloadMediaVersion"
	$var = RegRead($strKey, $strSubKey)
	
	If $var == "" Then
		Return "UNKNOWN"
	Else
		Return $var
	EndIf
EndFunc

;***************************************************************************
;** Function: 		GetDVDDrive()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to get preload media version from registry.
;** Return:
;**		Media version string from registry.
;**
;***************************************************************************
Func GetDVDDrive()
	$var = DriveGetDrive( "all" )
	If NOT @error Then
		;MsgBox(4096,"", "Found " & $var[0] & " drives")
		For $i = 1 To $var[0]
			;MsgBox(4096,"Drive " & $i, $var[$i])
			
			If StringUpper(DriveGetType($var[$i])) == "CDROM" Then
				Return $var[$i]
			EndIf
		Next
	EndIf
	
	Return "UNKNOWN"
EndFunc	

;***************************************************************************
;** Function: 		GetDVDLabel()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to get preload media version from registry.
;** Return:
;**		Media version string from registry.
;**
;***************************************************************************
Func GetDVDLabel($strDrive)
	$var = DriveGetLabel( $strDrive )
	
	Return $var
EndFunc

;***************************************************************************
;** Function: 		GetCorrectMedia()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to copy file from server.
;** Return:
;**		1 for no error occured, else 0.
;**
;***************************************************************************
Func CheckMediaVersion()
	$strPreloadMediaVersion = GetPreloadMediaVersion()
	If StringUpper($strPreloadMediaVersion) == "UNKNOWN" Then
		MsgBox(0, "ERROR", "Can't find PreloadMediaVersion registry.", 2)			
		Return 0
	EndIf
	
	$strDVDDrive = GetDVDDrive()
	$strDVDName = GetDVDLabel($strDVDDrive)
	While StringUpper($strPreloadMediaVersion) <> StringUpper($strDVDName)
		WriteLog(StringUpper($strPreloadMediaVersion) & StringUpper($strDVDName))
		
		; Close DVD window
		$strWindow = "(" & StringUpper($strDVDDrive) & ")"
		WinClose($strWindow)
		
		$strMsg = "Please insert the Preload Software DVD Volume Label " & $strPreloadMediaVersion &  @CRLF & "Press Retry to continue."
		$buttonPress = MsgBox(0x05, "ERROR - Wrong DVD", $strMsg)
		If $buttonPress == 2 Then
			Return 0
		EndIf
		
		; Get DVD label again
		$strDVDName = GetDVDLabel($strDVDDrive)
	WEnd		

	return 1
EndFunc

;***************************************************************************
;** Function: 		CopyFromServer(strScriptName, strSourceDir, strDestDir)
;** Parameters:
;**		strScriptName 	- script name call this function
;**		strSourceDir 	- source directory to copy
;**		strDestDir		- local destination directory to store files
;** Description: 				 
;**		This function is called to copy file from server.
;** Return:
;**		1 for no error occured, else 0.
;**
;** Usage - 
;**
;**	strSourceDir = "APPS\WinPM"
;**	strDestDir 	 = "APPS\WinPM"
;**	if !FileExist(strFile) then
;**		if !CopyFromServer(SCRIPT_NAME, strSourceDir, strDestDir)		
;**			...
;**		endif
;**	endif
;***************************************************************************
Func CopyFromServer($scriptName, $sourceDir,  $destDir)
		
		
	; Concate source/destination dir
	SplashTextOn($scriptName, "Copying files from server, please be patient...", 450, 50, -1, -1, 2, "", 10)	
	
	Opt("WinTitleMatchMode", 2)     ;1=start, 2=subStr, 3=exact, 4=advanced
	
	; Get DVD drive letter/label
	$strDVDDrive = GetDVDDrive()
	$strDVDName = GetDVDLabel($strDVDDrive)
	
	; WriteLog
	WriteLog("CopyFromServer() function.")
		
	; Get PreloadServerName/IP
	$strPreloadServer = GetPreLoadServerName()
	$strPreloadServerIP = GetPreloadServerIPAddress()
	
	; For DVD Preload
	$strPreloadtype = GetPreLoadType()
	If StringUpper($strPreloadServer) == "DVD" And StringUpper($strPreloadtype) == "DVD" Then
		If Not CheckMediaVersion() Then				
			; Close DVD window
			$strWindow = $strDVDDrive
			WinClose($strWindow)
			Return 0
		EndIf
	EndIf

	; NOTES: USE IP INSTEAD OF SERVER NAME FOR MAPPING
	; http://www.ultrabac.com/kb8/ubq000165.htm	
	; Summary:
	; 	If you get the error "The credentials supplied conflict with an existing set of credentials - 1219", 
	;	it's a bug in NT that very few people will encounter.
	; Details:
	; There are two articles on TechNet that discuss this. The second one has the actual solution 
	;	- use the IP address of the computer you're trying to back up instead of the computer name, 
	;		and it will force it to use TCP/IP protocol.
	; Article ID: Q15473 "Connect Network Drive" Caches First Credentials Supplied
	
	If StringUpper($strPreloadServer) == "DVD" And StringUpper($strPreloadtype) == "DVD" Then
		$sourceDir = $strDVDDrive & $sourceDir
		If ($destDir == @HomeDrive) Then
			$destDir = @HomeDrive & "\"
		Else
			$destDir = @HomeDrive & $destDir	
		EndIf
	Else
		$sourceDir = "\\" & $strPreloadServerIP & $sourceDir
		If ($destDir == @HomeDrive) Then
			$destDir = @HomeDrive & "\"
		Else
			$destDir = @HomeDrive & $destDir
		EndIf
		; Map drive
		$strMapDrive = "M:"
		$result = MapNetworkDrive($strMapDrive, $sourceDir)
		If Not $result Then
			Return 0
		EndIf
	EndIf

	; WriteLog
	WriteLog("CopyFromServer: " & $sourceDir & " " & $destDir)
	
	; CopyDir
	If ($scriptName == "VManager Launcher") Then		
		If Not RoboDirCopy($scriptName, $sourceDir,  $destDir) Then	
			Return 0
		EndIf
	Else
		If Not DirectoryCopy($scriptName, $sourceDir,  $destDir) Then			
			Return 0
		EndIf
	EndIf
	
	SplashOff()		
	
	Return 1
	
EndFunc

;***************************************************************************
;** Function: 		UnmapNetworkDrives()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to unmap all possible used mapped drive.
;** Return:
;**		1 for no error occured, else 0
;** Usage:
;**		blnResult = UnmapAllDrives()
;***************************************************************************
Func UnmapNetworkDrives()

	; Writelog
	WriteLog("UnmapNetworkDrives function.")

	$cmd = " /c NET USE * /DELETE /YES" 
	$val = RunWait(@ComSpec & $cmd, "", @SW_HIDE)

	WriteLog("UnmapNetworkDrives() Net use result : " & $val)
	
	; Get DVD drive letter/label
	$strDVDDrive = GetDVDDrive()
	
	; Close DVD window
	$strWindow = "(" & StringUpper($strDVDDrive) & ")"
	WinClose($strWindow)
		
	Return 1
	
EndFunc

;***************************************************************************
;** Function: 		MapNetworkDrive(ByRef $strMapDrive, $sourceDir) 	
;** Parameters:
;**		$sourceDir 	- source directory for wntAddDrive() function
;** Description: 				 
;**		This function is called by CopyFromServer() to map to server
;**		use MTN logon ID/password.
;** Return:
;**		1 for no error occured, else 0
;**
;** Usage
;**		MapNetworkDrive($sourceDir)
;**	
;***************************************************************************
Func MapNetworkDrive(ByRef $strMapDrive, $sourceDir) 	
	; WriteLog
	WriteLog("MapNetworkDrive() function.")
	
	; Unmap before starting mapping.
	UnmapNetworkDrives()
	
	; Map to network drive
	$strID = "MTN\script"
	$strPassword = "in84tel2"
	
	; Splash
	SplashTextOn("Mapping network drive.", "Please be patient...", 450, 50, -1, -1, 2, "", 10)	
	
	; Map to network drive
	;$val = RunNetUseCommand($strMapDrive, $sourceDir, $strPassword, $strID)
	$strMapDrive = DriveMapAdd("*", $sourceDir, 0, $strID, $strPassword)

	; Splash Off
	SplashOff()

	;WriteLog("MapNetworkDrive() Net use result : " & $val)
	;If $val == 0 Then
	If $strMapDrive == "" Then
		Return False
	Else
		Return True
	EndIf	
EndFunc	

;***************************************************************************
;** Function: 		MapLicensedDrive(ByRef $strMapDrive, $sourceDir, $strID, $strPassword) 	
;** Parameters:
;**		$sourceDir 	- source directory for wntAddDrive() function
;** Description: 				 
;**		This function is called by CopyFromServer() to map to server
;**		use MTN logon ID/password.
;** Return:
;**		1 for no error occured, else 0
;**
;** Usage
;**		MapLicensedDrive(ByRef $strMapDrive, $sourceDir, $strID, $strPassword) 	
;**	
;***************************************************************************
Func MapLicensedDrive(ByRef $strMapDrive, $sourceDir, $strID, $strPassword) 	
	; WriteLog
	WriteLog("MapLicensedDrive() function.")
	
	; Unmap before starting mapping.
	UnmapNetworkDrives()	
	
	; Splash
	SplashTextOn("Mapping network drive.", "Please be patient...", 450, 50, -1, -1, 2, "", 10)
	
	; Map to network drive
	;$val = RunNetUseCommand($strMapDrive, $sourceDir, $strPassword, $strID)
	$strMapDrive = DriveMapAdd("*", $sourceDir, 0, $strID, $strPassword)
			
	; Splash off
	SplashOff()
	
	;WriteLog("Net use result : " & $val)
	;If $val == 0 Then
	If $strMapDrive == "" Then
		Return False
	Else
		UpdateRegistryCredential()
		Return True
	EndIf	
	
EndFunc	

;***************************************************************************
;** Function: 		MapLicensedDriveMTNCredential($sourceDir) 	
;** Parameters:
;**		$sourceDir 	- source directory for wntAddDrive() function
;** Description: 				 
;**		This function is called by CopyFromServer() to map to server
;**		use MTN logon ID/password.
;** Return:
;**		1 for no error occured, else 0
;**
;** Usage
;**		MapLicensedDriveMTN(ByRef $strMapDrive, $sourceDir) 
;**	
;***************************************************************************
Func MapLicensedDriveMTN(ByRef $strMapDrive, $sourceDir) 	
	
	; WriteLog
	WriteLog("MapLicensedDriveMTN() function.")
	
	; Unmap before starting mapping.
	UnmapNetworkDrives()
	
	; Map to network drive
	$strID = "MTN\authscript"
	$strPassword = "aut91hr3"
		
	; Splash
	SplashTextOn("Mapping network drive.", "Please be patient...", 450, 50, -1, -1, 2, "", 10)
		
	; Map to network drive
	;$val = RunNetUseCommand($strMapDrive, $sourceDir, $strPassword, $strID)
	$strMapDrive = DriveMapAdd("*", $sourceDir, 0, $strID, $strPassword)

	; Splash off
	SplashOff()
	;WriteLog("Net use result : " & $val)
	;If $val == 0 Then
	If $strMapDrive == "" Then
		Return False
	Else
		Return True
	EndIf	
	
EndFunc	

;***************************************************************************
;** Function: 		GetUserCredential($sourceDir, ByRef $strID, ByRef $strPassword, ByRef $action)
;** Parameters:
;**		strScriptName 	- script name call this function
;**		strSourceDir 	- source directory to copy
;**		strDestDir		- local destination directory to store files
;** Description: 				 
;**		This function is called to copy file from server.
;** Return:
;**		1 for no error occured, else 0.
;**
;***************************************************************************
Global $btnOK
Global $msg
Func GetUserCredential($sourceDir, ByRef $strID, ByRef $strPassword, ByRef $action)
	;HotKeySet("{ENTER}", "_btnSubmit")
	
	$domain = "AMR"
	$strID = ""
	$strPassword = ""

	; == GUI generated with Koda ==);
	$frmLogon = GUICreate("NT Login ID/Password", 297, 258, 192, 125)
	$Group1 = GUICtrlCreateGroup("Select network domain", 16, 16, 265, 57)
	$radAMR = GUICtrlCreateRadio("AMR", 32, 40, 49, 17)
	GUICtrlSetState(-1, $GUI_CHECKED)
	$radGAR = GUICtrlCreateRadio("GAR", 88, 40, 65, 17)
	$radGER = GUICtrlCreateRadio("GER", 152, 40, 65, 17)
	$radCCR = GUICtrlCreateRadio("CCR", 216, 40, 57, 17)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	$btnOK = GUICtrlCreateButton("OK", 208, 176, 75, 25, $BS_DEFPUSHBUTTON)
	$btnCancel = GUICtrlCreateButton("Cancel", 208, 216, 75, 25, 0)
	$Label1	= GUICtrlCreateLabel("Enter your login ID:", 16, 96, 94, 17)
	$Label2 = GUICtrlCreateLabel("Enter your password:", 16, 144, 103, 17)
	$txtID = GUICtrlCreateInput("Your NT Login ID", 128, 88, 137, 21)
	GUICtrlSetState(-1, $GUI_FOCUS)
	GUICtrlSetState(-1, $GUI_ACCEPTFILES)
	$txtPassword = GUICtrlCreateInput("Password", 128, 136, 137, 21, $ES_PASSWORD)
	GUICtrlSetState(-1, $GUI_ACCEPTFILES)	
	GUISetState(@SW_SHOW)

	Do
		$msg = GuiGetMsg()
		GUISetState(@SW_SHOW)
		
		Switch $msg
			Case $btnOK		
				$strID = GUICtrlRead($txtID)	
				$strPassword = GUICtrlRead($txtPassword)
				$strID = $domain & "\" & $strID
				GUISetState(@SW_HIDE)
				$strMapDrive = "M:"
				If Not MapLicensedDrive($strMapDrive, $sourceDir, $strID, $strPassword) Then
					$answer = MsgBox(0x01, "WRONG UserID/Password", 'Please press "OK" to retry, "Cancel" to quit.')
					If $answer == 2 Then
						$msg = $GUI_EVENT_CLOSE
						$action = 0
					EndIf
				Else
					$msg = $GUI_EVENT_CLOSE
					$action = 1
				EndIf
				GUICtrlSetState($txtID, $GUI_FOCUS)
			Case $btnCancel
				$msg = $GUI_EVENT_CLOSE
				$action = 0
			Case $radAMR
				$domain = "AMR"
			Case $radGAR
				$domain = "GAR"
			Case $radGER
				$domain = "GER"
			Case $radCCR
				$domain = "CCR"

		EndSwitch
			
	Until $msg = $GUI_EVENT_CLOSE
	
	GUISetState(@SW_HIDE)
		
EndFunc

Func _btnSubmit()
    $msg = $btnOK
EndFunc

;***************************************************************************
;** Function: 		CopyFromLicensedServer(strScriptName, strSourceDir, strDestDir)
;** Parameters:
;**		strScriptName 	- script name call this function
;**		strSourceDir 	- source directory to copy
;**		strDestDir		- local destination directory to store files
;** Description: 				 
;**		This function is called to copy file from server.
;** Return:
;**		1 for no error occured, else 0.
;**
;** Usage - 
;**
;**		$strSourceDir = "\APPS\Licensed_APPS\SiSoftware_Sandra_2005_Pro"
;**		$strDestDir   = "\APPS\Licensed_APPS\Sandra"
;**		if !FileExist(strFile) then
;**			if !CopyFromLicensedServer(SCRIPT_NAME, strSourceDir, strDestDir)		
;**				...
;**			endif
;**		endif
;**
;***************************************************************************
Func CopyFromLicensedServer($scriptName, $sourceDir, $destDir)
		
	; Writelog
	WriteLog("CopyFromLicensedServer() function.")
		
	; Get PreloadServer Name and PreloadType from Registry.
	$strPreloadType   = StringUpper(GetPreLoadType())
	$strPreloadServer = StringUpper(GetPreLoadServerName())
	$strPreloadServerIP = GetPreloadServerIPAddress()
	
	If (($strPreloadServer == "DVD") And ($strPreloadtype == "DVD")) Then
		ShowMessage("DVD Preload Issue", "DVD doesn't have licensed applications.")
		Return 0
	EndIf
		
	; NOTES: USE IP INSTEAD OF SERVER NAME FOR MAPPING
	; http://www.ultrabac.com/kb8/ubq000165.htm	
	; Summary:
	; 	If you get the error "The credentials supplied conflict with an existing set of credentials - 1219", 
	;	it's a bug in NT that very few people will encounter.
	; Details:
	; There are two articles on TechNet that discuss this. The second one has the actual solution 
	;	- use the IP address of the computer you're trying to back up instead of the computer name, 
	;		and it will force it to use TCP/IP protocol.
	; Article ID: Q15473 "Connect Network Drive" Caches First Credentials Supplied
	$sourceDir = "\\" & $strPreloadServerIP & $sourceDir
	$destDir = @HomeDrive & $destDir		
		
	; WriteLog
	WriteLog("CopyFromLicensedServer: " & $sourceDir & " " & $destDir)	
	
	; Loop to make sure we got correct user credential.
	If Not CheckCredential() Then
		$strID = ""
		$strPassword = ""
		$action = 0
		While 1			
			GetUserCredential($sourceDir, $strID, $strPassword, $action)
			If $action == 0 Then
				Return 0
			ElseIf $action == 1 Then
				ExitLoop
			EndIf
		WEnd		
	EndIf
	
	; Map drive
	$strMapDrive = "M:"
	$result = MapLicensedDriveMTN($strMapDrive, $sourceDir)
	if Not $result Then
		Return 0
	EndIf
	
	; CopyDir
	If Not DirectoryCopy($scriptName, $sourceDir,  $destDir) Then
		Return 0
	EndIf
		
	Return 1
	
EndFunc

;***************************************************************************
;** Function: 		CopyFromLicensedServer(strScriptName, strSourceDir, strDestDir)
;** Parameters:
;**		strScriptName 	- script name call this function
;**		strSourceDir 	- source directory to copy
;**		strDestDir		- local destination directory to store files
;** Description: 				 
;**		This function is called to copy file from server.
;** Return:
;**		1 for no error occured, else 0.
;**
;** Usage - 
;**
;**		$strSourceDir = "\APPS\Licensed_APPS\SiSoftware_Sandra_2005_Pro"
;**		$strDestDir   = "\APPS\Licensed_APPS\Sandra"
;**		if !FileExist(strFile) then
;**			if !CopyFromLicensedServer(SCRIPT_NAME, strSourceDir, strDestDir)		
;**				...
;**			endif
;**		endif
;**
;***************************************************************************
Func MapToLicensedServer($scriptName, $sourceDir, $destDir, $drive)

EndFunc

;***************************************************************************
;** Function: 		HostToIp($hostName)
;** Parameters:
;**		$hostName		- host computer name want get IP address
;** Description: 				 
;**		This function is called to copy file from server.
;** Return:
;**		IP address for the host computer
;**
;***************************************************************************
Func HostToIp($hostName)
	
	$ipAddress = ""
	
	$junkFile = @HomeDrive & "\Scripts\junk.txt"
 	$cmd = " /c ping " & $hostName & " > " & $junkFile
 	$val = RunWait(@ComSpec & $cmd, "", @SW_HIDE)
	
	$file = FileOpen($junkFile, 0)

	While 1
		$line = FileReadLine($file)
		If @error = -1 Then ExitLoop

		$split = StringSplit($line, " ")
		For $i = 1 To $split[0] Step 1
			If StringInStr ( $split[$i], "[") Then
				$text = StringReplace($split[$i], "[", "")
				$ipAddress = StringReplace($text, "]", "")
				FileClose($file)
				FileDelete($junkFile)
				Return $ipAddress
			EndIf
		Next		
	Wend
	
	FileClose($file)
	FileDelete($junkFile)
	
	Return $ipAddress
EndFunc	

;***************************************************************************
;** Function: 		CheckIPAddress($ipAddress)
;** Parameters:
;**		$ipAddress		- IP address to check
;** Description: 				 
;**		This function is called to check if the input string has correct IP addresss.
;** Return:
;**		1 for no error, otherwise 0.
;**
;***************************************************************************
Func CheckIPAddress($ipAddress)
	;Check string is no longer than 15 characters
	If StringLen($ipAddress) > 15 Then
	  ;Message("NOTE","Not a valid ip address")
	  Return 0
	EndIf
	
	;Check there are 4 octets in the string
	$OctetCount = StringInStr($ipAddress, '.')
	If Not $OctetCount == 3 Then
	  ;MsgBox(0, 'Error', 'Malformed IP address')
	  Return 0
	EndIf

	$Octet = StringSplit($ipAddress, ".")
	For $index = 1 To $Octet[0]
		;MsgBox(0, $ipAddress, $Octet[$index])
		If (($Octet[$index] < 0) Or ($Octet[$index] > 255) And IsNumber($Octet[$index])) Then
			;MsgBox(0, 'Not Valid IP Address', $Octet[$index])
			Return 0
		EndIf
	Next
	
   Return 1
EndFunc

;***************************************************************************
;** Function: 		GetPreloadServerIPAddress()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to get PreloadServer IP address
;** Return:
;**		IP address of preload server, or "" for not found.
;**
;***************************************************************************
Func GetPreloadServerIPAddress()
	$strPreloadServer = GetPreLoadServerName()
	$strPreloadType = GetPreLoadType()

	If (StringUpper($strPreloadServer) == "DVD") And (StringUpper($strPreloadType) == "DVD") Then
		Return "DVD"
	EndIf
	
	;Check there are 4 octets in the string
	$intOctet = StringInStr($strPreloadServer, '.')
	If $intOctet > 1 Then
		If CheckIPAddress($strPreloadServer) Then
			Return $strPreloadServer
		Else
			Return ""
		EndIf
	EndIf

	Return HostToIp($strPreloadServer)

EndFunc	

;***************************************************************************
;** Function: 		UpdateRegistryCredential()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called write registry LISO GD.
;** Return:
;**		None
;**
;***************************************************************************
Func UpdateRegistryCredential()
	$strKey = "HKLM\SOFTWARE\Intel\MPG\TESTAPP"
	$strSubKey = "LISO"
	$var = RegWrite($strKey, $strSubKey, "REG_SZ", "GD")
EndFunc

;***************************************************************************
;** Function: 		CheckCredential()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to check if regity has user's credential GD.
;** Return:
;**		1 for GD, otherwiese 0.
;**
;***************************************************************************
Func CheckCredential()

	$strKey = "HKLM\SOFTWARE\Intel\MPG\TESTAPP"
	$strSubKey = "LISO"
	$var = RegRead($strKey, $strSubKey)
	;Writelog("TESTING:" & $var)
	If StringUpper($var) == "GD" Then	
		Return 1
	Else
		Return 0
	EndIf
EndFunc

;***************************************************************************
;** Function: 		RunNetUseCommand($strDrive, $sourceDir, $strPassword, $strID)
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called run DOS external command NET USE.
;** Return:
;**		None.
;**
;***************************************************************************
Func RunNetUseCommand(ByRef $strDrive, $sourceDir, $strPassword, $strID)
	
	; Check if drive in use
	Dim $strDriveArray
	$strDriveArray = _ArrayCreate("M:", "N:", "O:", "P:", "Q:", "R;", "S:", "T:", "U:", "V:", "W:", "X:", "Y:", "Z:", _
								"G:", "H:", "I:", "J:", "K:", "L:")

	$strLogicalDisk = ""
	$strDiskName = ""
	WMIGetLogicalDiskName($strLogicalDisk, $strDiskName)
	
	$blnFoundDrive = False
	For $intIndex = 0 To UBound($strDriveArray) - 1
		If Not StringInStr ($strLogicalDisk, $strDriveArray[$intIndex]) Then
			$strDrive = $strDriveArray[$intIndex]
			$blnFoundDrive = True
			ExitLoop
		EndIf
	Next
	
	If Not $blnFoundDrive Then
		MsgBox(0, "ERROR", "Can't find avaialbe logical drive for mapping.")
		Return False
	EndIf
	
	$cmd = " /c NET USE " & $strDrive & ' "' & $sourceDir & '"' & ' "' & $strPassword & '" ' & "/USER:" & '"' & $strID & '"'
	
	$val = 1
	For $index = 1 To 5
		$val = RunWait(@ComSpec & $cmd, "", @SW_HIDE)
		;Writelog($cmd & " RESULT " & $val)
		If $val == 0 Then
			ExitLoop
		Else
			Sleep(20)
		EndIf
	Next
	
	
	Return $val
EndFunc	

;***************************************************************************
;** Function: 		MapToServer(strScriptName, strSourceDir, strDestDir)
;** Parameters:
;**		strScriptName 	- script name call this function
;**		strSourceDir 	- source directory to copy
;**		strDestDir		- local destination directory to store files
;** Description: 				 
;**		This function is called to copy file from server.
;** Return:
;**		1 for no error occured, else 0.
;**
;** Usage - 
;**
;**	strSourceDir = "APPS\WinPM"
;**	strDestDir 	 = "APPS\WinPM"
;**	if !FileExist(strFile) then
;**		if !CopyFromServer(SCRIPT_NAME, strSourceDir, strDestDir)		
;**			...
;**		endif
;**	endif
;***************************************************************************
Func MapToServer($scriptName, $sourceDir, $destDir, ByRef $strMapDrive)
				
	; Concate source/destination dir
	SplashTextOn($scriptName, "Connecting to server. Please be patient...", 450, 50, -1, -1, 2, "", 10)	
	
	Opt("WinTitleMatchMode", 2)     ;1=start, 2=subStr, 3=exact, 4=advanced
	
	; Get DVD drive letter/label
	$strDVDDrive = GetDVDDrive()
	$strDVDName = GetDVDLabel($strDVDDrive)
	
	; WriteLog
	WriteLog("CopyFromServer() function.")
		
	; Get PreloadServerName/IP
	$strPreloadServer = GetPreLoadServerName()
	$strPreloadServerIP = GetPreloadServerIPAddress()
	
	; For DVD Preload
	$strPreloadtype = GetPreLoadType()
	If StringUpper($strPreloadServer) == "DVD" And StringUpper($strPreloadtype) == "DVD" Then
		If Not CheckMediaVersion() Then				
			; Close DVD window
			$strWindow = $strDVDDrive
			WinClose($strWindow)
			Return 0
		EndIf
	EndIf

	; NOTES: USE IP INSTEAD OF SERVER NAME FOR MAPPING
	; http://www.ultrabac.com/kb8/ubq000165.htm	
	; Summary:
	; 	If you get the error "The credentials supplied conflict with an existing set of credentials - 1219", 
	;	it's a bug in NT that very few people will encounter.
	; Details:
	; There are two articles on TechNet that discuss this. The second one has the actual solution 
	;	- use the IP address of the computer you're trying to back up instead of the computer name, 
	;		and it will force it to use TCP/IP protocol.
	; Article ID: Q15473 "Connect Network Drive" Caches First Credentials Supplied
	
	If StringUpper($strPreloadServer) == "DVD" And StringUpper($strPreloadtype) == "DVD" Then
		$sourceDir = $strDVDDrive & $sourceDir
		$destDir = @HomeDrive & $destDir		
	Else
		$sourceDir = "\\" & $strPreloadServerIP & $sourceDir
		$destDir = @HomeDrive & $destDir
		
		; Map drive
		$strMapDrive = "M:"
		$result = MapNetworkDrive($strMapDrive, $sourceDir)
		If Not $result Then
			Return False
		EndIf
	EndIf

	
	Return True
	
EndFunc

;***************************************************************************
;** Function: 		IsLANDriverInstalled()
;** Parameters:
;**	None
;** Description: 				 
;**	This function is called to check LAN driver installation.
;** Return:
;**	None
;***************************************************************************
Func IsLANDriverInstalled()
	; RunNowStatus
	If IsWinVista64SP1() Then
		$strKey    = "HKLM64\SOFTWARE\Intel\Prounstl\SupportedDevices\8086"
	Else
		$strKey    = "HKLM\SOFTWARE\Intel\Prounstl\SupportedDevices\8086"
	EndIf
	$strSubKey = "10BF"
	$strStatus = RegRead($strKey, $strSubKey)
	
	If $strStatus <> "" Then
		Return True
	Else
		Return False
	EndIf
EndFunc

;***************************************************************************
;** Function: 		InstallMontevinaLAN()
;** Parameters:
;**	None
;** Description: 				 
;**	This function is called to run setup.
;** Return:
;**	None
;***************************************************************************
Func InstallMontevinaLAN()
	
	; Set WinTitleMatchMode to 2 - 1=start, 2=subStr, 3=exact, 4=advanced
	Opt("WinTitleMatchMode", 2)     
	
	; Set RunErrorsFatal to 1 - 1=fatal, 0=silent set @error
	;Opt("RunErrorsFatal", 1) ;removed for AutoIt 3.3.0.0
	
	; Run setup
	If IsWinVista32SP1() Then
		$strFile = @HomeDrive & "\Drivers\NIC\intel\apps\setup\SETUPBD\Vista32\SetupBD.exe"
		$strDir  = @HomeDrive & "\Drivers\NIC\intel\apps\setup\SETUPBD\Vista32"
	Else
		$strFile = @HomeDrive & "\Drivers\NIC\intel\apps\setup\SETUPBD\Vistax64\SetupBD.exe"
		$strDir  = @HomeDrive & "\Drivers\NIC\intel\apps\setup\SETUPBD\Vistax64"
	EndIf
	If Not FileExists($strFile) Then
		Return False
	EndIf
	;Run($strFile, $strDir)
	Run(@ComSpec & " /c " & '"' & $strFile & '"', "", @SW_HIDE)
	
	; Installing Drivers
	$strTitle = "Installing Drivers"
	$strText  = "Install or update drivers for Intel® Network Connections."
	If Not WinWait($strTitle, $strText, 50) Then
		WriteLog("Step 1 of 2")
		Return False
	EndIf
 	ControlClick($strTitle, "OK", "OK")
	
	; Setup Progress
	$strText  = "Installing Intel(R) PRO Network Connections"	
	While True
		; For Vista ONLY.
		$strSecTitle = "Windows Security"
		If WinExists($strSecTitle) Then
			ControlFocus($strSecTitle, "&Install", "&Install")
			ControlClick($strSecTitle, "&Install", "&Install")
		EndIf
		
		; InstallShield Wizard Completed
		$strTitle = "Installing Drivers"
		$strFinish  = "Drivers for Intel® Network Connections were successfully installed."	
		If WinExists($strTitle, $strFinish) Then
			ExitLoop
		Else
			Sleep(30)
		EndIf
	WEnd

	; InstallShield Wizard Completed
	$strTitle = "Installing Drivers"
	$strText  = "Drivers for Intel® Network Connections were successfully installed."	
	If Not WinWait($strTitle, $strText, 50) Then
		WriteLog("Step 2 of 2")
		Return False
	EndIf
	ControlFocus($strTitle, "&Close", "&Close")
	ControlClick($strTitle, "&Close", "&Close")
		
	; Ping server
	$strPreloadServer = GetPreLoadServerName()
	For $index = 0 To 20
    	$blnResult = Ping($strPreloadServer, 250)
		If $blnResult Then
			ExitLoop
		Else
			Sleep(100)
		EndIf
	Next
		
	Return True
EndFunc