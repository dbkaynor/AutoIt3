#comments-start HEADER
;******************************************************************************************
;** Intel Corporation, MPG MPAD
;** Title			:  		incReg.wbt
;** Description	: 	
;**		Registry related include file.
;**
;** Revision: 	Rev 2.0.9
;******************************************************************************************
;******************************************************************************************
;** Revision History:
;** 
;** Update for Rev 2.0.0	- Dick Lin 02-23-2005
;**		- Initial release
;**
;** Update for Rev 2.0.1	- Dick Lin 08/09/2006
;**		- Fixed IsWinXPMediaCenter()for Vista.
;**
;** Update for Rev 2.1.0	- Dick Lin 08/15/2006
;**		- Updated to support AutoIt V3.2.0.1.
;**
;** Update for Rev 2.1.1	- Dick Lin 11/22/2006
;**		- Added IsWinXPProSP2() function.
;**
;** Update for Rev 2.1.2	- Dick Lin 02/05/2007
;**		- Modified IsWinXP64Family()/IsWinVistaFamily() to use AutoIt 3.2.2.0
;**
;** Update for Rev 2.1.3	- Dick Lin 04/19/2007
;**		- Added IsWinXP64SP2() function.
;**
;** Update for Rev 2.0.4	- Dick Lin 05/29/2007
;**		- Modified for Naming convention.
;**
;** Update for Rev 2.0.5	- Dick Lin 10/09/2007
;**		- Added IsWinVista32SP1() function.
;**		- Added IsWinVista64SP1() function.
;**		- Removed AutoIt version check.
;**
;** Update for Rev 2.0.6	- Dick Lin 02/22/2008
;**		- Added IsWinXPHomeSP3(), IsWinXPSP3() and IsWinXPProSP3() functions.
;**
;**	Update for Rev 2.0.7	- Jarek Szymanski 8/6/2008
;** 	- Added IsWinXPSP3Family()function 
;**
;**	Update for Rev 2.0.8	- Jarek Szymanski 2/17/2009
;** 	- Added IsWin7Family(), IsWin7_64(), IsWin7_32()function 
;**
;**	Update for Rev 2.0.9	- Jarek Szymanski 5/6/2009
;** 	- Changed @OSVersion for IsWinXP64Family() from WIN_2003 to WIN_XP due update to AutoIt 3.3.0.0
;**
;**	Update for Rev 2.0.10	- Jarek Szymanski 7/2/2009
;** 	- Updated OS recognition based on @OSbuild for IsWinVistaSP1Family(), IsWinVistaFamily(), IsWin7Family() due to new AutoIt compatibility
;**
;******************************************************************************************
#comments-end HEADER

#include-once

#include <Misc.au3>
#include <inet.au3>

;***************************************************************************
;** Function: 		IsSupportedOS()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to decide if it's a supported OS.
;** Return:
;**		1 for Win2k/WinXP or greater. Otherwise False.
;** Usage:
;**		if IsSupportedOS() then
;**			...
;**		endif
;**
;***************************************************************************	
Func IsSupportedOS()	
	Return _Iif(@OSType = "WIN32_NT", True, False)
EndFunc

;***************************************************************************
;** Function: 		IsWinXP64Family()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to check if it's Windows XP x64 family.
;** Return:
;**		1 for Windows Vista family, otherwise 0.
;** Usage:
;**		If IsWinXP64Family() then 
;**			...
;**		endif
;**
;***************************************************************************
Func IsWinXP64Family()

	Return _Iif((@OSType == "WIN32_NT") And (@OSVersion == "WIN_XP") And (@OSBuild == 3790), True, False)
EndFunc	

;***************************************************************************
;** Function: 		IsWinXP64SP1()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to check if it's Windows XP x64 Service Pack 1.
;** Return:
;**		1 for Windows XP x64 SP1, otherwise 0.
;** Usage:
;**		If IsWinXP64SP1() then 
;**			...
;**		endif
;**
;***************************************************************************\
Func IsWinXP64SP1()

	Return _Iif(IsWinXP64Family() And (@OSServicePack == "Service Pack 1"), True, False)
	
EndFunc	

;***************************************************************************
;** Function: 		IsWinXPFamily()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to decide if current windows version is WinXP.
;** Return:
;**		1 for Windows XP family. Otherwise 0.
;** Usage:
;**		If IsWinXPFamily() then 
;**			...
;**		endif
;**
;***************************************************************************	
Func IsWinXPFamily()

	$strDir = @HomeDrive & "\Program Files (x86)"
	Return _Iif((@OSVersion == "WIN_XP") And Not DirExists($strDir), True, False)

EndFunc
	
;***************************************************************************
;** Function: 		IsWinXPSP1()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to check if it's Windows XP SP1.
;** Return:
;**		1 for Windows XP SP2, otherwise 0
;** Usage:
;**		If IsWinXPFamily() then 
;**			...
;**		endif
;**
;***************************************************************************	
Func IsWinXPSP1()

	Return _Iif((@OSVersion == "WIN_XP") And (@OSBuild == 2600) And (@OSServicePack == "Service Pack 1"), True, False)
EndFunc	
	
;***************************************************************************
;** Function: 		IsWinXPSP2()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to check if it's Windows XP SP2.
;** Return:
;**		1 for Windows XP SP2, otherwise 0
;** Usage:
;**		If IsWinXPSP2() then 
;**			...
;**		endif
;**
;***************************************************************************	
Func IsWinXPSP2()
		Return _Iif((@OSVersion == "WIN_XP") And (@OSBuild == 2600) And (@OSServicePack == "Service Pack 2"), True, False)
EndFunc		

;***************************************************************************
;** Function: 		IsWinXPProSP2()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to check if it's Windows XP Pro SP2.
;** Return:
;**		1 for Windows XP Pro SP2, otherwise 0
;** Usage:
;**		If IsWinXPProSP2() then 
;**			...
;**		endif
;**
;***************************************************************************	
Func IsWinXPProSP2()
	$strOS = WMIGetOperatingSystemName()
	$strSP = WMIGetServicePackVersion()
	if (($strOS == "Microsoft Windows XP Professional") And ($strSP == "Service Pack 2")) Then
		Return True
	Else
		Return False
	EndIf
EndFunc

;***************************************************************************
;** Function: 		IsWinXPMediaCenter()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to check if OS is Windows XP Media Center.
;** Return:
;**		1 for Windows XP Media Center Edition, otherwise 0.
;** Usage:
;**		If IsWinXPMediaCenter() then 
;**			...
;**		endif
;**
;***************************************************************************
Func IsWinXPMediaCenter()
	If IsWinVistaFamily() Then
		Return False
	Else
		$strKey = "HKLM\Software\Microsoft\Windows\CurrentVersion\Media Center"
		$strSubKey = "Ident"
		$var = RegRead ($strKey, $strSubKey)

		If $var = "" Then	
			Return False
		Else
			Return True
		EndIf
	EndIf
EndFunc

;***************************************************************************
;** Function: 		IsWinXPTablet()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to check if OS is Windows XP Tablet.
;** Return:
;**		1 for WinXP Tablet PC, otherwise 0.
;** Usage:
;**		If IsWinXPTablet() then 
;**			...
;**		endif
;**
;***************************************************************************
Func IsWinXPTablet()
	If Not IsWinXPSP2() Then
		Return False
	EndIf
	
	$strKey = "HKLM\Software\Microsoft\Windows\CurrentVersion\Tablet PC"
	$strSubKey = "Ident"
	$var = RegRead ($strKey, $strSubKey)

	If $var = "" Then	
		Return False
	Else
		Return True
	EndIf
EndFunc

;***************************************************************************
;** Function: 		IsWinXPHome()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to check if OS is Windows XP Home.
;** Return:
;**		1 for Windows XP Home Edition, otherwise 0.
;** Usage:
;**		If IsWinXPHome() then 
;**			...
;**		endif
;**
;***************************************************************************
Func IsWinXPHome()
	$strOS = WMIGetOperatingSystemName()
	
	If $strOS == "Microsoft Windows XP Home Edition" Then
		Return True
	Else
		Return False
	EndIf

EndFunc

;***************************************************************************
;** Function: 		IsWin2KFamily()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to decide if current windows version is Win2K.
;** Return:
;**		1 for Windows 2000 family, otherwise 0.
;** Usage:
;**		If IsWin2KFamily() then 
;**			...
;**		endif
;***************************************************************************
Func IsWin2KFamily()
	
	Return _Iif(@OSVersion == "WIN_2000", True, False)
EndFunc

;***************************************************************************
;** Function: 		IsWin2KSP4()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to decide if current windows version is Win2K.
;** Return:
;**		1 for Windows 2000 SP4, Otherwise 0.
;** Usage:
;**		If IsWin2KSP4() then 
;**			...
;**		endif
;**
;***************************************************************************
Func IsWin2KSP4()

	Return _Iif((@OSVersion == "WIN_2000") And (@OSServicePack == "Service Pack 4"), True, False)
EndFunc

;***************************************************************************
;** Function: 		IsWinXP64SP1()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to check if it's Windows XP x64 Service Pack 1.
;** Return:
;**		1 for Windows XP x64 SP1, otherwise 0.
;** Usage:
;**		If IsWinXP64SP1() then 
;**			...
;**		endif
;**
;***************************************************************************\
Func IsWinXP64SP2()

	Return _Iif(IsWinXP64Family() And (@OSServicePack == "Service Pack 2"), True, False)
EndFunc	

;***************************************************************************
;** Function: 		IsWinVistaFamily()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to decide if current windows version is WinXP.
;** Return:
;**		1 for Windows Vista family, otherwise 0.
;** Usage:
;**		If IsWinVistaFamily() then 
;**			...
;**		endif
;**
;***************************************************************************
Func IsWinVistaFamily()
	$OS_build = StringTrimRight(@OSBuild, 3)
	
	;Return _Iif(((@OSType == "WIN32_NT") And (@OSVersion == "WIN_VISTA")), True, False)
	Return _Iif(((@OSType == "WIN32_NT") And ($OS_build == "6")), True, False)
EndFunc

;***************************************************************************
;** Function: 		IsWinVistaFamily()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to decide if current windows version is WinXP.
;** Return:
;**		1 for Windows Vista family, otherwise 0.
;** Usage:
;**		If IsWinVistaFamily() then 
;**			...
;**		endif
;**
;***************************************************************************
Func IsWinVistaSP1Family()
	$OS_build = StringTrimRight(@OSBuild, 3)
	$blnSP1 = StringInStr(@OSServicePack, "Service Pack 1")
	;Return _Iif($blnSP1 And (@OSType == "WIN32_NT") And (@OSVersion == "WIN_VISTA") And (@OSBuild == 6001), True, False)
	Return _Iif($blnSP1 And (@OSType == "WIN32_NT") And ($OS_build == "6") And (@OSBuild == 6001), True, False)
EndFunc

;***************************************************************************
;** Function: 		IsWinVista64()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to decide if current windows version is Windows Vista x64.
;** Return:
;**		1 for Windows Vista x64, otherwise 0.
;** Usage:
;**		If IsWin2KFamily() then 
;**			...
;**		endif
;***************************************************************************
Func IsWinVista64()
	$strDir = @HomeDrive & "\Program Files (x86)"
	If IsWinVistaFamily() And DirExists($strDir) Then
		Return True
	EndIf
	
	Return False

EndFunc

;***************************************************************************
;** Function: 		IsWinVista64SP1()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to decide if current windows version is Windows Vista x64.
;** Return:
;**		1 for Windows Vista x64, otherwise 0.
;** Usage:
;**		If IsWin2KFamily() then 
;**			...
;**		endif
;***************************************************************************
Func IsWinVista64SP1()
	
	$blnSP1 = StringInStr(@OSServicePack, "Service Pack 1")
	$strDir = @HomeDrive & "\Program Files (x86)"
	If ($blnSP1 And IsWinVistaFamily() And DirExists($strDir)) Then
		Return True
	EndIf
	
	Return False

EndFunc

;***************************************************************************
;** Function: 		IsWinVista32()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to decide if current windows version is Windows Vista x32.
;** Return:
;**		1 for Windows Vista x64, otherwise 0.
;** Usage:
;**		If IsWin2KFamily() then 
;**			...
;**		endif
;***************************************************************************
Func IsWinVista32()
	If (IsWinVistaFamily() And (Not IsWinVista64()) And (Not IsWinVista64SP1())) Then
		Return True
	EndIf
	
	Return False
EndFunc

;***************************************************************************
;** Function: 		IsWinVista32SP1()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to decide if current windows version is Windows Vista x32.
;** Return:
;**		1 for Windows Vista x64, otherwise 0.
;** Usage:
;**		If IsWin2KFamily() then 
;**			...
;**		endif
;***************************************************************************
Func IsWinVista32SP1()
	If (IsWinVistaSP1Family() And (Not IsWinVista64()) And (Not IsWinVista64SP1())) Then
		Return True
	EndIf
	
	Return False
EndFunc

;***************************************************************************
;** Function: 		IsWinXPHomeSP3()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to check if it's Windows XP SP1.
;** Return:
;**		1 for Windows XP SP2, otherwise 0
;** Usage:
;**		If IsWinXPFamily() then 
;**			...
;**		endif
;**
;***************************************************************************	
Func IsWinXPHomeSP3()
	
	$strOS = WMIGetOperatingSystemName()
	$blnSP3 = StringInStr(@OSServicePack, "Service Pack 3")
	
	If ($strOS == "Microsoft Windows XP Home Edition") And $blnSP3 Then
		Return True
	Else
		Return False
	EndIf
	
EndFunc

;***************************************************************************
;** Function: 		IsWinXPSP3()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to check if it's Windows XP SP1.
;** Return:
;**		1 for Windows XP SP2, otherwise 0
;** Usage:
;**		If IsWinXPFamily() then 
;**			...
;**		endif
;**
;***************************************************************************	
Func IsWinXPSP3()
	
	$strOS = WMIGetOperatingSystemName()
	$blnSP3 = StringInStr(@OSServicePack, "Service Pack 3")
	
	Return _Iif((@OSVersion == "WIN_XP") And (@OSBuild == 2600) And $blnSP3, True, False)
	
EndFunc

;***************************************************************************
;** Function: 		IsWinXPSP3()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to check if it's Windows XP SP1.
;** Return:
;**		1 for Windows XP SP2, otherwise 0
;** Usage:
;**		If IsWinXPFamily() then 
;**			...
;**		endif
;**
;***************************************************************************	
Func IsWinXPProSP3()
	
	$strOS = WMIGetOperatingSystemName()
	$blnSP3 = StringInStr(@OSServicePack, "Service Pack 3")
	
	Return _Iif(($strOS == "Microsoft Windows XP Professional") And (@OSBuild == 2600) And $blnSP3, True, False)
	
EndFunc

;***************************************************************************
;** Function: 		IsWinXPSP3Family()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to check for Windows XP SP3 family
;** Return:
;**		1 for Windows XP SP3, otherwise 0
;** Usage:
;**		If IsWinXPSP3Family() then 
;**			...
;**		endif
;**
;***************************************************************************	
Func IsWinXPSP3Family()
		Return _Iif((@OSVersion == "WIN_XP") And (@OSBuild == 2600) And (@OSServicePack == "Service Pack 3"), True, False)
EndFunc	


;***************************************************************************
;** Function: 		IsWin7Family()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to decide if current windows version is Win 7.
;** Return:
;**		1 for Windows 7 family, otherwise 0.
;** Usage:
;**		If IsWin7Family() then 
;**			...
;**		endif
;**
;***************************************************************************
Func IsWin7Family()
	$OS_build = StringTrimRight(@OSBuild, 3)
	
	;Return _Iif(((@OSType == "WIN32_NT") And (@OSVersion == "WIN_LONGHORN")), True, False)
	Return _Iif((@OSType == "WIN32_NT") And ($OS_build == "7"), True, False)
EndFunc

;***************************************************************************
;** Function: 		IsWin7_64()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to decide if current windows version is Windows 7 x64.
;** Return:
;**		1 for Windows 7 x64, otherwise 0.
;** Usage:
;**		If IsWin7_64() then 
;**			...
;**		endif
;***************************************************************************
Func IsWin7_64()
	$strDir = @HomeDrive & "\Program Files (x86)"
	If (IsWin7Family() And DirExists($strDir)) Then
		Return True
	EndIf
	
	Return False

EndFunc

;***************************************************************************
;** Function: 		IsWin7_32()
;** Parameters:
;**		None
;** Description: 				 
;**		This function is called to decide if current windows version is Windows 7 x32.
;** Return:
;**		1 for Windows 7 x32, otherwise 0.
;** Usage:
;**		If IsWin7_32() then 
;**			...
;**		endif
;***************************************************************************
Func IsWin7_32()
	If (IsWin7Family() And (Not IsWin7_64())) Then
		Return True
	EndIf
	
	Return False
EndFunc
