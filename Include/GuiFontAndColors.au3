; Get font and color information for the current windows theme!
;
; CyberSlug - 23 January 2004
;
; I reverse engineered _SystemFontSize and _SystemFontName from the Windows XP SP2 registry,
; so this might not work on all versions of Windows.

#include-once

Global Const $ActiveBorderColor				= _SystemColor("ActiveBorder")
Global Const $ActiveTitleColor				= _SystemColor("ActiveTitle")
Global Const $AppWorkSpaceColor				= _SystemColor("AppWorkSpace")
Global Const $BackgroundColor				= _SystemColor("Background")
Global Const $ButtonAlternateFaceColor		= _SystemColor("ButtonAlternateFace")
Global Const $ButtonDkShadowColor			= _SystemColor("ButtonDkShadow")
Global Const $ButtonFaceColor				= _SystemColor("ButtonFace")
Global Const $ButtonHilightColor			= _SystemColor("ButtonHilight")
Global Const $ButtonLightColor				= _SystemColor("ButtonLight")
Global Const $ButtonShadowColor				= _SystemColor("ButtonShadow")
Global Const $ButtonTextColor				= _SystemColor("ButtonText")
Global Const $GradientActiveTitleColor		= _SystemColor("GradientActiveTitle")
Global Const $GradientInactiveTitleColor	= _SystemColor("GradientInactiveTitle")
Global Const $GrayTextColor					= _SystemColor("GrayText")
Global Const $HilightColor					= _SystemColor("Hilight")
Global Const $HilightTextColor				= _SystemColor("HilightText")
Global Const $HotTrackingColor				= _SystemColor("HotTrackingColor")
Global Const $InactiveBorderColor			= _SystemColor("InactiveBorder")
Global Const $InactiveTitleColor			= _SystemColor("InactiveTitle")
Global Const $InactiveTitleTextColor		= _SystemColor("InactiveTitleText")
Global Const $InfoTextColor					= _SystemColor("InfoText")
Global Const $InfoWindowColor				= _SystemColor("InfoWindow")
Global Const $MenuColor						= _SystemColor("Menu")
Global Const $MenuTextColor					= _SystemColor("MenuText")
Global Const $ScrollbarColor				= _SystemColor("Scrollbar")
Global Const $TitleTextColor				= _SystemColor("TitleText")
Global Const $WindowColor					= _SystemColor("Window")
Global Const $WindowFrameColor				= _SystemColor("WindowFrame")
Global Const $WindowTextColor				= _SystemColor("WindowText")
Global Const $MenuHilightColor				= _SystemColor("MenuHilight")
Global Const $MenuBarColor					= _SystemColor("MenuBar")

Global Const $CaptionFontName		= _SystemFontName("CaptionFont")
Global Const $CaptionFontSize		= _SystemFontSize("CaptionFont")

Global Const $IconFontName			= _SystemFontName("IconFont")
Global Const $IconFontSize			= _SystemFontSize("IconFont")

Global Const $MenuFontName			= _SystemFontName("MenuFont")
Global Const $MenuFontSize			= _SystemFontSize("MenuFont")

Global Const $MessageFontName		= _SystemFontName("MessageFont")
Global Const $MessageFontSize		= _SystemFontSize("MessageFont")

Global Const $SmCaptionFontName		= _SystemFontName("SmCaptionFont")
Global Const $SmCaptionFontSize		= _SystemFontSize("SmCaptionFont")

Global Const $StatusFontName		= _SystemFontName("StatusFont")
Global Const $StatusFontSize		= _SystemFontSize("StatusFont")


Func _SystemColor($element)
	Local $tuple = RegRead("HKEY_CURRENT_USER\Control Panel\Colors", $element)
	If @error Then Return 0xFFFFFF  ;fail-safe value is white....
	; For example, registry key is '255 255 255' if white
	Local $t = StringSplit($tuple, ' ')
	Return Dec( Hex($t[1],2) & Hex($t[2],2) & Hex($t[3],2) )
EndFunc


Func _SystemFontSize($element)
	Local $regData = RegRead("HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics",$element)

	Local $sizeInfo = StringLeft($regData, 2)
	Local $hex = "f8,f7,f5,f4,f3,f1,f0,ef,ed,ec,eb,e9,e8,e7,e5,e4,e3,e1,e0"
	Local $pt = "06,07,08,09,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24"
	Local $fontSize = Number( StringMid($pt, StringInStr($hex,$sizeInfo), 2) )
	Return $fontSize ;FYI if element was not found, the return value will be 6
EndFunc


Func _SystemFontName($element)
	Local $regData = RegRead("HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics",$element)

	Local $binary = StringTrimLeft($regData, 56)
	Local $binName = Stringleft($binary, StringInStr($binary, "000000")-1)
	Local $commaSeparated = "" ;put a comma between each group of two characters
	For $i = 1 to StringLen($binName) Step 2
		$commaSeparated = $commaSeparated & StringMid($binName, $i, 2) & ","
	Next

	Local $split = StringSplit($commaSeparated, ",")
	local $fontName = ""
	For $i = 1 to $split[0]
		If $split[$i] = 20 Then
			$fontName = $fontName & Chr(Dec($split[$i])) ;" "
		Else
			$fontName = $fontName & Chr(Dec($split[$i]))
		EndIf
	Next
	If $fontName = "" Then $fontName = "Arial" ;failsafe font
	Return $fontName
EndFunc