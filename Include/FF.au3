; :wrap=none:collapseFolds=1:maxLineLen=80:mode=autoitscript:tabSize=8:folding=indent:
; created with jEdit4AutoIt
#include-once

#Region Copyright
#cs
	* FF.au3
	*
	* This program is free software; you can redistribute it and/or
	* modify it under the terms of the GNU General Public License
	* as published by the Free Software Foundation; either version 2
	* of the License, or any later version.
	*
	* This program is distributed in the hope that it will be useful,
	* but WITHOUT ANY WARRANTY; without even the implied warranty of
	* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
	* GNU General Public License for more details.
	*
	* You should have received a copy of the GNU General Public License
	* along with this program; if not, write to the Free Software
	* Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
#ce
#EndRegion Copyright

#Region Many thanks to:
#cs
	- Jonathan Bennett and the AutoIt Team for AutoIt v3 the freeware
	BASIC-like scripting language
	(and thanks John George Kemeny and Thomas Eugene Kurtz for BASIC)
	- Dale Hohm, for his great IE.au3, which I've used as base for the
	functions-names and the error handling.
	- Johannes Schirmer, for some functions and suggestions for the FF.au3
	- all others for their suggestions, bug reports ...
	- the FireFox-Team for his fantastic browser

	- and Massimiliano Mirra, without his work would this UDF impossible,
	for MozLab / MozRepl
#ce
#EndRegion Many thanks to:

#Region #CURRENT# ==============================================================
;_FFAction
;_FFClick
;_FFCmd
;_FFConnect
;_FFDialogWait
;_FFDisConnect
;_FFDispatchEvent
;_FFFormCheckBox
;_FFFormGetElementsLength
;_FFFormOptionSelect
;_FFFormRadioButton
;_FFFormReset
;_FFFormSubmit
;_FFFrameEnter
;_FFFrameLeave
;_FFGetLength
;_FFGetObjectInfo
;_FFGetPosition
;_FFGetValue
;_FFImageClick
;_FFImageGetBySize
;_FFImageMapClick
;_FFIsConnected
;_FFLinkClick
;_FFLinksGetAll
;_FFLoadWait
;_FFLoadWaitTimeOut
;_FFObj
;_FFObjDelete
;_FFObjGet
;_FFObjNew
;_FFOpenURL
;_FFPrefGet
;_FFPrefReset
;_FFPrefSet
;_FFQuit
;_FFReadHTML
;_FFReadText
;_FFSearch
;_FFSetValue
;_FFStart
;_FFTabAdd
;_FFTabClose
;_FFTabDuplicate
;_FFTabExists
;_FFTabGetSelected
;_FFTabSetSelected
;_FFTableWriteToArray
;_FFWindowClose
;_FFWindowGetHandle
;_FFWindowOpen
;_FFWindowSelect
;_FFWriteHTML
;_FFXPath
#EndRegion #CURRENT# ==============================================================

#Region Description
; ==============================================================================
; UDF ...........: FF.au3
Global Const $_FF_AU3VERSION = "0.6.0.1b-3"
; Description ...: An UDF for FireFox automation.
; Requirement ...: MozRepl AddOn:
;                  http://hyperstruct.net/projects/mozlab
;                  http://wiki.github.com/bard/mozrepl/home
; Author(s) .....: Thorsten Willert, Johannes Schirmer
; Date ..........: Tue Feb 16 12:56:50 CET 2010 @539 /Internet Time/
; FireFox Version: Firefox/3.6 (required 3.x.x)
; AutoIt Version : v3.3.0.0
; ==============================================================================
#cs
	V0.6.0.1b-3
	- Changed: Default values for _FFXpath: _FFXPath($sQuery, $sAttribute = "", $iReturnType = 9, $iFilter = 0)
	- Fixed: Different problems after _FFOpenURL
	- Fixed: Errorin _FFAu3Option()

	<V0.6.x.x
	For compatibily for older scripts, please use this UDF:
	http://thorsten-willert.de/Themen/FFau3/FF.au3/FFEx.au3

	ToDo:
	Textsuche (name, id, text ...) per Substring, String, RegEx
	Events fuer die _FFForm* Funktionen
	Update der Window-Funktionen
#ce
#EndRegion Description

#Region Global Constants
Global Const $_FF_PROC_NAME = "firefox.exe" ; Firefox process name
Global Const $_FF_COM_DELAY_MAX = 200 ; alternative connection delay in ms
Global Enum $_FF_ERROR_Success = 0, _ 	; No error
		$_FF_ERROR_GeneralError, _ 	; General error
		$_FF_ERROR_SocketError, _ 	; No socket
		$_FF_ERROR_InvalidDataType, _ 	; Invalid data type (IP, URL, Port ...)
		$_FF_ERROR_InvalidValue, _ 	; Invalid value in function-call
		$_FF_ERROR_SendRecv, _		; Send / Recv Error
		$_FF_ERROR_Timeout, _ 		; Connection / Send / Recv timeout
		$_FF_ERROR___UNUSED, _ 		;
		$_FF_ERROR_NoMatch, _ 		; No match for _FFAction-find/search _FFGetElement...
		$_FF_ERROR_RetValue, _		; Error echo from Repl e.g. _FFAction("fullscreen","true") <> "true"
		$_FF_ERROR_ReplException, _ 	; Exception from MozRepl / FireFox
		$_FF_ERROR_InvalidExpression ; Invalid expression in XPath query or RegEx
#EndRegion Global Constants

#Region Global Variables
Global $_FF_GLOBAL_SOCKET = -1 ; Socket
Global $_FF_CON_DELAY ; Connection Delay
Global $_FF_LOADWAIT_TIMEOUT = 120000
Global $_FF_LOADWAIT_STOP = True ; stop loading after $_FF_LOADWAIT_TIMEOUT
Global $_FF_COM_TRACE = True ; Trace communication to console (debugging)
Global $_FF_ERROR_MSGBOX = True ; Shows in compiled scripts error messages in msgboxes
Global $_FF_FRAME = 'top'
Global $_FF_SEARCH_MODE = 0 ; 0 = Substring / 1 = Compare
#EndRegion Global Variables

; #FUNCTION# ===================================================================
; Name ..........: _FFAction
; Description ...: Some standard actions to work with FireFox
; Beschreibung ..: Standardaktionen in FireFox
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFAction($sAction[, $vOption = ""[, $vOption2 = ""[, $bLoadWait = True]]])
; Parameter(s): .: $sAction     - one of the following actions:
;                               | about
;                               | alert
;                               | back
;                               | blank
;                               | copy
;                               | fullscreen / fs
;                               | hideall
;                               | home
;                               | maximize / max
;                               | minimize / min
;                               | presentationmode / pm
;                               | print
;                               | reload
;                               | resetconsole
;                               | restore
;                               | scrollXY
;                               | stop
;                               | zoom
;                  $vOption     - Optional: (Default = "") :
;                  $vOption2    - Optional: (Default = "") :
;                  $bLoadWait   - Optional: (Default = true) :
; Return Value ..: Success      - Return-value from MozRepl (sometimes an empty string!!!)
;                  Failure      - ""
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Fri Dec 04 16:39:49 CET 2009
; Link ..........: http://developer.mozilla.org/en/docs/XUL:tabbrowser, http://developer.mozilla.org/en/docs/XUL:Method:reloadWithFlags, https://developer.mozilla.org/En/DOM/Window.scrollBy, http://kb.mozillazine.org/About_Protocol_Links
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _FFAction($sAction, $vOption = "", $vOption2 = "", $bLoadWait = True)
	Local Const $sFuncName = "_FFAction"

	Local $sCommand
	Local $sActionL = StringLower($sAction)

	Select
		Case $sActionL = "back"
			$sCommand = "gBrowser.goBack()"
		Case $sActionL = "home"
			$sCommand = "gBrowser.goHome()"
		Case $sActionL = "forward"
			$sCommand = "gBrowser.goForward()"
		Case $sActionL = "copy"
			Return _FFCmd(".getSelection()", $_FF_CON_DELAY + 100)
		Case $sActionL = "reload" And $vOption <> ""
			$sCommand = "gBrowser.reloadWithFlags('" & $vOption & "')"
		Case $sActionL = "reload"
			$sCommand = "gBrowser.reload()"
		Case $sActionL = "fullscreen" Or $sActionL = "fs" And $vOption
			$sCommand = "fullScreen=true"
			$bLoadWait = False
		Case $sActionL = "fullscreen" Or $sActionL = "fs" And Not $vOption
			$sCommand = "fullScreen=false"
			$bLoadWait = False
		Case $sActionL = "minimize" Or $sActionL = "min"
			$bLoadWait = False
			$sCommand = "minimize()"
		Case $sActionL = "maximize" Or $sActionL = "max"
			$bLoadWait = False
			$sCommand = "maximize()"
		Case $sActionL = "restore"
			$bLoadWait = False
			$sCommand = "restore()"
		Case $sActionL = "stop"
			$bLoadWait = False
			$sCommand = "gBrowser.stop()"
		Case $sActionL = "print"
			$bLoadWait = False
			$sCommand = "content.print()"
		Case $sActionL = "find" Or $sActionL = "search"
			_FFLoadWait()
			$sCommand = "content.find()"
		Case $sActionL = "hideall" And $vOption
			Return _FFCmd("toggleAffectedChrome(true);")
		Case $sActionL = "hideall" And Not $vOption
			Return _FFCmd("toggleAffectedChrome(false);")
		Case $sActionL = "presentationmode" Or $sActionL = "pm" And $vOption
			Return _FFCmd("toggleAffectedChrome(true)" & @CRLF & "fullScreen=true")
		Case $sActionL = "presentationmode" Or $sActionL = "pm" And Not $vOption
			Return _FFCmd("toggleAffectedChrome(false)" & @CRLF & "fullScreen=false")
		Case $sActionL = "alert" And $vOption <> ""
			$sCommand = 'alert("' & __FFValue2JavaScript($vOption) & '")'
			$bLoadWait = False
		Case $sActionL = "chrome" And $vOption <> ""
			$sCommand = '.location.href="' & __FFChromeSelect($vOption) & '"'
		Case $sActionL = "blank"
			$vOption = "blank"
			ContinueCase
		Case $sActionL = "about"
			$sCommand = '.location.href="about:' & $vOption & '"'
		Case $sActionL = "scrollxy" And IsInt($vOption) And IsInt($vOption2)
			$sCommand = "window.content.scrollBy(" & $vOption & "," & $vOption2 & ");"
			$bLoadWait = False
		Case $sActionL = "zoom" And $vOption = ""
			$vOption = 1
			ContinueCase
		Case $sActionL = "zoom" And IsNumber($vOption)
			$sCommand = "gBrowser.selectedBrowser.markupDocumentViewer.fullZoom=" & $vOption
			$bLoadWait = False
		Case $sActionL = "resetconsole"
			$sCommand = ";" & @CRLF & "repl.home()"
			$bLoadWait = False
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "$sAction: " & $sAction & " or " & "$vOption :" & $vOption & " or " & "$vOption2 :" & $vOption2))
			Return ""
	EndSelect

	Local $sRet = _FFCmd($sCommand)
	If Not @error Then
		If $bLoadWait Then _FFLoadWait()
		Return $sRet
	Else
		SetError(@error)
		Return ""
	EndIf

EndFunc   ;==>_FFAction

; #FUNCTION# ===================================================================
; Name ..........: _FFClick
; Description ...: Simulates a click on an element.
; Beschreibung ..: Simuliert einen Klick auf ein Element.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFClick($sElement[, $sMode = "elements"[, $iIndex = 0[, $bLoadWait = True]]])
; Parameter(s): .: $sElement    - Element to click on
;                  $sMode       - Optional: (Default = "elements") :
;                               | elements
;                               | id
;                               | name
;                               | class
;                               | tag
;                  $iIndex      - Optional: (Default = 0) : Index if $sMode = class, name, tag
;                  $bLoadWait   - Optional: (Default = True) :
; Return Value ..: Success      - 1
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Fri Jan 22 14:50:40 CET 2010
; Link ..........:
; Related .......: _FFImageClick, _FFLinkClick. _FFImageMapClick
; Example .......: Yes
; ==============================================================================
Func _FFClick($sElement, $sMode = "elements", $iIndex = 0, $bLoadWait = True)
	Local Const $sFuncName = "_FFClick"

	If Not IsInt($iIndex) Then
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(int) $iIndex: " & $iIndex))
		Return 0
	EndIf

	If $sMode = Default Then $sMode = "elements"
	If $iIndex = Default Then $iIndex = 0

	Switch StringLower($sMode)
		Case "elements"
			If StringLeft($sElement, 7) = "OBJECT|" Then $sElement = StringMid($sElement, 8)
		Case "id"
			$sElement = ".getElementById('" & $sElement & "')"
		Case "name"
			$sElement = ".getElementsByName('" & $sElement & "')[" & $iIndex & "]"
		Case "class"
			$sElement = ".getElementsByClassName('" & $sElement & "')[" & $iIndex & "]"
		Case "tag"
			$sElement = ".getElementsByTagName('" & $sElement & "')[" & $iIndex & "]"
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(elements|id|name|class|tag) $sMode: " & $sMode))
			Return 0
	EndSwitch

	If StringLeft($sElement, 1) = "." Then $sElement = "FFau3.WCD" & $sElement

	Local $RetVal = _FFCmd("FFau3.simulateEvent(" & $sElement & ",'MouseEvents','click');")
	If Not @error And $RetVal <> "_FFCmd_Err" And $RetVal = 1 Then
		Sleep(25)
		If $bLoadWait Then Return _FFLoadWait()
		Return 1
	Else
		SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch, "$sElement: " & $sElement))
		Return 0
	EndIf

EndFunc   ;==>_FFClick

; #FUNCTION# ===================================================================
; Name ..........: _FFDialogWait
; Description ...: Waits for a browser-dialog-message (e.g. alert) and closes it.
; Beschreibung ..: Wartet auf ein Dialog-Fenster (z.B.: alert) und schließt es.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFDialogWait($sText[, $sTitle = ""[, $sClose = "close"[, $iTimeOut = 10000]]])
; Parameter(s): .: $sText       - Text in the dialog (substring)
;                  $sTitle      - Optional: (Default = "") : Window-title of the dialog (substring)
;                  $bClose      - Optional: (Default = "close") :
;                               | close = Just closes the dialog window
;                               | open = Do nothing
;                               | ok = Uses the "ok"-button
;                               | cancel = Uses the"cancel"-button
;                  $iTimeOut    - Optional: (Default = 10000) : Timeout for waiting (min. 1000ms)
; Return Value ..: Success      - 1
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Tue Apr 14 12:18:08 CEST 2009
; Link ..........:
; Related .......: _FFLoadWait
; Example .......: Yes
; ==============================================================================
Func _FFDialogWait($sText, $sTitle = "", $sClose = "close", $iTimeOut = 10000)
	Local Const $sFuncName = "_FFDialogWait"
	If $sTitle = "" And $sText = "" Then
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "Empty String: $sText and $sTitle"))
		Return 0
	EndIf
	If $sClose = Default Then $sClose = "close"
	If $iTimeOut < 1000 Then $iTimeOut = 1000

	Local $bFound = False, $b1 = False, $b2 = False

	Local $iTimer = TimerInit()
	While TimerDiff($iTimer) < $iTimeOut
		If _FFCmd('FFau3.obj=Components.classes["@mozilla.org/appshell/window-mediator;1"].getService(Components.interfaces.nsIWindowMediator).getMostRecentWindow("").window;' & _
				'FFau3.obj.document.documentURI') = "chrome://global/content/commonDialog.xul" Then
			If $sTitle <> "" Then
				If StringInStr(_FFCmd("FFau3.obj.title"), $sTitle) Then $b1 = True
			EndIf
			If $sText <> "" Then
				If StringInStr(_FFCmd("FFau3.obj.document.documentElement.textContent"), $sText) Then $b2 = True
			EndIf
			Select
				Case $sTitle = "" And $sText <> "" And Not $b1 And $b2
					$bFound = True
				Case $sTitle <> "" And $sText <> "" And $b1 And $b2
					$bFound = True
				Case $sTitle <> "" And $sText = "" And $b1 And Not $b2
					$bFound = True
				Case Else
					$bFound = False
			EndSelect
			If $bFound Then
				Switch StringLower($sClose)
					Case "close" ; just close the window
						_FFCmd("FFau3.obj.close()")
					Case "ok"
						_FFCmd("FFau3.obj.document.documentElement.acceptDialog()")
					Case "cancel"
						_FFCmd("FFau3.obj.document.documentElement.cancelDialog()")
					Case "open" ; do nothing
					Case Else
						SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "$sClose: " & $sClose))
						Return 0
				EndSwitch
				Return 1
			EndIf
		EndIf
		Sleep(500)
	WEnd

	SetError(__FFError($sFuncName, $_FF_ERROR_Timeout, Round(TimerDiff($iTimer)) & "ms > " & $iTimeOut & "ms"))
	Return 0
EndFunc   ;==>_FFDialogWait

; #FUNCTION# ===================================================================
; Name ..........: _FFImageClick
; Description ...: Clicks an image-link.
; Beschreibung ..: Klickt einen Image-Link.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFImageClick($vSearch[, $sMode = "src"[, $bLoadWait = True]])
; Parameter(s): .: $vSearch     - String to search or number (0-n) in index mode (the search is performed in the attributes of the image and the link)
;                  $sMode       - Optional: (Default = "src") :
;                               | alt
;                               | href
;                               | id
;                               | index
;                               | name
;                               | src
;                               | title
;                  $bLoadWait   - Optional: (Default = True) :
; Return Value ..: Success      - 1
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Fri Dec 04 14:55:57 CET 2009
; Link ..........:
; Related .......: _FFClick, _FFLinkClick, _FFImageMapClick
; Example .......: Yes
; ==============================================================================
Func _FFImageClick($vSearch, $sMode = "src", $bLoadWait = True)
	Local Const $sFuncName = "_FFImageClick"

	Local $sX
	$sMode = StringLower($sMode)
	$vSearch = __FFValue2JavaScript($vSearch)

	Switch $sMode
		Case "index"
			If Not IsInt($vSearch) Then
				SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(int) $vSearch: " & $vSearch))
				Return 0
			EndIf
			$sX = ".images[" & $vSearch & "]"
		Case "id", "src", "href", "alt", "name", "title"
			Switch $_FF_SEARCH_MODE
				Case 0
					$sX = "contains(@" & $sMode & ",'" & $vSearch & "')"
				Case 1
					$sX = "@" & $sMode & "='" & $vSearch & "'"
			EndSwitch
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(src|href|alt|name|title|id|index) $sMode: " & $sMode))
			Return 0
	EndSwitch

	If $sMode <> "index" Then
		If _FFClick(_FFXPath("//a//img[" & $sX & "] | //a[" & $sX & "]//img", "", 9)) Then
			If $bLoadWait Then Return _FFLoadWait()
			Return 1
		EndIf
	Else
		If _FFClick($sX) Then
			If $bLoadWait Then Return _FFLoadWait()
			Return 1
		EndIf
	EndIf

	SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch, "$vSearch: " & $vSearch))
	Return 0

EndFunc   ;==>_FFImageClick

; #FUNCTION# ===================================================================
; Name ..........: _FFImageMapClick
; Description ...: Clicks on area of an image-map
; Beschreibung ..: Klickt auf ein Area einer Image-map
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFImageMapClick($vArea[, $sModeArea = "index"[, $vMap = 0[, $sModeMap = "index"[, $bLoadWait = True]]]])
; Parameter(s): .: $vArea       - (substring)
;                  $sModeArea   - Optional: (Default = "index") :
;                               | alt
;                               | class
;                               | coords
;                               | href
;                               | id
;                               | index
;                               | name
;                               | title
;                  $vMap        - Optional: (Default = 0) : (substring)
;                  $sModeMap    - Optional: (Default = "index") :
;                               | class
;                               | id
;                               | index
;                               | name
;                               | title
;                  $bLoadWait   - Optional: (Default = True) :
; Return Value ..: Success      - 1
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Sep 28 10:39:55 CEST 2009
; Related .......: _FFClick, _FFLinkClick, _FFImageClick
; Example .......: Yes
; ==============================================================================
Func _FFImageMapClick($vArea, $sModeArea = "index", $vMap = 0, $sModeMap = "index", $bLoadWait = True)
	Local Const $sFuncName = "_FFImageMapClick"

	Local $sXPath

	$sModeArea = StringLower($sModeArea)
	$sModeMap = StringLower($sModeMap)

	Switch $sModeMap
		Case "index"
			If Not IsInt($vMap) Then
				SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(int) $vMap: " & $vMap))
				Return 0
			EndIf
			$sXPath = "//map[" & $vMap + 1 & "]"
		Case "id", "name", "title", "class"
			$sXPath = "//map[contains(@" & $sModeMap & ",'" & $vMap & "')]"
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(name|title|id|index|class) $sModeMap: " & $sModeMap))
			Return 0
	EndSwitch

	Switch $sModeArea
		Case "index"
			If Not IsInt($vMap) Then
				SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(int) $vArea: " & $vArea))
				Return 0
			EndIf
			$sXPath &= "//area[" & $vArea & "]"
		Case "id", "name", "title", "alt", "href", "class", "coords"
			$sXPath &= "//area[contains(@" & $sModeArea & ",'" & $vArea & "')]"
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(index|id|name|title|alt|href|coords|class) $sModeArea: " & $sModeArea))
			Return 0
	EndSwitch

	If _FFClick(_FFXPath($sXPath, "", 9)) Then
		If $bLoadWait Then Return _FFLoadWait()
		Return 1
	EndIf

	SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch))
	Return 0
EndFunc   ;==>_FFImageMapClick

; #FUNCTION# ===================================================================
; Name ..........: _FFImagesGetBySize
; Description ...: Returns an array with the index of matching images.
; Beschreibung ..: Liefert ein Array mit den Indizes der passenden Bilder.
; Syntax ........: _FFImagesGetBySize($iHeight, $iWidth[, $sMode = "eq"[, $iPercent = 0]])
; Parameter(s): .: $iHeight     - Heigth of the image(s)
;                  $iWidth      - Width of the image(s)
;                  $sMode       - Optional (Default = "eq") :
;                               | eq
;                               | lt
;                               | gt
;                  $iPercent    - Optional (Default = 0) : +/- tolerance in percent for $sMode = "eq" (min=0, max=100)
; Return Value ..: Success      - array[0] > 0
;                  Failure      - array[0] = 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Sep 28 10:40:47 CEST 2009
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _FFImagesGetBySize($iHeight, $iWidth, $sMode = "eq", $iPercent = 0)
	Local Const $sFuncName = "_FFImagesGetBySize"
	Local $aRet[1] = [0]

	If $iPercent < 0 Or $iPercent > 100 Then $iPercent = 0
	$sMode = StringLower($sMode)

	If Not IsInt($iHeight) Then
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(int) $iHeight: " & $iHeight))
		Return $aRet
	EndIf
	If Not IsInt($iWidth) Then
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(int) $iWidth: " & $iWidth))
		Return $aRet
	EndIf
	Switch $sMode
		Case "eq", "lt", "gt"
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(eq|lt|gt) $sMode: " & $sMode))
			Return $aRet
	EndSwitch

	Local $sCommand = StringFormat("FFau3.SearchImageBySize(%s,%s,'%s',%s);", $iHeight, $iWidth, $sMode, $iPercent)
	Local $iIndex = StringStripWS(_FFCmd($sCommand, 10000), 3)
	$aRet = StringSplit($iIndex, " ")
	If Not @error And $aRet[0] > 0 Then
		For $i = 1 To $aRet[0]
			$aRet[$i] = Int($aRet[$i])
		Next
	Else
		SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch))
		Return $aRet
	EndIf

	Return $aRet
EndFunc   ;==>_FFImagesGetBySize

; #FUNCTION# ===================================================================
; Name ..........: _FFLinkClick
; Description ...: Simulates a click on a link.
; Beschreibung ..: Simuliert einen Klick auf einen Link
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFLinkClick($vSearch[, $sMode = "href"[, $bLoadWait = True]])
; Parameter(s): .: $vSearch     - (Sub)String to search or number (0-n) in index mode
;                  $sMode       - Optional: (Default = "href") :
;                               | href
;                               | alt
;                               | name
;                               | title
;                               | id
;                               | index
;                  $bLoadWait   - Optional: (Default = True) :
; Return Value ..: Success      - 1
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Fri Jan 22 16:09:48 CET 2010
; Link ..........:
; Related .......: _FFClick, _FFImageClick, _FFImageMapClick
; Example .......: Yes
; ==============================================================================
Func _FFLinkClick($vSearch, $sMode = "href", $bLoadWait = True)
	Local Const $sFuncName = "_FFLinkClick"

	Local $sX
	$sMode = StringLower($sMode)
	If Not $sMode = "index" Then $vSearch = __FFValue2JavaScript($vSearch)

	Switch $sMode
		Case "index"
			If Not IsInt($vSearch) Then
				SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(int) $vSearch: " & $vSearch))
				Return 0
			EndIf
			$sX = ".links[" & $vSearch & "]"
		Case "text"
			$sX = "//a[contains(.,'" & $vSearch & "')]"
		Case "href", "name", "title", "id"
			Switch $_FF_SEARCH_MODE
				Case 0
					$sX = "//a[contains(@" & $sMode & ",'" & $vSearch & "')]"
				Case 1
					$sX = "//a[@" & $sMode & "='" & $vSearch & "']"
			EndSwitch
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(index|href|text|name|title|id) $sMode: " & $sMode))
			Return 0
	EndSwitch

	If $sMode <> "index" Then
		Return _FFClick(_FFXPath($sX, "", 9), "elements", 0, $bLoadWait)
	Else
		Return _FFClick($sX, "elements", 0, $bLoadWait)
	EndIf

	SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch, "$vSearch: " & $vSearch))
	Return 0

EndFunc   ;==>_FFLinkClick

; #FUNCTION# ===================================================================
; Name ..........: _FFFrameLeave
; Description ...: Leaves the currently entered frame
; Beschreibung ..: Verläßt den aktuellen Frame
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFFrameLeave()
; Parameter(s): .:
; Return Value ..: Success      - 1
;                  Failure      - 0
; Author(s) .....: Thorsten Willert
; Date ..........: Fri Sep 18 11:08:20 CEST 2009 @422 /Internet Time/
; Link ..........:
; Related .......: _FFFrameEnter
; Example .......: Yes
; ==============================================================================
Func _FFFrameLeave()
	;Local Const $sFuncName = "_FFFrameLeave"
	Local $bTrc = $_FF_COM_TRACE
	$_FF_COM_TRACE = False

	Local $sFrame = StringMid($_FF_FRAME, 1, StringInStr($_FF_FRAME, ".", 2, -1) - 1)
	_FFCmd("FFau3.WCD=window.content.wrappedJSObject." & $sFrame & ".document;")
	If Not @error Then
		$_FF_COM_TRACE = $bTrc
		$_FF_FRAME = $sFrame
		Return 1
	EndIf

	$_FF_COM_TRACE = $bTrc
	Return 0
EndFunc   ;==>_FFFrameLeave

; #FUNCTION# ===================================================================
; Name ..........: _FFFrameEnter
; Description ...: Selects a frame, which all subsequent commands apply on.
; Beschreibung ..: Wählt einen Frame aus, auf den alle nachfolgenden Befehle wirken.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFFrameEnter($vFrame[, $sMode = "index"])
; Parameter(s): .: $vFrame      - Frame name or id or number in index mode
;                  $sMode       - Optional: (Default = "index") :
;                               | name
;                               | id
;                               | index
; Return Value ..: Success      - 1
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Sat Oct 24 21:05:36 CEST 2009
; Link ..........:
; Related .......: _FFFrameLeave
; Example .......: Yes
; ==============================================================================
Func _FFFrameEnter($vFrame, $sMode = "index")
	Local Const $sFuncName = "_FFFrameEnter"

	Local $sFrame = $_FF_FRAME
	Local $bTrc = $_FF_COM_TRACE

	Switch StringLower($sMode)
		Case "index"
			If IsInt($vFrame) Then
				$sFrame &= ".frames[" & $vFrame & "]"
			Else
				SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(int) $vFrame: " & $vFrame))
				Return 0
			EndIf
		Case "name", "id"
			If IsString($vFrame) Then
				$sFrame &= "." & $vFrame
			Else
				SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(string) $vFrame: " & $vFrame))
				Return 0
			EndIf
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(index|name|id) $sMode: " & $sMode))
			Return 0
	EndSwitch

	$_FF_COM_TRACE = False
	_FFCmd("FFau3.WCD=window.content.wrappedJSObject." & $sFrame & ".document;")
	If Not @error Then
		$_FF_FRAME = $sFrame
		$_FF_COM_TRACE = $bTrc
		Return 1
	EndIf

	$_FF_COM_TRACE = $bTrc
	SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "Frame not found $vFrame: " & $vFrame))
	Return 0
EndFunc   ;==>_FFFrameEnter

; #FUNCTION# ===================================================================
; Name ..........: _FFGetLength
; Description ...: Returns the length of the elements in $sMode
; Beschreibung ..: Gibt die Anzahl der angegebenen Elemente zurück.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFGetLength([$sMode = "links"])
; Parameter(s): .: $sMode       - Optional: (Default = "links") :
;                               | links
;                               | images
;                               | forms
;                               | frames
;                               | anchors
;                               | tables
;                               | styleSheets
;                               | tabs
;                               | history
;                               | applets
;                               | embeds
;                               | plugins
;                               | tag
; Return Value ..: Success      - Length of the elements
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Mar 23 23:24:13 CET 2009 @975 /Internet Time/
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _FFGetLength($sMode = "links")
	Local Const $sFuncName = "_FFGetLength"

	$sMode = StringLower($sMode)

	Switch $sMode
		Case "links", "images", "forms", "anchors", "applets", "embeds"
			Return _FFCmd("." & $sMode & ".length")
		Case "tabs"
			Return _FFCmd("gBrowser.tabContainer.childNodes.length")
		Case "frames"
			Return _FFCmd("content.frames.length")
		Case "tables"
			Return _FFCmd(".getElementsByTagName('table').length")
		Case "history"
			Return _FFCmd("content.history.length")
		Case "stylesheets"
			Return _FFCmd(".styleSheets.length")
		Case "plugins"
			Return _FFCmd("navigator.plugins.length")
		Case Else
			Local $RetVal = _FFCmd(".getElementsByTagName('" & $sMode & "').length;")
			If Not @error And $RetVal <> "_FFCmd_Err" Then
				Return $RetVal
			Else
				SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch, "$sMode: " & $sMode))
				Return 0
			EndIf
	EndSwitch

	Return 0

EndFunc   ;==>_FFGetLength

; #FUNCTION# ===================================================================
; Name ..........: _FFLinksGetAll
; Description ...: Returns an array with informations about all existing links.
; Beschreibung ..: Gibt ein Array mit Informationen über alle vorhandenen Links zurück.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFLinksGetAll()
; Parameter(s): .:
; Return Value ..: Success      - 2 dim array with the link informations:
;                               | array[n][0] = href
;                               | array[n][1] = hash
;                               | array[n][2] = search
;                               | array[n][3] = name
;                               | array[n][4] = id
;                               | array[n][5] = text
;                               | array[n][6] = innerHTML
;                               | array[n][7] = target
;                               | array[n][8] = protocol
;                               | array[n][9] = port
;                  Failure      - array[0][0] = 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Tue Sep 22 12:00:34 CEST 2009 @458 /Internet Time/
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _FFLinksGetAll()
	Local Const $sFuncName = "_FFLinksGetAll"

	Local $aInfo, $sInfo, $aRet[1][9], $aTmp
	Local $sDelimiter

	Local $bTrc = $_FF_COM_TRACE
	$_FF_COM_TRACE = False

	If _FFGetLength() > 0 Then
		$sDelimiter = "FF" & Random(1000, 9999, 1)
		$sInfo = _FFCmd('FFau3.GetLinks("' & $sDelimiter & '");')
		$aInfo = StringSplit($sInfo, @CRLF)
		If @error Then
			$_FF_COM_TRACE = $bTrc
			SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch))
			Return $aRet[0][0] = 0
		EndIf
		ReDim $aRet[$aInfo[0] - 1][9]
		For $i = 1 To $aInfo[0] - 1
			$aTmp = StringSplit($aInfo[$i], $sDelimiter, 1)
			If @error Then
				SetError(__FFError($sFuncName, $_FF_ERROR_RetValue, "StringSplit: " & $aInfo[$i]))
				ExitLoop
			EndIf
			If UBound($aTmp) <> 11 Then
				$sDelimiter = "FF" & Random(1000, 9999, 1)
				$sInfo = _FFCmd("FFau3.GetLinkInfo(" & $i & " ,'top','" & $sDelimiter & "')")
				$aTmp = StringSplit($sInfo, $sDelimiter, 1)
			EndIf
			For $j = 1 To $aTmp[0] - 1
				$aRet[$i - 1][$j - 1] = $aTmp[$j]
			Next
		Next
		$_FF_COM_TRACE = $bTrc
		Return $aRet
	Else
		SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch))
		$aRet[0][0] = 0
	EndIf

	$_FF_COM_TRACE = $bTrc
	SetError(__FFError($sFuncName, $_FF_ERROR_GeneralError))
	Return $aRet[0][0] = 0
EndFunc   ;==>_FFLinksGetAll

; #FUNCTION# ===================================================================
; Name ..........: _FFConnect
; Description ...: Connects to FireFox / MozRepl
; Beschreibung ..: Stellt eine Verbindung mit Firefox / MozRepl her.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFConnect([$IP = "127.0.0.1"[, $iPort = 4242[, $iTimeOut = 60000]]])
; Parameter(s): .: $IP          - Optional: (Default = "127.0.0.1") :
;                  $iPort       - Optional: (Default = 4242) :
;                  $iTimeOut    - Optional: (Default = 60000) :
; Return Value ..: Success      - 1 / $_FF_GLOBAL_SOCKET > -1
;                  @EXTENDED    - 1 = Non Browser-window
;                               | 2 = Browser-window
;                  Failure      - 0 / $_FF_GLOBAL_SOCKET = -1
;                  @ERROR       -
;                  @EXTENDED    - 0
; Author(s) .....: Thorsten Willert
; Date ..........: Tue Aug 18 09:36:57 CEST 2009
; Link ..........: http://msdn.microsoft.com/en-us/library/ms740668(VS.85).aspx
; Related .......: _FFDisConnect, _FFIsConnected, _FFStart
; Example .......: Yes
; ==============================================================================
Func _FFConnect($IP = "127.0.0.1", $iPort = 4242, $iTimeOut = 60000)
	Local Const $sFuncName = "_FFConnect"
	Local $Extended = 2

	If $_FF_CON_DELAY > 100 Then Opt("TCPTimeout", $_FF_CON_DELAY)

	If $IP = Default Then $IP = "127.0.0.1"
	If $iPort = Default Then $iPort = 4242
	If $iTimeOut = Default Then $iTimeOut = 60000

	If Not __FFIsIP($IP) Then
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, $IP))
		Return 0
	EndIf
	If Not __FFIsPort($iPort) Then
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, $iPort))
		Return 0
	EndIf

	Local $iTCPErr
	Local $bTrace = $_FF_COM_TRACE
	$_FF_COM_TRACE = False

	ConsoleWrite("_FFConnect: OS:" & @TAB & @OSVersion & " " & @OSType & " " & @OSBuild & " " & @OSServicePack & @CRLF)
	ConsoleWrite("_FFConnect: AutoIt:" & @TAB & @AutoItVersion & @CRLF)
	ConsoleWrite("_FFConnect: FF.au3:" & @TAB & $_FF_AU3VERSION & @CRLF)
	ConsoleWrite("_FFConnect: IP:" & @TAB & $IP & @CRLF)
	ConsoleWrite("_FFConnect: Port:" & @TAB & $iPort & @CRLF)

	TCPStartup()
	If Not @error Then
		; delay for connection over network
		$_FF_CON_DELAY = Ping($IP)
		If @error Then
			SetError(__FFError($sFuncName, $_FF_ERROR_GeneralError, "Ping Error: " & @error))
			$_FF_CON_DELAY = $_FF_COM_DELAY_MAX
		EndIf
		$_FF_CON_DELAY += Round($_FF_CON_DELAY * 0.5)
		ConsoleWrite("_FFConnect: Delay: " & @TAB & $_FF_CON_DELAY & "ms" & @CRLF)

		; trying to connect
		Local $iTimeOutTimer = TimerInit()
		While 1
			$_FF_GLOBAL_SOCKET = TCPConnect($IP, $iPort)
			$iTCPErr = @error
			Sleep(500)
			If __FFIsSocket($_FF_GLOBAL_SOCKET) Then ExitLoop
			If (TimerDiff($iTimeOutTimer) > $iTimeOut) Then ; Profil laden
				SetError(__FFError($sFuncName, $_FF_ERROR_Timeout, "TCPConnect Error: " & $iTCPErr))
				$_FF_GLOBAL_SOCKET = -1
				ExitLoop
			EndIf
		WEnd
		If $_FF_GLOBAL_SOCKET <> -1 Then
			ConsoleWrite("_FFConnect: Socket: " & @TAB & $_FF_GLOBAL_SOCKET & @CRLF)
			If __FFWaitForRepl(10000) Then
				_FFCmd(".browserDOMWindow")
				If @error Then
					SetError(__FFError($sFuncName, $_FF_ERROR_GeneralError, "Warning: Connected to a non browser-window"))
					$Extended = 1
				EndIf
				ConsoleWrite("_FFConnect: Browser:" & @TAB & _FFCmd("navigator.userAgent") & @CRLF)
				__FFSendJavaScripts()
				$_FF_COM_TRACE = $bTrace
				Return SetError(0, $Extended, 1)
			EndIf
		Else
			TCPShutdown()
			$_FF_COM_TRACE = $bTrace
			Return SetError(__FFError($sFuncName, $_FF_ERROR_GeneralError, "Timeout: Can not connect to FireFox/MozRepl on: " & $IP & ":" & $iPort), 0, 0)
		EndIf
	Else
		$_FF_COM_TRACE = $bTrace
		Return SetError(__FFError($sFuncName, $_FF_ERROR_GeneralError, "TCPStartup Error: " & @error), 0, 0)
	EndIf

	$_FF_COM_TRACE = $bTrace
	Return SetError($_FF_ERROR_GeneralError, $Extended, 0)
EndFunc   ;==>_FFConnect

; #FUNCTION# ===================================================================
; Name ..........: _FFDisConnect
; Description ...: Disconnects from FireFox
; Beschreibung ..: Trennt die Verbindung mit FireFox / MozRepl
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFDisConnect()
; Parameter(s): .:
; Return Value ..: Success      - 1 / $_FF_GLOBAL_SOCKET = -1
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Wed Apr 15 09:54:09 CEST 2009
; Link ..........: http://msdn.microsoft.com/en-us/library/ms740668(VS.85).aspx
; Related .......: _FFConnect, _FFIsConnected, _FFStart
; Example .......: Yes
; ==============================================================================
Func _FFDisConnect()
	Local Const $sFuncName = "_FFDisConnect"

	Opt("TCPTimeout", 200)

	If __FFIsSocket($_FF_GLOBAL_SOCKET) Then
		__FFSend("repl.quit()")
		TCPCloseSocket($_FF_GLOBAL_SOCKET)
		If @error Then
			SetError(__FFError($sFuncName, $_FF_ERROR_GeneralError, "TCP Error: " & @error))
			Return 0
		EndIf
		TCPShutdown()
		If Not @error Then
			If $_FF_COM_TRACE Then ConsoleWrite("_FFDisConnect: disconnected" & @CRLF)
			$_FF_FRAME = "top"
			$_FF_GLOBAL_SOCKET = -1
			Return 1
		Else
			SetError(__FFError($sFuncName, $_FF_ERROR_GeneralError, "TCP Error: " & @error))
			Return 0
		EndIf
	Else
		SetError(__FFError($sFuncName, $_FF_ERROR_SocketError, $_FF_GLOBAL_SOCKET))
		Return 0
	EndIf

EndFunc   ;==>_FFDisConnect

; #FUNCTION# ===================================================================
; Name ..........: _FFIsConnected
; Description ...: Connection status to FireFox / MozRepl
; Beschreibung ..: Liefert den Status der Verbindung mit FireFox / MozRepl
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFIsConnected()
; Parameter(s): .:
; Return Value ..: Success      - 1
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Sep 28 11:03:44 CEST 2009
; Link ..........:
; Related .......: _FFConnect, _FFDisConnect, _FFStart
; Example .......: Yes
; ==============================================================================
Func _FFIsConnected()
	Local Const $sFuncName = "_FFIsConnected"

	If __FFIsSocket($_FF_GLOBAL_SOCKET) And _FFCmd("FFau3 != null?1:0") = 1 Then Return 1

	SetError(__FFError($sFuncName, $_FF_ERROR_SocketError, $_FF_GLOBAL_SOCKET))
	Return 0
EndFunc   ;==>_FFIsConnected

; #FUNCTION# ===================================================================
; Name ..........: _FFFormCheckBox
; Description ...: Checks or unchecks a CheckBox in a form.
; Beschreibung ..: Wählt eine Checkbox an oder ab.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFFormCheckBox($vBox[, $bChecked = True[, $iBoxNameIndex = 0[, $sBoxMode = "index"[, $vForm = 0[, $sFormMode = "index"[, $bCheckBox = True]]]]]])
; Parameter(s): .: $vBox        - Index name, id or value of the CheckBox
;                  $bChecked    - Optional: (Default = True) :
;                               | True
;                               | False
;                  $iBoxNameIndex - Optional: (Default = 0) : Index if $sBoxMode = "name"
;                  $sBoxMode    - Optional: (Default = "index") :
;                               | index
;                               | name
;                               | id
;                               | value
;                  $vForm       - Optional: (Default = 0) : Index, name or id of the Form
;                  $sFormMode   - Optional: (Default = "index") :
;                               | index
;                               | name
;                               | id
;                               | action
;                               | title
;                  $bCheckBox   - Optional: (Default = True) : Internal use if the function should work for RadioButtons
; Return Value ..: Success      - 1
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Fri Nov 06 18:10:36 CET 2009
; Link ..........:
; Related .......: _FFFormOptionSelect, _FFFormRadioButton, _FFFormReset, _FFFormSubmit
; Example .......: Yes
; ==============================================================================
Func _FFFormCheckBox($vBox, $bChecked = True, $iBoxNameIndex = 0, $sBoxMode = "index", $vForm = 0, $sFormMode = "index", $bCheckBox = True)
	Local Const $sFuncName = "_FFFormCheckBox"

	; Default parameters
	;If $bChecked = Default Then $bChecked = true
	If $iBoxNameIndex = Default Then $iBoxNameIndex = 0
	If $sBoxMode = Default Then $sBoxMode = "index"
	If $vForm = Default Then $vForm = 0
	If $sFormMode = Default Then $sFormMode = "index"

	Local $sType
	If $bCheckBox Then
		$sType = "checkbox"
	Else
		$sType = "radio"
	EndIf

	$sBoxMode = StringLower($sBoxMode)
	$sFormMode = StringLower($sFormMode)

	If Not IsBool($bChecked) Then
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(bool) $bChecked: " & $bChecked))
		Return 0
	EndIf

	Local $sCommand = "checked=" & __FFB2S($bChecked)
	Local $sForm, $sCheckBox

	Switch $sFormMode
		Case "index"
			If Not IsInt($vForm) Then
				SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(int) $vForm: " & $vForm))
				Return 0
			EndIf
			$sForm = $vForm + 1
		Case "id", "name", "action", "title"
			$sForm = StringFormat("@%s='%s'", $sFormMode, $vForm)
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(index|name|id) $sFormMode: " & $sFormMode))
			Return 0
	EndSwitch

	Switch $sBoxMode
		Case "index"
			If Not IsInt($vBox) Then
				SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(int) $vBox: " & $vBox))
				Return 0
			EndIf
			$sCheckBox = "position()=" & $vBox + 1
		Case "name"
			If Not IsInt($iBoxNameIndex) Then
				SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(int) $iBoxNameIndex: " & $iBoxNameIndex))
				Return 0
			EndIf
			$sCheckBox = StringFormat("@name='%s' and position()=%i", $vBox, $iBoxNameIndex + 1)
		Case "id", "value"
			$sCheckBox = StringFormat("@%s='%s'", $sBoxMode, $vBox)
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(index|name|id|value) $sBoxMode: " & $sBoxMode))
			Return 0
	EndSwitch

	_FFXPath(StringFormat("//form[%s]//input[@type='%s' and %s]", $sForm, $sType, $sCheckBox), $sCommand, 9)
	If Not @error Then
		Return 1
	Else
		SetError(@error)
		Return 0
	EndIf

EndFunc   ;==>_FFFormCheckBox

; #FUNCTION# ===================================================================
; Name ..........: _FFFormGetElementsLength
; Description ...: Length of the form elements
; Beschreibung ..: Gibt die Anzahl der Elemente eines Formulars zurück.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFFormGetElementsLength([$vForm = 0[, $sMode = "index"]])
; Parameter(s): .: $vForm       - Optional: (Default = 0) : Index (0-n), name or id
;                  $sMode       - Optional: (Default = "index") :
;                               | index
;                               | name
;                               | id
;                               | action
;                               | title
; Return Value ..: Success      - >0
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Fri Nov 06 22:04:28 CET 2009
; Link ..........:
; Related .......: _FFGetLength
; Example .......: Yes
; ==============================================================================
Func _FFFormGetElementsLength($vForm = 0, $sMode = "index")
	Local Const $sFuncName = "_FFFormGetElementsLength"

	If $vForm = Default Then $vForm = 0
	If $sMode = Default Then $sMode = "index"
	$sMode = StringLower($sMode)
	Local $sForm

	Switch $sMode
		Case "index"
			If IsInt($vForm) Then
				$sForm = $vForm + 1
			Else
				SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(int) $vForm: " & $vForm))
				Return 0
			EndIf
		Case "name", "id", "action", "title"
			$sForm = "@" & $sMode & "='" & $vForm & "'"
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(index|name|id|action|title) $sMode: " & $sMode))
			Return 0
	EndSwitch

	Local $bTrace = $_FF_COM_TRACE
	$_FF_COM_TRACE = False

	_FFXPath(StringFormat("//form[%s]", $sForm), "", 9)
	Local $sRet = _FFCmd("FFau3.xpath.elements.length")
	If Not @error Then
		$_FF_COM_TRACE = $bTrace
		Return $sRet
	Else
		$_FF_COM_TRACE = $bTrace
		SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch, "$vForm: " & $vForm))
		Return 0
	EndIf
EndFunc   ;==>_FFFormGetElementsLength

; #FUNCTION# ===================================================================
; Name ..........: _FFFormOptionSelect
; Description ...: Selects an element of an OptionSelect.
; Beschreibung ..: Wählt ein Element eines OptionSelect aus.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFFormOptionSelect([$vElement = 0[, $sElementMode = "index"[, $vOption = 0[, $sOptionMode = "index"[, $vForm = 0[, $sFormMode = "index"]]]]]])
; Parameter(s): .: $vElement    - Optional: (Default = 0) : index (0-n), name, id
;                  $sElementMode - Optional: (Default = "index") :
;                               | index
;                               | name
;                               | id
;                  $vOption     - Optional: (Default = 0) : index (0-n), name, id, text, value
;                  $sOptionMode - Optional: (Default = "index") :
;                               | index
;                               | name
;                               | id
;                               | text
;                               | value
;                  $vForm       - Optional: (Default = 0) : index (0-n), name, id
;                  $sFormMode   - Optional: (Default = "index") :
;                               | index
;                               | name
;                               | id
;                               | action
;                               | title
; Return Value ..: Success      - 1
;                  Failure      - 0 and sets
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Thu Feb 04 19:13:58 CET 2010
; Link ..........:
; Related .......: _FFFormCheckBox, _FFFormRadioButton, _FFFormReset, _FFFormSubmit
; Example .......: Yes
; ==============================================================================
Func _FFFormOptionSelect($vElement = 0, $sElementMode = "index", $vOption = 0, $sOptionMode = "index", $vForm = 0, $sFormMode = "index")
	Local Const $sFuncName = "_FFFormOptionSelect"

	; Default parameters
	If $vElement = Default Then $vElement = 0
	If $sElementMode = Default Then $sElementMode = "index"
	If $vOption = Default Then $vOption = 0
	If $sOptionMode = Default Then $sOptionMode = "index"
	If $vForm = Default Then $vForm = 0
	If $sFormMode = Default Then $sFormMode = "index"

	Local $sForm, $sElement, $sOption
	$sElementMode = StringLower($sElementMode)
	$sOptionMode = StringLower($sOptionMode)
	$sFormMode = StringLower($sFormMode)

	Switch $sFormMode
		Case "index"
			If Not IsInt($vForm) Then
				SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(int) $vForm: " & $vForm))
				Return 0
			EndIf
			$sForm = $vForm + 1
		Case "id", "name", "action", "title"
			$sForm = StringFormat("@%s='%s'", $sFormMode, $vForm)
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(index|name|id) $sFormMode: " & $sFormMode))
			Return 0
	EndSwitch

	Switch $sElementMode
		Case "index"
			If Not IsInt($vElement) Then
				SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(int) $vElement: " & $vElement))
				Return 0
			EndIf
			$sElement = "position()=" & $vElement + 1
		Case "id", "name"
			$sElement = StringFormat("@%s='%s'", $sElementMode, $vElement)
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(index|name|id) $sElementMode: " & $sElementMode))
			Return 0
	EndSwitch

	Switch $sOptionMode
		Case "index"
			If Not IsInt($vOption) Then
				SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(int) $vOption: " & $vOption))
				Return 0
			EndIf
			$sOption = "position()=" & $vOption + 1
		Case "text"
			$sOption = StringFormat("contains(.,'%s')", $vOption)
		Case "name", "id", "value"
			$sOption = StringFormat("@%s='%s'", $sOptionMode, $vOption)
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(index|text|name|id|value) $sOptionMode: " & $sOptionMode))
			Return 0
	EndSwitch

	Local $sCommand = StringFormat("//form[%s]//select[%s]//option[%s]", $sForm, $sElement, $sOption)
	_FFXPath($sCommand, "selected=true", 9)
	If Not @error Then
		Return 1
	Else
		SetError(@error)
		Return 0
	EndIf

EndFunc   ;==>_FFFormOptionSelect

; #FUNCTION# ===================================================================
; Name ..........: _FFFormRadioButton
; Description ...: Selects a RadioButton in a form
; Beschreibung ..: Wählt einen RadioButton aus.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFFormRadioButton($vRadioButton[, $iRadioButtonIndex = 0[, $sRadioButtonMode = "index"[, $vForm = 0[, $sFormMode = "index"]]]])
; Parameter(s): .: $vRadioButton - name, id or value
;                  $iRadioButtonIndex - Optional: (Default = 0) : RadioButton index (0-n)
;                  $sRadioButtonMode - Optional: (Default = "index") : RadioButton
;                               | index
;                               | name
;                               | id
;                               | value
;                  $vForm       - Optional: (Default = 0) : Form index or name
;                  $sFormMode   - Optional: (Default = "index") :
;                               | index
;                               | name
; Return Value ..: Success      - RadioButton state 1 or 0
;                  Failure      - -1
; Author(s) .....: Thorsten Willert
; Date ..........: Wed Mar 25 09:22:11 CET 2009 @390 /Internet Time/
; Link ..........:
; Related .......: _FFFormCheckBox, _FFFormOptionSelect, _FFFormReset, _FFFormSubmit
; Example .......: Yes
; ==============================================================================
Func _FFFormRadioButton($vRadioButton, $iRadioButtonIndex = 0, $sRadioButtonMode = "index", $vForm = 0, $sFormMode = "index")
	;Local Const $sFuncName = "_FFFormRadioButton"
	Local $RetVal = _FFFormCheckBox($vRadioButton, True, $iRadioButtonIndex, $sRadioButtonMode, $vForm, $sFormMode, False)
	SetError(@error)
	Return $RetVal
EndFunc   ;==>_FFFormRadioButton

; #FUNCTION# ===================================================================
; Name ..........: _FFFormReset
; Description ...: Resets a form
; Beschreibung ..: Setzt ein Formular zurück.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFFormReset([$vForm = 0[, $sMode = "index"]])
; Parameter(s): .: $vForm       - Optional: (Default = 0) : Index (0-n), name, id
;                  $sMode       - Optional: (Default = "index") :
;                               | Index
;                               | name
;                               | id
; Return Value ..: Success      - 1
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Fri Sep 18 06:01:42 CEST 2009 @209 /Internet Time/
; Link ..........:
; Related .......: _FFFormCheckBox, _FFFormOptionSelect, _FFFormSubmit, _FFFormRadioButton
; Example .......: Yes
; ==============================================================================
Func _FFFormReset($vForm = 0, $sMode = "index")
	Local Const $sFuncName = "_FFFormReset"

	Local $sCommand

	Switch StringLower($sMode)
		Case "index"
			If Not IsInt($vForm) Then
				SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(int) $vForm: " & $vForm))
				Return 0
			EndIf
			$sCommand = '.forms[' & $vForm & ']'
		Case "id"
			$sCommand = '.getElementById("' & $vForm & '")'
		Case "name"
			$sCommand = '.getElementsByName("' & $vForm & '")[0]'
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(index|id|name) $sMode: " & $sMode))
			Return 0
	EndSwitch

	Local $sRetVal = _FFCmd($sCommand & ".reset();")
	If Not @error And $sRetVal <> "_FFCmd_Err" Then Return 1

	SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch, "$vForm: " & $vForm))
	Return 0

EndFunc   ;==>_FFFormReset

; #FUNCTION# ===================================================================
; Name ..........: _FFFormSubmit
; Description ...: Submits a form
; Beschreibung ..: Sendet ein Formular ab.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFFormSubmit([$vForm = 0[, $sMode = "index"[, $sSubmitMode = "submit"[, $bLoadWait = True]]]])
; Parameter(s): .: $vForm       - Optional: (Default = 0) : Index-number, name or id
;                  $sSubmitMode - Optional: (Default = "auto") :
;                               | auto
;                               | click
;                               | submit
;                               | keydown
;                               | keypress
;                  $bLoadWait   - Optional: (Default = true) :
;                               | True
;                               | False
; Return Value ..: Success      - 1
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Tue Apr 07 17:56:46 CEST 2009 @706 /Internet Time/
; Link ..........:
; Related .......: _FFFormCheckBox, _FFFormOptionSelect, _FFFormReset, _FFFormRadioButton
; Example .......: Yes
; ==============================================================================
Func _FFFormSubmit($vForm = 0, $sMode = "index", $sSubmitMode = "submit", $bLoadWait = True)
	Local Const $sFuncName = "_FFFormSubmit"

	Local $sRetVal
	Local $iSubmit = -1, $iPassword = -1
	Local $bTrace
	$sMode = StringLower($sMode)

	If $sMode = "index" Then
		If IsInt($vForm) Then
			$vForm = 'forms[' & $vForm & ']'
		Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(int) $vForm: " & $vForm))
			Return 0
		EndIf
	ElseIf $sMode = "id" Then
		$vForm = "getElementById('" & $vForm & "')"
	ElseIf $sMode = "name" Then
		$vForm = "getElementsByName('" & $vForm & "')[0]"
	Else
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(index|id|name) $sMode: " & $sMode))
		Return 0
	EndIf

	Local $iLength = _FFCmd(StringFormat('.%s.elements.length;', $vForm))
	If $iLength = "_FFCmd_Err" Then
		SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch, "No Elements found in form: " & $vForm))
		Return 0
	EndIf

	$bTrace = $_FF_COM_TRACE
	$_FF_COM_TRACE = False
	If $iLength > 0 And ($sSubmitMode = "keypress" Or $sSubmitMode = "keydown" Or $sSubmitMode = "click" Or $sSubmitMode = "auto") Then
		For $i = 0 To $iLength - 1
			$sRetVal = _FFCmd(StringFormat('.%s.elements[%i].type;', $vForm, $i))
			If @error Then Return 0

			If $sRetVal = "submit" Then $iSubmit = $i
			If $sRetVal = "password" Then $iPassword = $i
			If $iSubmit > -1 And $iPassword > -1 Then ExitLoop
		Next
	EndIf
	$_FF_COM_TRACE = $bTrace

	Switch StringLower($sSubmitMode)
		Case "auto"
			ContinueCase
		Case "keypress"
			; fire "enter" with keypress on the password element
			_FFDispatchEvent(StringFormat('.%s.elements[%i]', $vForm, $iPassword), "keypress", 13)
			If Not @error Then
				If $bLoadWait Then
					If Not _FFLoadWait() Then Return 0
				EndIf
				Return 1
			Else
				If $sSubmitMode = "auto" Then ContinueCase
			EndIf
		Case "click"
			; clicks on the submit element
			$sRetVal = _FFCmd(StringFormat('.%s.elements[%i].click();', $vForm, $iSubmit), 3000)
			If Not @error And $sRetVal <> "_FFCmd_Err" Then
				If $bLoadWait Then
					If Not _FFLoadWait() And $sSubmitMode = "auto" Then ContinueCase
				EndIf
				Return 1
			Else
				If $sSubmitMode = "auto" Then ContinueCase
			EndIf
		Case "submit"
			; use form.submit()
			$sRetVal = _FFCmd(StringFormat('.%s.submit();', $vForm))
			If Not @error And $sRetVal <> "_FFCmd_Err" Then
				If $bLoadWait Then
					If Not _FFLoadWait() And $sSubmitMode = "auto" Then ContinueCase
				EndIf
				Return 1
			Else
				If $sSubmitMode = "auto" Then ContinueCase
			EndIf
		Case "keydown"
			; fire "enter" with keydown on the password element
			_FFDispatchEvent(StringFormat(".%s.elements[%i]", $vForm, $iPassword), "keydown", 13)
			If Not @error Then
				If $bLoadWait Then
					If Not _FFLoadWait() And $sSubmitMode = "auto" Then ContinueCase
				EndIf
				Return 1
			EndIf
		Case Else
			Return 0
	EndSwitch

	Return 0

EndFunc   ;==>_FFFormSubmit

; #FUNCTION# ===================================================================
; Name ..........: _FFGetValue
; Description ...: Gets an value of an element
; Beschreibung ..: Liefert das Value eines Elements
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFGetValue($sElement[, $sMode = "elements"[, $iIndex = 0[, $iFilter = 0]]])
; Parameter(s): .: $sElement    -
;                  $sMode       - Optional: (Default = "elements") :
;                               | elements: element or object from _FFXpath or _FFObj*
;                               | id
;                               | name
;                               | class
;                               | tag
;                  $iIndex      - Optional: (Default = 0) : index if $sMode = class, name, tag
;                  $iFilter     - Optional: (Default = 0) : you can add them:
;                               | 0 = No filter
;                               | 1 = Remove all non ASCII
;                               | 2 = Remove all double withespaces
;                               | 4 = Remove all double linefeeds
;                               | 8 = Remove all HTML-tags
;                               | 16 = simple html-tag / entities convertor
; Return Value ..: Success      - 1
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Wed Sep 23 22:55:36 CEST 2009
; Link ..........:
; Related .......: _FFSetValue, _FFXPath
; Example .......: Yes
; ==============================================================================
Func _FFGetValue($sElement, $sMode = "elements", $iIndex = 0, $iFilter = 0)
	Local Const $sFuncName = "_FFGetValue"

	If Not IsInt($iIndex) Then
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(int) $iIndex: " & $iIndex))
		Return 0
	EndIf
	If Not IsInt($iFilter) Then
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(int) $iFilter: " & $iFilter))
		Return 0
	EndIf

	Switch StringLower($sMode)
		Case "elements"
			If StringLeft($sElement, 7) = "OBJECT|" Then $sElement = StringMid($sElement, 8)
		Case "id"
			$sElement = ".getElementById('" & $sElement & "')"
		Case "name"
			$sElement = ".getElementsByName('" & $sElement & "')[" & $iIndex & "]"
		Case "class"
			$sElement = ".getElementsByClassName('" & $sElement & "')[" & $iIndex & "]"
		Case "tag"
			$sElement = ".getElementsByTagName('" & $sElement & "')[" & $iIndex & "]"
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(elements|id|name|class|tag) $sMode: " & $sMode))
			Return 0
	EndSwitch

	If StringLeft($sElement, 1) = "." Then $sElement = "window.content.document" & $sElement

	Local $sRetVal = _FFCmd($sElement & ".value")
	If Not @error And $sRetVal <> "_FFCmd_Err" Then
		If $iFilter > 0 Then __FFFilterString($sRetVal, $iFilter)
		Return $sRetVal
	Else
		SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch, "$sElement: " & $sElement))
		Return 0
	EndIf

EndFunc   ;==>_FFGetValue

; #FUNCTION# ===================================================================
; Name ..........: _FFDispatchEvent
; Description ...: Simulates a HTML event
; Beschreibung ..: Simuliert ein HTML-Ereignis an einem Element.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFDispatchEvent($sElement[, $sEventType = "change"[, $iKeyCode = 13]])
; Parameter(s): .: $sElement    - Element where the event is fired (or object from _FFXpath/_FFObj*)
;                  $sEventType  - Optional: (Default = "change") :
;                               | change
;                               | select
;                               | focus
;                               | blur
;                               | resize
;                               | keydown
;                               | keypress
;                               | keyup
;                  $iKeyCode    - ASCII-KeyCode (decimal)
; Return Value ..: Success      - 1
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Fri Dec 04 15:06:11 CET 2009
; Link ..........: http://www.w3.org/TR/DOM-Level-2-Events/events.html
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _FFDispatchEvent($sElement, $sEventType = "change", $iKeyCode = 13)
	Local Const $sFuncName = "_FFDispatchEvent"

	If $iKeyCode < 0 Or $iKeyCode > 255 Then
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(int) $iKeyCode: 0-255" & $iKeyCode))
		Return 0
	EndIf

	Local $sType

	$sEventType = StringLower($sEventType)
	If StringLeft($sElement, 7) = "OBJECT|" Then $sElement = StringMid($sElement, 8)

	Switch $sEventType
		Case "change", "select", "focus", "blur", "resize"
			$sType = "Event"
		Case "keydown", "keypress", "keyup"
			$sType = "KeyboardEvent"
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(change|select|focus|blur|resize|keydown|keypress|keyup) $sEventType: " & $sEventType))
			Return 0
	EndSwitch

	If StringLeft($sElement, 1) = "." Then $sElement = "window.content.document" & $sElement

	Local $RetVal = _FFCmd(StringFormat("FFau3.simulateEvent(%s,'%s','%s',%s);", $sElement, $sType, $sEventType, $iKeyCode))
	If Not @error And $RetVal <> "_FFCmd_Err" Then
		Return $RetVal
	Else
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "$sElement: " & $sElement))
		Return 0
	EndIf

	Return 0
EndFunc   ;==>_FFDispatchEvent

; #FUNCTION# ===================================================================
; Name ..........: _FFLoadWait
; Description ...: Wait while the page is loading.
; Beschreibung ..: Wartet bis die Seite geladen ist.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFLoadWait([$iTimeOut = $_FF_LOADWAIT_TIMEOUT[, $bStop = $_FF_LOADWAIT_STOP]])
; Parameter(s): .: $iTimeOut    - Optional: (Default = $_FF_LOADWAIT_TIMEOUT) : Timeout while waiting in ms
;                  $bStop       - Optional: (Default = $_FF_LOADWAIT_STOP) : Stops loading after TimeOut
; Return Value ..: Success      - 1
;                  @EXTENDED    - loading time in ms
;                  Failure      - 0
;                  @ERROR       -
;                  @EXTENDED    - loading time in ms
; Author(s) .....: Thorsten Willert
; Date ..........: Thu Feb 04 19:20:36 CET 2010
; Link ..........: https://developer.mozilla.org/en/nsIWebProgress
; Related .......: _FFDialogWait
; Example .......: Yes
; ==============================================================================
Func _FFLoadWait($iTimeOut = $_FF_LOADWAIT_TIMEOUT, $bStop = $_FF_LOADWAIT_STOP)
	Local Const $sFuncName = "_FFLoadWait"

	Local $iLoadingTime = 0
	If $iTimeOut < 1000 Then $iTimeOut = 1000

	Local $bTrace = $_FF_COM_TRACE
	$_FF_COM_TRACE = False

	If Not $_FF_COM_TRACE Then ConsoleWrite("_FFLoadWait: ")

	Local $TimeOutTimer = TimerInit()

	While TimerDiff($TimeOutTimer) < $iTimeOut
		If Not $_FF_COM_TRACE Then ConsoleWrite(".")
		If _FFCmd("FFau3.tmp=window.getBrowser().webProgress;FFau3.tmp.isLoadingDocument && FFau3.tmp.busyFlags ? false: true;") Then
			$iLoadingTime = Round(TimerDiff($TimeOutTimer))
			ConsoleWrite(" loaded in " & $iLoadingTime & "ms" & @CRLF)
			Sleep($_FF_CON_DELAY * 1.4)
			$_FF_COM_TRACE = $bTrace
			Return SetError(0, $iLoadingTime - $_FF_CON_DELAY * 1.4, 1)
		EndIf
		Sleep(250)
	WEnd

	If $bStop Then _FFAction("stop")

	$_FF_COM_TRACE = $bTrace
	ConsoleWrite(@CRLF & @CRLF)

	Return SetError(__FFError($sFuncName, $_FF_ERROR_Timeout, "Can not check site status."), $iLoadingTime, 0)
EndFunc   ;==>_FFLoadWait

; #FUNCTION# ===================================================================
; Name ..........: _FFGetObjectInfo
; Description ...: Returns a 2-dimensional array with every information of an object.
; Beschreibung ..: Liefert ein 2-dimensionales Array mit allen Informationen über ein Objekt.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFGetObjectInfo([$sObject = "window"])
; Parameter(s): .: $sObject     - Optional: (Default = "window") :
; Return Value ..: Success      - Array with object information
;                  Failure      -
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Mar 23 23:05:55 CET 2009 @962 /Internet Time/
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _FFGetObjectInfo($sObject = "window")
	Local Const $sFuncName = "_FFGetObjectInfo"

	Local $aRet[1][2], $aTmp
	If StringLeft($sObject, 1) = "." Then $sObject = "window.content.document" & $sObject

	Local $aInfo = StringSplit(_FFCmd("repl.inspect(" & $sObject & ");"), @CRLF, 2)
	If Not @error And IsArray($aInfo) And $aInfo <> "_FFCmd_Err" Then
		ReDim $aRet[UBound($aInfo)][2]
		For $i = 0 To UBound($aInfo) - 1
			$aTmp = StringRegExp($aInfo[$i], "<object>.(.*?)=(.*?)$", 3)
			If Not @error And IsArray($aTmp) Then
				$aRet[$i][0] = $aTmp[0]
				$aRet[$i][1] = $aTmp[1]
			EndIf
		Next
		Return $aRet
	Else
		SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch, $sObject))
		Return $aRet[0][0] = 0
	EndIf
EndFunc   ;==>_FFGetObjectInfo

; #FUNCTION# ===================================================================
; Name ..........: _FFGetPosition
; Description ...: Returns an array with the position of an element and/or the position and size of the FireFox content area, relative to the client area of the window
; Beschreibung ..: Liefert ein Array mit der Position eines Elementes und/oder die Position und Größe des FireFox-Inhalts, relativ zur Client Area des Fensters
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFGetPosition([$sElement = ""])
; Parameter(s): .: $sElement    - Optional: (Default = "") : Element or FFau3 - object.
;                                 An empty string returns only the position and size of the content area
; Return Value ..: Success      - array with the position and sizes
;                               | [0] = X position of the element (relative to the document)
;                               | [1] = Y position of the element (relative to the document)
;                               | [2] = X position of the content area (relative to the window client)
;                               | [3] = Y position of the content area (relative to the window client)
;                               | [4] = width of the content area
;                               | [6] = height of the content area
;                  Failure      -
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Thu Feb 04 18:49:07 CET 2010
; Link ..........: https://developer.mozilla.org/en/Accessibility/AT-APIs/MSAA/FindWindow
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _FFGetPosition($sElement = "")
	Local Const $sFuncName = "_FFGetPosition"

	Local $aElement[2] = [-1, -1]

	If StringLeft(_FFCmd(".location.href"), 7) = "chrome:" Then
		SetError(__FFError($sFuncName, $_FF_ERROR_GeneralError, "Can't get position from chrome-address"))
		Return
	EndIf

	If $sElement <> "" Then
		If StringLeft($sElement, 7) = "OBJECT|" Then $sElement = StringMid($sElement, 8)
		If StringLeft($sElement, 1) = "." Then $sElement = "window.content.document" & $sElement

		Local $sPos = _FFCmd("FFau3.findPos(" & $sElement & ")")
		If $sPos <> "_FFCmd_Err" Then
			$aElement = StringSplit($sPos, ' ', 2)
		Else
			SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch, "$sElement not found: " & $sElement))
			Return
		EndIf
	EndIf

	Local $aContent = ControlGetPos(_FFCmd(".title"), "", "[CLASS:MozillaContentWindowClass]")
	If @error Then
		SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch, "Cant find 'MozillaContentWindowClass'-control " & @error))
		Local $aContent[4] = [-1, -1, -1, -1]
	EndIf

	Local $aRet[6] = [$aElement[0], $aElement[1], $aContent[0], $aContent[1], $aContent[2], $aContent[3]]

	Return $aRet
EndFunc   ;==>_FFGetPosition

; #FUNCTION# ===================================================================
; Name ..........: _FFOpenURL
; Description ...: Opens a new URL
; Beschreibung ..: Öffnet eine URL.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFOpenURL($sURL[, $bLoadWait = True])
; Parameter(s): .: $sURL        - URL or chrome (shortcut)
;                  $bLoadWait   - Optional: (Default = True) :
; Return Value ..: Success      - 1
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Tue Feb 16 12:49:03 CET 2010
; Link ..........:
; Related .......: _FFTabAdd
; Example .......: Yes
; ==============================================================================
Func _FFOpenURL($sURL, $bLoadWait = True)
	Local Const $sFuncName = "_FFOpenURL"

	If Not __FFCheckURL($sURL) Then
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(URL) $sURL: " & $sURL))
		Return 0
	EndIf

	If StringLeft($sURL, 7) = "chrome:" Then $sURL = __FFChromeSelect(StringMid($sURL, 8))

	ConsoleWrite("_FFOpenURL: " & $sURL & @CRLF)

	Local $sRetVal = _FFCmd(StringFormat(".location.href='%s'", $sURL))
	If Not @error And $sRetVal = $sURL Then
		If $bLoadWait Then
			$sRetVal = _FFLoadWait()
		Else
			$sRetVal = 1
		EndIf
		$_FF_FRAME = 'top'
		_FFCmd("FFau3.WCD=window.content.top.document;")
		Return $sRetVal
	Else
		SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch, $sURL))
		Return 0
	EndIf

EndFunc   ;==>_FFOpenURL

; #FUNCTION# ===================================================================
; Name ..........: _FFPrefReset
; Description ...: Resets a config entry
; Beschreibung ..: Setzt einen Eintrag in der Config zurück
; Syntax ........: _FFPrefReset($sName)
; Parameter(s): .: $sName       - Name of the config entry
; Return Value ..: Success      - 1
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Modified ......: Wed Mar 25 18:06:53 CET 2009 @754 /Internet Time/
; Link ..........: https://developer.mozilla.org/en/nsIPrefBranch, http://kb.mozillazine.org/About:config_entries
; Related .......: _FFPrefSet, _FFPrefGet
; Example .......: Yes
; ==============================================================================
Func _FFPrefReset($sName)
	Local Const $sFuncName = "_FFPrefReset"

	Local $iType = _FFCmd('FFau3.obj=Components.classes["@mozilla.org/preferences-service;1"].getService(Components.interfaces.nsIPrefBranch);FFau3.obj.getPrefType("' & $sName & '");')
	If @error Then
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "$sName: " & $sName))
		Return 0
	EndIf

	If $iType <> 0 Then
		If Not _FFCmd('FFau3.obj.prefHasUserValue("' & $sName & '");') Then Return 1

		_FFCmd('FFau3.obj.clearUserPref("' & $sName & '");')
		If Not @error Then Return 1
	Else
		SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch, "$sName: " & $sName))
	EndIf

	Return 0
EndFunc   ;==>_FFPrefReset

; #FUNCTION# ===================================================================
; Name ..........: _FFPrefGet
; Description ...: Gets a config entry
; Beschreibung ..: Ließt einen Wert aus der Config aus
; Syntax ........: _FFPrefGet($sName)
; Parameter(s): .: $sName       - Name of the config entry
; Return Value ..: Success      - Value of the entry
;                  Failure      - 0 and sets @error
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Modified ......: Wed Mar 25 18:06:59 CET 2009 @754 /Internet Time/
; Link ..........: https://developer.mozilla.org/en/nsIPrefBranch, http://kb.mozillazine.org/About:config_entries
; Related .......: _FFPrefSet, _FFPrefReset
; Example .......: Yes
; ==============================================================================
Func _FFPrefGet($sName)
	Local Const $sFuncName = "_FFPrefGet"

	Local $sCommand
	Local $iType = _FFCmd('FFau3.obj=Components.classes["@mozilla.org/preferences-service;1"].getService(Components.interfaces.nsIPrefBranch); FFau3.obj.getPrefType("' & $sName & '");')
	If @error Then
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, $sName))
		Return 0
	EndIf

	Switch $iType
		Case 0
			SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch, "$sName: " & $sName))
			Return 0
		Case 32 ; PREF_STRING
			$sCommand = "getCharPref"
		Case 64 ; PREF_INT
			$sCommand = "getIntPref"
		Case 128 ; PREF_BOOL
			$sCommand = "getBoolPref"
	EndSwitch

	$sCommand = StringFormat('FFau3.obj.%s("%s");', $sCommand, $sName)

	Return _FFCmd($sCommand)
EndFunc   ;==>_FFPrefGet

; #FUNCTION# ===================================================================
; Name ..........: _FFPrefSet
; Description ...: Sets a config entry
; Beschreibung ..: Setzt einen Wert in der Config.
; Syntax ........: _FFPrefSet($sName, $vValue)
; Parameter(s): .: $sName       - Name of the config entry
;                  $vValue      - Value of the config entry
; Return Value ..: Success      - 1 and sets @EXTENDED
;                  Failure      - 0
;                  @ERROR       -
;                  @EXTENDED    - Old value
; Author(s) .....: Thorsten Willert
; Modified ......: Wed Mar 25 18:07:07 CET 2009 @754 /Internet Time/
; Link ..........: https://developer.mozilla.org/en/nsIPrefBranch, http://kb.mozillazine.org/About:config_entries
; Related .......: _FFPrefGet, _FFPrefReset
; Example .......: Yes
; ==============================================================================
Func _FFPrefSet($sName, $vValue)
	Local Const $sFuncName = "_FFPrefSet"

	Local $sCommand
	Local $iType = _FFCmd('FFau3.obj=Components.classes["@mozilla.org/preferences-service;1"].getService(Components.interfaces.nsIPrefBranch);FFau3.obj.getPrefType("' & $sName & '");')
	If @error Then
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, $sName))
		Return 0
	EndIf
	Local $vOldValue = _FFPrefGet($sName)

	Switch $iType
		Case 0
			SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch, "$sName: " & $sName))
			Return 0
		Case 32 ; PREF_STRING
			If Not IsString($vValue) Then
				SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(string) $vValue: " & $vValue))
				Return 0
			EndIf
			$sCommand = StringFormat('FFau3.obj.%s("%s","%s");', "setCharPref", $sName, $vValue)
		Case 64 ; PREF_INT
			If Not IsInt($vValue) Then
				SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(int) $vValue: " & $vValue))
				Return 0
			EndIf
			$sCommand = StringFormat('FFau3.obj.%s("%s",%s);', "setIntPref", $sName, $vValue)
		Case 128 ; PREF_BOOL
			If Not IsBool($vValue) Then
				SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(bool) $vValue: " & $vValue))
				Return 0
			EndIf
			$sCommand = StringFormat('FFau3.obj.%s("%s",%s);', "setBoolPref", $sName, __FFB2S($vValue))
	EndSwitch

	_FFCmd($sCommand)
	If Not @error Then
		Local $vRetVal = _FFPrefGet($sName)
		If Not @error And $vRetVal = $vValue Then
			SetExtended($vOldValue)
			Return 1
		Else
			SetError(__FFError($sFuncName, $_FF_ERROR_RetValue, "$vValue <> " & $vRetVal))
			Return 0
		EndIf
	EndIf

	Return 0
EndFunc   ;==>_FFPrefSet

; #FUNCTION# ===================================================================
; Name ..........: _FFQuit
; Description ...: Quits FireFox
; Beschreibung ..: Schließt FireFox
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFQuit()
; Parameter(s): .:
; Return Value ..: Success      - 1 / $_FF_GLOBAL_SOCKET = -1
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Sat Mar 14 14:44:55 CET 2009 @614 /Internet Time/
; Link ..........:
; Related .......: _FFDisConnect
; Example .......: Yes
; ==============================================================================
Func _FFQuit()
	Local Const $sFuncName = "_FFQuit"

	If $_FF_GLOBAL_SOCKET = -1 Then Return
	If Not __FFIsSocket($_FF_GLOBAL_SOCKET) Then
		SetError(__FFError($sFuncName, $_FF_ERROR_SocketError))
		Return 0
	EndIf

	TCPSend($_FF_GLOBAL_SOCKET, "close()" & @CRLF)
	If Not @error Then
		$_FF_FRAME = "top"
		$_FF_GLOBAL_SOCKET = -1
		If $_FF_COM_TRACE Then ConsoleWrite('__FFQuit: Closing FireFox ...' & @CRLF)
		Return 1
	Else
		Return 0
	EndIf

EndFunc   ;==>_FFQuit

; #FUNCTION# ===================================================================
; Name ..........: _FFReadHTML
; Description ...: Returns the HTML-source code of a page.
; Beschreibung ..: Liefert den HTMl-Quellcode einer Seite.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFReadHTML([$sMode = "body"[, $iFilter = 0]])
; Parameter(s): .: $sMode       - Optional: (Default = "body") :
;                               | body = innerHTML from <body>
;                               | html = innerHTML from <html>
;                  $iFilter     - Optional: (Default = 0) : you can add them:
;                               | 0 = No filter
;                               | 1 = Remove all non ASCII
;                               | 2 = Remove all double withespaces
;                               | 4 = Remove all double linefeeds
;                               | 8 = Remove all HTML-tags
;                               | 16 = simple html-tag / entities convertor
; Return Value ..: Success      - String with the HTML-source code
;                  Failure      - ""
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Mar 23 22:55:56 CET 2009 @955 /Internet Time/
; Link ..........:
; Related .......: _FFReadText, _FFXPath
; Example .......: Yes
; ==============================================================================
Func _FFReadHTML($sMode = "body", $iFilter = 0)
	Local Const $sFuncName = "_FFReadHTML"

	Local $sCommand

	Switch StringLower($sMode)
		Case "body"
			$sCommand = ".body"
		Case "html"
			$sCommand = ".documentElement"
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(body|html) $sMode: " & $sMode))
			Return 0
	EndSwitch

	Local $sRetVal = _FFCmd($sCommand & ".innerHTML;", 10000)
	If Not @error And $sRetVal <> "_FFCmd_Err" Then
		If $iFilter > 0 Then $sRetVal = __FFFilterString($sRetVal, $iFilter)
		Return $sRetVal
	Else
		SetError(1)
		Return ""
	EndIf
EndFunc   ;==>_FFReadHTML

; #FUNCTION# ===================================================================
; Name ..........: _FFReadText
; Description ...: Returns the text of a page.
; Beschreibung ..: Liefert den Text einer Seite.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFReadText([$iFilter = 0])
; Parameter(s): .: $iFilter     - Optional: (Default = 0) : you can add them:
;                               | 0 = No filter
;                               | 1 = Remove all non ASCII
;                               | 2 = Remove all double withespaces
;                               | 4 = Remove all double linefeeds
;                               | 8 = Remove all HTML-tags
;                               | 16 = simple html-tag / entities convertor
; Return Value ..: Success      - String
;                  Failure      - ""
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Mar 23 22:55:05 CET 2009 @954 /Internet Time/
; Link ..........:
; Related .......: _FFReadHTML, _FFXPath
; Example .......: Yes
; ==============================================================================
Func _FFReadText($iFilter = 0)
	;Local Const $sFuncName = "_FFReadText"

	Local $sRetVal = _FFCmd(".documentElement.textContent;", 10000)
	If Not @error And $sRetVal <> "_FFCmd_Err" Then
		If $iFilter > 0 Then $sRetVal = __FFFilterString($sRetVal, $iFilter)
		Return $sRetVal
	Else
		SetError(1)
		Return ""
	EndIf
EndFunc   ;==>_FFReadText

; #FUNCTION# ===================================================================
; Name ..........: _FFSearch
; Description ...: Searches a string
; Beschreibung ..: Sucht einen Text.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFSearch($sSearchString[, $bCaseSensitive = False[, $bBackwards = False[, $bWrapAround = True[, $bWholeWord = False[, $bSearchInFrames = True]]]]])
; Parameter(s): .: $sSearchString - String to search
;                  $bCaseSensitive - Optional: (Default = false) :
;                  $bBackwards  - Optional: (Default = false) :
;                  $bWrapAround - Optional: (Default = true) :
;                  $bWholeWord  - Optional: (Default = false) :
;                  $bSearchInFrames - Optional: (Default = true) :
; Return Value ..: Success      - 1
;                  Failure      - 0 and sets
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Thu Dec 03 20:53:30 CET 2009
; Link ..........:
; Related .......: _FFAction
; Example .......: Yes
; ==============================================================================
Func _FFSearch($sSearchString, $bCaseSensitive = False, $bBackwards = False, $bWrapAround = True, $bWholeWord = False, $bSearchInFrames = True)
	Local Const $sFuncName = "_FFSearch"

	If $sSearchString = "" Then
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "Empty String $sSearchString"))
		Return ""
	EndIf

	If $bCaseSensitive = Default Then $bCaseSensitive = False
	If $bBackwards = Default Then $bBackwards = False
	If $bWrapAround = Default Then $bWrapAround = True
	If $bWholeWord = Default Then $bWholeWord = False
	If $bSearchInFrames = Default Then $bSearchInFrames = True

	Local $sCommand = 'content.find("' & __FFValue2JavaScript($sSearchString) & '", ' & _
			__FFB2S($bCaseSensitive) & ', ' & _
			__FFB2S($bBackwards) & ', ' & _
			__FFB2S($bWrapAround) & ', ' & _
			__FFB2S($bWholeWord) & ', ' & _
			__FFB2S($bSearchInFrames) & ', false);'

	Local $RetVal = _FFCmd($sCommand)
	If Not @error And $RetVal <> "_FFCmd_Err" Then
		Return 1
	Else
		SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch, "$sSearchString: " & $sSearchString))
		Return 0
	EndIf
EndFunc   ;==>_FFSearch

; #FUNCTION# ===================================================================
; Name ..........: _FFCmd
; Description ...: Send and receive data from Firefox.
; Beschreibung ..: Sendet und empfängt Daten von FireFox.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFCmd($sArg[, $iTimeOut = 30000[, $bTry = True[, $bRec = True]]])
; Parameter(s): .: $sArg        - Any JavaScript
;                  $iTimeOut    - Optional: (Default = 30000) :
;                  $bTry        - Optional: (Default = True) : adds try/catch to $sArg
;                  $bRec        - Optional: (Default = True) : record command
; Return Value ..: Success      - String
;                  Failure      - ""
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Fri Sep 18 06:03:57 CEST 2009 @211 /Internet Time/
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _FFCmd($sArg, $iTimeOut = 30000, $bTry = True)
	Local Const $sFuncName = "_FFCmd"
	Local $sRet

	If StringLeft($sArg, 1) = "." Then $sArg = "window.content.document" & $sArg

	If $bTry And Not StringInStr(StringLower($sArg), "try") Then $sArg = "try{" & $sArg & "}catch(e){'_FFCmd_Err';};"

	Local $sArgWrapped = StringReplace($sArg, "window.content.document", "window.content.wrappedJSObject.document")

	$sArg = StringReplace($sArg, "window.content.document", "window.content." & $_FF_FRAME & ".document")
	$sArgWrapped = StringReplace($sArgWrapped, "window.content.wrappedJSObject.document", "window.content.wrappedJSObject." & $_FF_FRAME & ".document")

	If (Not StringInStr($sArg, ".frames[")) Or StringInStr($sArg, "evaluate") Or (Not StringRegExp($sArg, "\.[a-zA-z]+\(.*?\)(\[.+\])?\.", 0)) Then
		If __FFSend($sArg) Then

			$sRet = __FFRecv($iTimeOut)

			If Not @error Or String($sRet) <> "_FFCmd_Err" Then
				Return $sRet
			ElseIf StringInStr($sArgWrapped, "wrappedJSObject") Then
				__FFSend($sArgWrapped)
				$sRet = __FFRecv($iTimeOut)
				If Not @error And String($sRet) <> "_FFCmd_Err" Then Return $sRet
			EndIf
		EndIf
	Else
		__FFSend($sArgWrapped)
		$sRet = __FFRecv($iTimeOut)
		If Not @error And String($sRet) <> "_FFCmd_Err" Then Return $sRet
	EndIf

	SetError(__FFError($sFuncName, $_FF_ERROR_RetValue, $sRet))
	Return ""
EndFunc   ;==>_FFCmd

; #FUNCTION# ===================================================================
; Name ..........: _FFSetValue
; Description ...: Sets a value of an element
; Beschreibung ..: Setzt den Wert (value) eines Elements
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFSetValue($sValue, $sElement[, $sMode = "elements"[, $iIndex = 0]])
; Parameter(s): .: $sElement    -
;                  $sMode       - Optional: (Default = "elements") :
;                               | elements: element or object from _FFXpath or _FFObj*
;                               | id
;                               | name
;                               | class
;                               | tag
;                  $iIndex      - Optional: (Default = 0) : index if $sMode = class, name, tag
; Return Value ..: Success      - 1
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Wed Sep 23 22:55:36 CEST 2009
; Link ..........:
; Related .......: _FFSetValue, _FFXPath
; Example .......: Yes
; ==============================================================================
Func _FFSetValue($sValue, $sElement, $sMode = "elements", $iIndex = 0)
	Local Const $sFuncName = "_FFSetValue"

	If Not IsInt($iIndex) Then
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(int) $iIndex: " & $iIndex))
		Return 0
	EndIf

	Switch StringLower($sMode)
		Case "elements"
			If StringLeft($sElement, 7) = "OBJECT|" Then $sElement = StringMid($sElement, 8)
		Case "id"
			$sElement = ".getElementById('" & $sElement & "')"
		Case "name"
			$sElement = ".getElementsByName('" & $sElement & "')[" & $iIndex & "]"
		Case "class"
			$sElement = ".getElementsByClassName('" & $sElement & "')[" & $iIndex & "]"
		Case "tag"
			$sElement = ".getElementsByTagName('" & $sElement & "')[" & $iIndex & "]"
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(elements|id|name|class|tag) $sMode: " & $sMode))
			Return 0
	EndSwitch

	__FFValue2JavaScript($sValue)

	Local $sRetVal = _FFCmd($sElement & ".value='" & $sValue & "'")
	If Not @error And $sRetVal <> "_FFCmd_Err" Then
		Return 1
	Else
		SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch, "$sElement: " & $sElement))
		Return 0
	EndIf

EndFunc   ;==>_FFSetValue

; #FUNCTION# ===================================================================
; Name ..........: _FFStart
; Description ...: Starts FireFox
; Beschreibung ..: Startet FireFox
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFStart([$sURL = "about:blank"[, $sProfile = "default"[, $iMode = 1[, $bHide = False[, $IP = "127.0.0.1"[, $iPort = 4242]]]]]])
; Parameter(s): .: $sURL        - Optional: (Default = "about:blank") : Start URL
;                  $sProfile    - Optional: (Default = "default") : User profile
;                  $iMode       - Optional: (Default = 1) : Start mode:
;                               | 0 = Connect to existing process
;                               | 1 = Start new process
;                               | 2 = Connect to an existing process, on failure start a new process
;                               | +8 = Starting FF with parameter -no-remote
;                  $bHide       - Optional: (Default = False) : Start in hidden mode
;                  $IP          - Optional: (Default = "127.0.0.1") : IP of MozRepl
;                  $iPort       - Optional: (Default = 4242) : Port number of MozRepl
; Return Value ..: Success      - 1
;                  @EXTENDED    - PID
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Thu Apr 30 22:10:07 CEST 2009
; Link ..........:
; Related .......: _FFQuit, _FFConnect, _FFDisConnect, _FFIsConnected
; Example .......: Yes
; ==============================================================================
Func _FFStart($sURL = "about:blank", $sProfile = "default", $iMode = 1, $bHide = False, $IP = "127.0.0.1", $iPort = 4242)
	Local Const $sFuncName = "_FFStart"

	; Default parameters
	If $sURL = Default Then $sURL = "about:blank"
	If $IP = Default Then $IP = "127.0.0.1"
	If $iPort = Default Then $iPort = 4242

	If Not __FFCheckURL($sURL) Then
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(URL) $sURL: " & $sURL))
		Return 0
	EndIf
	If Not __FFIsIP($IP) Then
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(IP) $IP: " & $IP))
		Return 0
	EndIf
	If Not __FFIsPort($iPort) Then
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(PORT) $iPort: " & $iPort))
		Return 0
	EndIf

	; Default parameters
	If $sProfile = Default Then $sProfile = "default"
	If $iMode = Default Then $iMode = 1
	If $bHide = Default Then $bHide = False

	Local $PID = -1, $bNoRemote = False

	If $iMode >= 8 Then
		$iMode -= 8
		$bNoRemote = True
	EndIf

	Switch $iMode
		Case 2
			ContinueCase
		Case 0
			If ProcessExists($_FF_PROC_NAME) Then _FFConnect($IP, $iPort)
			If $_FF_GLOBAL_SOCKET > -1 Then
				_FFOpenURL($sURL)
			ElseIf $iMode = 2 Then
				ContinueCase
			EndIf
		Case 1
			__FFStartProcess($sURL, True, $sProfile, $bNoRemote, $bHide, $iPort)
			$PID = @extended
			If ProcessExists($_FF_PROC_NAME) Then _FFConnect($IP, $iPort)
			If $_FF_GLOBAL_SOCKET > -1 Then _FFLoadWait()
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "$iMode: " & $iMode))
			Return 0
	EndSwitch

	If $bHide Then
		Local $hWin = _FFWindowGetHandle()
		WinSetState($hWin, "", @SW_HIDE)
	EndIf

	SetExtended($PID)
	Return 1
EndFunc   ;==>_FFStart

; #FUNCTION# ===================================================================
; Name ..........: _FFTabAdd
; Description ...: Opens a new tab.
; Beschreibung ..: Öffnet einen neuen Tab
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFTabAdd([$sURL = "about:blank"[, $bSelect = True[, $bLoadWait = True]]])
; Parameter(s): .: $sURL        - Optional: (Default = "about:blank") :
;                  $bSelect     - Optional: (Default = true) : Select new tab
;                  $bLoadWait   - Optional: (Default = true) : Wait while the page is loading
; Return Value ..: Success      - 1
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Mar 23 22:44:43 CET 2009 @947 /Internet Time/
; Link ..........: https://developer.mozilla.org/en/XUL/tabs
; Related .......: _FFTabDuplicate, _FFTabClose, _FFTabCloseAll, _FFTabExists, _FFTabSetSelected, _FFTabGetSelected
; Example .......: Yes
; ==============================================================================
Func _FFTabAdd($sURL = "about:blank", $bSelect = True, $bLoadWait = True)
	Local Const $sFuncName = "_FFOpenURL"

	If $sURL = Default Then $sURL = "about:blank"
	If Not __FFCheckURL($sURL) Then
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(URL) $sURL: " & $sURL))
		Return 0
	EndIf

	If $bSelect Then
		_FFCmd("gBrowser.selectedTab = gBrowser.addTab('" & $sURL & "')", 3000)
		If Not @error Then
			If $bLoadWait Then Return _FFLoadWait()
			Return 1
		EndIf
	Else
		_FFCmd("gBrowser.loadOneTab('" & $sURL & "',null,null,null,true)", 3000)
		If Not @error Then
			If $bLoadWait Then Return _FFLoadWait()
			Return 1
		EndIf
	EndIf

	Return 0

EndFunc   ;==>_FFTabAdd

; #FUNCTION# ===================================================================
; Name ..........: _FFTabClose
; Description ...: Closes one or all existing tabs.
; Beschreibung ..: Schließt einen oder alle vorhandenen Tabs.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFTabClose([$vTab = -1[, $sMode = "index"]])
; Parameter(s): .: $vTab        - Optional: (Default = -1 = current tab) :
;                               | index (0-n)
;                               | label
;                               | prev
;                               | next
;                               | first
;                               | last
;                               | all (except current)
;                  $sMode       - Optional: (Default = "index") :
;                               | index
;                               | label
;                               | key or keyword
; Return Value ..: Success      - 1
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Aug 03 16:13:15 CEST 2009
; Link ..........: https://developer.mozilla.org/en/XUL/tabs
; Related .......: _FFTabAdd, _FFTabDuplicate, _FFTabCloseAll, _FFTabExists, _FFTabSetSelected, _FFTabGetSelected
; Example .......: Yes
; ==============================================================================
Func _FFTabClose($vTab = -1, $sMode = "index")
	Local Const $sFuncName = "_FFTabClose"

	Local $sCommand
	;Local $bWarn
	Local $iLength = _FFGetLength("tabs")

	Switch StringLower($sMode)
		Case "index"
			If $vTab = -1 Then
				$sCommand = "gBrowser.removeCurrentTab();"
			ElseIf $vTab > -1 And $vTab < $iLength Then
				$sCommand = "gBrowser.removeTab(gBrowser.mTabs[" & $vTab & "]);"
			Else
				SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(int) $vTab: " & $vTab))
				Return 0
			EndIf
		Case "label"
			If _FFTabExists($vTab) Then
				$sCommand = "gBrowser.removeTab(gBrowser.mTabs[" & @extended & "]);"
			Else
				SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch, "$vTab: " & $vTab))
				Return 0
			EndIf
		Case "key", "keyword"
			Switch StringLower($vTab)
				Case "prev"
					$sCommand = "gBrowser.removeTab(gBrowser.mTabs[gBrowser.tabContainer.selectedIndex-1]);"
				Case "next"
					$sCommand = "gBrowser.removeTab(gBrowser.mTabs[gBrowser.tabContainer.selectedIndex+1]);"
				Case "first"
					$sCommand = "gBrowser.removeTab(gBrowser.mTabs[0]);"
				Case "last"
					$sCommand = "gBrowser.removeTab(gBrowser.mTabs[gBrowser.tabContainer.childNodes.length-1]);"
				Case "all"
					; browser.tabs.warnOnClose
					$sCommand = "gBrowser.removeAllTabsBut(gBrowser.selectedTab);"
				Case Else
					SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(prev|next|first|last|all) $vTab: " & $vTab))
					Return 0
			EndSwitch
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(index|label|key|keyword) $sMode: " & $sMode))
			Return 0
	EndSwitch

	If _FFCmd($sCommand & "gBrowser.tabContainer.childNodes.length", 3000) < $iLength - 1 Then
		Sleep(100)
		Return 1
	EndIf

	Return 0
EndFunc   ;==>_FFTabClose

; #FUNCTION# ===================================================================
; Name ..........: _FFTabDuplicate
; Description ...: Duplicates a tab
; Beschreibung ..: Dupliziert einen Tab
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFTabDuplicate([$vTab = -1[, $sMode = "index"[, $bSelect = False]]])
; Parameter(s): .: $vTab        - Optional: (Default = -1) : Tab to duplicate, -1 = current tab
;                  $sMode       - Optional: (Default = "index") :
;                               | index
;                               | label
;                  $bToFront    - Optional: (Default = false) : Bring duplicated tab to front
; Return Value ..: Success      - Lenght of tabs >= 1
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Aug 03 16:13:08 CEST 2009
; Link ..........: https://developer.mozilla.org/en/XUL/tabs
; Related .......: _FFTabAdd, _FFTabClose, _FFTabCloseAll, _FFTabExists, _FFTabSetSelected, _FFTabGetSelected
; Example .......: Yes
; ==============================================================================
Func _FFTabDuplicate($vTab = -1, $sMode = "index", $bSelect = False)
	Local Const $sFuncName = "_FFTabDuplicate"

	If $sMode = "index" And Not IsInt($vTab) Then
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(int) $vTab: " & $vTab))
		Return 0
	EndIf

	If $vTab = Default Then $vTab = -1
	If $sMode = Default Then $sMode = "index"

	Local $iRet
	Local $sCommand

	Switch StringLower($sMode)
		Case "index"
			If $vTab = -1 Then
				$sCommand = "gBrowser.duplicateTab(gBrowser.mCurrentTab);"
			ElseIf $vTab >= 0 And $vTab < _FFGetLength("tabs") Then
				$sCommand = "gBrowser.duplicateTab(gBrowser.mTabs[" & $vTab & "]);"
			Else
				SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "$vTab: " & $vTab))
				Return 0
			EndIf
		Case "label"
			If _FFTabExists($vTab) Then
				$sCommand = "gBrowser.duplicateTab(gBrowser.mTabs[" & @extended & "]);"
			Else
				SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch, "$vTab: " & $vTab))
				Return 0
			EndIf
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(index|label) $sMode: " & $sMode))
			Return 0
	EndSwitch

	$iRet = _FFCmd($sCommand & "gBrowser.tabContainer.childNodes.length;", 3000)

	If $bSelect Then _FFTabSetSelected("last")

	Return $iRet

EndFunc   ;==>_FFTabDuplicate

; #FUNCTION# ===================================================================
; Name ..........: _FFTabExists
; Description ...: Checks if a tab exists
; Beschreibung ..: Prüft ob ein Tab existiert.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFTabExists($sLabel)
; Parameter(s): .: $sLabel      - Label of the tab (substring or RegEx)
; Return Value ..: Success      - 1
;                  @EXTENDED    - Index of the tab
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Jul 27 08:47:03 CEST 2009
; Link ..........: https://developer.mozilla.org/en/XUL/tabs
; Related .......: _FFTabAdd, _FFTabDuplicate, _FFTabClose, _FFTabCloseAll, _FFTabSetSelected, _FFTabGetSelected
; Example .......: Yes
; ==============================================================================
Func _FFTabExists($sLabel)
	Local Const $sFuncName = "_FFTabExists"

	_FFLoadWait()
	Local $RetVal = Int(_FFCmd("FFau3.SearchTab('" & $sLabel & "')", 3000))
	Switch @error
		Case $_FF_ERROR_Success
			If $RetVal > -1 Then
				SetExtended($RetVal)
				Return 1
			ElseIf $RetVal = -3 Then
				SetError(__FFError($sFuncName, $_FF_ERROR_InvalidExpression, "$sLabel: " & $sLabel))
				Return 0
			Else
				SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch, "$sLabel: " & $sLabel))
				Return 0
			EndIf
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_GeneralError))
			Return 0
	EndSwitch
EndFunc   ;==>_FFTabExists

; #FUNCTION# ===================================================================
; Name ..........: _FFTabSetSelected
; Description ...: Select a tab
; Beschreibung ..: Wählt ein Tab aus.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFTabSetSelected([$vTab = 0[, $sMode = "index"]])
; Parameter(s): .: $vTab        - Optional: (Default = 0) :
;                               | index (0-n)
;                               | label = Substring or RegEx
;                               | prev
;                               | next
;                               | first
;                               | last
;                  $sMode       - Optional: (Default = "index") :
;                               | index
;                               | label
;                               | otherwise the keywords in $vTab are used
; Return Value ..: Success      -
;                  Failure      - -1
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Aug 03 16:12:59 CEST 2009
; Link ..........: https://developer.mozilla.org/en/XUL/tabs
; Related .......: _FFTabAdd, _FFTabDuplicate, _FFTabClose, _FFTabCloseAll, _FFTabExists, _FFTabGetSelected
; Example .......: Yes
; ==============================================================================
Func _FFTabSetSelected($vTab = 0, $sMode = "index")
	Local Const $sFuncName = "_FFTabSetSelected"

	Local $sCommand = "gBrowser.tabContainer"

	If $sMode = "index" Then
		$sCommand &= ".selectedIndex = " & $vTab
	ElseIf $sMode = "label" Then
		If _FFTabExists($vTab) Then
			$sCommand &= ".selectedIndex = " & @extended
		Else
			Return -1
		EndIf
	Else
		Switch StringLower($vTab)
			Case "prev"
				$sCommand &= ".advanceSelectedTab( -1, true )"
			Case "next"
				$sCommand &= ".advanceSelectedTab( 1, true )"
			Case "first"
				$sCommand &= ".selectedIndex = 0"
			Case "last"
				$sCommand &= ".selectedIndex = gBrowser.tabContainer.childNodes.length -1"
			Case Else
				SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(index|label|prev|next|first|last) $sMode: " & $sMode))
				Return -1
		EndSwitch
	EndIf

	Local $RetVal = _FFCmd($sCommand, 3000)
	Sleep(50)
	Return $RetVal

EndFunc   ;==>_FFTabSetSelected

; #FUNCTION# ===================================================================
; Name ..........: _FFTabGetSelected
; Description ...: Returns the name or the index of the selected tab.
; Beschreibung ..: Gibt den Namen (label) oder den Index des aktuellen Tabs zurück.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFTabGetSelected([$sMode = "index"])
; Parameter(s): .: $sMode       - Optional: (Default = "index") :
;                               | index
;                               | label
; Return Value ..: Success      - Tab index 0-n or label
;                  Failure      - -1
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Thu Apr 23 08:52:13 CEST 2009 @327 /Internet Time/
; Link ..........: https://developer.mozilla.org/en/XUL/tabs
; Related .......: _FFTabAdd, _FFTabDuplicate, _FFTabClose, _FFTabCloseAll, _FFTabExists, _FFTabSetSelected
; Example .......: Yes
; ==============================================================================
Func _FFTabGetSelected($sMode = "index")
	Local Const $sFuncName = "_FFTabGetSelected"

	_FFLoadWait()
	Switch StringLower($sMode)
		Case "index"
			Return Int(_FFCmd("gBrowser.tabContainer.selectedIndex", 3000))
		Case "label"
			Return _FFCmd("gBrowser.mTabs[gBrowser.tabContainer.selectedIndex].label", 3000)
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(index|label) $sMode: " & $sMode))
			Return -1
	EndSwitch

EndFunc   ;==>_FFTabGetSelected

; #FUNCTION# ===================================================================
; Name ..........: _FFTableWriteToArray
; Description ...: Writes a HTML-table to an array
; Beschreibung ..: Kopiert eine HTML-Tabelle in ein Array.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFTableWriteToArray([$vTable = 0[, $sMode = "index"[, $sReturnMode = "text"[, $iFilter = 0[, $bTransponse = False]]]]])
; Parameter(s): .: $vTable      - Optional: (Default = 0) : Table selection by $sMode
;                  $sMode       - Optional: (Default = "index") :
;                               | index (0-n)
;                               | name
;                               | id
;                  $sReturnMode - Optional: (Default = "text") :
;                               | text
;                               | html
;                  $iFilter     - Optional: (Default = 0) :
;                               | 0 = No filter
;                               | 1 = Remove all non ASCII
;                               | 2 = Remove all double withespaces
;                               | 4 = Remove all double linefeeds
;                               | 8 = Remove all HTML-tags
;                               | 16 = simple html-tag / entities convertor
;                  $bTransponse - Optional: (Default = False) : Switch Rows and Columns
; Return Value ..: Success      - Array containing contents of the table
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert, Johannes Schirmer
; Date ..........: Fri Sep 18 17:38:05 CEST 2009 @693 /Internet Time/
; Link ..........:
; Related .......: _FFXpath
; Example .......: Yes
; ==============================================================================
Func _FFTableWriteToArray($vTable = 0, $sMode = "index", $sReturnMode = "text", $iFilter = 0, $bTransponse = False)
	Local $sFuncName = "_FFTableWriteToArray"

	Local $sTable, $sRnd

	Local $bTrace = $_FF_COM_TRACE
	$_FF_COM_TRACE = False

	Switch StringLower($sMode)
		Case "index"
			If IsInt($vTable) And $vTable < _FFGetLength("tables") Then
				$sTable = "//table[" & $vTable + 1 & "]"
			Else
				SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(INT) $vTable: " & $vTable))
				Return 0
			EndIf
		Case "name"
			$sTable = "//table[@name='" & $vTable & "']"
		Case "id"
			$sTable = "//table[@id='" & $vTable & "']"
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(index|name|id) $sMode: " & $sMode))
			Return 0
	EndSwitch

	Switch StringLower($sReturnMode)
		Case "text"
			$sReturnMode = "textContent"
		Case "html"
			$sReturnMode = "innerHTML"
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(text|html) $sReturnMode: " & $sReturnMode))
			Return 0
	EndSwitch

	If Not IsInt($iFilter) Then
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(int) $iFilter: " & $iFilter))
		Return 0
	EndIf

	If Not IsBool($bTransponse) Then
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(boolean) $bTransponse: " & $bTransponse))
		Return 0
	EndIf

	Local $iRows = _FFXPath($sTable & "//tr", Default, 10)
	If @error Or $iRows = 0 Then
		SetError(__FFError($sFuncName, $_FF_ERROR_GeneralError, $sTable & " not found."))
		Return 0
	EndIf

	Local $aCols

	If _FFXPath($sTable & "//@colspan", Default, 10) > 0 Or _FFXPath($sTable & "//@rowspan", Default, 10) > 0 Then
		Local $iColsTmp[$iRows + 1]
		Local $iCols = 0, $iSum, $iCol
		For $i = 1 To $iRows
			$iSum = 0
			$aCols = _FFXPath($sTable & "//tr[" & $i & "]//td", "colSpan", 6)
			$iColsTmp[$i - 1] = $aCols[0]
			For $j = 1 To $aCols[0]
				$iSum += $aCols[$j]
			Next
			If $iSum > $iCols Then $iCols = $iSum
		Next

		Local $aTable[$iRows][$iCols]
		Local $iRow = 0
		Local $sCmd, $aTmp

		For $i = 0 To $iRows
			$sRnd = "FF" & Random(100, 999, 1)
			$iCol = 0
			If $iColsTmp[$i] < $iCols Then
				For $j = 0 To $iColsTmp[$i] - 1
					$sCmd = "FFau3.xpath=window.content.document.evaluate('//tr[" & $i + 1 & "]//td[" & $j + 1 & "]',window.content.document,null,XPathResult.FIRST_ORDERED_NODE_TYPE,null).singleNodeValue;FFau3.xpath." & $sReturnMode & "+'" & $sRnd & "'+FFau3.xpath.colSpan"
					$aTmp = StringSplit(_FFCmd($sCmd), $sRnd, 3)
					$aTable[$iRow][$iCol] = __FFFilterString($aTmp[0], $iFilter)
					$iCol += $aTmp[1]
				Next

			Else
				Local $aRows = _FFXPath($sTable & "//tr[" & $i + 1 & "]//td", $sReturnMode, 6, $iFilter)
				If $aRows[0] > 0 Then
					For $j = 1 To $aRows[0]
						$aTable[$i][$j - 1] = $aRows[$j]
					Next
				EndIf
			EndIf

			$iRow += 1
		Next
	Else
		$iCols = _FFXPath($sTable & "//tr[last()]//td", Default, 10)
		Local $aTable[$iRows][$iCols]

		If $iCols > $iRows Then
			For $i = 0 To $iRows - 1
				$aCols = _FFXPath($sTable & "//tr[" & $i + 1 & "]//td", $sReturnMode, 6, $iFilter)
				If $aCols[0] > 0 Then
					For $j = 0 To $iCols - 1
						$aTable[$i][$j] = $aCols[$j + 1]
					Next
				EndIf
			Next
		Else
			For $i = 0 To $iCols - 1
				$aRows = _FFXPath($sTable & "//tr//td[" & $i + 1 & "]", $sReturnMode, 6, $iFilter)
				If $aRows[0] > 0 Then
					For $j = 0 To $iRows - 1
						$aTable[$j][$i] = $aRows[$i + 1]
					Next
				EndIf
			Next
		EndIf
	EndIf

	If $bTransponse Then
		Local $iD1 = UBound($aTable, 1), $iD2 = UBound($aTable, 2), $aTmp[$iD2][$iD1]
		For $i = 0 To $iD2 - 1
			For $j = 0 To $iD1 - 1
				$aTmp[$i][$j] = $aTable[$j][$i]
			Next
		Next
		$aTable = $aTmp
	EndIf

	$_FF_COM_TRACE = $bTrace

	Return $aTable
EndFunc   ;==>_FFTableWriteToArray

; #FUNCTION# ===================================================================
; Name ..........: _FFWindowClose
; Description ...: Closes a browser window
; Beschreibung ..: Schließt ein Browser Fenster.
; Syntax ........: _FFWindowClose([$sSearch = ""[, $sSearchMode = "title"]])
; Parameter(s): .: $sSearch     - Optional: (Default = "") : "" = current window, Substring to search
;                  $sSearchMode - Optional: (Default = "title") :
;                               | "title" window title
;                               | "label" tab label
;                               | "href" current url
; Return Value ..: Success      - 1
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Mar 23 22:33:01 CET 2009 @939 /Internet Time/
; Link ..........: https://developer.mozilla.org/En/DOM/Window, https://developer.mozilla.org/En/NsIWindowMediator
; Related .......: _FFWindowSelect, _FFWindowOpen
; Example .......: Yes
; ==============================================================================
Func _FFWindowClose($sSearch = "", $sSearchMode = "title")
	Local Const $sFuncName = "_FFWindowClose"

	If $sSearch = "" Then
		_FFCmd("window.close();")
		Return @error
	EndIf

	Switch StringLower($sSearchMode)
		Case "title", "label", "href" ; current href
			Local $sRetVal = _FFCmd(StringFormat("FFau3.CloseWin('%s','%s')", $sSearch, $sSearchMode))
			If Not @error And $sRetVal = 1 Then
				Return 1
			Else
				SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch, "$sSearch: " & $sSearch))
			EndIf
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(title|label|href) $sSearchMode: " & $sSearchMode))
	EndSwitch

	Return 0
EndFunc   ;==>_FFWindowClose

; #FUNCTION# ===================================================================
; Name ..........: _FFWindowGetHandle
; Description ...: Returns the window-handle (hwnd) from the current browser-window.
; Beschreibung ..: Liefert das Window-handle (hwnd) des aktuellen Browser Fenster.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFWindowGetHandle()
; Parameter(s): .:
; Return Value ..: Success      - Window handle
;                  Failure      - ""
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Sun Oct 18 21:35:12 CEST 2009
; Link ..........:
; Related .......: _FFWindowClose, _FFWindowSelect, _FFWindowOpen
; Example .......: Yes
; ==============================================================================
Func _FFWindowGetHandle()
	Local Const $sFuncName = "_FFWindowGetHandle"
	; in the future to handle with XPCOM =#=
	Local $bErr = False
	Local $hWindow = ""

	BlockInput(1) ; =#=
	If Not @error Then
		Local $sTitleRnd = "FFAU3" & Random(1000000000, 9999999999, 1)
		_FFCmd("FFau3.tmp=document.title;document.title='" & $sTitleRnd & "'")
		Local $iMatchMode = AutoItSetOption("WinTitleMatchMode", 1)
		Sleep(500)
		$hWindow = WinGetHandle($sTitleRnd)
		If @error Then $bErr = True
		BlockInput(0)
		AutoItSetOption("WinTitleMatchMode", $iMatchMode)
		_FFCmd("document.title=FFau3.tmp")
		ConsoleWrite("_FFWindowGetHandle: " & $hWindow & @CRLF)
	EndIf
	BlockInput(0)

	If $bErr Then SetError(__FFError($sFuncName, $_FF_ERROR_GeneralError))
	Return $hWindow
EndFunc   ;==>_FFWindowGetHandle

; #FUNCTION# ===================================================================
; Name ..........: _FFWindowSelect
; Description ...: Selects the active window for all other functions
; Beschreibung ..: Wählt das aktive Browser Fenster für alle nachfolgenden Befehle.
; Syntax ........: _FFWindowSelect([$sSearch = ""[, $sSearchMode = "title"[, $bActivate = True]]])
; Parameter(s): .: $sSearch     - Optional: (Default = "") : "" = most recent window, Substring to search
;                  $sSearchMode - Optional: (Default = "title") :
;                               | "title" window title
;                               | "label" tab label
;                               | "href" current url
;                  $bActivate   - Optional: (Default = True) :
; Return Value ..: Success      - 1
;                  @EXTENDED    - Window handle
;                  Failure      - 0
;                  @ERROR       -
;                  @EXTENDED    - ""
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Mar 23 22:32:05 CET 2009 @938 /Internet Time/
; Link ..........: https://developer.mozilla.org/En/DOM/Window, https://developer.mozilla.org/En/NsIWindowMediator
; Related .......: _FFWindowOpen, _FFWindowClose, _FFWindowGetHandle
; Example .......: Yes
; ==============================================================================
Func _FFWindowSelect($sSearch = "", $sSearchMode = "title", $bActivate = True)
	Local Const $sFuncName = "_FFWindowSelect"

	Local $hWin = ""

	If $sSearch = "" Then
		Local $sCommand = 'Components.classes["@mozilla.org/appshell/window-mediator;1"].getService(Components.interfaces.nsIWindowMediator).getMostRecentWindow("navigator:browser")'
		_FFCmd("repl.enter(" & $sCommand & ")")
		If Not @error Then
			$hWin = _FFWindowGetHandle()
			If Not @error And $bActivate Then WinActivate($hWin)
			SetExtended($hWin)
			Return 1
		EndIf
		SetExtended($hWin)
		Return 0
	EndIf

	Switch StringLower($sSearchMode)
		Case "title", "label", "href"
			Local $sRetVal = _FFCmd(StringFormat("FFau3.SelectWin('%s','%s','navigator:browser');", $sSearch, $sSearchMode))
			If Not @error And $sRetVal = 1 Then
				$hWin = _FFWindowGetHandle()
				If Not @error And $bActivate Then WinActivate($hWin)
				SetExtended($hWin)
				Return 1
			Else
				SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch, "$sSearch: " & $sSearch), $hWin)
			EndIf
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(title|label|href) $sSearchMode: " & $sSearchMode), $hWin)
	EndSwitch

	SetExtended($hWin)
	Return 0
EndFunc   ;==>_FFWindowSelect

; #FUNCTION# ===================================================================
; Name ..........: _FFWindowOpen
; Description ...: Opens a new browser window
; Beschreibung ..: Öffnet ein neues Browser Fenster.
; Syntax ........: _FFWindowOpen([$sURL = "about:blank"[, $bActivate = True[, $bLoadWait = True]]])
; Parameter(s): .: $sURL        - Optional: (Default = "about:blank") :
;                  $bActivate   - Optional: (Default = True)
;                  $bLoadWait   - Optional: (Default = True) :
; Return Value ..: Success      - 1 and sets
;                  @EXTENDED    - Window handle
;                  Failure      - 0 and sets
;                  @ERROR       -
;                  @EXTENDED    - ""
; Author(s) .....: Thorsten Willert
; Date ..........: Fri Nov 13 18:31:06 CET 2009
; Link ..........: https://developer.mozilla.org/En/DOM/Window, https://developer.mozilla.org/En/NsIWindowMediator
; Related .......: _FFWindowSelect, _FFWindowClose, _FFWindowGetHandle
; Example .......: Yes
; ==============================================================================
Func _FFWindowOpen($sURL = "about:blank", $bActivate = True, $bLoadWait = True)
	Local Const $sFuncName = "_FFWindowOpen"

	If Not __FFCheckURL($sURL) Then
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(URL) $sURL: " & $sURL))
		Return 0
	EndIf

	Local $bTrace = $_FF_COM_TRACE
	$_FF_COM_TRACE = False

	ConsoleWrite("_FFWindowOpen: " & $sURL & @CRLF)

	_FFCmd("window.open('" & $sURL & "');")
	If $bLoadWait Then _FFLoadWait()
	If Not @error Then
		If _FFWindowSelect() Then
			Local $hWin = @extended
			If Not @error And $bActivate Then WinActivate($hWin)
			__FFSendJavaScripts()
			$_FF_COM_TRACE = $bTrace
			SetExtended($hWin)
			Return 1
		EndIf
	EndIf

	$_FF_COM_TRACE = $bTrace
	SetExtended("")
	Return 0
EndFunc   ;==>_FFWindowOpen

; #FUNCTION# ===================================================================
; Name ..........: _FFWriteHTML
; Description ...: Replaces the html of body element
; Beschreibung ..: Ersetzt das HTML des Body Elements.
; Syntax ........: _FFWriteHTML([$sHTML = ""[, $sMode = "body"]])
; Parameter(s): .: $sHTML       - Optional: (Default = "") :
;                  $sMode       - Optional: (Default = "body") :
;                               | body
;                               | html
; Return Value ..: Success      - 1
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Mar 23 22:30:01 CET 2009 @937 /Internet Time/
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _FFWriteHTML($sHTML = "", $sMode = "body")
	Local Const $sFuncName = "_FFWriteHTML"

	Local $sCommand

	Switch StringLower($sMode)
		Case "body"
			$sCommand = "body"
		Case "html"
			$sCommand = "documentElement"
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(body|html) $sMode: " & $sMode))
			Return 0
	EndSwitch

	$sCommand = StringFormat("content.document.%s.innerHTML='" & $sHTML & "';", $sCommand)

	Local $sRetVal = _FFCmd($sCommand)
	If Not @error And $sRetVal <> "_FFCmd_Err" Then Return 1

	Return 0
EndFunc   ;==>_FFWriteHTML

; #FUNCTION# ===================================================================
; Name ..........: _FFXPath
; Description ...: Returns and sets values due to a XPath-query.
; Beschreibung ..: Abfrage und setzen von Werten anhand einer XPath-Abfrage.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFXPath($sQuery[, $sAttribute = ""[, $iReturnType = 9[, $iFilter = 0]]])
; Parameter(s): .: $sQuery      - XPath query
;                  $sAttribute  - Optional: (Default = "") : attribute OR searchstring for SEARCH
;                  $iReturnType - Optional: (Default = 0) :
;                               | 0 = ANY_TYPE
;                               | 1 = NUMBER_TYPE (Returns numberValue)
;                               | 2 = STRING_TYPE (Return stringValue)
;                               | 3 = BOOLEAN_TYPE (Returns booleanValue)
;                               | 4 = UNORDERED_NODE_ITERATOR_TYPE - not implemented
;                               | 5 = ORDERED_NODE_ITERATOR_TYPE - not implemented
;                               | 6 = UNORDERED_NODE_SNAPSHOT_TYPE (Returns arrray)
;                               | 7 = ORDERED_NODE_SNAPSHOT_TYPE (Returns arrray)
;                               | 8 = ANY_UNORDERED_NODE_TYPE (Returns singleNodeValue) (Object in FFau3.xpath)
;                               | 9 = FIRST_ORDERED_NODE_TYPE (Returns singleNodeValue) (Object in FFau3.xpath)
;                               | 10 = COUNT (Returns the number of matches)
;                               | 11 = CONTAINS (stringValue contains substring in $sAttribute) (Returns booleanValue)
;                               | 12 = STARTS-WITH (stringValue start-with string in $sAttribute) (Returns booleanValue)
;                               | 13 = SUBSTRING-AFTER (stringValue substring-after in $sAttribute)
;                               | 14 = SUBSTRING-BEFORE (stringValue substring-bevore in $sAttribute)
;                  $iFilter     - Optional: (Default = 0) :
;                               | 0 = No filter
;                               | 1 = Remove all non ASCII
;                               | 2 = Remove all double withespaces
;                               | 4 = Remove all double linefeeds
;                               | 8 = Remove all HTML-tags
;                               | 16 = simple html-tag / entities convertor
; Return Value ..: Success      - Value $iReturnType 0-5 and 8-10
;                               | Value $iReturnType 8-9 + object FFau3.xpath
;                               | Array $iReturnType 6-7 or if an attribute is set the number of the matches.
;                  Failure      - 0 and sets
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Tue Feb 16 19:43:55 CET 2010
; Link ..........: http://www.w3.org/TR/xpath, https://developer.mozilla.org/en/XPath, https://developer.mozilla.org/En/XPathResult, https://developer.mozilla.org/en/Introduction_to_using_XPath_in_JavaScript
; Related .......: _FFGetValue, _FFSetValue
; Example .......: Yes
; ==============================================================================
Func _FFXPath($sQuery, $sAttribute = "", $iReturnType = 9, $iFilter = 0)
	Local Const $sFuncName = "_FFXPath"

	If $sQuery = "" Then
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "Empty String $sQuery"))
		Return 0
	EndIf
	If Not IsInt($iFilter) Then
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(int) $iFilter: " & $iFilter))
		Return 0
	EndIf
	If $iReturnType < 10 And Not StringInStr($sAttribute, "=") And StringRegExp($sAttribute, "[\W]+") Then
		SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "$sAttribute: " & $sAttribute))
		Return 0
	EndIf

	If $sAttribute = Default Then $sAttribute = "textContent"
	If $iReturnType = Default Then $iReturnType = 0
	$sQuery = StringReplace($sQuery, '"', "'")

	Local $sResult, $sValue

	Switch StringLower($sAttribute)
		Case "textcontent"
			$sAttribute = "textContent"
		Case "innerhtml"
			$sAttribute = "innerHTML"
	EndSwitch

	Switch $iReturnType
		Case 0
			$sResult = "ANY_TYPE"
		Case 1
			$sResult = "NUMBER_TYPE"
			$sValue = ".numberValue"
		Case 2
			$sResult = "STRING_TYPE"
			$sValue = ".stringValue"
		Case 3
			$sResult = "BOOLEAN_TYPE"
			$sValue = ".booleanValue"
			#cs  - not implemented yet
				Case 4
				$sResult = "UNORDERED_NODE_ITERATOR_TYPE"
				Case 5
				$sResult = "ORDERED_NODE_ITERATOR_TYPE"
			#ce
		Case 6
			$sResult = "UNORDERED_NODE_SNAPSHOT_TYPE"
			If $sAttribute <> "" And StringLeft($sAttribute, 1) <> "." Then $sValue &= "." & $sAttribute
		Case 7
			$sResult = "ORDERED_NODE_SNAPSHOT_TYPE"
			If $sAttribute <> "" And StringLeft($sAttribute, 1) <> "." Then $sValue &= "." & $sAttribute
		Case 8
			$sResult = "ANY_UNORDERED_NODE_TYPE"
			$sValue = ".singleNodeValue"
			If $sAttribute <> "" And StringLeft($sAttribute, 1) <> "." Then $sValue &= "." & $sAttribute
		Case 9
			$sResult = "FIRST_ORDERED_NODE_TYPE"
			$sValue = ".singleNodeValue"
			If $sAttribute <> "" And StringLeft($sAttribute, 1) <> "." Then $sValue &= "." & $sAttribute
			; no datatypes
		Case 10
			$sQuery = "count(" & $sQuery & ")"
			$sResult = "NUMBER_TYPE"
			$sValue = ".numberValue"
		Case 11
			$sQuery = "contains(" & $sQuery & ",'" & $sAttribute & "')"
			$sResult = "BOOLEAN_TYPE"
			$sValue = ".booleanValue"
		Case 12
			$sQuery = "starts-with(" & $sQuery & ",'" & $sAttribute & "')"
			$sResult = "BOOLEAN_TYPE"
			$sValue = ".booleanValue"
		Case 13
			$sQuery = "substring-after(" & $sQuery & ",'" & $sAttribute & "')"
			$sResult = "STRING_TYPE"
			$sValue = ".stringValue"
		Case 14
			$sQuery = "substring-before(" & $sQuery & ",'" & $sAttribute & "')"
			$sResult = "STRING_TYPE"
			$sValue = ".stringValue"
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(int) $iReturnType: 0-9 10-14: " & $iReturnType))
			Return 0
	EndSwitch

	Local $sErrVal = "_FFXPath_Error: "
	Local $sCommand

	If $iReturnType < 6 Or $iReturnType >= 8 Then
		$sCommand = 'try{FFau3.xpath=FFau3.WCD.evaluate("' & $sQuery & '",FFau3.WCD,null,XPathResult.' & $sResult & ',null)'
		$sCommand &= $sValue & ";}catch(e){'" & $sErrVal & "'+e;};"

	ElseIf $iReturnType = 6 Or $iReturnType = 7 Then

		Local $sDelimiter = "FF" & Random(100, 999, 1)
		$sCommand = 'try{with(FFau3){'
		$sCommand &= 'this.obj=FFau3.WCD.evaluate("' & $sQuery & '",FFau3.WCD,null,XPathResult.' & $sResult & ',null);'
		$sCommand &= 'this.tmp=this.obj.snapshotLength;'
		If Not StringInStr($sAttribute, "=") Then
			$sCommand &= 'for(var i=0;i<this.obj.snapshotLength;i++){this.tmp+="' & $sDelimiter & '"+this.obj.snapshotItem(i)' & $sValue & ';};'
		Else
			$sCommand &= 'for(var i=0;i<this.obj.snapshotLength;i++){this.obj.snapshotItem(i)' & $sValue & ';};'
		EndIf
		$sCommand &= "this.tmp;}}catch(e){'" & $sErrVal & "' +e;};"
	EndIf

	Local $sRetVal = _FFCmd("FFau3.xpath=null;" & $sCommand)
	If Not StringInStr($sRetVal, $sErrVal) Then
		__FFFilterString($sRetVal, $iFilter)
		If $iFilter > 0 Then __FFFilterString($sRetVal, $iFilter)

		If $iReturnType = 6 Or $iReturnType = 7 Then
			Return StringSplit($sRetVal, $sDelimiter, 3)
		Else
			If $iReturnType < 10 And $sAttribute = "" Then Return "OBJECT|FFau3.xpath"
			Return StringStripWS($sRetVal, 3)
		EndIf
	Else
		If StringInStr($sRetVal, "TypeError") And StringInStr($sRetVal, "is null") Then
			SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch, "singleNodeValue is null $sQuery: " & $sQuery & " - " & $sAttribute))
		ElseIf StringInStr($sRetVal, "NS_ERROR_DOM_INVALID_EXPRESSION_ERR") Then
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidExpression, "NS_ERROR_DOM_INVALID_EXPRESSION_ERR $sQuery: " & $sQuery))
		Else
			SetError(__FFError($sFuncName, $_FF_ERROR_RetValue, $sRetVal))
		EndIf
		Return 0
	EndIf
EndFunc   ;==>_FFXPath

; #FUNCTION# ===================================================================
; Name ..........: _FFObjGet
; Description ...: Returns an object (string to use with the other _FFObj* functions)
; Beschreibung ..: Liefert ein Object zurueck. (String zur Verwendung mit den anderen _FFObj* Funktionen)
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFObjGet($sElement[, $sMode = "id"[, $iIndex = 0]])
; Parameter(s): .: $sElement    -
;                  $sMode       - Optional: (Default = "id") :
;                               | ID
;                               | Name + index
;                               | Class + index
;                               | Tag-name + index
;                  $iIndex      - Optional: (Default = 0) :
; Return Value ..: Success      - String (object to use in the _FFObj* functions)
;                  Failure      - Epmty string
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Fri Oct 16 19:08:02 CEST 2009
; Link ..........:
; Related .......: _FFObj, _FFObjNew, _FFObjDelete, _FFClick
; Example .......: Yes
; ==============================================================================
Func _FFObjGet($sElement, $sMode = "id", $iIndex = 0)
	Local Const $sFuncName = "_FFObjGet"

	Local $sRet

	Switch StringLower($sMode)
		Case "id"
			$sRet = ".getElementById('" & $sElement & "')"
		Case "name"
			$sRet = ".getElementsByName('" & $sElement & "')[" & $iIndex & "]"
		Case "class"
			$sRet = ".getElementsByClassName('" & $sElement & "')[" & $iIndex & "]"
		Case "tag"
			$sRet = ".getElementsByTagName('" & $sElement & "')[" & $iIndex & "]"
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(id|name|class|tag) $sMode: " & $sMode))
			Return ""
	EndSwitch

	Local $RetVal = _FFCmd("window.content.document" & $sRet & "?1:0;")
	If Not @error And $RetVal = 1 Then
		Return "OBJECT|" & $sRet
	Else
		SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch, "$sElement not found: " & $sElement))
		Return ""
	EndIf
EndFunc   ;==>_FFObjGet

; #FUNCTION# ===================================================================
; Name ..........: _FFObj
; Description ...: Returns or sets an value of an attribute.
; Beschreibung ..: Liefert oder setzt den Wert eines Attributes.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFObj($sObject[, $sAttrib = ""[, $vValue = ""]])
; Parameter(s): .: $sObject     - Object or element in FireFox
;                  $sAttrib     - Optional: (Default = "") :
;                  $vValue      - Optional: (Default = "") :
; Return Value ..: Success      -
;                  Failure      - ""
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Fri Oct 16 19:07:54 CEST 2009
; Link ..........:
; Related .......: _FFObjGet, _FFObjNew, _FFObjDelete
; Example .......: Yes
; ==============================================================================
Func _FFObj($sObject, $sAttrib = "", $vValue = "")
	Local Const $sFuncName = "_FFObj"

	If StringLeft($sObject, 7) = "OBJECT|" Then $sObject = StringMid($sObject, 8)
	If StringLeft($sObject, 1) = "." Then $sObject = "window.content.document" & $sObject
	If Not (StringInStr($sObject, "window.content") Or StringInStr($sObject, "FFau3.")) Then $sObject = "FFau3." & $sObject

	If $sAttrib <> "" Then
		If Not StringInStr($sAttrib, "(") Then
			_FFCmd($sObject & ".hasAttribute('" & $sAttrib & "')")
			If @error And $vValue <> "" Then
				;SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "$sAttrib not found: " & $sAttrib))
				_FFCmd($sObject & "setAttributeNode('" & $sAttrib & "')")
			EndIf
		EndIf
		$sObject &= "." & $sAttrib

		If IsInt($vValue) Then
			$sObject &= "=" & $vValue
		ElseIf IsBool($vValue) Then
			$sObject &= "=" & __FFB2S($vValue)
		ElseIf $vValue <> "" Then
			$sObject &= "='" & $vValue & "'"
		EndIf
	EndIf

	Local $RetVal = _FFCmd($sObject)
	If Not @error And $RetVal <> "_FFCmd_Err" Then
		Return $RetVal
	Else
		SetError(__FFError($sFuncName, $_FF_ERROR_NoMatch, "$sObject not found: " & $sObject))
		Return ""
	EndIf
EndFunc   ;==>_FFObj

; #FUNCTION# ===================================================================
; Name ..........: _FFObjNew
; Description ...: Creates a new object in FireFox.
; Beschreibung ..: Erzeugt ein neues Object in FireFox.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFObjNew($sNewObject, $sObject)
; Parameter(s): .: $sNewObject  -
;                  $sObject     -
; Return Value ..: Success      - 1
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Tue Apr 07 15:14:59 CEST 2009
; Link ..........:
; Related .......: _FFObjGet, _FFObj, _FFObjDelete
; Example .......: Yes
; ==============================================================================
Func _FFObjNew($sNewObject, $sObject)
	If StringRegExp($sNewObject, "[\W]+") Then Return 0
	Local $RetVal = _FFCmd("try{FFau3." & $sNewObject & "=" & $sObject & "}catch(e){};FFau3." & $sNewObject & "?1:0;", 3000, False)
	Return SetError(Number($RetVal), @extended, $RetVal)
EndFunc   ;==>_FFObjNew

; #FUNCTION# ===================================================================
; Name ..........: _FFObjDelete
; Description ...: Deletes an object in FireFox.
; Beschreibung ..: Loescht ein Objekt in FireFox.
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFObjDelete($sObject)
; Parameter(s): .: $sObject     -
; Return Value ..: Success      - 1
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Wed Apr 08 14:21:26 CEST 2009 @556 /Internet Time/
; Link ..........:
; Related .......: _FFObjGet, _FFObjNew, _FFObj
; Example .......: Yes
; ==============================================================================
Func _FFObjDelete($sObject)
	Local $sFFau3 = "simulateevent getlinkinfo getlinks searchimagebysize searchlink searchtab closewin selectwin"
	If StringInStr($sFFau3, StringLower($sObject)) Then Return 0
	Return _FFCmd("delete FFau3." & $sObject & ";!FFau3." & $sObject & "?1:0;")
EndFunc   ;==>_FFObjDelete

; #FUNCTION# ===================================================================
; Name ..........: _FFAu3Option
; Description ...: Sets and get options for the FF.au3
; Beschreibung ..: Abfrage und Setzen von Optionen für die FF.au3
; AutoIt Version : V3.3.0.0
; Syntax ........: _FFAu3Option($sOption[, $vValue = ""])
; Parameter(s): .: $sOption     - LoadWaitTimeOut
;                               | SearchMode
;                               | ComTrace
;                               | ErrorMsgBox
;                  $vValue      - Optional: (Default = "") : if noe value is given, the current value is returned
;                               | SearchMode 0 = SubString, 1 = Compare
;                               | LoadWaitTimeOut (int / min. 1000)
;                               | LoadWaitStop (bool) stop loading after LoadWaitTimeOut
;                               | ComTrace (bool)
;                               | ErrorMsgBox (bool)
; Return Value ..: Success      - 1 / current value
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Feb 15 22:14:23 CET 2010
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _FFAu3Option($sOption, $vValue = "")
	Local Const $sFuncName = "_FFOption"

	Switch $sOption
		Case "SearchMode"
			If $vValue == "" Then Return $_FF_SEARCH_MODE
			If Not IsInt($vValue) Then
				SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(int) $vValue: " & $vValue))
				Return 0
			EndIf
			If $vValue > 1 Then $vValue = 0
			$_FF_SEARCH_MODE = $vValue
		Case "LoadWaitTimeOut"
			If $vValue == "" Then Return $_FF_LOADWAIT_TIMEOUT
			If Not IsInt($vValue) Then
				SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(int) $vValue: " & $vValue))
				Return 0
			EndIf
			If $vValue < 1000 Then $vValue = 1000
			$_FF_LOADWAIT_TIMEOUT = $vValue
		Case "LoadWaitStop"
			If $vValue == "" Then Return $_FF_LOADWAIT_STOP
			If Not IsBool($vValue) Then
				SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(bool) $vValue: " & $vValue))
				Return 0
			EndIf
			$_FF_LOADWAIT_STOP = $vValue
		Case "ComTrace"
			If $vValue == "" Then Return $_FF_COM_TRACE
			If Not IsBool($vValue) Then
				SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(bool) $vValue: " & $vValue))
				Return 0
			EndIf
			$_FF_COM_TRACE = $vValue
		Case "ErrorMsgBox"
			If $vValue == "" Then Return $_FF_ERROR_MSGBOX
			If Not IsBool($vValue) Then
				SetError(__FFError($sFuncName, $_FF_ERROR_InvalidDataType, "(bool) $vValue: " & $vValue))
				Return 0
			EndIf
			$_FF_ERROR_MSGBOX = $vValue
		Case Else
			SetError(__FFError($sFuncName, $_FF_ERROR_InvalidValue, "(SearchMode|LoadWaitTimeOut|ComTrace|ErrorMsgBox) $sOption" & $sOption))
	EndSwitch

	Return 1
EndFunc   ;==>_FFAu3Option

;===============================================================================
;
; Internal Functions with names starting with two underscores will not be documented
; as user functions
;
; #INTERNAL_USE_ONLY# ==========================================================
; Name ..........: __FFRecv
; Description ...:
; AutoIt Version : V3.3.0.0
; Syntax ........: __FFRecv([$iTimeOut = 10000])
; Parameter(s): .: $iTimeOut    - Optional: (Default = 10000) : TimeOut for __FFWaitForRepl
; Return Value ..: Success      - Return value from MozRepl
;                  Failure      - ""
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Mar 23 22:27:52 CET 2009 @936 /Internet Time/
; ==============================================================================
Func __FFRecv($iTimeOut = 30000)
	Local $sRet = __FFWaitForRepl($iTimeOut)
	Local $iErr = @error
	$sRet = StringStripWS($sRet, 3)

	; Removing leading and trailing "
	If StringLeft($sRet, 1) = '"' Then $sRet = StringTrimLeft($sRet, 1)
	If StringRight($sRet, 1) = '"' Then $sRet = StringTrimRight($sRet, 1)

	; String to bool
	If $sRet = "true" Then
		$sRet = 1
	ElseIf $sRet = "false" Then
		$sRet = 0
	EndIf

	If $_FF_COM_TRACE Then ConsoleWrite("__FFRecv: " & $sRet & @CRLF)

	SetError($iErr)
	Return $sRet
EndFunc   ;==>__FFRecv

; #INTERNAL_USE_ONLY# ==========================================================
; Name ..........: __FFSend
; Description ...:
; AutoIt Version : V3.3.0.0
; Syntax ........: __FFSend($sCommand)
; Parameter(s): .: $sCommand    - String to send
; Return Value ..: Success      - 1
;                  Failure      - 0
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Mar 23 22:26:09 CET 2009 @934 /Internet Time/
; ==============================================================================
Func __FFSend($sCommand)
	Local Const $sFuncName = "__FFSend"

	If Not __FFIsSocket($_FF_GLOBAL_SOCKET) Then
		SetError(__FFError($sFuncName, $_FF_ERROR_SocketError))
		Return 0
	EndIf

	TCPSend($_FF_GLOBAL_SOCKET, $sCommand & @CRLF)
	If Not @error Then
		If $_FF_COM_TRACE Then ConsoleWrite("__FFSend: " & $sCommand & @CRLF)
		Return 1
	Else
		SetError(__FFError($sFuncName, $_FF_ERROR_SendRecv, "TCPSend: " & $sCommand))
		Return 0
	EndIf
EndFunc   ;==>__FFSend

; #INTERNAL_USE_ONLY# ==========================================================
; Name ..........: __FFChromeSelect
; Description ...: Some shortcuts for chrome
; AutoIt Version : V3.3.0.0
; Syntax ........: __FFChromeSelect($sOpt)
; Parameter(s): .: $sOpt        - Shortcut
; Return Value ..: Success      - crome: + ....xul
;                  Failure      - crome: + $sOpt
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Sat Feb 14 12:33:16 CET 2009 @523 /Internet Time/
; ==============================================================================
Func __FFChromeSelect($sOpt)
	Local $sXUL

	Switch StringLower($sOpt)
		Case "bookmarks"
			$sXUL = "//browser/content/bookmarks/bookmarksPanel.xul"
		Case "history"
			$sXUL = "//browser/content/history/history-panel.xul"
		Case "extensions"
			$sXUL = "//mozapps/content/extensions/extensions.xul?type=extensions"
		Case "advanced-scripts"
			$sXUL = "//browser/content/preferences/advanced-scripts.xul"
		Case "changeaction"
			$sXUL = "//browser/content/preferences/changeaction.xul"
		Case "prefs"
			$sXUL = "//browser/content/preferences/preferences.xul"
		Case "colors"
			$sXUL = "//browser/content/preferences/colors.xul"
		Case "connection"
			$sXUL = "//browser/content/preferences/connection.xul"
		Case "cookies"
			$sXUL = "//browser/content/preferences/cookies.xul"
		Case "downloads"
			$sXUL = "//mozapps/content/downloads/downloads.xul"
		Case "downloadactions"
			$sXUL = "//browser/content/preferences/downloadactions.xul"
		Case "fonts"
			$sXUL = "//browser/content/preferences/fonts.xul"
		Case "languages"
			$sXUL = "//browser/content/preferences/languages.xul"
		Case "permissions"
			$sXUL = "//browser/content/preferences/permissions.xul"
		Case "sanitize"
			$sXUL = "//browser/content/preferences/sanitize.xul"
		Case "fireftp"
			$sXUL = "//fireftp/content/fireftp.xul"
		Case Else
			$sXUL = $sOpt
	EndSwitch

	Return "chrome:" & $sXUL
EndFunc   ;==>__FFChromeSelect

; #FUNCTION# ===================================================================
; Name ..........: __FFFilterString
; Description ...: Filter for strings
; AutoIt Version : V3.3.0.0
; Syntax ........: __FFFilterString(ByRef $sString[, $iMode = 1])
; Parameter(s): .: $sString     - String to filter
;                  $iMode       - Optional: (Default = 0) : removes nothing
;                               - 0 = no filter
;                               - 1 = removes non ascii characters
;                               - 2 = removes all double whitespaces
;                               - 4 = removes all double linefeeds
;                               - 8 = removes all html-tags
;                               - 16 = simple html-tag / entities convertor
; Return Value ..: Success      - Filterd String
;                  Failure      - Input String
; Author(s) .....: Thorsten Willert
; Date ..........: Wed Apr 08 20:58:17 CEST 2009 @832 /Internet Time/
; ==============================================================================
Func __FFFilterString(ByRef $sString, $iMode = 0)
	If $iMode = 0 Then Return $sString
	;16 simple HTML tag converter
	If $iMode >= 16 And $iMode < 32 Then
		Local $sRepl[2][6] = [["<br>", "&nbsp;", "&gt;", "&lt;", "&amp;", "&quot;"], _
				[@CRLF, " ", ">", "<", "&", '"']]
		$sString = StringRegExpReplace($sString, "<p.*>", @CRLF & @CRLF)
		For $i = 0 To UBound($sRepl, 2) - 1
			$sString = StringReplace($sString, $sRepl[0][$i], $sRepl[1][$i])
		Next
		$iMode -= 16
	EndIf
	;8 Tag filter
	If $iMode >= 8 And $iMode < 16 Then
		$sString = StringRegExpReplace($sString, "<[^>]*>", "")
		$iMode -= 8
	EndIf
	; 4 remove all double cr, lf
	If $iMode >= 4 And $iMode < 8 Then
		$sString = StringRegExpReplace($sString, "([ \t]*[\n\r]+[ \t]*)", @CRLF)
		$sString = StringRegExpReplace($sString, "[\n\r]+", @CRLF)
		$iMode -= 4
	EndIf
	; 2 remove all double withespaces
	If $iMode = 2 Or $iMode = 3 Then
		$sString = StringRegExpReplace($sString, "[[:blank:]]+", " ")
		$sString = StringRegExpReplace($sString, "\n[[:blank:]]+", @CRLF)
		$sString = StringRegExpReplace($sString, "[[:blank:]]+\n", "")
		$iMode -= 2
	EndIf
	; 1 remove all non ASCII
	If $iMode = 1 Then
		$sString = StringRegExpReplace($sString, "[^\x00-\x7F]", " ")
	EndIf

	Return $sString
EndFunc   ;==>__FFFilterString

; #INTERNAL_USE_ONLY# ==========================================================
; Name ..........: __FFError
; Description ...: Writes Error to the console and show message-boxes if the script is compiled
; AutoIt Version : V3.3.0.0
; Syntax ........: __FFError($sWhere, ByRef $i_FF_ERROR[, $sMessage = ""])
; Parameter(s): .: $i_FF_ERROR  - Error Const
;                  $sMessage    - Optional: (Default = "") : Additional Information
; Return Value ..: Success      - Error Const from $i_FF_ERROR
; Author(s) .....: Thorsten Willert
; Date ..........: Sat Jul 18 11:52:36 CEST 2009
; ==============================================================================
Func __FFError($sWhere, Const ByRef $i_FF_ERROR, $sMessage = "")
	Local $sOut, $sMsg
	Sleep(200)

	Switch $i_FF_ERROR
		Case $_FF_ERROR_Success
			Return $_FF_ERROR_Success
		Case $_FF_ERROR_GeneralError
			$sOut = "General Error"
		Case $_FF_ERROR_SocketError
			$sOut = "Socket Error"
		Case $_FF_ERROR_InvalidDataType
			$sOut = "Invalid data type"
		Case $_FF_ERROR_InvalidValue
			$sOut = "Invalid value"
		Case $_FF_ERROR_Timeout
			$sOut = "Timeout"
		Case $_FF_ERROR_NoMatch
			$sOut = "No match"
		Case $_FF_ERROR_RetValue
			$sOut = "Error return value"
		Case $_FF_ERROR_SendRecv
			$sOut = "Error TCPSend / TCPRecv"
		Case $_FF_ERROR_ReplException
			$sOut = "MozRepl Exception"
		Case $_FF_ERROR_SendRecv
			$sOut = "Error TCPSend / TCPRecv"
		Case $_FF_ERROR_InvalidExpression
			$sOut = "Invalid Expression"
	EndSwitch
	If $sMessage = "" Then
		$sMsg = $sWhere & " ==> " & $sOut & @CRLF
		ConsoleWriteError($sMsg)
		If @Compiled Then
			If $_FF_ERROR_MSGBOX And $i_FF_ERROR < 6 Then MsgBox(16, "FF.au3 Error:", $sMsg)
			DllCall("kernel32.dll", "none", "OutputDebugString", "str", $sMsg)
		EndIf

	Else
		$sMsg = $sWhere & " ==> " & $sOut & ": " & $sMessage & @CRLF
		ConsoleWriteError($sMsg)
		If @Compiled Then
			If $_FF_ERROR_MSGBOX And $i_FF_ERROR < 6 Then MsgBox(16, "FF.au3 Error:", $sMsg)
			DllCall("kernel32.dll", "none", "OutputDebugString", "str", $sMsg)
		EndIf
	EndIf

	Return $i_FF_ERROR
EndFunc   ;==>__FFError

; #INTERNAL_USE_ONLY# ==========================================================
; Name ..........: __FFIsIP
; Description ...: IP4 check
; AutoIt Version : V3.3.0.0
; Syntax ........: __FFIsIP(ByRef $IP[, $bSubstring = False])
; Parameter(s): .: $IP          - IP4 Address
; Return Value ..: Success      - 1
;                  Failure      - 0
; Author(s) .....: Thorsten Willert
; Date ..........: Sun Feb 15 00:21:41 CET 2009
; ==============================================================================
Func __FFIsIP(ByRef $IP, $bSubstring = False)
	If $bSubstring Then
		Return StringRegExp(StringStripWS($IP, 3), "(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9])\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9]|0)\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9]|0)\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[0-9])")
	Else
		Return StringRegExp(StringStripWS($IP, 3), "^(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9])\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9]|0)\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9]|0)\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[0-9])$")
	EndIf
EndFunc   ;==>__FFIsIP

; #INTERNAL_USE_ONLY# ==========================================================
; Name ..........: __FFIsPort
; Description ...: Port check
; AutoIt Version : V3.3.0.0
; Syntax ........: __FFIsPort(ByRef $iPort)
; Parameter(s): .: $iPort       - Port number 1024-65535
; Return Value ..: Success      - 1
;                  Failure      - 0
; Author(s) .....: Thorsten Willert
; Date ..........: 28. August 2007
; ==============================================================================
Func __FFIsPort(ByRef $iPort)
	Return (IsInt($iPort) And $iPort >= 1024 And $iPort <= 65535)
EndFunc   ;==>__FFIsPort

; #INTERNAL_USE_ONLY# ==========================================================
; Name ..........: __FFIsSocket
; Description ...: Socket check
; AutoIt Version : V3.3.0.0
; Syntax ........: __FFIsSocket(ByRef $Socket)
; Parameter(s): .: $Socket      -
; Return Value ..: Success      - 1
;                  Failure      - 0
; Author(s) .....: Thorsten Willert
; Date ..........: Tue Apr 07 11:08:05 CEST 2009 @422 /Internet Time/
; ==============================================================================
Func __FFIsSocket(ByRef $Socket)
	Return ($Socket > 0 And IsInt($Socket))
EndFunc   ;==>__FFIsSocket

; #INTERNAL_USE_ONLY# ==========================================================
; Name ..........: __FFIsURL
; Description ...: URL check
; AutoIt Version : V3.3.0.0
; Syntax ........: __FFIsURL(ByRef $URL)
; Parameter(s): .: $URL         - http/https/ftp/about/chrome/file
; Return Value ..: Success      - 1
;                  Failure      - 0
; Author(s) .....: Thorsten Willert
; Date ..........: Tue Jul 07 11:32:45 CEST 2009
; Link ..........: http://geekswithblogs.net/casualjim/archive/2005/12/01/61722.aspx
; ==============================================================================
Func __FFIsURL(ByRef $URL)
	Return (StringRegExp($URL, '^^((ht|f)tp(s?)\:\/\/|~/|/)([\w]+:\w+@)?([a-zA-Z]{1}([\w\-]+\.)+([\w]{2,5}))(:[\d]{1,5})?((/?\w+/)+|/?)(\w+\.[\w]{3,4})?((\?\w+=\w+)?(&\w+=\w+)*)?') Or _
			StringRegExp($URL, '^^((ht|f)tp(s?)\:\/\/|~/|/)(\d{1,3}\.){3}\d{1,3}(:[\d]{1,5})?(/.*)?$') Or _
			StringLeft($URL, 6) = "about:" Or _
			StringLeft($URL, 7) = "chrome:" Or _
			StringLeft($URL, 10) = "localhost:" Or _
			StringLeft($URL, 8) = "file:///")
EndFunc   ;==>__FFIsURL

Func __FFCheckURL(ByRef $URL)
	If Not __FFIsURL($URL) Then
		If Not StringRegExp($URL, '^((ht|f)tp(s?)\:\/\/|~/|/)') Then $URL = "http://" & $URL
	EndIf

	Return __FFIsURL($URL)
EndFunc   ;==>__FFCheckURL

; #INTERNAL_USE_ONLY# ==========================================================
; Name ..........: __FFB2S
; Description ...: Bool to string
; AutoIt Version : V3.3.0.0
; Syntax ........: __FFB2S($bBool)
; Parameter(s): .: $bBool       - true / false
; Return Value ..: Success      - String "true" or "false"
;                  Failure      - ""
; Author(s) .....: Thorsten Willert
; Date ..........: Sun Feb 15 00:26:46 CET 2009
; ==============================================================================
Func __FFB2S($bBool)
	If $bBool Then
		Return "true"
	Else
		Return "false"
	EndIf
EndFunc   ;==>__FFB2S

; #INTERNAL_USE_ONLY# ==========================================================
; Name ..........: __FFValue2JavaScript
; Description ...:
; AutoIt Version : V3.3.0.0
; Syntax ........: __FFValue2JavaScript(ByRef $sValue)
; Parameter(s): .: $sValue      -
; Return Value ..: Success      -
;                  Failure      -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Apr 13 15:42:21 CEST 2009
; ==============================================================================
Func __FFValue2JavaScript(ByRef $sValue)
	If $sValue = "" Then Return ""
	Local $sTmp = "", $c

	$sValue = StringRegExpReplace($sValue, '([\\"''])', '\\$1')
	$sValue = StringReplace($sValue, @CRLF, "\n")
	$sValue = StringReplace($sValue, @CR, "\r")
	$sValue = StringReplace($sValue, @TAB, "\t")

	If StringRegExp($sValue, "[^\x00-\x7F]", 0) Then
		For $i = 0 To StringLen($sValue)
			$c = StringMid($sValue, $i, 1)
			If StringRegExp($c, "[^\x00-\x7F]", 0) Then $c = StringFormat("\\u%0000s", Hex(AscW($c), 4))
			$sTmp &= $c
		Next
		$sValue = $sTmp
	EndIf

	Return $sValue

EndFunc   ;==>__FFValue2JavaScript

; #INTERNAL_USE_ONLY# ==========================================================
; Name ..........: __FFSendJavaScripts
; Description ...: JavaScript functions
; AutoIt Version : V3.3.0.0
; Syntax ........: __FFSendJavaScripts([$bRefresh = true[, $bTrace = False]])
; Parameter(s): .: $bRefresh    - Optional: (Default = false) : Send again to FF
; Return Value ..: Success      - 1
;                  Failure      - 0
; Author(s) .....: Thorsten Willert
; Date ..........: Wed Dec 02 17:50:27 CET 2009
; Link ..........: http://developer.mozilla.org/en/docs/DOM:event.initMouseEvent
; ==============================================================================
Func __FFSendJavaScripts($bRefresh = True, $bTrace = False)

	Local $bTrc = $_FF_COM_TRACE
	$_FF_COM_TRACE = $bTrace

	If _FFCmd("repl.search(/FFau3/i)", 3000) = "FFau3" And Not $bRefresh Then
		$_FF_COM_TRACE = $bTrace
		Return 1
	EndIf
	If $bRefresh Then _FFCmd('try{delete FFau3;}catch(e){e};')

	ConsoleWrite("__FFSendJavaScripts: Sending functions to FireFox .")

	_FFCmd('FFau3=new Object();FFau3.tmp="";FFau3.obj=null;FFau3.WCD=window.content.top.document;')
	ConsoleWrite(".")

	; sEvent = MouseEvents | KeyboardEvent | Event
	; sType:
	; 	MouseEvents: click
	; 	KeyboardEvent: keyup | keydown | keypress
	; 	Event: change | select | focus | blur | resize
	; args: oObject, sEvent, sType[, Key[, X[ ,Y]]]
	; key = dec-ASCII-code
	; X = clientX
	; Y = clientY
	_FFCmd('FFau3.simulateEvent=function simulateEvent(a,b,c){try{var d=document.createEvent(b);switch(b){case"MouseEvents":d.initMouseEvent(c,true,true,window,0,0,0,simulateEvent.arguments[4],simulateEvent.arguments[5],false,false,false,false,0,null);break;case"KeyboardEvent":d.initKeyEvent(c,true,true,null,false,false,false,false,simulateEvent.arguments[3],0);break;case"Event":d.initEvent(c,true,true);break}a.dispatchEvent(d);return 1}catch(e){return-3}return 0};')
	#cs
	_FFCmd('FFau3.simulateEvent=function simulateEvent(oObject,sEvent,sType){' & _
			'try{ var evt=document.createEvent(sEvent);' & _
			'switch(sEvent){' & _
			' case "MouseEvents":' & _
			'  evt.initMouseEvent(sType,true,true,window,0,0,0,simulateEvent.arguments[4],simulateEvent.arguments[5],false,false,false,false,0,null);break;' & _
			' case "KeyboardEvent":' & _
			'  evt.initKeyEvent(sType,true,true,null,false,false,false,false,simulateEvent.arguments[3],0);break;' & _
			' case "Event":' & _
			'  evt.initEvent(sType,true,true);break; }' & _
			'oObject.dispatchEvent(evt);' & _
			'return 1;' & _
			'}catch(e){return -3;}return 0;};')
	#ce
	ConsoleWrite(".")

	; GetLinkInfo(): JavaScript to get all informations about a link
	_FFCmd('FFau3.GetLinkInfo=function GetLinkInfo(i,a){var b="";with(FFau3.WCD){try{b=links[i].href+a;b+=links[i].hash+a;b+=links[i].search+a;b+=links[i].name+a;b+=links[i].id+a;b+=links[i].text+a;b+=links[i].innerHTML+a;b+=(links[i].hasAttribute("target")?links[i].target+a:a);b+=links[i].protocol+a;b+=links[i].port;return b.replace(/[\s]+/g," ")}catch(e){return-3}}};')
	#cs
	_FFCmd('FFau3.GetLinkInfo = function GetLinkInfo(i,sLimiter)' & _
			'{ var info=""; with(FFau3.WCD) {' & _
			'try {' & _
			'info = links[i].href + sLimiter;' & _
			'info += links[i].hash + sLimiter;' & _
			'info += links[i].search + sLimiter;' & _
			'info += links[i].name + sLimiter;' & _
			'info += links[i].id + sLimiter;' & _
			'info += links[i].text + sLimiter;' & _
			'info += links[i].innerHTML + sLimiter;' & _
			'info += (links[i].hasAttribute("target") ? links[i].target + sLimiter : sLimiter);' & _
			'info += links[i].protocol + sLimiter;' & _
			'info += links[i].port;' & _
			'return info.replace(/[\s]+/g," ");' & _
			'}catch(e){return -3;}' & _
			'}};')
	#ce
	ConsoleWrite(".")

	; GetLinks(): JavaScript to get an information of all links
	_FFCmd('FFau3.GetLinks=function GetLinks(a){var b="";with(FFau3.WCD){try{for(var i=0;i<links.length;i++){b+=this.GetLinkInfo(i,a)+"\n"}return b}catch(e){return-3}}};')
	#cs
	_FFCmd('FFau3.GetLinks = function GetLinks(sLimiter)' & _
			'{ var res=""; with(FFau3.WCD) {try{' & _
			'for (var i=0;i<links.length;i++){res += this.GetLinkInfo(i,sLimiter) + "\n";}' & _
			'return res; } catch(e) {return -3;}}};')
	#ce
	ConsoleWrite(".")

	; SearchImageBySize(): JavaScript
	_FFCmd('FFau3.SearchImageBySize=function SearchImageBySize(a,b,c,d){var e="";with(FFau3.WCD){for(var i=0;i<images.length;i++){var f=images[i].width;var g=images[i].height;switch(c){case"eq":if((f<=b+b/100*d&&g<=a+a/100*d&&f>=b-b/100*d&&g>=a-a/100*d))e+=" "+i;break;case"lt":if((f<=b&&g<=a)&&(f>=b-b/100*d&&g>=a-a/100*d))e+=" "+i;break;case"gt":if((f>=b&&g>=a)&&(f<=b+b/100*d&&g<=a+a/100*d))e+=" "+i;break;default:return-1}}return e}};')
	#cs
	_FFCmd('FFau3.SearchImageBySize = function SearchImageBySize(iH, iW, sMode, iP){' & _
			'var sRet=""; with(FFau3.WCD) {' & _
			'for (var i=0;i<images.length;i++){' & _
			'var iWidth=images[i].width;' & _
			'var iHeight=images[i].height;' & _
			'switch (sMode) {' & _
			'case "eq": if ((iWidth<=iW+iW/100*iP && iHeight<=iH+iH/100*iP && iWidth>=iW-iW/100*iP && iHeight>=iH-iH/100*iP)) sRet += " " + i; break;' & _
			'case "lt": if ((iWidth<=iW && iHeight<=iH) && (iWidth>=iW-iW/100*iP && iHeight>=iH-iH/100*iP)) sRet += " " + i; break;' & _
			'case "gt": if ((iWidth>=iW && iHeight>=iH) && (iWidth<=iW+iW/100*iP && iHeight<=iH+iH/100*iP)) sRet += " " + i; break;' & _
			'default: return -1}' & _
			'} return sRet; } };')
	#ce
	ConsoleWrite(".")

	_FFCmd('FFau3.findPos=function findPos(a){try{var b=c=0;if(a.offsetParent){do{b+=a.offsetLeft;c+=a.offsetTop}while(a=a.offsetParent);return b+" "+c}}catch(e){return-3}}')
	#cs
	_FFCmd("FFau3.findPos = function findPos(obj){try{" & _
			"var curleft=c=0;" & _
			"if (obj.offsetParent) {" & _
			"do {curleft+=obj.offsetLeft;" & _
			"  c+=obj.offsetTop;" & _
			"} while (obj=obj.offsetParent);" & _
			"return curleft+' '+c;" & _
			"}}catch(e){return -3}}")
	#ce
	ConsoleWrite(".")

	; SearchTab(); JavaScript to search a tab by label
	_FFCmd('FFau3.SearchTab=function SearchTab(a){try{for(var i=0;i<gBrowser.mTabs.length;i++){if(RegExp(a).test(gBrowser.mTabs[i].label))return i};return-1}catch(e){return-3}};')
	#cs
	_FFCmd('FFau3.SearchTab = function SearchTab(RegEx)' & _
			'{try{' & _
			'for (var i=0;i<gBrowser.mTabs.length;i++){' & _
			'if (RegExp(RegEx).test(gBrowser.mTabs[i].label)) return i;};' & _
			'return -1;}catch(e){return -3;}};')
	#ce
	ConsoleWrite(".")

	; CloseWin()
	_FFCmd('FFau3.CloseWin=function CloseWin(a,b){var c;var d=Components.classes["@mozilla.org/appshell/window-mediator;1"].getService(Components.interfaces.nsIWindowMediator).getEnumerator("navigator:browser");while(d.hasMoreElements()){c=d.getNext();switch(b){case"title":if(c.title.indexOf(a)!=-1){c.close();return 1};break;case"label":if(FFau3.SearchTab(a)>-1){c.close();return 1};break;case"href":if(c.content.document.location.href.indexOf(a)!=-1){c.close();return 1}break;default:return-1}}return-1};')
	#cs
	_FFCmd('FFau3.CloseWin = function CloseWin(sSearch,sMode){' & _
			'var win;var enum = Components.classes["@mozilla.org/appshell/window-mediator;1"].getService(Components.interfaces.nsIWindowMediator).getEnumerator("navigator:browser");' & _
			'while(enum.hasMoreElements() ){' & _
			'win=enum.getNext();' & _
			'switch(sMode){' & _
			'case "title": if (win.title.indexOf(sSearch) != -1){win.close();return 1;};break;' & _
			'case "label": if (FFau3.SearchTab(sSearch)>-1){win.close();return 1;};break;' & _
			'case "href": if (win.content.document.location.href.indexOf(sSearch) != -1){win.close();return 1;}break;' & _
			'default: return -1;' & _
			'}}return -1;};')
	#ce
	ConsoleWrite(".")

	; SelectWin()
	_FFCmd('FFau3.SelectWin=function SelectWin(a,b,c){var d;var e=Components.classes["@mozilla.org/appshell/window-mediator;1"].getService(Components.interfaces.nsIWindowMediator).getEnumerator(c);while(e.hasMoreElements()){d=e.getNext();switch(b){case"title":if(d.title.indexOf(a)!=-1){repl.enter(d);return 1};break;case"label":if(FFau3.SearchTab(a)>-1){repl.enter(d);return 1};break;case"href":if(d.content.document.location.href.indexOf(a)!=-1){repl.enter(d);return 1};break;default:return-1}}return-1};')
	#cs
	_FFCmd('FFau3.SelectWin = function SelectWin(sSearch,sMode,sType){' & _
			'var win;var enum = Components.classes["@mozilla.org/appshell/window-mediator;1"].getService(Components.interfaces.nsIWindowMediator).getEnumerator(sType);' & _
			'while(enum.hasMoreElements() ){' & _
			'win=enum.getNext();' & _
			'switch(sMode){' & _
			'case "title": if (win.title.indexOf(sSearch) != -1){repl.enter(win);return 1;};break;' & _
			'case "label": if (FFau3.SearchTab(sSearch)>-1){repl.enter(win);return 1;};break;' & _
			'case "href": if (win.content.document.location.href.indexOf(sSearch) != -1){repl.enter(win);return 1;};break;' & _
			'default: return -1};' & _
			'}return -1;};')
	#ce
	ConsoleWrite(". done" & @CRLF)

	$_FF_COM_TRACE = $bTrc
	Return 1

EndFunc   ;==>__FFSendJavaScripts

; #INTERNAL_USE_ONLY# ==========================================================
; Name ..........: __FFWaitForRepl
; Description ...:
; AutoIt Version : V3.3.0.0
; Syntax ........: __FFWaitForRepl($iTimeOut)
; Parameter(s): .: $iTimeOut
; Return Value ..: Success      - Return value from MozRepl
;                  Failure      - ""
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Fri Dec 04 11:48:08 CET 2009
; ==============================================================================
Func __FFWaitForRepl($iTimeOut)
	Local Const $sFuncName = @CRLF & "__FFWaitForRepl"

	Local $recv, $sRet = "", $TimeOutTimer = TimerInit()
	If $iTimeOut < 200 Then $iTimeOut = 200

	While TimerDiff($TimeOutTimer) < $iTimeOut

		; connection delay
		Sleep($_FF_CON_DELAY)

		$recv = TCPRecv($_FF_GLOBAL_SOCKET, 4096)
		;ConsoleWrite($recv & @CRLF)
		; TCP error
		If @error Then
			SetError(__FFError($sFuncName, $_FF_ERROR_SendRecv, "TCPRecv :" & @error))
			Return ""
		EndIf
		$sRet &= $recv

		; error from MozRepl
		If StringRegExp($recv, "!!!(.*?)(TypeError|Exception|ReferenceError):?") Then
			$recv = StringLeft($recv, StringInStr($recv, "location") - 1)
			Sleep(200)
			SetError(__FFError($sFuncName, $_FF_ERROR_ReplException, StringStripWS($recv, 3)))
			__FFSend(";") ; MozRepl-Reset
			Sleep(200)
			Return ""
		ElseIf StringInStr($recv, "....>") Then
			__FFSend(";") ; MozRepl-Reset
			Sleep(200)
			SetError(__FFError($sFuncName, $_FF_ERROR_RetValue, "MozRepl ....>"))
			Return ""
		ElseIf StringInStr($recv, "beginning of the line to force evaluation") Then
			Sleep(500) ; first connection delay
		EndIf

		; multiple connections to MozRepl (e.g repl2?> ...)
		If StringRegExp($recv, "repl[\d]?>") Then Return StringRegExpReplace($sRet, "repl[\d]?>", "")
	WEnd

	; Timeout
	SetError(__FFError($sFuncName, $_FF_ERROR_Timeout, Round(TimerDiff($TimeOutTimer)) & "ms > " & $iTimeOut & "ms $iTimeOut"))
	Return ""

EndFunc   ;==>__FFWaitForRepl

; #INTERNAL_USE_ONLY# ==========================================================
; Name ..........: __FFStartProcess
; Description ...: Starts the firefox.exe
; AutoIt Version : V3.3.0.0
; Syntax ........: __FFStartProcess([$sURL = "about:blank"[, $bNewWin = False[, $sProfile = "default"[, $bNoRemote = False[, $bHide = False[, $iPort = 4242[, $iTimeOut = 30000]]]]]]])
; Parameter(s): .: $sProfile    - Optional: (Default = "default") :
;                  $bNewWin     - Optional: (Default = false) :
;                  $bNoRemote   - Optional: (Default = false) :
;                  $bHide       - Optional: (Default = False) :
;                  $iPort       - Optional: (Default = 4242) :
;                  $iTimeOut    - Optional: (Default = 30000) : min. 2000ms
; Return Value ..: Success      - 1
;                  Failure      - 0
;                  @ERROR       -
;                  @EXTENDED    - PID from the firefox.exe
; Author(s) .....: Thorsten Willert
; Date ..........: Wed Nov 04 16:01:59 CET 2009
; ==============================================================================
Func __FFStartProcess($sURL = "about:blank", $bNewWin = False, $sProfile = "default", $bNoRemote = False, $bHide = False, $iPort = 4242, $iTimeOut = 30000)
	Local Const $sFuncName = "__FFStartProcess"

	Local $PID = -1
	Local $sNoRemote = "", $sNewWin
	Local $sProcName = $_FF_PROC_NAME

	If $sProfile = "default" Then
		$sProfile = ''
	Else
		$sProfile = ' -P "' & $sProfile & '"'
	EndIf
	If $bNoRemote Then $sNoRemote = "-no-remote"
	If $bNewWin Then $sNewWin = "-new-window"
	$sURL = '"' & $sURL & '"'
	If $iTimeOut < 2000 Then $iTimeOut = 2000

	Local $sHKLM = "HKEY_LOCAL_MACHINE\SOFTWARE\Mozilla\Mozilla Firefox"
	Local $sFFExe = RegRead($sHKLM & "\" & RegRead($sHKLM, "CurrentVersion") & "\Main", "PathToExe")
	If @error Then
		SetError(__FFError($sFuncName, $_FF_ERROR_GeneralError, "Error reading registry entry for FireFox." & @CRLF & _
				"HKEY_LOCAL_MACHINE\SOFTWARE\Mozilla\Mozilla Firefox\*CurrentVersion*\Main\PathToExe" & @CRLF & _
				"Error from RegRead: " & @error))
		Return 0
	EndIf

	Local $sCommand = StringFormat('"%s" %s %s %s -repl %i %s', $sFFExe, $sNewWin, $sURL, $sNoRemote, $iPort, $sProfile)
	$PID = Run($sCommand)
	If $bHide Then BlockInput(1)

	Local $iTimeOutTimer = TimerInit()
	While 1
		Sleep(2000)
		If ProcessExists($sProcName) Then ExitLoop
		If (TimerDiff($iTimeOutTimer) > $iTimeOut) Then
			SetError(__FFError($sFuncName, $_FF_ERROR_Timeout, "Browser process not exists: " & $sProcName))
			BlockInput(0)
			Return 0
		EndIf
	WEnd

	If $bHide Then
		Local $WINTITLE_MATCH_MODE = AutoItSetOption("WinTitleMatchMode", 4)
		WinWaitActive("[CLASS:MozillaUIWindowClass]")
		Sleep(500)
		WinSetState("[CLASS:MozillaUIWindowClass]", "", @SW_MINIMIZE)
		BlockInput(0)
		AutoItSetOption("WinTitleMatchMode", $WINTITLE_MATCH_MODE)
	Else
		Sleep(1000)
	EndIf

	If $_FF_COM_TRACE Then ConsoleWrite('__FFStartProcess: "' & $sCommand & @CRLF)

	SetExtended($PID)
	Return 1
EndFunc   ;==>__FFStartProcess