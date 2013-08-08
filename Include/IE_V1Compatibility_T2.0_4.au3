#include-once
#cs
	Title:   Internet Explorer Automation UDF Library for AutoIt3
	Filename:  IE_V1Compatibility.au3
	Description: This file includes finctions deprecated from T1.0 to T2.0 of the IE.aus UDF library.  Include it along with
				 IE.au3 to allow T1 scripts to run with the T2 library.  It can only be used along with IT.au3 T2.0-4 or higher.
	Author:   DaleHohm
	Version:  T2.0-4
	Last Update: 6/26/06
	Requirements: IE.au3 T2.0-4 or higher, AutoIt3 Beta with COM support (3.1.1.63 or higher), 
				  Developed/Tested on WindowsXP Pro with Internet Explorer
#ce
Global $__IEAU3V1Compatibility = True
#region Deprecated Functions
;===============================================================================
;
; Function Name:    _IEFrameGetSrcByIndex()
; Description:		Obtain the URL references within a frame by 0-based index
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):  On Success - Returns an object variable pointing to
;                   On Failure - 0  and sets @ERROR = 1
; Author(s):        Dale Hohm
;
;===============================================================================
;
Func _IEFrameGetSrcByIndex($o_object, $i_index)
	If IsObj($o_object.document.parentwindow.frames.item ($i_index)) Then
		SetError(0)
		Return $o_object.document.parentwindow.frames.item ($i_index).src
	Else
		SetError(1)
		Return 0
	EndIf
EndFunc   ;==>_IEFrameGetSrcByIndex

;===============================================================================
;
; Function Name:    _IEFrameGetSrcByName()
; Description:		Obtain the URL references within a frame by name
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):  On Success - Returns an object variable pointing to
;                   On Failure - 0  and sets @ERROR = 1
; Author(s):        Dale Hohm
;
;===============================================================================
;
Func _IEFrameGetSrcByName($o_object, $s_Name, $i_index = 0)
	If IsObj($o_object.document.parentwindow.frames ($s_Name, $i_index)) Then
		SetError(0)
		Return $o_object.document.parentwindow.frames ($s_Name, $i_index).src
	Else
		SetError(1)
		Return 0
	EndIf
EndFunc   ;==>_IEFrameGetSrcByName

;===============================================================================
;
; Function Name:    _IEFormGetNameByIndex()
; Description:		Obtain the name of a form by its 0-based index
; Parameter(s):
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):  On Success - Returns an object variable pointing to
;                   On Failure - 0  and sets @ERROR = 1
; Author(s):        Dale Hohm
;
;===============================================================================
;
Func _IEFormGetNameByIndex($o_object, $i_index)
	; $o_object - IE object
	; return object reference to specific form
	If IsObj($o_object.document.forms.item ($i_index)) Then
		SetError(0)
		Return $o_object.document.forms.item ($i_index).name
	Else
		SetError(1)
		Return 0
	EndIf
EndFunc   ;==>_IEFormGetNameByIndex

;===============================================================================
;
; Function Name:    _IEFormElementGetTypeByIndex()
; Description:		Obtain the type of a givien form element within a form by 0-based index
;					(button, checkbox, fileUpload, hidden, image, password, radio, reset, submit, or text)
; Parameter(s):		$o_object 	- form object
;					$i_index	- 0-based index of form element within form
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):  On Success - Returns an object variable pointing to
;                   On Failure - 0  and sets @ERROR = 1
; Author(s):        Dale Hohm
;
;===============================================================================
;
Func _IEFormElementGetTypeByIndex($o_object, $i_index)
	If IsObj($o_object.elements.item ($i_index)) Then
		SetError(0)
		Return $o_object.elements.item ($i_index).type
	Else
		SetError(1)
		Return 0
	EndIf
EndFunc   ;==>_IEFormElementGetTypeByIndex

;===============================================================================
;
; Function Name:    _IEFormElementOptionGetCount()
; Description:		Get count of Options within a Select drop-down form element
; Parameter(s):		$o_object 	- Select Element object
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):  On Success 	- Returns an object variable pointing to
;                   On Failure 	- 0  and sets @ERROR = 1
; Author(s):        Dale Hohm
;
;===============================================================================
;
Func _IEFormElementOptionGetCount($o_object)
	If Not IsObj($o_object) Then
		SetError(1)
		Return 0
	EndIf
	;
	SetError(0)
	Return $o_object.options.length
EndFunc   ;==>_IEFormElementOptionGetCount

Func _IEClickLinkByIndex($o_object, $i_index, $f_wait = 1)
	Return _IELinkClickByIndex($o_object, $i_index, $f_wait)
EndFunc   ;==>_IEClickLinkByIndex

Func _IEClickLinkByText($o_object, $s_linkText, $i_index = 0, $f_wait = 1)
	Return _IELinkClickByText($o_object, $s_linkText, $i_index, $f_wait)
EndFunc   ;==>_IEClickLinkByText

Func _IEDocumentGetObj($o_object)
	Return _IEDocGetObj($o_object)
EndFunc   ;==>_IEDocumentGetObj

Func _IEClickImg($o_object, $s_linkText, $s_mode = "src", $i_index = 0, $f_wait = 1)
	Return _IEImgClick($o_object, $s_linkText, $s_mode, $i_index, $f_wait)
EndFunc   ;==>_IEClickImg

Func _IEGetProperty($o_object, $s_property)
	Return _IEPropertyGet($o_object, $s_property)
EndFunc   ;==>_IEGetProperty

;===============================================================================
;
; Function Name:    _IEFormGetCount()
; Description:		Returns the number of Forms in the specified document
; Parameter(s):		$o_object	- Object variable of an InternetExplorer.Application, Window or Frame object
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):  On Success - Returns an integer denoting the number of Forms
;                   On Failure - Returns 0 and sets @ERROR = 1
; Author(s):        Dale Hohm
;
;===============================================================================
;
Func _IEFormGetCount($o_object)
	If Not IsObj($o_object) Then
		SetError(1)
		Return 0
	EndIf
	;
	SetError(0)
	Return $o_object.document.forms.length
EndFunc   ;==>_IEFormGetCount

;===============================================================================
;
; Function Name:    _IEFormGetObjByIndex()
; Description:		Returns an object variable reference to a Form by 0-based index
; Parameter(s):		$o_object	- Object variable of an InternetExplorer.Application, Window or Frame object
;					$i_index	- 0-based index of the Form you wish to match
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):  On Success - Returns an object variable pointing to the Form object
;                   On Failure - Returns 0 and sets @ERROR = 1
; Author(s):        Dale Hohm
;
;===============================================================================
;
Func _IEFormGetObjByIndex($o_object, $i_index)
	If Not IsObj($o_object) Then
		SetError(1)
		Return 0
	EndIf
	;
	SetError(0)
	Return $o_object.document.forms.item ($i_index)
EndFunc   ;==>_IEFormGetObjByIndex

;===============================================================================
;
; Function Name:    _IEFormElementGetCount()
; Description:		Returns the number of Element within a given Form
; Parameter(s):		$o_object	- Object variable of an InternetExplorer.Application, Form object
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):  On Success - Returns an integer denoting the number of Forms Elements
;                   On Failure - Returns 0 and sets @ERROR = 1
; Author(s):        Dale Hohm
;
;===============================================================================
;
Func _IEFormElementGetCount($o_object)
	If Not IsObj($o_object) Then
		SetError(1)
		Return 0
	EndIf
	;
	SetError(0)
	Return $o_object.elements.length
EndFunc   ;==>_IEFormElementGetCount

;===============================================================================
;
; Function Name:    _IEFormElementGetObjByIndex()
; Description:		Returns an object variable reference to a Form Element by 0-based index
; Parameter(s):		$o_object	- Object variable of an InternetExplorer.Application, Form object
;					$i_index	- 0-based index of the Form Element you wish to match
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):  On Success - Returns an object variable pointing to the Form Element object
;                   On Failure - Returns 0 and sets @ERROR = 1
; Author(s):        Dale Hohm
;
;===============================================================================
;
Func _IEFormElementGetObjByIndex($o_object, $i_index)
	If Not IsObj($o_object) Then
		SetError(1)
		Return 0
	EndIf
	;
	SetError(0)
	Return $o_object.elements ($i_index)
EndFunc   ;==>_IEFormElementGetObjByIndex

;===============================================================================
;
; Function Name:    _IEFrameGetCount()
; Description:		Returns the number of Frames (standard or iFrame) in the specified document
; Parameter(s):		$o_object 	- Object variable of an InternetExplorer.Application, Window or Frame object
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):  On Success - Returns an integer denoting the number of Frames
;                   On Failure - Returns 0 and sets @ERROR = 1
; Author(s):        Dale Hohm
;
;===============================================================================
;
Func _IEFrameGetCount($o_object)
	If Not IsObj($o_object) Then
		SetError(1)
		Return 0
	EndIf
	;
	SetError(0)
	Return $o_object.document.parentwindow.frames.length
EndFunc   ;==>_IEFrameGetCount

;===============================================================================
;
; Function Name:    _IEFrameGetObjByIndex()
; Description:		Returns an object reference to a Window within the specified Frame
; Parameter(s):		$o_object 	- Object variable of an InternetExplorer.Application, Window or Frame object
;					$i_index	- 0-based index of the Frame you wish to match
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):  On Success 	- Returns an object variable pointing to the Window object in a Frame
;                   On Failure 	- Returns 0 and sets @ERROR = 1
; Author(s):        Dale Hohm
;
;===============================================================================
;
Func _IEFrameGetObjByIndex($o_object, $i_index)
	If Not IsObj($o_object) Then
		SetError(1)
		Return 0
	EndIf
	;
	SetError(0)
	;	Return $o_object.document.parentWindow.frames.item ($i_index)
	Return $o_object.document.frames.item ($i_index)
EndFunc   ;==>_IEFrameGetObjByIndex

;===============================================================================
;
; Function Name:    _IEFrameGetNameByIndex()
; Description:		Returns an object name of a Frame by 0-based index
; Parameter(s):		$o_object	- Object variable of an InternetExplorer.Application, Window or Frame object
;					$i_index	- 0-based index of the Frame you wish to match
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):  On Success - Returns the string name of a Frame
;                   On Failure - Returns 0 and sets @ERROR = 1
; Author(s):        Dale Hohm
;
;===============================================================================
;
Func _IEFrameGetNameByIndex($o_object, $i_index)
	If Not IsObj($o_object) Then
		SetError(1)
		Return 0
	EndIf
	;
	SetError(0)
	Return $o_object.document.parentwindow.frames.item ($i_index).name
EndFunc   ;==>_IEFrameGetNameByIndex

;===============================================================================
;
; Function Name:    _IETableGetCount()
; Description:		Returns the number of Tables in the specified document
; Parameter(s):		$o_object	- Object variable of an InternetExplorer.Application, Window or Frame object
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):  On Success - Returns an integer denoting the number of Tables
;                   On Failure - Returns 0 and sets @ERROR = 1
; Author(s):        Dale Hohm
;
;===============================================================================
;
Func _IETableGetCount($o_object)
	If Not IsObj($o_object) Then
		SetError(1)
		Return 0
	EndIf
	;
	SetError(0)
	Return $o_object.document.GetElementsByTagName ("table").length
EndFunc   ;==>_IETableGetCount

;===============================================================================
;
; Function Name:    _IETableGetObjByIndex()
; Description:		Returns an object reference to a Table in a document by 0-based index
; Parameter(s):		$o_object	- Object variable of an InternetExplorer.Application, Window or Frame object
;					$i_index	- 0-based index of the Table you wish to match
; Requirement(s):   AutoIt3 Beta with COM support (post 3.1.1)
; Return Value(s):  On Success - Returns the collection object containing the Tables
;                   On Failure - Returns 0 and sets @ERROR = 1
; Author(s):        Dale Hohm
;
;===============================================================================
;
Func _IETableGetObjByIndex($o_object, $i_index)
	If Not IsObj($o_object) Then
		SetError(1)
		Return 0
	EndIf
	;
	SetError(0)
	Return $o_object.document.GetElementsByTagName ("table").item ($i_index)
EndFunc   ;==>_IETableGetObjByIndex

#endregion
