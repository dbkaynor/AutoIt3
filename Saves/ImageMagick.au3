Dim $img
Dim $ret

; Initialize error handler
$oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")

$img = ObjCreate("ImageMagickObject.MagickImage.1")

$ret = $img.Convert("C:\Temp\windows2.jpg", _
		"-resize", "640x400", _
		"-format", "bmp", _
		"C:\Temp\New.bmp")


Func MyErrFunc()
	$HexNumber = Hex($oMyError.number, 8)
	MsgBox(0, "COM Error Test", "We intercepted a COM Error !" & @CRLF & @CRLF & _
			"err.description is: " & @TAB & $oMyError.description & @CRLF & _
			"err.windescription:" & @TAB & $oMyError.windescription & @CRLF & _
			"err.number is: " & @TAB & $HexNumber & @CRLF & _
			"err.lastdllerror is: " & @TAB & $oMyError.lastdllerror & @CRLF & _
			"err.scriptline is: " & @TAB & $oMyError.scriptline & @CRLF & _
			"err.source is: " & @TAB & $oMyError.source & @CRLF & _
			"err.helpfile is: " & @TAB & $oMyError.helpfile & @CRLF & _
			"err.helpcontext is: " & @TAB & $oMyError.helpcontext _
			)
	SetError(1) ; to check for after this function returns
EndFunc   ;==>MyErrFunc

Const $ERROR_SUCCESS = 0

Dim $img
Dim $info
Dim $msgs
Dim $elem
Dim $sMsgs

; Initialize error handler
$oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")

; This is the simplest sample I could come up with. It creates
; the ImageMagick COM object and then sends a copy of the IM
; logo out to a JPEG image files on disk.
;
$img = ObjCreate("ImageMagickObject.MagickImage.1")
;
; The methods for the IM COM object are identical to utility
; command line utilities. You have convert, composite, identify,
; mogrify, and montage. We did not bother with animate, and
; display since they have no purpose in this context.
;
; The argument list is exactly the same as the utility programs
; as a list of strings. In fact you should just be able to
; copy and past - do simple editing and it will work. See the
; other samples for more elaborate command sequences and the
; documentation for the utility programs for more details.
;
$sMsgs = $img.Convert("logo:", "-format", "%m,%h,%w", "logo.jpg")
;
; By default - the string returned is the height, width, and the
; type of the image that was output. You can control this using
; the -format "xxxxxx" command as documented by identify.
;


MsgBox(0, "info: ", "Return = " & $sMsgs)