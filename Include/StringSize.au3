#include-once

; #INDEX# ============================================================================================================
; Title .........: _StringSize
; AutoIt Version : v3.2.12.1 or higher
; Language ......: English
; Description ...: Returns size of rectangle required to display string - width can be chosen
; Remarks .......:
; Note ..........:
; Author(s) .....: Melba23
; ====================================================================================================================

;#AutoIt3Wrapper_au3check_parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6

; #CURRENT# ==========================================================================================================
; _StringSize: Returns size of rectangle required to display string - maximum permitted width can be chosen
; ====================================================================================================================

; #INTERNAL_USE_ONLY#=================================================================================================
; _StringSize_Error: Returns from error condition after DC and GUI clear up
; ====================================================================================================================

; #FUNCTION# =========================================================================================================
; Name...........: _StringSize
; Description ...: Returns size of rectangle required to display string - maximum permitted width can be chosen
; Syntax ........: _StringSize($sText[, $iSize[, $iWeight[, $iAttrib[, $sName[, $iWidth]]]]])
; Parameters ....: $sText   - String to display
;                  $iSize   - [optional] Font size in points - default AutoIt GUI default
;                  $iWeight - [optional] Font weight (400 = normal) - default AutoIt GUI default
;                  $iAttrib - [optional] Font attribute (0-Normal, 2-Italic, 4-Underline, 8 Strike - default AutoIt
;                  $sName   - [optional] Font name - default AutoIt GUI default
;                  $iWidth  - [optional] Width of rectangle - default is unwrapped width of string
; Requirement(s) : v3.2.12.1 or higher
; Return values .: Success - Returns array with details of rectangle required for text:
;                  |$array[0] = String formatted with @CRLF at required wrap points
;                  |$array[1] = Height of single line in selected font
;                  |$array[2] = Width of rectangle required to hold formatted string
;                  |$array[3] = Height of rectangle required to hold formatted string
;                  Failure - Returns 0 and sets @error:
;                  |1 - Incorrect parameter type (@extended = parameter index)
;                  |2 - Failure to create GUI to test label size
;                  |3 - DLL call error - extended set to indicate which
;                  |4 - Font too large for chosen width - longest word will not fit
; Author ........: Melba23
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
;=====================================================================================================================
Func _StringSize($sText, $iSize = Default, $iWeight = Default, $iAttrib = Default, $sName = "", $iWidth = 0)

    Local $avSize_Info[4], $aRet, $iLine_Width = 0, $iLast_Word, $iWrap_Count
    Local $hLabel_Handle, $hFont, $hDC, $oFont, $tSize = DllStructCreate("int X;int Y")

    ; Check parameters are correct type
    If Not IsString($sText) Then Return SetError(1, 1, 0)
    If Not IsNumber($iSize) And $iSize <> Default   Then Return SetError(1, 2, 0)
    If Not IsInt($iWeight)  And $iWeight <> Default Then Return SetError(1, 3, 0)
    If Not IsInt($iAttrib)  And $iAttrib <> Default Then Return SetError(1, 4, 0)
    If Not IsString($sName) Then Return SetError(1, 5, 0)
    If Not IsNumber($iWidth) Then Return SetError(1, 6, 0)

    ; Create GUI to contain test labels, set to required font parameters
    Local $hGUI = GUICreate("", 1200, 500, 10, 10)
        If $hGUI = 0 Then Return SetError(2, 0, 0)
        GUISetFont($iSize, $iWeight, $iAttrib, $sName)

    ; Store unwrapped text
    $avSize_Info[0] = $sText

    ; Ensure EoL is @CRLF and break text into lines
    If StringInStr($sText, @CRLF) = 0 Then $sText = StringRegExpReplace($sText, "[\x0a|\x0d]", @CRLF)
    Local $asLines = StringSplit($sText, @CRLF, 1)

    ; Draw label with unwrapped lines to check on max width
    Local $hText_Label = GUICtrlCreateLabel($sText, 10, 10)
    Local $aiPos = ControlGetPos($hGUI, "", $hText_Label)

    GUISetState(@SW_HIDE)

    GUICtrlDelete($hText_Label)

    ; Store line height for this font size after removing label padding (always 8)
    $avSize_Info[1] = ($aiPos[3] - 8)/ $asLines[0]
    ; Store width and height of this label
    $avSize_Info[2] = $aiPos[2]
    $avSize_Info[3] = $aiPos[3] - 4 ; Reduce margin

    ; Check if wrapping is required
    If $aiPos[2] > $iWidth And $iWidth > 0 Then

        ; Set returned text element to null
        $avSize_Info[0] = ""

        ; Set width element to max allowed
        $avSize_Info[2] = $iWidth

        ; Set line count to zero
        Local $iLine_Count = 0

        ; Take each line in turn
        For $j = 1 To $asLines[0]

            ; Size this line unwrapped
            $hText_Label = GUICtrlCreateLabel($asLines[$j], 10, 10)
            $aiPos = ControlGetPos($hGUI, "", $hText_Label)
            GUICtrlDelete($hText_Label)

            ; Check wrap status
            If $aiPos[2] < $iWidth Then
                ; No wrap needed so count line and store
                $iLine_Count += 1
                $avSize_Info[0] &= $asLines[$j] & @CRLF
            Else
                ; Wrap needed so need to count wrapped lines

                ; Create label to hold line as it grows
                $hText_Label = GUICtrlCreateLabel("", 0, 0)
                ; Initialise Point32 method
                $hLabel_Handle = ControlGetHandle($hGui, "", $hText_Label)
                ; Get DC with selected font
                $aRet = DllCall("User32.dll", "hwnd", "GetDC", "hwnd", $hLabel_Handle)
                If @error Or $aRet[0] = 0 Then Return SetError(3, 1, 0)
                $hDC = $aRet[0]
                $aRet = DllCall("user32.dll", "lparam", "SendMessage", "hwnd", $hLabel_Handle, "int", 0x0031, "wparam", 0, "lparam", 0) ; $WM_GetFont
                If @error Or $aRet[0] = 0 Then Return SetError(3, _StringSize_Error(2, $hLabel_Handle, $hDC, $hGUI), 0)
                $hFont = $aRet[0]
                $aRet = DllCall("GDI32.dll", "hwnd", "SelectObject", "hwnd", $hDC, "hwnd", $hFont)
                If @error Or $aRet[0] = 0 Then Return SetError(3, _StringSize_Error(3, $hLabel_Handle, $hDC, $hGUI), 0)
                $oFont = $aRet[0]

                ; Zero counter
                $iWrap_Count = 0

                While 1

                    ; Set line width to 0
                    $iLine_Width = 0
                    ; Initialise pointer for end of word
                    $iLast_Word = 0

                    For $i = 1 To StringLen($asLines[$j])

                        ; Is this just past a word ending?
                        If StringMid($asLines[$j], $i, 1) = " " Then $iLast_Word = $i - 1
                        ; Increase line by one character
                        Local $sTest_Line = StringMid($asLines[$j], 1, $i)
                        ; Place line in label
                        GUICtrlSetData($hText_Label, $sTest_Line)

                        ; Get line length into size structure
                        $iSize = StringLen($sTest_Line)
                        DllCall("GDI32.dll", "int", "GetTextExtentPoint32", "hwnd", $hDC, "str", $sTest_Line, "int", $iSize, "ptr", DllStructGetPtr($tSize))
                        If @error Then Return SetError(3, _StringSize_Error(4, $hLabel_Handle, $hDC, $hGUI), 0)
                        $iLine_Width = DllStructGetData($tSize, "X")

                        ; If too long exit the loop
                        If $iLine_Width >= $iWidth - Int($iSize / 2) Then ExitLoop
                    Next

                    ; End of the line of text?
                    If $i > StringLen($asLines[$j]) Then
                        ; Yes, so add final line to count
                        $iWrap_Count += 1
                        ; Store line
                        $avSize_Info[0] &= $sTest_Line & @CRLF
                        ExitLoop
                    Else
                        ; No, but add line just completed to count
                        $iWrap_Count += 1
                        ; Check at least 1 word completed or return error
                        If $iLast_Word = 0 Then Return SetError(4, _StringSize_Error(0, $hLabel_Handle, $hDC, $hGUI), 0)
                        ; Store line up to end of last word
                        $avSize_Info[0] &= StringLeft($sTest_Line, $iLast_Word) & @CRLF
                        ; Strip string to point reached
                        $asLines[$j] = StringTrimLeft($asLines[$j], $iLast_Word)
                        ; Trim leading whitespace
                        $asLines[$j] = StringStripWS($asLines[$j], 1)
                        ; Repeat with remaining characters in line
                    EndIf

                WEnd

                ; Add the number of wrapped lines to the count
                $iLine_Count += $iWrap_Count

                ; Clean up
                DllCall("User32.dll", "int", "ReleaseDC", "hwnd", $hLabel_Handle, "hwnd", $hDC)
                If @error Then Return SetError(3, _StringSize_Error(5, $hLabel_Handle, $hDC, $hGUI), 0)
                GUICtrlDelete($hText_Label)

            EndIf

        Next

        ; Convert lines to pixels and add reduced margin
        $avSize_Info[3] = ($iLine_Count * $avSize_Info[1]) + 4

    EndIf

    ; Clean up
    GUIDelete($hGUI)

    ; Return array
    Return $avSize_Info

EndFunc ; => _StringSize

; #INTERNAL_USE_ONLY#============================================================================================================
; Name...........: _StringSize_Error
; Description ...: Returns from error condition after DC and GUI clear up
; Syntax ........: _StringSize_Error($iExtended, $hLabel_Handle, $hDC, $hGUI)
; Parameters ....: $iExtended - required extended value to return
;                  $hLabel_Handle, $hDC, $hGUI - variables as set in _StringSize function
; Author ........: Melba23
; Modified.......:
; Remarks .......: This function is used internally by _StringSize
; ===============================================================================================================================
Func _StringSize_Error($iExtended, $hLabel_Handle, $hDC, $hGUI)

    ; Release DC if created
    DllCall("User32.dll", "int", "ReleaseDC", "hwnd", $hLabel_Handle, "hwnd", $hDC)
    ; Delete GUI
    GUIDelete($hGUI)
    ; Return with extended set
    Return $iExtended

EndFunc ; => _StringSize_Error