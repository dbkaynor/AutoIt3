#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Run_Tidy=y
#Tidy_Parameters=/tc 4 /kv 2
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=y
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****

#cs  This is a list of functions in this UDF
    Func _Help($msg)
    Func _GuiDisable($choice)
    Func _AutoSplit($InString, $delimiters, $flag, $index)
    Func _About($WindowName, $SystemS='', $MessageS='')
    Func _FileInfo($file
    Func _FormatedFileGetTime($file, $Type)
    Func _ChoseTextEditor()
    Func _ShowState($input)
    Func _TrueFalse($input)
    Func _Commify($Number)
    Func _AddSlash2PathString($sPath)
    Func _CleanUpPath($sPath)
    Func _SystemLocalTime()
    Func _IPAddress($IPAddress)
    Func _TestIP($IPAddress)
    Func _IpPad($IPAddress)
    Func _IPUnPad($IPAddress)
    Func _CheckIPClass($AddressToTest)
    Func _RemoveBlankLines(ByRef $Array)
    Func _CheckNumericString($NumerToCheck)
    Func _TrimArray(ByRef $Array,$Mode=7)
    Func _FileListToArrayR($sPath, $sFilter = "*", $iFlag = 0, $iRecurse = 0, $iBaseDir = 1, $sExclude = "", $i_Options = 1)
    Func _ArrayDeleteDupes1(ByRef $arrItems)
    Func _FileListToArrayFolders1($sPathF, ByRef $sFileStringF, $sFilterF, $iRecurseF, $sExcludeF = "")
    Func _FileListToArrayFiles1($sPathF, ByRef $sFileStringF, $sFilterF, $sExcludeF = "")
    Func _RecursiveFileSearchC($startDir, $RFSpattern = "*", $Exclude = "", $depth = 0)
    Func _FileListToArrayFiles1($sPathF, ByRef $sFileStringF, $sFilterF, $sExcludeF = "")
    Func _FileListToArrayRecAll1($sPathF, ByRef $sFileStringF, $sFilterF, $sExcludeF = "")
    Func _FileListToArrayRecFiles1($sPathF, ByRef $sFileStringF, $sFilterF, $sExcludeF = "")
    Func _FileListToArrayBrief2a($sPathF, ByRef $sFileStringF, $sFilterF, $sExcludeF = "")
    Func _FileListToArrayBrief1a($sPathF, ByRef $sFileStringF, $sFilterF, $sExcludeF = "")
    Func _Debug($DebugMSG, $Log_filename = '', $ShowMsgBox = False, $Timeout = 0, $Verbose = False)
    Func _StartDebugViewer($DebugMSG)
    Func _CheckWindowLocation($WindowName, $Center = False)
    Func _SetWindowPosition($KeyName, $WindowName, $ReadLine)
    Func _SaveWindowPosition($KeyName,debug	$WindowName, $FileName)
    Func _ComputeStats($InDataArray, $ResultsDataArray)
#ce

#include-once

Opt("MustDeclareVars", 1) ; 0=no, 1=require pre-declare
Opt("GUICoordMode", 1) ; 0=relative, 1=absolute, 2=cell
;Opt("GUIResizeMode", 1) ; 0=no resizing, <1024 special resizing

#include <Array.au3>
#include <Constants.au3>
#include <Date.au3>
#include <File.au3>
#include <GUIConstants.au3>
#include <misc.au3>
#include <String.au3>

Global $AuxPath = @ScriptDir & "\AUXFiles\"
Global $UtilPath = @ScriptDir & "\AUXFiles\Utils\"

Global $RFSarray[1]

;-----------------------------------------------
Func _Help($msg)
    MsgBox(64, "Help", $msg)
EndFunc   ;==>_Help
;-----------------------------------------------

;-----------------------------------------------
; Will either disable, enable, or toggle all gui items
; The gui item in $DoNotDisable will be enabled
; $MaxGUIItems specifies how many gui items to operate on
Func _GuiDisable($choice, $DoNotDisable = '', $MaxGUIItems = 100)
    Static $LastState
    Local $Setting

    Switch $choice
        Case "Enable"
            $Setting = $GUI_ENABLE
        Case "Disable"
            $Setting = $GUI_DISABLE
        Case "Toggle"
            If $LastState = $GUI_DISABLE Then
                $Setting = $GUI_ENABLE
            Else
                $Setting = $GUI_DISABLE
            EndIf
        Case Else
            ;Func _Debug($DebugMSG, $Log_filename = '', $ShowMsgBox = False, $Timeout = 0, $Verbose = False)
            _Debug(@ScriptLineNumber & " Invalid choice at GuiDisable " & $choice, '', True)
    EndSwitch

    For $x = 0 To $MaxGUIItems
        GUICtrlSetState($x, $Setting)
    Next

    If IsNumber($DoNotDisable) Then
        ConsoleWrite(@ScriptLineNumber & ": " & $DoNotDisable & @LF)
        GUICtrlSetState($DoNotDisable, $GUI_ENABLE)
    EndIf

EndFunc   ;==>_GuiDisable
;-----------------------------------------------
Func _AutoSplit($InString, $delimiters, $flag, $index)
    Local $ta = StringSplit($InString, $delimiters, $flag)
    Return $ta[$index]
EndFunc   ;==>_AutoSplit
;-----------------------------------------------
Func _About($WindowName, $SystemS = '', $MessageS = '')
    Local $D = WinGetPos($WindowName)
    Local $MsgString = "About ERROR. Window name not found " & $WindowName
    If IsArray($D) = True Then
        $MsgString = StringFormat("%s" & @LF & @LF & "%s" & @LF & "WinPOS: %d  %d " & @LF & "WinSize: %d %d " & @LF & "Desktop: %d %d" & @LF & "%s", _
                $SystemS, $WindowName, $D[0], $D[1], $D[2], $D[3], @DesktopWidth, @DesktopHeight, $MessageS)
    EndIf
    _Debug($MsgString, '', True)
EndFunc   ;==>_About
;-----------------------------------------------
;Returns a message box abd/or a string with formated file information
Func _FileInfo($file, $MsgBox = True)
    Local $FileInfoString = StringFormat("%s %s %s %s %s %s %s %s %s", _
            FileGetLongName($file), _
            @LF & FileGetShortName($file), _
            @LF & "Attributes:    " & FileGetAttrib($file), _
            @LF & "Size:          " & FileGetSize($file), _
            @LF & "Version:       " & FileGetVersion($file), _
            @LF & "Modified Time: " & _FormatedFileGetTime($file, 0), _
            @LF & "Create time:   " & _FormatedFileGetTime($file, 1), _
            @LF & "Access Time:   " & _FormatedFileGetTime($file, 2))
    If $MsgBox Then MsgBox(64, 'File info', $FileInfoString)
    Return $FileInfoString
EndFunc   ;==>_FileInfo
;-----------------------------------------------
Func _FormatedFileGetTime($file, $type)
    Local $FT = FileGetTime($file, $type)
    Return StringFormat("%s-%s-%s %s:%s:%s", $FT[0], $FT[1], $FT[2], $FT[3], $FT[4], $FT[5])
EndFunc   ;==>_FormatedFileGetTime
;-----------------------------------------------
;Returns a string pointing to the text editor
Func _ChoseTextEditor()
    Const $edit1 = "c:\program files\notepad++\notepad++.exe"
    Const $edit2 = "c:\program files (x86)\notepad++\notepad++.exe"
    Const $edit3 = "notepad.exe"

    If FileExists($edit1) Then
        Return $edit1
    ElseIf FileExists($edit2) Then
        Return $edit2
    Else
        Return $edit3
    EndIf
EndFunc   ;==>_ChoseTextEditor
;-----------------------------------------------
;Returns "CHECKED" or "UNCHECKED" as appropriate
Func _ShowState($input)
    If $input = $GUI_CHECKED Then Return "CHECKED"
    If $input = $GUI_UNCHECKED Then Return "UN_CHECKED"
    Return 'ShowState error'
EndFunc   ;==>_ShowState
;-----------------------------------------------
;Returns "TRUE" or "FALSE" as appropriate
Func _TrueFalse($input)
    If $input = True Then Return "TRUE"
    If $input = False Then Return "FALSE"
    Return 'TrueFalse error'
EndFunc   ;==>_TrueFalse
;-----------------------------------------------
;Insert commas into a string in the correct places
Func _Commify($Number)
    Do
        $Number = StringRegExpReplace($Number, "(?<=[0-9])(?=(?:[0-9]{3})+(?![0-9]))", ",")
    Until @extended = 0
    Return $Number
EndFunc   ;==>_Commify
;-----------------------------------------------
;Adds a '\' to the end of a path string if needed
Func _AddSlash2PathString($sPath)
    If StringRight($sPath, 1) <> "\" Then $sPath = $sPath & "\"
    Return $sPath
EndFunc   ;==>_AddSlash2PathString
;-----------------------------------------------
;Changes  "/" to "\" and  "\\" to "\"
Func _CleanUpPath($sPath)
    $sPath = StringReplace($sPath, "/", "\")
    Return StringReplace($sPath, "\\", "\")
EndFunc   ;==>_CleanUpPath
;-----------------------------------------------
; Returns the current time/date string
Func _SystemLocalTime()
    Local $tTime = _Date_Time_GetLocalTime()
    Return (_Date_Time_SystemTimeToDateTimeStr($tTime) & "  ")
EndFunc   ;==>_SystemLocalTime
;-----------------------------------------------
;This function expands an IP address string an populates array $Results with the values
Func _IPAddress($IPAddress)
    Local $Results[1] ;The array to hold the final results
    Local $array = StringSplit($IPAddress, "+")
    Local $T

    If $array[0] <> 2 Then
        _Debug(@ScriptLineNumber & " No count value found. Testing and returning the value. ", '', True)
        $T = _TestIP($array[1])
        If StringInStr($T, "ERROR0") = 0 Then
            _ArrayAdd($Results, "TestIP failed " & $array[1] & " Return " & $T)
            Return $Results
        EndIf
        _ArrayAdd($Results, $array[1])
        _ArrayDelete($Results, 0) ; This returns the count entry
        Return $Results
    EndIf

    $T = _TestIP($array[1])
    If StringInStr($T, "ERROR0") = 0 Then
        _ArrayAdd($Results, "TestIP failed " & $array[1] & " Return " & $T)
        Return $Results
    EndIf

    Local $IPArray = StringSplit($array[1], ".")
    Local $Count = $array[2]
    Local $IPAddressHEX = Hex($IPArray[1], 2) & Hex($IPArray[2], 2) & Hex($IPArray[3], 2) & Hex($IPArray[4], 2)
    Local $IPAddressDEC = Dec($IPAddressHEX)

    For $x = 0 To $Count
        Local $tmp3 = Hex($IPAddressDEC)
        Local $IPout = Dec(StringMid($tmp3, 1, 2)) & "." & Dec(StringMid($tmp3, 3, 2)) & "." & Dec(StringMid($tmp3, 5, 2)) & "." & Dec(StringMid($tmp3, 7, 2))
        _ArrayAdd($Results, $IPout)
        $IPAddressDEC += 1
    Next
    _ArrayDelete($Results, 0) ; deletes a blank entry at the begining

    Return $Results

EndFunc   ;==>_IPAddress

;-----------------------------------------------
;This function tests an IP address. It must have four octets and be in the correct range (0<>255)
Func _TestIP($IPAddress)
    Local $IPArray = StringSplit($IPAddress, ".") ;This is the ipaddress octets split on .

    If $IPArray[0] <> 4 Then
        Return "ERROR1  Not enough octets. 4 Required, " & $IPArray[0] & " Found. "
    EndIf

    _ArrayDelete($IPArray, 0) ; This returns the count entry

    For $T In $IPArray ;verify that the octet values are within range
        If $T < 0 Or $T > 255 Then
            Return "ERROR2 octet out of range (0 to 255). " & $T
        EndIf
    Next

    Return "ERROR0" ;good address
EndFunc   ;==>_TestIP
;-----------------------------------------------
;This function puts leading 0's on the octets. It also tests for four octets. Returns a string.
Func _IpPad($IPAddress)
    Local $IPArray = StringSplit($IPAddress, ".") ;This is the ipaddress octets split on .

    If $IPArray[0] <> 4 Then Return "ERROR1  Not enough octets. 4 Required, " & $IPArray[0] & " Found."

    Local $ReturnStr = StringFormat("%03d.%03d.%03d.%03d", $IPArray[1], $IPArray[2], $IPArray[3], $IPArray[4])

    Return $ReturnStr ;good address
EndFunc   ;==>_IpPad
;-----------------------------------------------
;This function puts leading 0's on the octets. It also tests for four octets. Returns the new string.
Func _IPUnPad($IPAddress)
    Local $IPArray = StringSplit($IPAddress, ".") ;This is the ipaddress octets split on .

    If $IPArray[0] <> 4 Then Return "ERROR1  Not enough octets. 4 Required, " & $IPArray[0] & " Found."

    Local $ReturnStr = StringFormat("%d.%d.%d.%d", $IPArray[1], $IPArray[2], $IPArray[3], $IPArray[4])

    Return $ReturnStr ;good address
EndFunc   ;==>_IPUnPad
;-----------------------------------------------
;Returns the class of an ip address
Func _CheckIPClass($AddressToTest)
    _Debug(@ScriptLineNumber & " CheckIPClass")
    Local $octets = StringSplit($AddressToTest, ".")
    _Debug(@ScriptLineNumber & " " & $octets[1])
    If $octets[1] = 127 Then
        Return 'Loopback'
    ElseIf $octets[1] >= 1 And $octets[1] <= 126 Then
        Return 'Class A'
    ElseIf $octets[1] >= 128 And $octets[1] <= 191 Then
        Return 'Class B'
    ElseIf $octets[1] >= 192 And $octets[1] <= 223 Then
        Return 'Class C'
    ElseIf $octets[1] >= 224 And $octets[1] <= 239 Then
        Return 'Class D'
    ElseIf $octets[1] >= 240 And $octets[1] <= 254 Then
        Return 'Class E'
    ElseIf $octets[1] = 255 And $octets[2] = 255 And $octets[3] = 255 And $octets[4] = 255 Then
        Return 'Broadcast'
    Else
        Return 'Undefined'
    EndIf
EndFunc   ;==>_CheckIPClass
;-----------------------------------------------
;This function removes blank lines from an array
Func _RemoveBlankLines(ByRef $array)
    Local $TmpArray[1]
    For $x = 0 To UBound($array) - 1
        If StringLen(StringStripWS($array[$x], 2)) <> 0 Then _ArrayAdd($TmpArray, $array[$x])
    Next
    _ArrayDelete($TmpArray, 0)
    ;_ArrayDisplay($TmpArray, @ScriptLineNumber)
    $array = $TmpArray
EndFunc   ;==>_RemoveBlankLines
;-----------------------------------------------
;This function will take in a numeric string verify it and format and return the result
;It will handle integer, decimal and floating point numbers
;Commas are removed from the input string before processing
Func _CheckNumericString($NumerToCheck)
    _Debug(@ScriptLineNumber & " CheckNumber  >>" & $NumerToCheck & "<<")
    $NumerToCheck = StringRegExpReplace($NumerToCheck, "[,]", "", 0)
    If StringIsDigit($NumerToCheck) = 1 Then Return $NumerToCheck
    If StringIsFloat($NumerToCheck) = 1 Then Return $NumerToCheck

    Local $array = StringSplit($NumerToCheck, "eE")
    If $array[0] <> 2 Then
        SetError(1)
        Return "ERROR 1: " & $NumerToCheck
    EndIf
    If (StringIsInt($array[1]) = 0) And (StringIsFloat($array[1]) = 0) Then
        SetError(2)
        Return "ERROR 2: " & $NumerToCheck
    EndIf
    If (StringIsInt($array[2]) = 0) And (StringIsFloat($array[2]) = 0) Then
        SetError(3)
        Return "ERROR 3: " & $NumerToCheck
    EndIf
    Local $result = $array[1] * (10 ^ $array[2])
    If StringInStr($result, "#") <> 0 Then
        SetError(4)
        Return "ERROR 4: " & $result
    EndIf
    Return $result
EndFunc   ;==>_CheckNumericString
;-----------------------------------------------
;Trims white spaces from all strings in an array. Mode determines how it is done
Func _TrimArray(ByRef $array, $Mode = 7)
    ;_ArrayDisplay($Array, "$Array " & @ScriptLineNumber)
    Local $Count = 0
    While True
        $Count += 1
        If $Count >= UBound($array) Then ExitLoop
        $array[$Count] = StringStripWS($array[$Count], $Mode)
    WEnd
    ;_ArrayDisplay($Array, "$Array " & @ScriptLineNumber)

EndFunc   ;==>_TrimArray

;-----------------------------------------------
; #FUNCTION# ====================================================================================================================
;
; Description:      lists all files and folders in a specified path
; Syntax:          Func _FileListToArrayR($sPath, $sFilter = "*", $iFlag = 0, $iRecurse = 0, $iBaseDir = 1, $sExclude = "", $i_Options = 1)
; Parameter(s):    	$s_Path = Path to generate filelist for
;					$s_Filter = The filter to use. Search the Autoit3 manual for the word "WildCards" For details
;                   $i_Flag = determines whether to return file or folders or both
;						$i_Flag=0(Default) Return both files and folders
;                       $i_Flag=1 Return files Only
;						$i_Flag=2 Return Folders Only
;					$i_Recurse = Indicate whether recursion to subfolders required
;						$i_Recurse=0(Default) No recursion to subfolders
;                       $i_Recurse=1 recursion to subfolders
;					$i_BaseDir = Indicate whether base directory name included in returned elements
;						$i_BaseDir=0 base directory name not included
;                       $i_BaseDir=1 (Default) base directory name included
;					$s_Exclude= The Exclude filter to use.  "WildCards" For details
; 					$i_Options= $i_ReturnAsString  $i_deleteduplicate
;
; Requirement(s):   None
; Return Value(s):  On Success - Returns an array containing the list of files and folders in the specified path
;                        On Failure - Returns the an empty string "" if no files are found and sets @Error on errors
;						@Error=1 Path not found or invalid
;						@Error=2 Invalid $s_Filter
;                       @Error=3 Invalid $i_Flag
;                       @Error=4 No File(s) Found

;
;===============================================================================
Func _FileListToArrayR($sPath, $sFilter = "*", $iFlag = 0, $iRecurse = 0, $iBaseDir = 1, $sExclude = "", $i_Options = 1)
    ;ConsoleWrite(@ScriptLineNumber & "  New version of _FileListToArray" & @lf)
    ;Declare local variables
    Local $sFileString, $asList[1], $sep = "|", $sFileString1, $sFilter1 = $sFilter;$hSearch, $sFile,
    Local $i_ReturnAsString = BitAND($i_Options, 2)
    Local $i_deleteduplicate = BitAND($i_Options, 1)
    ;Set default filter to wildcard
    If $sFilter = -1 Or $sFilter == "" Or $sFilter = Default Then $sFilter = "*"

    ;Strip trailing slash from search path
    If StringRight($sPath, 1) == "\" Then $sPath = StringTrimRight($sPath, 1)

    ;Ensure search path exists
    If Not FileExists($sPath) Then Return SetError(1, 1, "")

    ;Return error if special characters are found in filter
    If (StringInStr($sFilter, "\")) Or (StringInStr($sFilter, "/")) Or (StringInStr($sFilter, ":")) Or (StringInStr($sFilter, ">")) Or (StringInStr($sFilter, "<")) Or (StringStripWS($sFilter, 8) = "") Then Return SetError(2, 2, "")

    ;Only allow 0,1,2 for flag options
    If Not ($iFlag = 0 Or $iFlag = 1 Or $iFlag = 2) Then Return SetError(3, 3, "");~     $sFilter = StringReplace("*" & $sFilter & "*", "**", "*")

    ;Determine seperator character
    If StringInStr($sFilter, ';') Then $sep = ";" ;$sFilter &= ';'
    If StringInStr($sFilter, ',') Then $sep = "," ;$sFilter &= ';'

    ;Append pipe to file filter if no semi-colons and pipe symbols are found
    $sFilter &= $sep

    ;Declare local variables, Implode file filter
    Local $aFilterSplit = StringSplit(StringStripWS($sFilter, 8), $sep), $sHoldSplit, $arFolders[2] = [$sPath, ""];~     $cw = ConsoleWrite("UBound($aFilterSplit) =" & UBound($aFilterSplit) & @LF)

    If $sExclude <> "" Then $sExclude = "(?i)(" & StringReplace(StringReplace(StringReplace($sExclude, ".", "\."), "*", ".*"), "?", ".") & ")" ;change the filters to RegExp filters

    ;ConsoleWrite("$sExclude=" & $sExclude & @LF)
    ;exit
    ;If recursion is desired, build an array of all sub-folders in search path (eliminates the need to run a conditional statement against FileAttrib)
    If $iRecurse Then;$cw = ConsoleWrite("UBound($aFilterSplit) =" & UBound($aFilterSplit) & @LF)

        ConsoleWrite(@ScriptLineNumber & " recursion" & @LF)

        ;if folders only, build string ($sFileString1) of foldernames within search path, recursion and exclusion options are passed from main function
        If $iFlag = 2 Then _FileListToArrayFolders1($sPath, $sFileString1, "*", $iRecurse, $sExclude)

        ;if not folders only,  Build string ($sFileString1) of foldernames within search path, recursion (not exclusion, as would exclude some folders from subsequent filesearch) options are passed from main function
        If $iFlag <> 2 And StringTrimRight($sFilter, 1) <> "*" And StringTrimRight($sFilter, 1) <> "*.*" Then
            _FileListToArrayFolders1($sPath, $sFileString1, "*", $iRecurse, "")

            ;Implode folder string
            $arFolders = StringSplit(StringTrimRight($sFileString1, 1), "*")

            ;Store search path in first element
            $arFolders[0] = $sPath
        EndIf
    EndIf

    If $iFlag <> 2 And (StringTrimRight($sFilter, 1) == "*" Or StringTrimRight($sFilter, 1) == "*.*") And $iRecurse Then
        If $iFlag = 1 Then
            _FileListToArrayRecFiles1($sPath, $sFileString, "*")
        ElseIf $iFlag = 0 Then
            _FileListToArrayRecAll1($sPath, $sFileString, "*")
        EndIf
    Else;If ($iFlag <> 2) then

        ;_ArrayDisplay($arFolders,"$arFolders")
        ;Loop through folder array
        For $iCF = 0 To UBound($arFolders) - 1;    $cw = ConsoleWrite("$iCF=" & $iCF & " $arFolders[$iCF]    =" & @LF & $arFolders[$iCF] & @LF)

            ;Verify folder name isn't just whitespace
            If StringStripWS($arFolders[$iCF], 8) = '' Then ContinueLoop

            ;Loop through file filters
            For $iCC = 1 To UBound($aFilterSplit) - 1

                ;Verify file filter isn't just whitespace
                If StringStripWS($aFilterSplit[$iCC], 8) = '' Then ContinueLoop

                ;Append asterisk to file filter if a period is leading
                If StringLeft($aFilterSplit[$iCC], 1) == "." Then $aFilterSplit[$iCC] = "*" & $aFilterSplit[$iCC] ;, "**", "*")

                ;Replace multiple asterisks in file filter
                $sFilter = StringReplace("*" & $sFilter & "*", "**", "*")
                Select; options for not recursing; quicker than filtering after for single directory

                    ;Below needs work, _FileListToArrayBrief1a and _FileListToArrayBrief2a
                    ;should be consolidated with an option passed for the files / folders flag [says Ultima -but slower?]

                    ;Fastest, Not $iRecurse with with files and folders(? was written files only; just Not $iBaseDir), not recursed
                    Case Not $iRecurse And Not $iFlag And Not $iBaseDir
                        _FileListToArrayBrief2a($arFolders[$iCF], $sFileString, $aFilterSplit[$iCC], $sExclude)

                        ;Not $iRecurse and  And $iBaseDir ;fast, with files and folders, not recursed
                    Case Not $iFlag
                        _FileListToArrayBrief1a($arFolders[$iCF], $sFileString, $aFilterSplit[$iCC], $sExclude)

                        ;Fast, with files only,  not recursed
                    Case $iFlag = 1
                        _FileListToArrayFiles1($arFolders[$iCF], $sFileString, $aFilterSplit[$iCC], $sExclude)

                        ;Folders only , not recursed
                    Case Not $iRecurse And $iFlag = 2
                        _FileListToArrayFolders1($arFolders[$iCF], $sFileString, $aFilterSplit[$iCC], $iRecurse, $sExclude)
                EndSelect;$cw = ConsoleWrite("$iCC=" & $iCC & " $sFileString    =" & @LF & $sFileString & @LF)

                ;Append pipe symbol and current file filter onto $sHoldSplit ???????
                If $iCF = 0 Then $sHoldSplit &= $sep & $aFilterSplit[$iCC]; $cw = ConsoleWrite("$iCC=" & $iCC & " $sFileString    =" & @LF & $sFileString & @LF)
            Next

            ;Replace multiple asterisks
            If $iCF = 0 Then $sFilter = StringReplace(StringTrimLeft($sHoldSplit, 1), "**", "*");,$cw = ConsoleWrite("$iCC=" & $iCC & " $sFilter    =" & @LF & $sFilter & @LF)
        Next
    EndIf
    ;Below needs work....

    ;If recursive, folders-only, and filter ins't a wildcard
    If $iRecurse And ($iFlag = 2) And StringTrimRight($sFilter, 1) <> "*" And StringTrimRight($sFilter, 1) <> "*.*" And Not StringInStr($sFilter, "**") Then ; filter folders -------------------

        ;Trim trailing character
        $sFileString1 = StringTrimRight(StringReplace($sFileString1, "*", @LF), 1)

        ;Change the filters to RegExp filters
        $sFilter1 = StringReplace(StringReplace(StringReplace($sFilter1, ".", "\."), "*", ".*"), "?", ".")
        Local $pattern = '(?m)(^(?i)' & $sFilter1 & '$)' ;, $cw = ConsoleWrite("$sFilter    =" & @LF & $sFilter1 & @LF), $cw = ConsoleWrite("$pattern    =" & @LF & $pattern & @LF)
        $asList = StringRegExp($sFileString1, $pattern, 3)

        ;If only relative file / folder names are desired
        If (Not $iBaseDir) Then

            ; past ARRAY.AU3 DEPENDENCY
            $sFileString1 = _ArrayToString($asList, "*")
            $sFileString1 = StringReplace($sFileString1, $sPath & "\", "", 0, 2)
            $asList = StringSplit($sFileString1, "*")
        EndIf
    ElseIf $iRecurse And ($iFlag = 2) Then
        $sFileString = StringStripCR($sFileString1)
    EndIf;If UBound($asList) > 1 Then ConsoleWrite("$asList[1]     =" & @LF & $asList[1] & @LF);~

    ;past ARRAY.AU3 DEPENDENCY
    If IsArray($asList) And UBound($asList) > 0 And $asList[0] <> "" And Not IsNumber($asList[0]) Then _ArrayInsert($asList, 0, UBound($asList))
    If IsArray($asList) And UBound($asList) > 1 And $asList[0] <> "" Then Return $asList
    If (Not $iBaseDir) Or (Not $iRecurse And Not $iFlag And Not $iBaseDir) Then $sFileString = StringReplace($sFileString, $sPath & "\", "", 0, 2)
    If $i_ReturnAsString Then Return StringTrimRight($sFileString, 1)
    Local $arReturn = StringSplit(StringTrimRight($sFileString, 1), "*");~     local $a=ConsoleWrite("$sFileString :"&@lf&StringReplace($sFileString,"|",@lf)&@lf),$timerstamp1=TimerInit()
    If $i_deleteduplicate And IsArray($arReturn) And UBound($arReturn) > 1 And $arReturn[1] <> "" And Not (UBound($aFilterSplit) = 3 And $aFilterSplit[2] == "") Then _ArrayDeleteDupes1($arReturn);and  $arFolders[1]<>""
    Return $arReturn;~     Return StringSplit(StringTrimRight($sFileString, 1), "*")
EndFunc   ;==>_FileListToArrayR
;-----------------------------------------------
;===============================================================================
;
; Description:  _ArrayDeleteDupes1; deletes duplicates in an Array 1D
; Syntax:           _ArrayDeleteDupes1(ByRef $ar_Array)
; Parameter(s):    	$ar_Array = 1d Array
; Requirement(s):   None
; Return Value(s):  On Success - Returns a sorted array with no duplicates
;                        On Failure -
;						@Error=1 P
;						@Error=2
;
; Author(s):        randallc
;===============================================================================
Func _ArrayDeleteDupes1(ByRef $arrItems)
    If @OSType = "WIN32_WINDOWS" Then Return 0
    Local $i = 0, $objDictionary = ObjCreate("Scripting.Dictionary")
    For $strItem In $arrItems
        If Not $objDictionary.Exists($strItem) Then
            $objDictionary.Add($strItem, $strItem)
        EndIf
    Next
    ReDim $arrItems[$objDictionary.Count]
    For $strKey In $objDictionary.Keys
        $arrItems[$i] = $strKey
        $i += 1
    Next
    $arrItems[0] = $objDictionary.Count - 1
    Return 1
EndFunc   ;==>_ArrayDeleteDupes1
;===============================================================================
;
; Description:      Helper  self-calling func for  _FileListToArray wrapper; lists all  folders in a specified path
; Syntax:           _FileListToArrayFolders1($s_PathF, ByRef $s_FileStringF, $s_FilterF,  $i_RecurseF)
; Parameter(s):    	$s_PathF = Path to generate filelist for
;					$s_FileStringF = The string for lists all folders only in a specified path
;					$s_FilterF = The filter to use. Search the Autoit3 manual for the word "WildCards" For details
;					$i_RecurseF = Indicate whether recursion to subfolders required
;						$i_RecurseF=0(Default) No recursion to subfolders
;                       $i_RecurseF=1 recursion to subfolders
;					$sExcludeF= The Exclude filter to use.  "WildCards" For details
; Requirement(s):   None
; Return Value(s):  On Success - Returns an array containing the list of files and folders in the specified path
;                        On Failure - Returns the an empty string "" if no files are found and sets @Error on errors
;						@Error=1 Path not found or invalid
;						@Error=2 Invalid $s_Filter
;                       @Error=3 Invalid $i_Flag
;                 @Error=4 No File(s) Found
;
; Author(s):        randallc; modified from SolidSnake, SmoKe_N, GEOsoft and big_daddy
;===============================================================================
Func _FileListToArrayFolders1($sPathF, ByRef $sFileStringF, $sFilterF, $iRecurseF, $sExcludeF = "")
    Local $hSearch = FileFindFirstFile($sPathF & "\" & $sFilterF), $sPathF2, $sFileF
    If $hSearch = -1 Then Return SetError(4, 4, "")
    If $sExcludeF == "" Then
        While 1
            $sFileF = FileFindNextFile($hSearch)
            If @error Then ExitLoop

            $sPathF2 = $sPathF & "\" & $sFileF
            If StringInStr(FileGetAttrib($sPathF2), "D") Then ;directories only wanted; and  the attrib shows is  directory
                $sFileStringF &= $sPathF2 & "*" ;this writes the filename to the delimited string with * as delimiter
                If $iRecurseF = 1 Then _FileListToArrayFolders1($sPathF2, $sFileStringF, $sFilterF, $iRecurseF)
            EndIf
        WEnd
    Else
        While 1
            $sFileF = FileFindNextFile($hSearch)
            If @error Then ExitLoop

            $sPathF2 = $sPathF & "\" & $sFileF; if folders only and this pattern matches exclude pattern, no further list or subdir
            If StringRegExp($sPathF2, $sExcludeF) Then ContinueLoop
            If StringInStr(FileGetAttrib($sPathF2), "D") Then ;directories only wanted; and  the attrib shows is  directory
                $sFileStringF &= $sPathF2 & "*" ;this writes the filename to the delimited string with * as delimiter with * as delimiter
                If $iRecurseF = 1 Then _FileListToArrayFolders1($sPathF2, $sFileStringF, $sFilterF, $iRecurseF, $sExcludeF)
            EndIf
        WEnd
    EndIf
    FileClose($hSearch)
EndFunc   ;==>_FileListToArrayFolders1
;===============================================================================
; Returns an array of files in a folder tree.
Func _RecursiveFileSearchC($startDir, $RFSpattern = "*", $Exclude = "", $depth = 0)
;~ 	If StringRight($startDir, 1) <> "\"  Then $startDir &= "\"
    If StringRight($startDir, 1) == "\" Then $startDir = StringTrimRight($startDir, 1)

    If $depth = 0 Then
        ;change filters to RegExp filters
        If $RFSpattern <> "" Then $RFSpattern = "(?i)(^" & StringReplace(StringReplace(StringReplace($RFSpattern, ".", "\."), "*", ".*"), "?", ".") & "$)" ;change the filters to RegExp filters
        If $Exclude <> "" Then $Exclude = "(?i)(^" & StringReplace(StringReplace(StringReplace($Exclude, ".", "\."), "*", ".*"), "?", ".") & "$)" ;change the filters to RegExp filters

        ;Get count of all files in subfolders
        Local $RFSfilecount = DirGetSize($startDir, 1)
        Global $RFSarray[$RFSfilecount[1] + 1]
    EndIf

    Local $search = FileFindFirstFile($startDir & "\*.*")
    If @error Then Return

    ;Search through all files and folders in directory
    While 1
        Local $next = FileFindNextFile($search)
        If @error Then ExitLoop

        ;If folder, recurse
        If StringInStr(FileGetAttrib($startDir & "\" & $next), "D") Then
            _RecursiveFileSearchC($startDir & "\" & $next, $RFSpattern, $Exclude, $depth + 1)
        Else
            If StringRegExp($next, $RFSpattern, 0) And Not StringRegExp($next, $Exclude, 0) Then
                ;Append filename to array
                $RFSarray[$RFSarray[0] + 1] = $startDir & "\" & $next

                ;Increment filecount
                $RFSarray[0] += 1
            EndIf
        EndIf
    WEnd
    FileClose($search)

    If $depth = 0 Then
        ReDim $RFSarray[$RFSarray[0] + 1]
        Return $RFSarray
    EndIf
EndFunc   ;==>_RecursiveFileSearchC
;===============================================================================
;
; Description:      Helper  func for  _FileListToArray wrapper; lists all files in a specified path
; Syntax:           _FileListToArrayFiles1($s_PathF, ByRef $s_FileStringF, $s_FilterF) ;quick as not recursive
; Parameter(s):    	$s_PathF = Path to generate filelist for
;					$s_FileStringF = The string for lists all files and folders in a specified path
;					$s_FilterF = The filter to use. Search the Autoit3 manual for the word "WildCards" For details
;					$sExcludeF= The Exclude filter to use.  "WildCards" For details
; Requirement(s):   None
; Return Value(s):  On Success - Returns an array containing the list of files and folders in the specified path
;                        On Failure - Returns the an empty string "" if no files are found and sets @Error on errors
;						@Error=1 Path not found or invalid
;						@Error=2 Invalid $s_Filter
;                 @Error=4 No File(s) Found
;
; Author(s):        randallc; modified from SolidSnake, SmoKe_N, GEOsoft and big_daddy
;===============================================================================
Func _FileListToArrayFiles1($sPathF, ByRef $sFileStringF, $sFilterF, $sExcludeF = "")
    Local $hSearch = FileFindFirstFile($sPathF & "\" & $sFilterF), $sPathF2, $sFileF
    If $hSearch = -1 Then Return SetError(4, 4, "")
    If $sExcludeF == "" Then
        While 1
            $sFileF = FileFindNextFile($hSearch)
            If @error Then ExitLoop

            $sPathF2 = $sPathF & "\" & $sFileF;directories not wanted; and  the attrib shows not  directory
            If Not StringInStr(FileGetAttrib($sPathF2), "D") Then
                $sFileStringF &= $sPathF2 & "*" ;this writes the filename to the delimited string with * as delimiter
            EndIf
        WEnd
    Else
        While 1
            $sFileF = FileFindNextFile($hSearch)
            If @error Then ExitLoop

            $sPathF2 = $sPathF & "\" & $sFileF;directories not wanted; and  the attrib shows not  directory; and filename [only]  does not match exclude
            If Not StringInStr(FileGetAttrib($sPathF2), "D") _
                    And Not StringRegExp($sFileF, $sExcludeF) Then $sFileStringF &= $sPathF2 & "*" ;this writes the filename to the delimited string with * as delimiter
        WEnd
    EndIf
    FileClose($hSearch)
EndFunc   ;==>_FileListToArrayFiles1
;===============================================================================
;
; Description:      Helper  self-calling func for  _FileListToArray wrapper; lists all files and folders in a specified path, recursive
; Syntax:           _FileListToArrayRecFiles1($s_PathF, ByRef $s_FileStringF, $s_FilterF) ; recursive
; Parameter(s):    	$s_PathF = Path to generate filelist for
;					$s_FileStringF = The string for lists all files and folders in a specified path
;					$s_FilterF = The filter to use. Search the Autoit3 manual for the word "WildCards" For details
;					$sExcludeF= The Exclude filter to use.  "WildCards" For details
; Requirement(s):   None
; Return Value(s):  On Success - Returns an array containing the list of files and folders in the specified path
;                        On Failure - Returns the an empty string "" if no files are found and sets @Error on errors
;						@Error=1 Path not found or invalid
;						@Error=2 Invalid $s_Filter
;                 @Error=4 No File(s) Found
;
; Author(s):        randallc; modified from SolidSnake, SmoKe_N, GEOsoft and big_daddy
;===============================================================================
Func _FileListToArrayRecAll1($sPathF, ByRef $sFileStringF, $sFilterF, $sExcludeF = "")
    Local $hSearch = FileFindFirstFile($sPathF & "\" & $sFilterF), $sPathF2, $sFileF
    If $hSearch = -1 Then Return SetError(4, 4, "")
    If $sExcludeF == "" Then
        While 1
            $sFileF = FileFindNextFile($hSearch)
            If @error Then ExitLoop

            $sPathF2 = $sPathF & "\" & $sFileF
            $sFileStringF &= $sPathF2 & "*" ;this writes the filename to the delimited string with * as delimiter
            If StringInStr(FileGetAttrib($sPathF2), "D") Then _FileListToArrayRecAll1($sPathF2, $sFileStringF, $sFilterF);, $iFlagF, $iRecurseF)
        WEnd
    Else
        While 1
            $sFileF = FileFindNextFile($hSearch)
            If @error Then ExitLoop

            $sPathF2 = $sPathF & "\" & $sFileF
            If StringInStr(FileGetAttrib($sPathF2), "D") Then
                $sFileStringF &= $sPathF2 & "*" ;this writes the directoryname
                _FileListToArrayRecAll1($sPathF2, $sFileStringF, $sFilterF, $sExcludeF);, $iFlagF, $iRecurseF)
            Else ;if not directory, check Exclude match
                If Not StringRegExp($sFileF, $sExcludeF) Then $sFileStringF &= $sPathF2 & "*" ;this writes the filename to the delimited string with * as delimiter
            EndIf
        WEnd
    EndIf
    FileClose($hSearch)
EndFunc   ;==>_FileListToArrayRecAll1
;===============================================================================
;
; Description:      Helper  self-calling func for  _FileListToArray wrapper; lists all files  in a specified path, recursive
; Syntax:           _FileListToArrayRecFiles1($s_PathF, ByRef $s_FileStringF, $s_FilterF) ; recursive
; Parameter(s):    	$s_PathF = Path to generate filelist for
;					$s_FileStringF = The string for lists all files and folders in a specified path
;					$s_FilterF = The filter to use. Search the Autoit3 manual for the word "WildCards" For details
;					$sExcludeF= The Exclude filter to use.  "WildCards" For details
; Requirement(s):   None
; Return Value(s):  On Success - Returns an array containing the list of files and folders in the specified path
;                        On Failure - Returns the an empty string "" if no files are found and sets @Error on errors
;						@Error=1 Path not found or invalid
;						@Error=2 Invalid $s_Filter
;                 @Error=4 No File(s) Found
;
; Author(s):        randallc; modified from SolidSnake, SmoKe_N, GEOsoft and big_daddy
;===============================================================================
Func _FileListToArrayRecFiles1($sPathF, ByRef $sFileStringF, $sFilterF, $sExcludeF = "")
    Local $hSearch = FileFindFirstFile($sPathF & "\" & $sFilterF), $sPathF2, $sFileF
    If $hSearch = -1 Then Return SetError(4, 4, "")
    If $sExcludeF == "" Then
        While 1
            $sFileF = FileFindNextFile($hSearch)
            If @error Then ExitLoop

            $sPathF2 = $sPathF & "\" & $sFileF
            If StringInStr(FileGetAttrib($sPathF2), "D") Then
                _FileListToArrayRecFiles1($sPathF2, $sFileStringF, $sFilterF);, $iFlagF, $iRecurseF)
            Else
                $sFileStringF &= $sPathF2 & "*" ;this writes the filename to the delimited string with * as delimiter
            EndIf
        WEnd
    Else
        While 1
            $sFileF = FileFindNextFile($hSearch)
            If @error Then ExitLoop

            $sPathF2 = $sPathF & "\" & $sFileF
            If StringInStr(FileGetAttrib($sPathF2), "D") Then
                _FileListToArrayRecFiles1($sPathF2, $sFileStringF, $sFilterF, $sExcludeF);, $iFlagF, $iRecurseF)
            Else
                If Not StringRegExp($sFileF, $sExcludeF) Then $sFileStringF &= $sPathF2 & "*" ;this writes the filename to the delimited string with * as delimiter
            EndIf
        WEnd
    EndIf
    FileClose($hSearch)
EndFunc   ;==>_FileListToArrayRecFiles1
;===============================================================================
;
; Description:      Helper  func for  _FileListToArray wrapper; ;Fastest, Not $iRecurse with with files and folders(? was written files only; just Not $iBaseDir), not recursed
; Syntax:           _FileListToArrayBrief2a($s_PathF, ByRef $s_FileStringF, $s_FilterF) ;quick as not recursive
; Parameter(s):    	$s_PathF = Path to generate filelist for
;					$s_FileStringF = The string for lists all files and folders in a specified path
;					$s_FilterF = The filter to use. Search the Autoit3 manual for the word "WildCards" For details
;					$sExcludeF= The Exclude filter to use.  "WildCards" For details
; Requirement(s):   None
; Return Value(s):  On Success - Returns an array containing the list of files and folders in the specified path
;                        On Failure - Returns the an empty string "" if no files are found and sets @Error on errors
;						@Error=1 Path not found or invalid
;						@Error=2 Invalid $s_Filter
;                 @Error=4 No File(s) Found
;
; Author(s):        randallc; modified from SolidSnake, SmoKe_N, GEOsoft and big_daddy
;===============================================================================
Func _FileListToArrayBrief2a($sPathF, ByRef $sFileStringF, $sFilterF, $sExcludeF = "")
    Local $hSearch = FileFindFirstFile($sPathF & "\" & $sFilterF), $sFileF
    If $hSearch = -1 Then Return SetError(4, 4, "")
    If $sExcludeF == "" Then
        While 1
            $sFileF = FileFindNextFile($hSearch)
            If @error Then ExitLoop

            $sFileStringF &= $sFileF & "*" ;this writes the filename to the delimited string with * as delimiter; only time no full path included
        WEnd
    Else
        While 1
            $sFileF = FileFindNextFile($hSearch)
            If @error Then ExitLoop

            ;If not StringRegExp($sFileF,$sExcludeF) then $sFileStringF &= $sPathF2 & "*"
            If Not StringRegExp($sFileF, $sExcludeF) Then $sFileStringF &= $sFileF & "*" ;this writes the filename to the delimited string with * as delimiter; only time no full path included
        WEnd
    EndIf
    FileClose($hSearch)
EndFunc   ;==>_FileListToArrayBrief2a
;===============================================================================
;
; Description:      Helper  func for  _FileListToArray wrapper; lists all files and folders in a specified path, not recursive
; Syntax:           _FileListToArrayBrief1a($s_PathF, ByRef $s_FileStringF, $s_FilterF) ;quick as not recursive
; Parameter(s):    	$s_PathF = Path to generate filelist for
;					$s_FileStringF = The string for lists all files and folders in a specified path
;					$s_FilterF = The filter to use. Search the Autoit3 manual for the word "WildCards" For details
;					$sExcludeF= The Exclude filter to use.  "WildCards" For details
; Requirement(s):   None
; Return Value(s):  On Success - Returns an array containing the list of files and folders in the specified path
;                        On Failure - Returns the an empty string "" if no files are found and sets @Error on errors
;						@Error=1 Path not found or invalid
;						@Error=2 Invalid $s_Filter
;                 @Error=4 No File(s) Found
;
; Author(s):        randallc; modified from SolidSnake, SmoKe_N, GEOsoft and big_daddy
;===============================================================================

Func _FileListToArrayBrief1a($sPathF, ByRef $sFileStringF, $sFilterF, $sExcludeF = "")
    Local $hSearch = FileFindFirstFile($sPathF & "\" & $sFilterF), $sFileF
    If $hSearch = -1 Then Return SetError(4, 4, "")
    If $sExcludeF == "" Then
        While 1
            $sFileF = FileFindNextFile($hSearch)
            If @error Then ExitLoop

            $sFileStringF &= $sPathF & "\" & $sFileF & "*" ;this writes the filename to the delimited string with * as delimiter [remo]
        WEnd
    Else
        While 1
            $sFileF = FileFindNextFile($hSearch)
            If @error Then ExitLoop

            If Not StringRegExp($sFileF, $sExcludeF) Then $sFileStringF &= $sPathF & "\" & $sFileF & "*" ;this writes the filename to the delimited string with * as delimiter [remo]
        WEnd
    EndIf
    FileClose($hSearch)
EndFunc   ;==>_FileListToArrayBrief1a
;-----------------------------------------------
Func _Debug($DebugMSG, $Log_filename = '', $ShowMsgBox = False, $Timeout = 0, $Verbose = False)
    ConsoleWrite(@ScriptLineNumber & " " & StringInStr($DebugMSG, "-1") & @LF)
    ConsoleWrite($DebugMSG)

    If $Verbose Then
        $DebugMSG = "DEBUG >> " & @ScriptName & "  " & $DebugMSG & @LF
    Else
        If StringInStr($DebugMSG, "-1") = 1 Then $DebugMSG = StringReplace($DebugMSG, "-1", "", 1)
        $DebugMSG = $DebugMSG & @LF
    EndIf

    DllCall("kernel32.dll", "none", "OutputDebugString", "str", $DebugMSG)

    If $Log_filename Then _FileWriteLog($Log_filename, $DebugMSG)
    If $ShowMsgBox = True Then MsgBox(48, @ScriptName & " Debug", $DebugMSG, $Timeout)
EndFunc   ;==>_Debug
;-----------------------------------------------
;WinGetProcess ( "title" [, "text"] )
Func _StartDebugViewer($Clear = True)
    Opt("WinTitleMatchMode", 2)
    Local $result = WinGetProcess('DebugView')
    Local $DebugPath = $UtilPath & 'Dbgview.exe'
    ConsoleWrite(@ScriptLineNumber & " " & $DebugPath & "  " & $result & @CRLF)
    If $result = -1 Then ShellExecute($DebugPath)
    If $Clear Then _Debug("DBGVIEWCLEAR")
EndFunc   ;==>_StartDebugViewer
;-----------------------------------------------
;This will check the position of a window and if it is off of the desktop will put it in the center of the main screen
;If center is true the window will be moved to the center of the main screen
;WinMove ( "title", "text", x, y [, width [, height[, speed]]] )
Func _CheckWindowLocation($WindowName, $Position = 'null')
    Local $array = WinGetPos($WindowName, "")

    If IsArray($array) = 0 Then
        MsgBox(16, "_CheckWindowLocation error", 'Window not found: ' & $WindowName)
        Return -1
    EndIf

    Local $X_Position = $array[0]
    Local $Y_Position = $array[1]
    Local $Width = $array[2]
    Local $Height = $array[3]

    Switch $Position
        Case 'null'; This option puts the window onto the screen is any portion is off screen
            ;First check the X position
            If $X_Position < 0 Or @DesktopWidth < $X_Position + $Width Then
                WinMove($WindowName, "", (@DesktopWidth / 2) - ($Width / 2), $Y_Position)
            EndIf
            ;Now check the Y position
            If $Y_Position < 0 Or @DesktopHeight < $Y_Position + $Height Then
                WinMove($WindowName, "", $X_Position, @DesktopHeight / 2 - $Height / 2)
            EndIf
        Case 'NW'
            WinMove($WindowName, "", 0, 0)
        Case 'W'
            WinMove($WindowName, "", 0, @DesktopHeight / 2 - $Height / 2)
        Case 'SW'
            WinMove($WindowName, "", 0, @DesktopHeight - $Height)
        Case 'N'
            WinMove($WindowName, "", @DesktopWidth / 2 - ($Width / 2), 0)
        Case 'center'
            WinMove($WindowName, "", (@DesktopWidth / 2) - ($Width / 2), @DesktopHeight / 2 - $Height / 2)
        Case 'S'
            WinMove($WindowName, "", @DesktopWidth / 2 - ($Width / 2), @DesktopHeight - $Height)
        Case 'NE'
            WinMove($WindowName, "", @DesktopWidth - $Width, 0)
        Case 'E'
            WinMove($WindowName, "", @DesktopWidth - $Width, @DesktopHeight / 2 - $Height)
        Case 'SE'
            WinMove($WindowName, "", @DesktopWidth - $Width, @DesktopHeight - $Height)
        Case Else
            MsgBox(16, '_CheckWindowLocation error', 'Unknown position specified')
            Return -2
    EndSwitch

    Return 1
EndFunc   ;==>_CheckWindowLocation
;-----------------------------------------------
;This is used to restore the window position when loading an options file
Func _SetWindowPosition($KeyName, $WindowName, $ReadLine)
    If StringInStr($ReadLine, $KeyName) Then
        Local $F = StringMid($ReadLine, StringInStr($ReadLine, ":") + 1)
        $F = StringSplit($F, " ", 2)

        If WinMove($WindowName, "", $F[0], $F[1], $F[2], $F[3]) = 0 Then MsgBox(16, "SetWindowPosition error", $WindowName)

    EndIf
EndFunc   ;==>_SetWindowPosition
;-----------------------------------------------
;This is used to save the window position when creating an options file
Func _SaveWindowPosition($KeyName, $WindowName, $FileName)
    Local $F = WinGetPos($WindowName, "")
    FileWriteLine($FileName, $KeyName & $F[0] & " " & $F[1] & " " & $F[2] & " " & $F[3])
EndFunc   ;==>_SaveWindowPosition
;-----------------------------------------------
; These are constants for use with ComputeStats
Const $NumberOfDataPoints = 0
Const $MinimumValue = 1
Const $MaximumValue = 2
Const $MeansValue = 3
Const $MedianValue = 4
Const $ModeValue = 5
Const $StandardDeviationValue = 6

Func _ComputeStats($InDataArray, ByRef $ResultsDataArray)
    ;_ArrayDisplay($InDataArray, @ScriptLineNumber)
    Local $ModeArray1[1]
    Local $ModeArray2[1]
    Local $MedianArray[1]
    Local $STDArray[1]

    $ResultsDataArray[$NumberOfDataPoints] = UBound($InDataArray)
    $ResultsDataArray[$MinimumValue] = 9e19
    $ResultsDataArray[$MaximumValue] = 0
    $ResultsDataArray[$MeansValue] = 0
    $ResultsDataArray[$ModeValue] = 0
    $ResultsDataArray[$StandardDeviationValue] = 0

    Local $TotalOfAllValues = 0

    For $i = 0 To UBound($InDataArray) - 1
        ;Get the data for median value
        _ArrayAdd($MedianArray, $InDataArray[$i])
        _ArrayAdd($STDArray, $InDataArray[$i])
        $TotalOfAllValues = $TotalOfAllValues + $InDataArray[$i]

        ;Calculate the mode
        Local $T = _ArraySearch($ModeArray1, $InDataArray[$i])
        If $T = -1 Then ;Did not find the value
            _ArrayAdd($ModeArray1, $InDataArray[$i])
            _ArrayAdd($ModeArray2, 1)
        Else ;Did find the value
            $ModeArray2[$T] = $ModeArray2[$T] + 1
        EndIf

        ;Calulate minimum and maximum values
        ConsoleWrite(@ScriptLineNumber & " " & $InDataArray[$i] & @CRLF)
        If $InDataArray[$i] < $ResultsDataArray[$MinimumValue] Then $ResultsDataArray[$MinimumValue] = $InDataArray[$i]
        If $InDataArray[$i] > $ResultsDataArray[$MaximumValue] Then $ResultsDataArray[$MaximumValue] = $InDataArray[$i]
    Next
    ;Finish means calulation
    $ResultsDataArray[$MeansValue] = $TotalOfAllValues / $ResultsDataArray[$NumberOfDataPoints]

    ;Finish the mode calculation
    Local $biggest = 0
    Local $biggestValue = ''
    For $i = 0 To UBound($ModeArray2) - 1
        Local $current = $ModeArray2[$i]
        If $biggest < $current Then
            $biggest = $current
            $biggestValue = $ModeArray1[$i]
        EndIf
    Next
    $ResultsDataArray[$ModeValue] = $biggestValue & " (" & $biggest & " instances)"

    ;Finish the median
    _ArraySort($MedianArray)
    If Mod($NumberOfDataPoints, 2) = 0 Then
        Local $a1 = $MedianArray[Floor(UBound($MedianArray) / 2)]
        Local $a2 = $MedianArray[Ceiling(UBound($MedianArray) / 2)]
        $ResultsDataArray[$MedianValue] = ($a1 + $a2) / 2
    Else
        $ResultsDataArray[$MedianValue] = $MedianArray[UBound($MedianArray) / 2]
    EndIf

    ; Standard deviation
    _ArrayDelete($STDArray, 0)
    Local $sumofall = 0
    For $i = 0 To UBound($STDArray) - 1
        $sumofall = $sumofall + ($STDArray[$i] - $ResultsDataArray[$MeansValue]) ^ 2
    Next
    $sumofall = $sumofall / UBound($STDArray) - 1
    $ResultsDataArray[$StandardDeviationValue] = Sqrt(Abs($sumofall))
    ;_ArrayDisplay($ResultsDataArray, @ScriptLineNumber)
EndFunc   ;==>ComputeStats
;-----------------------------------------------
