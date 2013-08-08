#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_icon=../icons/HotSun.ico
#AutoIt3Wrapper_outfile=C:\Program Files\AutoIt3\Dougs\FTP.exe
#AutoIt3Wrapper_Res_Comment=A FTP
#AutoIt3Wrapper_Res_Description=FTP
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Fileversion=0.0.0.1
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=Y
#AutoIt3Wrapper_Res_ProductVersion=666
#AutoIt3Wrapper_Res_LegalCopyright=Copyright 2010 Douglas B Kaynor
#AutoIt3Wrapper_Res_SaveSource=y
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Field=AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Res_Field=Compile date|%longdate% %time%
#AutoIt3Wrapper_Res_Field=Developer|Douglas Kaynor
#AutoIt3Wrapper_Res_Field=Email|doug@kaynor.net
#AutoIt3Wrapper_Res_Field=Made By|Douglas Kaynor
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=y
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Run_Tidy=y
#Tidy_Parameters=/tc 4 /kv 2
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****
;-----------------------------------------------

#cs
    This area is used to store things to do, bugs, and other notes

    Fixed:

    Todo:

#ce

Opt("MustDeclareVars", 1)


#include <FTPEx.au3>
#include <array.au3>

Global $server = 'localhost'
Global $username = 'kaynor'
Global $pass = 'ronyak'
Global $dir = 'photo'
Global $files[1]
Global $folders[1]

RecursiveFTP()

Func RecursiveFTP()
    Local $Result
    Local $FTP_Session = _FTP_Open('MyFTP Control')
    If $FTP_Session = 0 Then
        MsgBox(16, @ScriptLineNumber & " FTP open errror", "FTP open errror")
        Exit
    EndIf

    ConsoleWrite(@ScriptLineNumber & " FTP_Session: " & $FTP_Session & @CRLF)

    $Result = _FTP_Connect($FTP_Session, $server, $username, $pass)
    If $Result = 0 Then
        MsgBox(16, @ScriptLineNumber & " FTP Connect error", "FTP Connect error" & @CRLF & $server & @CRLF & $username & @CRLF & $pass & @CRLF & @error)
        Exit
    EndIf

    ConsoleWrite(@ScriptLineNumber & " Result:" & $Result & @CRLF)

    ConsoleWrite(@ScriptLineNumber & " FTP_L:" & _FTP_DirList($FTP_Session) & @CRLF)

    _ArrayAdd($folders, $dir)
    _ArrayDelete($folders, 0)
    ;-----------------------------------------------
    Local $count = 0

    While UBound($folders) > $count

        $Result = _FTP_DirSetCurrent($FTP_Session, $folders[$count])
        ConsoleWrite(@ScriptLineNumber & " " & $folders[$count] & "   " & $count & @CRLF)
        If $Result = 0 Then
            MsgBox(16, @ScriptLineNumber & " Set directory errror", "Set directory errror" & @CRLF & $folders[$count] & @CRLF & $count & @CRLF & $Result & @CRLF & @error)
            ;_ArrayDisplay($folders, @ScriptLineNumber)
            Exit
        EndIf

        Local $RawList = _FTP_ListToArrayEx($FTP_Session, 0)
        ;_ArrayDisplay($RawList, @ScriptLineNumber)

        ; this parses the last returned $Rawlist
        For $I = 1 To UBound($RawList) - 1
            ; ConsoleWrite(@ScriptLineNumber & " raw " & $dir & "\" & $RawList[$I][0] & @CRLF)

            If $RawList[$I][2] = 16 Then ;folders
                _ArrayAdd($folders, $dir & "\" & $RawList[$I][0])
            ElseIf $RawList[$I][2] = 128 Then ;files
                _ArrayAdd($files, $dir & "\" & $RawList[$I][0] & " ~ " & $RawList[$I][1] & " ~ " & $RawList[$I][2] & " ~ " & $RawList[$I][3])
            EndIf
        Next
        $count = $count + 1
    WEnd

    ;clean up the arrays
    _ArrayDelete($folders, 0)
    _ArrayDelete($files, 0)

    ;_ArrayDisplay($folders, @ScriptLineNumber)

    #cs
        local $counter = 0
        While $counter < UBound($folders)
        ConsoleWrite(@ScriptLineNumber & " " & $counter & " " & UBound($folders) & @CRLF)
        ConsoleWrite(@ScriptLineNumber & $folders[$counter] & @CRLF)
        $counter = $counter + 1
        WEnd
    #ce

    ;-----------------------------------------------

    _FTP_Close($FTP_Session)

    _ArrayDisplay($folders, @ScriptLineNumber & " folders")
    _ArrayDisplay($files, @ScriptLineNumber & " files")

EndFunc   ;==>RecursiveFTP
