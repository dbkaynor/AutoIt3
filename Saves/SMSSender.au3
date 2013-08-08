#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_icon=SMS.ico
#AutoIt3Wrapper_outfile=SMS.exe
#AutoIt3Wrapper_Res_Comment=By: Isaac Flaum
#AutoIt3Wrapper_Res_Description=Allows computer to send SMS messages to phones
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_Res_LegalCopyright=Isaac Flaum 2008
#AutoIt3Wrapper_Res_Language=1033
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <GuiScrollBars.au3>
#include<file.au3>
#include<misc.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <ScrollBarConstants.au3>
#include <EditConstants.au3>
$s_SmtpServer = "smtp.gmail.com" ; address for the smtp-server to use - REQUIRED
$s_FromName = "" ; name from who the email was sent
$s_FromAddress = "" ;  address from where the mail should come
$s_ToAddress = "" ; destination address of the email - REQUIRED
$s_Subject = "" ; subject from the email - can be anything you want it to be
$as_Body = "" ; the messagebody from the mail - can be left blank but then you get a blank mail
$s_AttachFiles = "" ; the file you want to attach- leave blank if not needed
$s_CcAddress = "" ; address for cc - leave blank if not needed
$s_BccAddress = "" ; address for bcc - leave blank if not needed
$s_Username = "" ; username for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
$s_Password = "" ; password for the account used from where the mail gets sent  - Optional (Needed for eg GMail)
$IPPort = 465 ; port used for sending the mail
$ssl = 1 ; enables/disables secure socket layer sending - put to 1 if using httpS

Global $oMyRet[2]
Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")
GUICreate("SMS", 160, 180)
$send = GUICtrlCreateButton("Send", 5, 150, 150, 25)
$phonenum = GUICtrlCreateLabel("Phone Number:", 5, 5, 80, 20)
$phonesubject = Guictrlcreatelabel("Subject:", 5, 30, 80, 20)
$subject = guictrlcreateinput("", 50, 30, 105, 20)
guictrlsetstate(-1, $GUI_DISABLE)
$to = GUICtrlCreateInput("", 85, 5, 70, 20)
$attatchment = GUICtrlCreateButton("Attatchment", 5, 55, 80, 20)
$carrier = GUICtrlCreateCombo("Carrier", 90, 55, 65, 20)
GUICtrlSetData(-1, "Verizon|AT&T|Sprint|T-Mobile", "Carrier")
$numchar = GUICtrlCreateLabel("Characters Left: 160", 5, 130, 100, 20)
$txt = GUICtrlCreateEdit("", 5, 85, 150, 40, BitOR($ES_AUTOVSCROLL,$ES_AUTOHSCROLL,$ES_WANTRETURN))
guisetbkcolor(0x99cc66)
GUISetState(@SW_SHOW)
While 1
    sleep(10)
    $msg = GUIGetMsg()
    $readtxt = GUICtrlRead($txt)
$length = StringLen($readtxt)
$actualchar = Execute(160 - $length)
GUICtrlSetData($numchar, "Characters Left: " & $actualchar)
    Select
        Case ($msg = $attatchment)
            $s_AttachFiles = FileOpenDialog("Attatchment", @DesktopDir, "Files (*.mid;*.gif;*.jpg;*.mp3)")
            guictrlsetstate($subject, $GUI_ENABLE)
        Case ($msg = -3)
            Exit
        Case ($msg = $send)
            $s_Subject = guictrlread($subject)
            $as_Body = GUICtrlRead($txt)
            $s_ToAddress = GUICtrlRead($to)
            $wacarrier = GUICtrlRead($carrier)
            Select
                Case ($wacarrier = "Verizon")
                    If $s_AttachFiles = "" Then
                        $s_ToAddress = $s_ToAddress & "@vtext.com"
                    Else
                        $s_ToAddress = $s_ToAddress & "@vzwpix.com"
                    EndIf
                Case ($wacarrier = "AT&T")
                    If $s_AttachFiles = "" Then
                        $s_ToAddress = $s_ToAddress & "@txt.att.net"
                    Else
                        $s_ToAddress = $s_ToAddress & "@mobile.att.net"
                    EndIf
                Case ($wacarrier = "Sprint")
                    If $s_AttachFiles = "" Then
                        $s_ToAddress = $s_ToAddress & "@messaging.sprintpcs.com"
                    Else
                        $s_ToAddress = $s_ToAddress & "@messaging.sprintpcs.com"
                    EndIf
                Case ($wacarrier = "T-Mobile")
                    If $s_AttachFiles = "" Then
                        $s_ToAddress = $s_ToAddress & "@tmomail.net"
                    Else
                        $s_ToAddress = $s_ToAddress & "@tmomail.net"
                    EndIf
            EndSelect
            $sendprog = ProgressOn("Sending", "Sending Message...")
            $rc = _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject, $as_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Username, $s_Password, $IPPort, $ssl)
            If @error Then
                MsgBox(0, "Error sending message", "Error code:" & @error & "  Rc:" & $rc)
            EndIf
            
        EndSelect
    
WEnd

Func _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject = "", $as_Body = "", $s_AttachFiles = "", $s_CcAddress = "", $s_BccAddress = "", $s_Username = "", $s_Password = "", $IPPort = 25, $ssl = 0)
    ProgressSet(10)
    $objEmail = ObjCreate("CDO.Message")
    $objEmail.From = '"' & $s_FromName & '" <' & $s_FromAddress & '>'
    ProgressSet(20)
    $objEmail.To = $s_ToAddress
    Local $i_Error = 0
    Local $i_Error_desciption = ""
    ProgressSet(30)
    If $s_CcAddress <> "" Then $objEmail.Cc = $s_CcAddress
    If $s_BccAddress <> "" Then $objEmail.Bcc = $s_BccAddress
    $objEmail.Subject = $s_Subject
    If StringInStr($as_Body, "<") And StringInStr($as_Body, ">") Then
        $objEmail.HTMLBody = $as_Body
    Else
        $objEmail.Textbody = $as_Body & @CRLF
    EndIf
    If $s_AttachFiles <> "" Then
        Local $S_Files2Attach = StringSplit($s_AttachFiles, ";")
        For $x = 1 To $S_Files2Attach[0]
            $S_Files2Attach[$x] = _PathFull($S_Files2Attach[$x])
            If FileExists($S_Files2Attach[$x]) Then
                $objEmail.AddAttachment($S_Files2Attach[$x])
            Else
                $i_Error_desciption = $i_Error_desciption & @LF & 'File not found to attach: ' & $S_Files2Attach[$x]
                SetError(1)
                Return 0
            EndIf
        Next
    EndIf
    ProgressSet(40)
    $objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    $objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $s_SmtpServer
    $objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $IPPort
    ProgressSet(60)
    ;Authenticated SMTP
    If $s_Username <> "" Then
        $objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        $objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
        $objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
    EndIf
    If $ssl Then
        $objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    EndIf
    ProgressSet(80)
    ;Update settings
    $objEmail.Configuration.Fields.Update
    ; Sent the Message
    $objEmail.Send
    If @error Then
        SetError(2)
        Return $oMyRet[1]
    EndIf
    ProgressSet(100)
    ProgressOff()
EndFunc   ;==>_INetSmtpMailCom
;
;
; Com Error Handler