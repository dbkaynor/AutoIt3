Opt("MustDeclareVars", 1)
Local $perl = "c:\perl\bin\perl.exe"
Local $script = "C:\Perl\Projects\_CurrentProjects\SavedScripts\TkPath.pl"
Local $folder = "C:\Perl\Projects\_CurrentProjects\"

If FileExists($perl) = False Then
	MsgBox(0, "Error", $perl & " does not exist", 5)
	Exit
EndIf
If FileExists($script) = False Then
	MsgBox(0, "Error", $script & " does not exist", 5)
	Exit
EndIf
If FileExists($folder) = False Then
	MsgBox(0, "Error", $folder & " does not exist", 5)
	Exit
EndIf

ShellExecute($perl, $script, $folder, "", @SW_HIDE)