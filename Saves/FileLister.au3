;Recursive File Lister
#include-once
#include <Array.au3>

Global $WorkingFolder = "\\chakotay\SoftVal\iAMT\Drivers_7.0\MEI_SOL_Installer"
ToolTip("Start working", 0, 0)

Global $Results[1]
ScanFolder($WorkingFolder)
ToolTip("Done working", 0, 0)

_ArrayDisplay($results,@ScriptLineNumber & " results")

Func ScanFolder($SourceFolder)
	Local $Search
	Local $File
	Local $FileAttributes
	Local $FullFilePath

	$Search = FileFindFirstFile($SourceFolder & "\*.*")
	While 1
		If $Search = -1 Then
			ExitLoop
		EndIf

		$File = FileFindNextFile($Search)
		If @error Then ExitLoop

		$FullFilePath = $SourceFolder & "\" & $File
		$FileAttributes = FileGetAttrib($FullFilePath)

		If StringInStr($FileAttributes, "D") Then
			ScanFolder($FullFilePath)
		Else
			_ArrayAdd($Results, $FullFilePath)
		EndIf
			WEnd
	FileClose($Search)


EndFunc   ;==>ScanFolder
