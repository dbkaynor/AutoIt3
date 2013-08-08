
Local $dbname = 'BirdID.mdb'
;$tblname = *
$fldname = 'BIRD NAME'

Dim $_output

Local $query
Local $adoCon = ObjCreate("ADODB.Connection")
If Not IsObj($adoCon) Then MsgBox(14, 'ERROR ' & @ScriptLineNumber, 'Not an object')
$adoCon.Open("Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & $dbname)

local $query = "SELECT * FROM  * where 1=0"
$adoRs = $adoCon.Execute($query)
While Not $adoRs.EOF
	$_output = $_output & $adoRs.fields("title").value & @CRLF
	$adoRs.MoveNext
WEnd
$adoCon.Close
MsgBox(0,"Guest List",$_output)
