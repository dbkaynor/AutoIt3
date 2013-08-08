Opt("MustDeclareVars", 1)
Dim $Message1
Dim $Message2
Dim $Message3
Dim $Message4
Dim $Title

$Title = "Help me if I am lost!"
$Message1 = "If you have found me and you are not my owner please contact:"
$Message2 = "Doug Kaynor, 19515 SW Alexander ST, Aloha, Oregon 97006-2308"
$Message3 = "Phone 503-591-9115 or 503-314-3321"
$Message4 = "Thanks for your honesty in advance. - Doug"

MsgBox(266304, $Title, $Message1 & @CRLF & $Message2 & @CRLF & $Message3 & @CRLF & $Message4, 5)

ShellExecute("explorer.exe", @ScriptDir)