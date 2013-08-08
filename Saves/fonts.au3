#include <GuiConstants.au3>
#include "GuiFontAndColors.au3"

GuiCreate("Example", 300, 200)
GuiSetBkColor($BACKGROUNDCOLOR)
GuiSetFont($MESSAGEFONTSIZE, 400, 0, $MESSAGEFONTNAME)

GuiCtrlCreateLabel(" GUI has Desktop background color", 10, 10)
    GuiCtrlSetBkColor(-1, $INFOWINDOWCOLOR);label has tooltip color
    GuiCtrlSetColor(-1, $BUTTONTEXTCOLOR);label's text color

$button = GuiCtrlCreateButton("Click Me", 100, 100)

GuiSetState()
While 1
    $msg = GuiGetMsg()
    If $msg = $GUI_EVENT_CLOSE Then ExitLoop
    If $msg = $button Then
        MsgBox(4096, "Note...","Message boxes use the correct font automatically")
    EndIf
WEnd