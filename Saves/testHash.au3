  #include <_HashTable.au3>
; Same script as Garys modified for the _HashTable.au3
  _Main()
 
  Func _Main()
    ; Initialize the hash table
       _HashTableInit()
      Local $vKey, $sItem, $sMsg
 
    ; Add keys with items
      _DebugPrint('_AddItem("One", "Same")' & @TAB & _HT_AddItem("One", "Same"))
      _DebugPrint('_AddItem("Two", "Car")' & @TAB & _HT_AddItem("Two", "Car"))
      _DebugPrint('_AddItem("Three", "House")' & @TAB & _HT_AddItem("Three", "House"))
      _DebugPrint('_AddItem("Three", "House")' & @TAB & _HT_AddItem("Three", "House"))
      _DebugPrint('_AddItem("Four", "Boat")' & @TAB & _HT_AddItem("Four", "Boat"))
    ; Display items
      MsgBox(0x0, 'Items Count: ' &  _HT_GetCount(), $sItem, 3)
      If _HT_FindItem('One') Then
        ; Display item
          MsgBox(0x0, 'Item One', _HT_GetItem('One'), 2)
        ; Set an item
          _HT_UpdateItem('One', 'Changed')
        ; Display item
          MsgBox(0x20, 'Did Item One Change?', _HT_GetItem('One'), 3)
        ; Remove key
          _HT_RemoveItem('One')
        ;
      EndIf
 
    ; Store items into a variable
      For $vKey In $a_HT_TableKeys
            If $vKey <> "" Then $sItem &= $vKey & " : " & _HT_GetItem($vKey) & @CRLF
      Next
 
    ; Display items
      MsgBox(0x0, 'Items Count: ' &  _HT_GetCount(), $sItem, 3)
 
    ; Add items into an array
      $aArray =  _HT_ToArray()
 
    ; Display items in the array
      For $i = 0 To  _HT_GetCount() - 1
          MsgBox(0x0, 'Array [ ' & $i & ' ]', $aArray[$i], 2)
      Next
 
      _DebugPrint('_ItemRemove("Two")' & @TAB & _HT_RemoveItem("Two"))
      _DebugPrint('_ItemRemove("Three")' & @TAB & _HT_RemoveItem("Three"))
      _DebugPrint('_ItemRemove("Three")' & @TAB & _HT_RemoveItem("Three"))
      _DebugPrint('_ItemRemove("Four")' & @TAB & _HT_RemoveItem("Four"))
       _DebugPrint('$sz_HT_Deleted' &@TAB & $sz_HT_Deleted)
    ; use keys like an array index
      For $x = 1 To 3
          _HT_AddItem($x, "")
      Next
      $sItem = ""
      _HT_UpdateItem(2, "My Custom Item")
      _HT_UpdateItem(1, "This is the 1st item")
      _HT_UpdateItem(3, "This is the last item")
      For $vKey In $a_HT_TableKeys
            If $vKey <> "" Then $sItem &= $vKey & " : " & _HT_GetItem($vKey) & @CRLF
      Next
    ; Display items
      MsgBox(0x0, 'Items Count: ' & _HT_GetCount(), $sItem, 3)
 
      $sItem = ""
    
      _HT_ChangeKey(2, "My New Key")
       if @error Then _DebugPrint(StringFormat("error changing key\t%d\t%d",@error,@extended))
      For $vKey In $a_HT_TableKeys
            If $vKey <> "" Then $sItem &= $vKey & " : " & _HT_GetItem($vKey) & @CRLF
      Next
    ; Display items
      MsgBox(0x0, 'Items Count: ' & _HT_GetCount(), $sItem, 3)
    
      _HT_RemoveAll()
      MsgBox(0x0, 'Items Count',_HT_GetCount(), 3)
    
 
  EndFunc ;==>_Main
  Func _DebugPrint($s_Text)
      ConsoleWrite( _
              "!===========================================================" & @LF & _
              "+===========================================================" & @LF & _
              "-->" & $s_Text & @LF & _
              "+===========================================================" & @LF)
  EndFunc ;==>_DebugPrint