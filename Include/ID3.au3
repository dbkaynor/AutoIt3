#include-once
#include <Array.au3>
#include <String.au3>
#include <File.au3>
Opt("MustDeclareVars", 1)

Global $ID3Filenames = "", $AlbumArtFilename = ""
Global $ID3BufferArray[1] = [0];use to buffer ID3 Data


;===============================================================================
; Function Name:    _ID3ReadTag($Filename, $iVersion = 0, $sFilter = -1, $iReturnArray = 0)
; Description:      Reads ID3 Tag data from mp3 file and stores tag data in an array
;						Reads ID3v1.0, ID3v1.1, ID3v2.3, ID3v2.2($sFilter will not work for ID3v2.2)
; Parameter(s):     $Filename 		- Filename of mp3 file include full path
;					$iVersion		- ID3 Version to read (Default: 0 => ID3v1 & ID3v2)
;										0 => Read ID3v1 & ID3v2
;										1 => Read ID3v1
;										2 => Read ID3v2
;					$sFilter		- ID3v2 Frame Filter (Default: -1 => all frames)
;										If used format should be a pipe delimted string
;										of the Frame IDs (ie. "TIT2|TALB|TPE1|TYER|APIC")
;					$iReturnArray	- Return Array of Field Data (Default: 0 => array not returned)
;										0 => Do not return Array
;										1 => Return Array
; Requirement(s):   None
; Return Value(s):  On Success - Returns the Array of tag fields and data or 1 depends on $iReturnArray.
;                   On Failure - 0 and Sets @error = -1 for end-of-file, @error = 1 for other errors
;Notes:	(Frame ID Definitions - not all may work)
;~ AENC Audio encryption
;~ APIC Attached picture
;~ COMM Comments
;~ COMR Commercial frame
;~ ENCR Encryption method registration
;~ EQUA Equalization
;~ ETCO Event timing codes
;~ GEOB General encapsulated object
;~ GRID Group identification registration
;~ IPLS Involved people list
;~ LINK Linked information
;~ MCDI Music CD identifier
;~ MLLT MPEG location lookup table
;~ OWNE Ownership frame
;~ PRIV Private frame
;~ PCNT Play counter
;~ POPM Popularimeter
;~ POSS Position synchronisation frame
;~ RBUF Recommended buffer size
;~ RVAD Relative volume adjustment
;~ RVRB Reverb
;~ SYLT Synchronized lyric/text
;~ SYTC Synchronized tempo codes
;~ TALB Album/Movie/Show title
;~ TBPM BPM (beats per minute)
;~ TCOM Composer
;~ TCON Content type
;~ TCOP Copyright message
;~ TDAT Date
;~ TDLY Playlist delay
;~ TENC Encoded by
;~ TEXT Lyricist/Text writer
;~ TFLT File type
;~ TIME Time
;~ TIT1 Content group description
;~ TIT2 Title/songname/content description
;~ TIT3 Subtitle/Description refinement
;~ TKEY Initial key
;~ TLAN Language(s)
;~ TLEN Length
;~ TMED Media type
;~ TOAL Original album/movie/show title
;~ TOFN Original filename
;~ TOLY Original lyricist(s)/text writer(s)
;~ TOPE Original artist(s)/performer(s)
;~ TORY Original release year
;~ TOWN File owner/licensee
;~ TPE1 Lead performer(s)/Soloist(s)
;~ TPE2 Band/orchestra/accompaniment
;~ TPE3 Conductor/performer refinement
;~ TPE4 Interpreted, remixed, or otherwise modified by
;~ TPOS Part of a set
;~ TPUB Publisher
;~ TRCK Track number/Position in set
;~ TRDA Recording dates
;~ TRSN Internet radio station name
;~ TRSO Internet radio station owner
;~ TSIZ Size
;~ TSRC ISRC (international standard recording code)
;~ TSSE Software/Hardware and settings used for encoding
;~ TYER Year
;~ TXXX User defined text information frame
;~ UFID Unique file identifier
;~ USER Terms of use
;~ USLT Unsychronized lyric/text transcription
;~ WCOM Commercial information
;~ WCOP Copyright/Legal information
;~ WOAF Official audio file webpage
;~ WOAR Official artist/performer webpage
;~ WOAS Official audio source webpage
;~ WORS Official internet radio station homepage
;~ WPAY Payment
;~ WPUB Publishers official webpage
;~ WXXX User defined URL link frame
;===============================================================================
Func _ID3ReadTag($Filename, $iVersion = 0, $sFilter = -1, $iReturnArray = 0)

	If Not (FileExists($Filename)) Then Return 0
	Dim $ID3BufferArray[1] = [0]  ;clear Buffer
	Switch $iVersion
		Case 0
			_ReadID3v2($Filename, $ID3BufferArray, $sFilter)
			_ReadID3v1($Filename, $ID3BufferArray)
		Case 1
			_ReadID3v1($Filename, $ID3BufferArray)
		Case 2
			_ReadID3v2($Filename, $ID3BufferArray, $sFilter)
	EndSwitch

	;Add Filename to buffer
	;****************************************************************************************
	_ArrayAdd($ID3BufferArray,'Filename' & '|' & $Filename) 
	$ID3BufferArray[0] += 1
	SetError(0)
	;****************************************************************************************

	;Setup simple array to return (IF $iReturnArray = 1)
	;****************************************************************************************
	If $iReturnArray = 1 Then
		;Return a 1D array with just frame string data (original)
		Local $ID3Array[1] = [0]
		For $i = 3 To $ID3BufferArray[0]
			$ID3FrameArray = StringSplit($ID3BufferArray[$i], "|")
			_ArrayAdd($ID3Array, $ID3FrameArray[1] & "|" & _ID3GetTagField($ID3FrameArray[1]))
			$ID3Array[0] += 1
		Next
		Return $ID3Array
	Elseif $iReturnArray = 2 Then
		;Return a 2D array with a fields of the frames
		Local $ID3FullArray[$ID3BufferArray[0]][5]
		For $i = 1 To $ID3BufferArray[0]
			$ID3FrameArray = StringSplit($ID3BufferArray[$i], "|")
			$ID3FullArray[$i-1][0] = $ID3FrameArray[1]
			$ID3FullArray[$i-1][1] = $ID3FrameArray[2]
			;_ArrayAdd($ID3FullArray,$ID3FrameArray[1] & "|" & _ID3GetTagField($ID3FrameArray[1]))
			;$ID3FullArray[0] += 1
		Next
		Return $ID3FullArray
		;Return $ID3BufferArray
	Else
		Return 1
	EndIf
	;****************************************************************************************
EndFunc


;===============================================================================
; Function Name:    _ID3GetTagField()
; Description:      Reads the ID3 Tag Data from an mp3 file and stores it in a Buffer and returns the Field Requested
; Parameter(s):     $Filename 			- Filename of mp3 file include full path
;					$sFieldIDRequest	- ID3 Field ID String of the Field to return (ie. "TIT2" for ID3v2 Title or "Title" for ID3v1 Title)
; Requirement(s):   None
; Return Value(s):  On Success - Returns Field String. If multiple fields found string is delimited with @CRLF.
;						@error = 0; @extended = number of fields found (ID3v2 can have multiple COMM Fields)
;                   On Failure - Returns Empty String meaning $sFieldIDRequest did not match any IDs in the mp3 File
;						@error = 1; @extended = 0
;===============================================================================
Func _ID3GetTagField($sFieldIDRequest) ;add support for ID3v2.2


 	;Check Buffer; If buffer is empty then set error and return
	If $ID3BufferArray[0] == 0 Then
		SetError(1)
		Return ""
	EndIf

	Local $sFieldID, $sFieldString = "", $TagFound = False, $NumFound = 0

	;get specified field from buffer
	For $i = 1 To $ID3BufferArray[0]
		$aBufferFrameString = StringSplit($ID3BufferArray[$i],"|")
		If $aBufferFrameString[0] > 1 Then
			$sFieldID = $aBufferFrameString[1]

			If $sFieldID == $sFieldIDRequest Then
				$TagFound = True
				$sFieldString = _GetID3FrameString($ID3BufferArray[$i])
;~ 				If $TagFound Then
;~ 					$sFieldString = $sFieldString & @CRLF & $aBufferFrameString[2]
;~ 				Else
;~ 					$sFieldString = $aBufferFrameString[2]
;~ 				EndIf
				;$NumFound += 1
				;$TagFound = True
			EndIf

		EndIf
	Next

	;Set Error and Extended Macros
	SetError(Not ($TagFound), $NumFound)

	Return $sFieldString

EndFunc

;updates the field in the buffer and updates the Frame Size in buffer
Func _ID3SetTagField($sFieldIDRequest, $sFieldValue)

	Local $FrameID = ""
	Local $FrameIDFound = False, $TagSize = 0
	Local $ArrayIndex = 0, $TagSizeIndex = 0

	;Get specified field from buffer
	;****************************************************************************************
	For $i = 1 To $ID3BufferArray[0]
		$aBufferFrameString = StringSplit($ID3BufferArray[$i],"|")
		If $aBufferFrameString[0] > 1 Then
			$FrameID = $aBufferFrameString[1]
			If $FrameID == "TagSize" Then
				$TagSize = Number($aBufferFrameString[2])
				$TagSizeIndex = $i
			EndIf
			If StringCompare($FrameID, $sFieldIDRequest, 1) == 0 Then
				$FrameIDFound = True
				$ArrayIndex = $i
				ExitLoop
			EndIf
		EndIf
	Next
	;MsgBox(0,"$FrameIDFound",$FrameID & " - " & $FrameIDFound)
	;****************************************************************************************

	;Change Frame Data and Size (Need to add all Frame Types)
	;****************************************************************************************
	If Not $FrameIDFound Then
		; this adds a new field with empty data, the data and ID are then added below
		_ArrayAdd($ID3BufferArray, "")
		$ID3BufferArray[0] += 1
		$ArrayIndex = $ID3BufferArray[0]
		;MsgBox(0,"New Field",$FrameID & " - " & $sFieldIDRequest
		$FrameID = $sFieldIDRequest
	EndIf

		If (StringMid($FrameID,1,1) == "T") and (StringLen($FrameID) == 4) and ($FrameID <> "TXXX") Then
;~ 			<Header for 'Text information frame', ID: "T000" - "TZZZ",
;~      		excluding "TXXX" described in 4.2.2.>
;~      		Text encoding                $xx
;~      		Information                  <text string according to encoding>
			$bFrameData = Binary("0x00") ;Text Encoding
		$bFrameData &= StringToBinary($sFieldValue)
		$ID3BufferArray[$ArrayIndex] = $FrameID & "|" & $bFrameData & "|" & BinaryLen($bFrameData)
		Elseif $FrameID == "TXXX" Then
;~ 			<Header for 'User defined text information frame', ID: "TXXX">
;~      		Text encoding     $xx
;~      		Description       <text string according to encoding> $00 (00)
;~      		Value             <text string according to encoding>
		$bFrameData = Binary("0x00") ;Text Encoding
		$bFrameData = Binary("0x00") ;Description
		$bFrameData &= StringToBinary($sFieldValue)
		$ID3BufferArray[$ArrayIndex] = $FrameID & "|" & $bFrameData & "|" & BinaryLen($bFrameData)

		ElseIf (StringMid($FrameID,1,1) == "W") and (StringLen($FrameID) == 4) and ($FrameID <> "WXXX") Then
;~ 			<Header for 'URL link frame', ID: "W000" - "WZZZ", excluding "WXXX"
;~      		described in 4.3.2.>
;~      		URL              <text string>
		$bFrameData = StringToBinary($sFieldValue)
		$ID3BufferArray[$ArrayIndex] = $FrameID & "|" & $bFrameData & "|" & BinaryLen($bFrameData)
	ElseIf $FrameID == "WXXX" Then
;~ 			<Header for 'User defined URL link frame', ID: "WXXX">
;~      		Text encoding     $xx
;~      		Description       <text string according to encoding> $00 (00)
;~      		URL               <text string>
		$bFrameData = Binary("0x00") ;Text Encoding
		$bFrameData = Binary("0x00") ;Description
		$bFrameData &= StringToBinary($sFieldValue)
		$ID3BufferArray[$ArrayIndex] = $FrameID & "|" & $bFrameData & "|" & BinaryLen($bFrameData)

	Else
			
		Switch $FrameID
			Case "COMM"
;~ 					<Header for 'Comment', ID: "COMM">
;~      				Text encoding          $xx
;~      				Language               $xx xx xx
;~      				Short content descrip. <text string according to encoding> $00 (00)
;~      				The actual text        <full text string according to encoding>
				$bFrameData = Binary("0x00") ;Text Encoding
				$bFrameData &= StringToBinary("eng") ;Language (eng ONLY)
				$bFrameData &= Binary("0x00") ;Short content descrip. (NONE)
					if isbinary($sFieldValue) then
					$bFrameData &= $sFieldValue
				Else
					$bFrameData &= StringToBinary($sFieldValue)
				EndIf
				$ID3BufferArray[$ArrayIndex] = $FrameID & "|" & $bFrameData & "|" & BinaryLen($bFrameData)
			Case "APIC"
;~ 					 <Header for 'Attached picture', ID: "APIC">
;~ 						Text encoding      $xx
;~ 						MIME type          <text string> $00
;~ 						Picture type       $xx
;~ 						Description        <text string according to encoding> $00 (00)
;~ 						Picture data       <binary data>
				$bFrameData = Binary("0x00") ;Text Encoding
				$bFrameData &= StringToBinary("image/jpg") & Binary("0x00") ;MIME type (JPEG ONLY)
				$bFrameData &= Binary("0x03") ;Picture type ($03  Cover (front) ONLY)
				$bFrameData &= StringToBinary("Cover Image") & Binary("0x00") ;Description
				$FrameSize = FileGetSize($sFieldValue) + BinaryLen($bFrameData)
				$ID3BufferArray[$ArrayIndex] = $FrameID & "|" & $bFrameData & "|" & $sFieldValue & "|" & $FrameSize
			Case "USLT"
;~ 					<Header for 'Unsynchronised lyrics/text transcription', ID: "USLT">
;~      				Text encoding        $xx
;~      				Language             $xx xx xx
;~      				Content descriptor   <text string according to encoding> $00 (00)
;~      				Lyrics/text          <full text string according to encoding>
				$bFrameData = Binary("0x00") ;Text Encoding
				$bFrameData &= StringToBinary("eng") ;Language (eng ONLY)
				$bFrameData &= Binary("0x00") ;Content descriptor (NONE)
				$FrameSize = FileGetSize($sFieldValue) + BinaryLen($bFrameData)
				$ID3BufferArray[$ArrayIndex] = $FrameID & "|" & $bFrameData & "|" & $sFieldValue & "|" & $FrameSize
			Case "UFID" ;Not Implimented Yet
;~ 					<Header for 'Unique file identifier', ID: "UFID">
;~      				Owner identifier        <text string> $00
;~      				Identifier              <up to 64 bytes binary data>
				Case "Title" Or "Artist" Or "Album" Or "Year" or "Track" Or "Comment" or "Genre"
				$ID3BufferArray[$ArrayIndex] = $FrameID & "|" & $sFieldValue

		EndSwitch

	EndIf
	;****************************************************************************************


EndFunc

Func _ID3RemoveField($sFieldIDRequest)

	Local $FrameID = ""
	Local $FrameIDFound = False, $TagSize = 0
	Local $ArrayIndex = 0, $TagSizeIndex = 0

	;Get specified field from buffer
	;****************************************************************************************
	For $i = 1 To $ID3BufferArray[0]
		$aBufferFrameString = StringSplit($ID3BufferArray[$i],"|")
		If $aBufferFrameString[0] > 1 Then
			$FrameID = $aBufferFrameString[1]
			If $FrameID == "TagSize" Then
				$TagSize = Number($aBufferFrameString[2])
				$TagSizeIndex = $i
			EndIf
			If $FrameID == $sFieldIDRequest Then
				$FrameIDFound = True
				$ArrayIndex = $i
				ExitLoop
			EndIf
		EndIf
	Next
	;****************************************************************************************

	If $FrameIDFound Then
		_ArrayDelete($ID3BufferArray, $ArrayIndex)
		$ID3BufferArray[0] -= 1
	EndIf
	
	
EndFunc


;actually writes the tag
Func _ID3WriteTag($Filename, $iFlag = 0)
	Dim $ZPAD = Binary("0x00")
	Switch $iFlag
		Case 0 ;Wite both ID3v1 and ID3v2 Tags
			$TagFile = StringTrimRight($Filename,4) & "_ID3TAG.mp3"
			_FileCreate($TagFile)

			Local $OldTagSize = 0, $ID3v1Tag = StringToBinary("TAG"), $ID3v1_Artist = ""
			Local $ID3v1_Title = "", $ID3v1_Album = "", $ID3v1_Year = "", $ID3v1_Comment = ""
			Local $ID3v1_Genre = "", $ID3v1_Track = ""

			$hTagFile = Fileopen($TagFile,2) ;erase all

			;Write ID3
			FileWrite($hTagFile, StringToBinary("ID3"))

			;Write Version (Assume ID3v2.3)
			;****************************************************************************************
			FileWrite($hTagFile, Binary("0x" & Hex(Number(3), 2)))
			FileWrite($hTagFile, Binary("0x" & Hex(Number(0), 2)))
			;****************************************************************************************

			;Write Flags (Assume all zero)
			;****************************************************************************************
			FileWrite($hTagFile, Binary("0x00"))
			;****************************************************************************************

			;Calculate new TagSize
			;****************************************************************************************
			Dim $NewTagSize = 0
			For $i = 1 To $ID3BufferArray[0]
				$aBufferFrameString = StringSplit($ID3BufferArray[$i],"|")
				If $aBufferFrameString[0] > 2 Then
					$NewTagSize += Number($aBufferFrameString[$aBufferFrameString[0]]) + 10
				EndIf
				If $aBufferFrameString[1] == "ZPAD" Then
					$NewTagSize += Number($aBufferFrameString[$aBufferFrameString[0]])
				EndIf
			Next
;~ 			MsgBox(0,"$NewTagSize",$NewTagSize)
			;****************************************************************************************

			;Write TagSize (4 byte number)
			;****************************************************************************************
			$sTagSize = $NewTagSize
			$iTagSize = Hex($NewTagSize)
			$bTagSize = _HexToBin_ID3($iTagSize)
			$bTagSize = _StringReverse($bTagSize)
			$bTagSize = StringLeft($bTagSize, 28)
			$TagHeaderBin = StringMid($bTagSize,1,7) & "0" & StringMid($bTagSize,8,7) & "0" & _
					StringMid($bTagSize, 15, 7) & "0" & StringMid($bTagSize, 22, 7) & "0"
			$TagHeaderBin = _StringReverse($TagHeaderBin)
			$TagHeader = _BinToHex_ID3($TagHeaderBin)
			FileWrite($hTagFile, Binary("0x" & $TagHeader))
			;****************************************************************************************

			;Write Frames to File
			;****************************************************************************************
			For $i = 1 To $ID3BufferArray[0]
				$aBufferFrameString = StringSplit($ID3BufferArray[$i], "|")
				;MsgBox(0,"$aBufferFrameString[0]",$aBufferFrameString[0])
				If $aBufferFrameString[1] == "ZPAD" Then
					For $iZPAD = 2 to Number($aBufferFrameString[$aBufferFrameString[0]])
						$ZPAD &= Binary("0x00")
					Next
					$NewTagSize += Number($aBufferFrameString[$aBufferFrameString[0]])
				ElseIf $aBufferFrameString[1] == "TagSize" Then
					$OldTagSize = Number($aBufferFrameString[2])
					;MsgBox(0,"$OldTagSize",$OldTagSize)
				ElseIf $aBufferFrameString[1] == "Title" Then
					$ID3v1_Title = StringToBinary(StringLeft($aBufferFrameString[2], 30))
					for $iPAD = 1 to (30 - BinaryLen($ID3v1_Title))
						$ID3v1_Title &= Binary("0x00")
					Next
				ElseIf $aBufferFrameString[1] == "Artist" Then
					$ID3v1_Artist = StringToBinary(StringLeft($aBufferFrameString[2], 30))
					for $iPAD = 1 to (30 - BinaryLen($ID3v1_Artist))
						$ID3v1_Artist &= Binary("0x00")
					Next
				ElseIf $aBufferFrameString[1] == "Album" Then
					$ID3v1_Album = StringToBinary(StringLeft($aBufferFrameString[2], 30))
					for $iPAD = 1 to (30 - BinaryLen($ID3v1_Album))
						$ID3v1_Album &= Binary("0x00")
					Next
				ElseIf $aBufferFrameString[1] == "Year" Then
					$ID3v1_Year = StringToBinary(StringLeft($aBufferFrameString[2], 4))

				ElseIf $aBufferFrameString[1] == "Comment" Then
					;MsgBox(0,"$aBufferFrameString[2]",$aBufferFrameString[2])
					$ID3v1_Comment = StringToBinary(StringLeft($aBufferFrameString[2], 28)) ;28 bytes for ID3v1.1
					for $iPAD = 1 to (28 - BinaryLen($ID3v1_Comment))
						$ID3v1_Comment &= Binary("0x00")
					Next
				ElseIf $aBufferFrameString[1] == "Track" Then
					$ID3v1_Track = StringLeft($aBufferFrameString[2], 3)
				ElseIf $aBufferFrameString[1] == "Genre" Then
					$ID3v1_Genre = $aBufferFrameString[2]
				EndIf

				If $aBufferFrameString[0] > 2 Then

					If $aBufferFrameString[1] == "APIC" Then

						FileWrite($hTagFile, StringToBinary($aBufferFrameString[1]))
						FileWrite($hTagFile, Binary("0x" & Hex($aBufferFrameString[$aBufferFrameString[0]], 8))) ;FrameSize
						FileWrite($hTagFile, Binary("0x" & Hex(0, 2))) ;FrameFlag1
						FileWrite($hTagFile, Binary("0x" & Hex(0, 2))) ;FrameFlag2
						FileWrite($hTagFile, Binary($aBufferFrameString[2]))
						$PicFile_h = FileOpen($aBufferFrameString[3], 16) ;force binary
						$WriteError = FileWrite($hTagFile,FileRead($PicFile_h))
						FileClose($PicFile_h)


					ElseIf $aBufferFrameString[1] == "USLT" Then

						FileWrite($hTagFile, StringToBinary($aBufferFrameString[1]))
						FileWrite($hTagFile, Binary("0x" & Hex($aBufferFrameString[$aBufferFrameString[0]], 8))) ;FrameSize
						FileWrite($hTagFile, Binary("0x" & Hex(0, 2))) ;FrameFlag1
						FileWrite($hTagFile, Binary("0x" & Hex(0, 2))) ;FrameFlag2
						FileWrite($hTagFile, Binary($aBufferFrameString[2]))
						$LyricsFile_h = FileOpen($aBufferFrameString[3], 16) ;force binary
						$LyricData = FileRead($LyricsFile_h)
						$WriteError = FileWrite($hTagFile,$LyricData)
						FileClose($LyricsFile_h)

					Else

						FileWrite($hTagFile, StringToBinary($aBufferFrameString[1]))
						FileWrite($hTagFile, Binary("0x" & Hex($aBufferFrameString[$aBufferFrameString[0]], 8))) ;FrameSize
						FileWrite($hTagFile, Binary("0x" & Hex(0, 2))) ;FrameFlag1
						FileWrite($hTagFile, Binary("0x" & Hex(0, 2))) ;FrameFlag2
						FileWrite($hTagFile, Binary($aBufferFrameString[2])) ;Frame Data

					EndIf
					$NewTagSize += Number($aBufferFrameString[$aBufferFrameString[0]]) + 10

				EndIf
			Next
			;****************************************************************************************

			;Write Tag Padding
			;****************************************************************************************
			FileWrite($hTagFile, $ZPAD)
			;****************************************************************************************

			;Write MP3 Data
			;****************************************************************************************
			$hFile = FileOpen($Filename,16) 		;read force binary
			FileRead($hFile, $OldTagSize) ;read past Old Tag
			FileWrite($hTagFile, FileRead($hFile, FileGetSize($Filename) - 128 - $OldTagSize)) ;Write the mp3 data
			;****************************************************************************************

			;Write ID3v1.1 Tag
			;****************************************************************************************
			$ID3v1Tag &= $ID3v1_Title
			$ID3v1Tag &= $ID3v1_Artist
			$ID3v1Tag &= $ID3v1_Album
			$ID3v1Tag &= $ID3v1_Year
			$ID3v1Tag &= $ID3v1_Comment
			$ID3v1Tag &= Binary("0x00")
			$ID3v1Tag &= Binary("0x" & Hex(Number($ID3v1_Track), 2))
			$ID3v1Tag &= Binary("0x" & Hex(_GetGenreID($ID3v1_Genre), 2))

			;check tag as proper length
			;MsgBox(0,"ID3v1 Tag Length = ", BinaryLen($ID3v1Tag))
			FileWrite($hTagFile, $ID3v1Tag)
			;****************************************************************************************

			FileClose($hFile)
			FileClose($hTagFile)

			FileCopy($TagFile, $Filename, 1)
			FileDelete($TagFile)
	EndSwitch



EndFunc


;===============================================================================
; Function Name:    _ID3DeleteFiles()
; Description:      Deletes any files created by ID3.au3 (ie. AlbumArt.jpeg and SongLyrics.txt)
; Parameter(s):     None
; Requirement(s):   None
; Return Value(s):  On Success - Returns 1.
;                   On Failure - Returns 0
;===============================================================================
Func _ID3DeleteFiles()

	If $ID3Filenames == "" Then Return 1
	$aID3File = StringSplit($ID3Filenames,"|")
	For $i = 1 To $aID3File[0]
		If FileExists($aID3File[$i]) Then
			$ret = FileDelete($aID3File[$i])
			If $ret == 0 Then Return 0
		EndIf
	Next
	$ID3Filenames = ""

	Return 1

EndFunc

Func _ID3GetFrameDescription($sFrameID)
	Local $sFrameD = ""
	Switch $sFrameID
	Case "UFID"
		$sFrameD = "This frame's purpose is to be able to identify the audio file in a " & _
		"database that may contain more information relevant to the content. " & _
		"Since standardisation of such a database is beyond this document, all " & _
		"frames begin with a null-terminated string with a URL [URL] " & _
		"containing an email address, or a link to a location where an email " & _
		"address can be found, that belongs to the organisation responsible " & _
		"for this specific database implementation. Questions regarding the " & _
		"database should be sent to the indicated email address. The URL " & _
		"should not be used for the actual database queries. The string " & _
		'"http://www.id3.org/dummy/ufid.html" should be used for tests. ' & _
		"Software that isn't told otherwise may safely remove such frames. The " & _
		"'Owner identifier' must be non-empty (more than just a termination). " & _
		"The 'Owner identifier' is then followed by the actual identifier, " & _
		'which may be up to 64 bytes. There may be more than one "UFID" frame ' & _
		"in a tag, but only one with the same 'Owner identifier'."
	Case "TALB"
		$sFrameD = "The 'Album/Movie/Show title' frame is intended for the title of the " & _
		"recording(/source of sound) which the audio in the file is taken " & _
		"from."
	Case "TBPM"
		$sFrameD = "The 'BPM' frame contains the number of beats per minute in the " & _
		"mainpart of the audio. The BPM is an integer and represented as a " & _
		"numerical string."
	EndSwitch
	
	Return $sFrameD
	
EndFunc
; -----------------------------------------------------------------------------



;-------------------  Main File Reading Functions -----------------------------
Func _ReadID3v2($Filename, ByRef $aID3V2Tag, $sFilter = -1)

	Local $ZPAD = 0, $BytesRead = 0
	Local $hFile = FileOpen($Filename, 16) ;mode = Force binary
	Local $iFilterNum

	If $sFilter <> -1 Then
		$aFilter = StringSplit($sFilter,"|")
		$iFilterNum = $aFilter[0]
	EndIf

	;Read Header Bytes (First 10 Bytes of File)
	;****************************************************************************************
	Local $ID3v2Header = FileRead($hFile, 10)
	_ArrayAdd($aID3V2Tag, "Header" & "|" & $ID3v2Header)
	$aID3V2Tag[0] += 1
	$BytesRead = 10 ;ID3 Header has been read (Header Size = 10 Bytes)
;~ 	MsgBox(0,"$ID3v2Header",$ID3v2Header)
	;****************************************************************************************

	;Read "ID3" string
	;****************************************************************************************
	If Not (BinaryToString(BinaryMid($ID3v2Header, 1, 3)) == "ID3") Then
		FileClose($hFile)
		SetError(1)
		;MsgBox(0,"ID3 string not found in Header",BinaryToString(BinaryMid($ID3v2Header,1,3)))
		Return 0
	EndIf
	;****************************************************************************************

	;GetTagVer
	;****************************************************************************************
	Local $FrameIDLen ;equals 3 or 4 depending on ID3v2.2 or ID3v2.3+
	Local $ID3v2Version = String(Number(BinaryMid($ID3v2Header, 4, 1))) & "." & String(Number(BinaryMid($ID3v2Header, 5, 1)))
	If $sFilter == -1 Then
		_ArrayAdd($aID3V2Tag,"Version" & "|" & "ID3v2." & $ID3v2Version)
		$aID3V2Tag[0] += 1
	EndIf
;~ 	MsgBox(0,"$ID3v2Version",$ID3v2Version)
	If StringInStr($ID3v2Version, "2.") Then
		$FrameIDLen = 3
		If $sFilter <> -1 Then
			_ConvertFilterToID3v2_2($sFilter)
		EndIf
	Else
		$FrameIDLen = 4
	EndIf
	;****************************************************************************************

	;GetTagFlags
	;****************************************************************************************
	Local $TagFlagsBin = BinaryMid($ID3v2Header, 6, 1)
	Local $TagFlags = _HexToBin_ID3(StringTrimLeft($TagFlagsBin, 2))
	Local $Unsynchronisation = StringMid($TagFlags,1,1)
	Local $ExtendedHeader = StringMid($TagFlags,2,1)
	Local $ExperimentalIndicator = StringMid($TagFlags,3,1)
	Local $Footer = StringMid($TagFlags,4,1)
	If Not $TagFlags == "00000000" Then
		;MsgBox(0,"$ID3TagFlags", $TagFlags)
	EndIf
	If $sFilter == -1 Then
		_ArrayAdd($aID3V2Tag,"Unsynchronisation" & "|" & $Unsynchronisation)
		_ArrayAdd($aID3V2Tag,"ExtendedHeader" & "|" & $ExtendedHeader)
		_ArrayAdd($aID3V2Tag,"ExperimentalIndicator" & "|" & $ExperimentalIndicator)
		_ArrayAdd($aID3V2Tag,"Footer" & "|" & $Footer)
		$aID3V2Tag[0] += 4
	EndIf
	;****************************************************************************************


	;GetTagSize
	;****************************************************************************************
	Local $TagSizeBin = ""
	$TagHeaderBin = _HexToBin_ID3(StringTrimLeft(BinaryMid($ID3v2Header,7,4),2)) ;Removes 0x
	For $i = 1 To 33 Step 8
		$TagSizeBin &= StringMid($TagHeaderBin, $i + 1, 7) ;Skip most significant bit in every byte
	Next
	Local $TagSize = Dec(_BinToHex_ID3($TagSizeBin)) + 10 ;add 10 to include header
	If ($sFilter == -1) or StringInStr($sFilter,"TagSize") Then
		_ArrayAdd($aID3V2Tag, "TagSize" & "|" & $TagSize)
		$aID3V2Tag[0] += 1
		If $aID3V2Tag[0] == ($iFilterNum + 1) Then ;Plus one needed to account for Header
			FileClose($hFile)
			Return 1
		EndIf
	EndIf
;~ 	MsgBox(0,"$TagSize",$TagSize)
	;****************************************************************************************

	;Read in all of the tag and close the file might be faster
	Local $ZPadding, $FrameIDFristHex, $FrameID, $FrameSizeHex, $FrameSize, $FrameFlag1, $FrameFlag2, $FoundTag, $index
	Local $FrameHeader

	;Get Rest Of Tag
	While $BytesRead < $TagSize

		;Initialize ZPAD & Read first Hex value to test for new Frame or NULL Bytes
		;****************************************************************************************
		$ZPadding = 0
		$FrameIDFristHex = StringTrimLeft(FileRead($hFile, 1), 2)
		$BytesRead += 1
;~ 		MsgBox(0,"Chr(Dec($FrameIDFristHex))",Chr(Dec($FrameIDFristHex)))
		;****************************************************************************************

		;check for NULL
		If $FrameIDFristHex == "00" Then
			;Count how many NULL Bytes at the end of the Tag then Exit Loop
			;****************************************************************************************
			$ZPadding += 1
			While $FrameIDFristHex == "00"
				$FrameIDFristHex = StringTrimLeft(FileRead($hFile, 1), 2)
				$BytesRead += 1
				$ZPadding += 1
				If $BytesRead >= $TagSize Then
					ExitLoop
				EndIf
			WEnd
			$ZPAD = $ZPadding
			ExitLoop
			;****************************************************************************************

		Else

			;Read Frame ID String
			;****************************************************************************************
			$FrameID = Chr(Dec($FrameIDFristHex)) & BinaryToString(FileRead($hFile, $FrameIDLen - 1))
			$BytesRead += $FrameIDLen - 1
			;****************************************************************************************

			;Check for a valid frameID string
			If StringIsAlNum($FrameID) Then

				;Read Frame Header & Size (Differant for 2.2 and 2.3+)
				;****************************************************************************************
				If $FrameIDLen == 4 Then
					$bFrameHeader = FileRead($hFile,6)
					$BytesRead += 6
					$FrameSizeHex = StringTrimLeft(BinaryMid($bFrameHeader, 1, 4), 2)
					$FrameSize = _HexToUint32_ID3($FrameSizeHex)
					$FrameFlag1 = _HexToBin_ID3(StringTrimLeft(BinaryMid($bFrameHeader,5,1),2))
					$FrameFlag2 = _HexToBin_ID3(StringTrimLeft(BinaryMid($bFrameHeader,6,1),2))
				ElseIf $FrameIDLen == 3 Then
					$bFrameHeader = FileRead($hFile, 3) ;only frame size
					$BytesRead += 3
					$FrameSizeHex = StringTrimLeft(BinaryMid($bFrameHeader, 1, 3), 2)
					$FrameSize = Dec($FrameSizeHex)
				EndIf
;~ 				MsgBox(0,$FrameID,$FrameID)
				;****************************************************************************************


				;Add FrameID and Frame Data to array
				;****************************************************************************************
				If $sFilter == -1 Then
					Switch $FrameID
						Case "APIC" ;Picture (Write to File)
							$FrameData = _GetAlbumArt($hFile, $FrameSize)
						Case "USLT" ;Lyrics (Write to File)
							$FrameData = _GetSongLyrics($hFile, $FrameSize)
						Case Else
							$FrameData = FileRead($hFile, $FrameSize)
					EndSwitch
					_ArrayAdd($aID3V2Tag, $FrameID & "|" & $FrameData & "|" & $FrameSize)
					$aID3V2Tag[0] += 1
					$BytesRead += $FrameSize
				Else
					If $aID3V2Tag[0] == ($iFilterNum + 1) Then ;Plus one needed to account for Header
						ExitLoop
					EndIf
					If StringInStr($sFilter, $FrameID) Then
						Switch $FrameID
							Case "APIC" ;Picture (Write to File)
								$FrameData = _GetAlbumArt($hFile, $FrameSize)
							Case "USLT" ;Lyrics (Write to File)
								$FrameData = _GetSongLyrics($hFile, $FrameSize)
							Case Else
								$FrameData = FileRead($hFile, $FrameSize)
						EndSwitch
						_ArrayAdd($aID3V2Tag, $FrameID & "|" & $FrameData & "|" & $FrameSize)
						$aID3V2Tag[0] += 1
					Else
						FileRead($hFile, $FrameSize)
					EndIf
					$BytesRead += $FrameSize
				EndIf
				;****************************************************************************************
			EndIf
		EndIf
	WEnd

	;Read MPEG Header & Write to Array MPEG & ZPAD
	;****************************************************************************************
	If ($sFilter == -1) or StringInStr($sFilter,"MPEG") or StringInStr($sFilter,"ZPAD") Then
		Local $MPEGHeaderCheck = $FrameIDFristHex & StringTrimLeft(FileRead($hFile, 50), 2)
		Local $index = StringInStr($MPEGHeaderCheck, "FF")
		Local $MPEGHeaderHex = StringMid($MPEGHeaderCheck, $index, 8)
		;check MPEG Header
		If _CheckMPEGHeader($MPEGHeaderHex) and (StringInStr($sFilter,"MPEG") or ($sFilter == -1)) Then
			_ArrayAdd($aID3V2Tag, "MPEG" & "|" & $MPEGHeaderHex)
			$aID3V2Tag[0] += 1
		EndIf
		If StringInStr($sFilter,"ZPAD") or ($sFilter == -1) Then
			_ArrayAdd($aID3V2Tag, "ZPAD" & "|" & $ZPAD)
			$aID3V2Tag[0] += 1
		EndIf
	EndIf
	;****************************************************************************************

	FileClose($hFile)
	Return 1

EndFunc

Func _ReadID3v1($Filename, ByRef $aID3V1Tag)

	Local $hfile = FileOpen($Filename,16) ;open in binary mode
	FileRead($hfile,FileGetSize($Filename)-128)
	Local $ID3v1Tag = FileRead($hfile)
	FileClose($hfile)

	Local $ID3v1ID = BinaryToString(BinaryMid($ID3v1Tag, 1, 3))
	;MsgBox(0,"ID3v1 Tag ID",$ID3v1ID)
	If Not ($ID3v1ID == "TAG") Then
		FileClose($hfile)
		SetError(-1)
		Return 0
	EndIf

	Local $Title, $Artist, $Album, $Year, $Comment, $Track, $GenreID, $Genre

	$Title = BinaryToString(BinaryMid($ID3v1Tag, 4, 30))
	_ArrayAdd($aID3V1Tag, "Title" & "|" & $Title)
	$aID3V1Tag[0] += 1

	$Artist = BinaryToString(BinaryMid($ID3v1Tag, 34, 30))
	_ArrayAdd($aID3V1Tag, "Artist" & "|" & $Artist)
	$aID3V1Tag[0] += 1

	$Album = BinaryToString(BinaryMid($ID3v1Tag, 64, 30))
	_ArrayAdd($aID3V1Tag, "Album" & "|" & $Album)
	$aID3V1Tag[0] += 1
	;MsgBox(0,"$Album",$Album)
	$Year = BinaryToString(BinaryMid($ID3v1Tag, 94, 4))
	_ArrayAdd($aID3V1Tag, "Year" & "|" & $Year)
	$aID3V1Tag[0] += 1
	;MsgBox(0,"$Year",$Year)
	$Comment = BinaryToString(BinaryMid($ID3v1Tag, 98, 28))
	_ArrayAdd($aID3V1Tag, "Comment" & "|" & $Comment)
	$aID3V1Tag[0] += 1
	;MsgBox(0,"$Comment",$Comment)
	$Track = Dec(StringTrimLeft(BinaryMid($ID3v1Tag, 126, 2), 2))
;~ 	MsgBox(0,"$Track",$Track)
	If $Track < 1000 And $Track > 0 Then
		_ArrayAdd($aID3V1Tag, "Track" & "|" & $Track)
		$aID3V1Tag[0] += 1
	Else
		_ArrayAdd($aID3V1Tag, "Track" & "|" & "")
		$aID3V1Tag[0] += 1
	EndIf

	$GenreID = Dec(StringTrimLeft(BinaryMid($ID3v1Tag, 128, 1), 2))
	$Genre = _GetGenreByID($GenreID)
	_ArrayAdd($aID3V1Tag, "Genre" & "|" & $Genre)
	$aID3V1Tag[0] += 1

	If $Track == 0 Then
		_ArrayAdd($aID3V1Tag,"Version1" & "|" & "ID3v1.0")
		$aID3V1Tag[0] += 1
	Else
		_ArrayAdd($aID3V1Tag,"Version1" & "|" & "ID3v1.1")
		$aID3V1Tag[0] += 1
	EndIf

	FileClose($hfile)
	Return 1

EndFunc

Func _GetID3FrameString($sFrameData)
	
	Local $bTagFieldFound = False, $sFrameString = ""
	Local $bText_Encoding_Description_Byte
	Local $bUnicode_BOM 
	Local $BinaryToString_Flag = 1 ;Default binary data is taken to be ANSI
;~ 								2, binary data is taken to be UTF16 Little Endian (BOM = FF FE)
;~ 								3, binary data is taken to be UTF16 Big Endian (BOM = FE FF)
;~ 								4, binary data is taken to be UTF8 (BOM = EF BB BF)

	$aFrameData = StringSplit($sFrameData,"|")
	$FrameID = $aFrameData[1]

	If (StringMid($FrameID,1,1) == "T") and (StringLen($FrameID) == 4) and ($FrameID <> "TXXX") and ($FrameID <> "TCON") Then
		$bTagFieldFound = True
		$bFrameData = Binary($aFrameData[2])
		$bText_Encoding_Description_Byte = Int(BinaryMid($bFrameData,1,1))
		;MsgBox(0,"$bText_Encoding_Description_Byte",$bText_Encoding_Description_Byte)
		If $bText_Encoding_Description_Byte <> 0 Then
			$bUnicode_BOM = BinaryMid($bFrameData,2,2)
			If $bUnicode_BOM = "0xFFFE" Then
				$BinaryToString_Flag = 2 ;UTF16 Little Endian
			EndIf
			If $bUnicode_BOM = "0xFEFF" Then
				$BinaryToString_Flag = 3 ;UTF16 Big Endian
			EndIf
			$sFrameString = BinaryToString(BinaryMid($bFrameData,4),$BinaryToString_Flag)
		Else
		$sFrameString = BinaryToString(BinaryMid($bFrameData, 2))
		EndIf
	Elseif $FrameID == "TXXX" Then
		;User defined text information frame
		$bTagFieldFound = True
		$bFrameData = Binary($aFrameData[2])
		$bText_Encoding_Description_Byte = BinaryMid($bFrameData,1,1)
		$ByteIndex = 2
				
		$Description = ""
		$Byte = ""
				
		$BinaryToString_Flag = 1
		If $bText_Encoding_Description_Byte <> 0 Then
			If BinaryMid($bFrameData,$ByteIndex,2) = "0xFFFE" Then
				$BinaryToString_Flag = 2 ;UTF16 Little Endian
				$ByteIndex += 2
			EndIf
			If BinaryMid($bFrameData,$ByteIndex,2) = "0xFEFF" Then
				$BinaryToString_Flag = 3 ;UTF16 Big Endian
				$ByteIndex += 2
			EndIf
		EndIf
				
		While BinaryMid($bFrameData,$ByteIndex,1) <> "0x00"
			;$Short_Content_Descrip &= BinaryToString($Byte)
			$Byte = binary($Byte & BinaryMid($bFrameData,$ByteIndex,1))
		$ByteIndex += 1
		WEnd
		While BinaryMid($bFrameData,$ByteIndex,1) == "0x00"
			$ByteIndex += 1
		WEnd
		$Description = BinaryToString($Byte,$BinaryToString_Flag)
		;MsgBox(0,"$Short_Content_Descrip",$Short_Content_Descrip)
				
		$BinaryToString_Flag = 1
		If $bText_Encoding_Description_Byte <> 0 Then
			If BinaryMid($bFrameData,$ByteIndex,2) = "0xFFFE" Then
				$BinaryToString_Flag = 2 ;UTF16 Little Endian
			EndIf
			If BinaryMid($bFrameData,$ByteIndex,2) = "0xFEFF" Then
				$BinaryToString_Flag = 3 ;UTF16 Big Endian
			EndIf
			$ByteIndex += 2
		EndIf
		$sFrameString = BinaryToString(BinaryMid($bFrameData,$ByteIndex),$BinaryToString_Flag)
		If $Description <> "" then
			$sFrameString = $Description & "-" & $sFrameString
		EndIf
;~ 		MsgBox(0,"$sFrameString",$sFrameString)
	Elseif $FrameID == "TCON" Then;Content Type/Genre
		$bTagFieldFound = True
		$bFrameData = Binary($aFrameData[2])
		$bText_Encoding_Description_Byte = Int(BinaryMid($bFrameData,1,1))
		;MsgBox(0,"$bText_Encoding_Description_Byte",$bText_Encoding_Description_Byte)
		If $bText_Encoding_Description_Byte <> 0 Then
			$bUnicode_BOM = BinaryMid($bFrameData,2,2)
			If $bUnicode_BOM = "0xFFFE" Then
				$BinaryToString_Flag = 2 ;UTF16 Little Endian
			EndIf
			If $bUnicode_BOM = "0xFEFF" Then
				$BinaryToString_Flag = 3 ;UTF16 Big Endian
			EndIf
			$Genre = BinaryToString(BinaryMid($bFrameData,4),$BinaryToString_Flag)
		Else
			$Genre = BinaryToString(BinaryMid($bFrameData,2))
		EndIf
		;$Genre = BinaryToString($bFrameData)
		
		If StringMid($Genre, 1, 1) == "(" Then ;check if first char is "("
			$closeparindex = StringInStr($Genre,")")
			$GenreID = StringMid($Genre,2,$closeparindex-1)
			$Genre = _GetGenreByID($GenreID)
		EndIf ;If no "(" then return the whole field as is
		$sFrameString = $Genre

	ElseIf (StringMid($FrameID,1,1) == "W") and (StringLen($FrameID) == 4) and ($FrameID <> "WXXX") Then
		$bTagFieldFound = True
		$bFrameData = Binary($aFrameData[2])
		$sFrameString = BinaryToString($bFrameData)
		MsgBox(0,"$sFrameString",$sFrameString)
	ElseIf $FrameID == "WXXX" Then
		;User Defined URL
		$bTagFieldFound = True
		$bFrameData = Binary($aFrameData[2])
		$bText_Encoding = BinaryMid($bFrameData,1,1)
		$ByteIndex = 2
		$Description = ""
		$Byte = BinaryMid($bFrameData, $ByteIndex, 1)
		$ByteIndex += 1
		While $Byte <> "0x00"
			$Description &= BinaryToString($Byte)
			$Byte = BinaryMid($bFrameData, $ByteIndex, 1)
			$ByteIndex += 1
		WEnd
		$sFrameString = BinaryToString(BinaryMid($bFrameData, $ByteIndex))

	Else

		Switch $FrameID
			Case "COMM", "COM" ;Comment
				$bTagFieldFound = True
				$bFrameData = Binary($aFrameData[2])
				$bText_Encoding_Description_Byte = BinaryMid($bFrameData,1,1)
				$Language = BinaryToString(BinaryMid($bFrameData,2,3))
				if $Language <> "eng" Then
					;MsgBox(0,"Text may not be in English",$Language)
					;text may not be in English
				EndIf
				$ByteIndex = 5

				$Short_Content_Descrip = ""
				$Byte = ""
				
				$BinaryToString_Flag = 1
				If $bText_Encoding_Description_Byte <> 0 Then
					If BinaryMid($bFrameData,$ByteIndex,2) = "0xFFFE" Then
						$BinaryToString_Flag = 2 ;UTF16 Little Endian
						$ByteIndex += 2
					EndIf
					If BinaryMid($bFrameData,$ByteIndex,2) = "0xFEFF" Then
						$BinaryToString_Flag = 3 ;UTF16 Big Endian
						$ByteIndex += 2
					EndIf
				EndIf
				
				While BinaryMid($bFrameData,$ByteIndex,1) <> "0x00"
					;$Short_Content_Descrip &= BinaryToString($Byte)
					$Byte = binary($Byte & BinaryMid($bFrameData,$ByteIndex,1))
					$ByteIndex += 1
				WEnd
				While BinaryMid($bFrameData,$ByteIndex,1) == "0x00"
					$ByteIndex += 1
				WEnd
				$Short_Content_Descrip = BinaryToString($Byte,$BinaryToString_Flag)
				;MsgBox(0,"$Short_Content_Descrip",$Short_Content_Descrip)
				
				$BinaryToString_Flag = 1
				If $bText_Encoding_Description_Byte <> 0 Then
					If BinaryMid($bFrameData,$ByteIndex,2) = "0xFFFE" Then
						$BinaryToString_Flag = 2 ;UTF16 Little Endian
					EndIf
					If BinaryMid($bFrameData,$ByteIndex,2) = "0xFEFF" Then
						$BinaryToString_Flag = 3 ;UTF16 Big Endian
					EndIf
					$ByteIndex += 2
				EndIf
				$sFrameString = BinaryToString(BinaryMid($bFrameData,$ByteIndex),$BinaryToString_Flag)
				If $Short_Content_Descrip <> "" then
					$sFrameString = $Short_Content_Descrip & "-" & $sFrameString
				EndIf
;~ 				MsgBox(0,"$sFrameString",$sFrameString)
			Case "APIC"
				;Text encoding $xx
				;MIME type <text string> $00
				;Picture type $xx
				;Description <text string according to encoding> $00 (00)
				;Picture data <binary data>
				$bTagFieldFound = True
				$sFrameString = $aFrameData[3]
				;MsgBox(0,"$sFrameString",$sFrameString)
			Case "USLT"
				;Text encoding $xx
				;Language $xx xx xx
				;Content descriptor <text string according to encoding> $00 (00)
				;Lyrics/text <full text string according to encoding>
				$bTagFieldFound = True
				$sFrameString = $aFrameData[3]
			Case "UFID"
				;Owner identifier <text string> $00
				;Identifier <up to 64 bytes binary data>
				$bTagFieldFound = True
				$bFrameData = Binary($aFrameData[2])
				$sFrameString = BinaryToString(BinaryMid($bFrameData, 2))
			Case "Artist" ;ID3v1
				$bTagFieldFound = True
				$sFrameString = $aFrameData[2]
			Case "Title" ;ID3v1
				$bTagFieldFound = True
				$sFrameString = $aFrameData[2]
			Case "Album" ;ID3v1
				$bTagFieldFound = True
				$sFrameString = $aFrameData[2]
			Case "Track" ;ID3v1
				$bTagFieldFound = True
				$sFrameString = $aFrameData[2]
			Case "Year" ;ID3v1
				$bTagFieldFound = True
				$sFrameString = $aFrameData[2]
			Case "Genre" ;ID3v1
				$bTagFieldFound = True
				$sFrameString = $aFrameData[2]
			Case "Comment" ;ID3v1
				$bTagFieldFound = True
				$sFrameString = $aFrameData[2]
			Case "ZPAD" ;ID3v2
				$bTagFieldFound = True
				$sFrameString = $aFrameData[2]
			Case "MPEG" ;MPEG
				$bTagFieldFound = True
				$sFrameString = $aFrameData[2]
			Case "NCON" ;??http://www.id3.org/Compliance_Issues --- non-standard NCON frame in tag
				$bTagFieldFound = True
				$sFrameString &= "Non-Standard NCON Frame"
			Case "PRIV" ;Private frame
				$bTagFieldFound = True
;~ 				$Owner_Identifier = FileRead($hFile,1)
;~ 				$LengthToRead -= 1
;~ 				$FrameString &= FileRead($hFile,$LengthToRead)
;~ 			Case "PCNT", "CNT" ;Play counter
;~ 				$bTagFieldFound = True
;~ 				$FrameString &= _HexToUint32_ID3(StringTrimLeft($aFrameData[2],2)) ;Not Tested
;~ 			Case "RVA" ;Relative volume adjustment ID3v2.2 only
;~ 				$bTagFieldFound = True
;~ 				$FrameString &= "TBP - Relative Volume Adjustment" ;Not Tested
;~ 				FileRead($hFile,$LengthToRead)
			Case Else
				;MsgBox(0,"$FrameID Unknown",$FrameID)
				$sFrameString &= "Undefined FrameID"

		EndSwitch

	EndIf


	If $bTagFieldFound == False Then
		SetError(1)
	EndIf

	Return $sFrameString

EndFunc

;--------------------------  Support Functions --------------------------------

Func _CheckMPEGHeader($MPEGFrameSyncHex)

	$MPEGFrameSyncUint32 = _HexToUint32_ID3($MPEGFrameSyncHex)

	If $MPEGFrameSyncUint32 > _HexToUint32_ID3("FFE00000") Then
		If $MPEGFrameSyncUint32 < _HexToUint32_ID3("FFFFEC00") Then
			If Not (StringMid($MPEGFrameSyncHex, 4, 1) == "0") Then
				If Not (StringMid($MPEGFrameSyncHex, 4, 1) == "1") Then
					If Not (StringMid($MPEGFrameSyncHex, 4, 1) == "9") Then
						;valid MPEG Header Found
						Return 1
					EndIf
				EndIf
			EndIf
		EndIf
	EndIf

	Return 0

EndFunc

Func _GetAlbumArt($hFile, $FrameLen)

	Local $LengthToRead = $FrameLen, $AlbumArtFilename = @ScriptDir & "\" & "AlbumArt", $bReturn

	$bText_Encoding = FileRead($hFile,1)
	$LengthToRead -= 1

	$MIME_Type = ""
	$Byte = FileRead($hFile,1)
	$bMIME_Type = $Byte
	$LengthToRead -= 1
	While $Byte <> "0x00"
		$MIME_Type &= BinaryToString($Byte)
		$Byte = FileRead($hFile, 1)
		$bMIME_Type &= $Byte
		$LengthToRead -= 1
	WEnd
;~ 	MsgBox(0,"$bMIME_Type",$bMIME_Type)

	$bPicture_Type = FileRead($hFile,1)
	$LengthToRead -= 1

	$Description = ""
	$Byte = FileRead($hFile, 1)
	$bDescription = $Byte
	$LengthToRead -= 1
	While $Byte <> "0x00"
		$Description &= BinaryToString($Byte)
		$Byte = FileRead($hFile, 1)
		$bDescription &= $Byte
		$LengthToRead -= 1
	WEnd

;~ 	MsgBox(0,"$Description",$Description)

	If StringInStr($MIME_Type, "jpg") Or StringInStr($MIME_Type, "jpeg") Then
		$AlbumArtFilename &= ".jpg"
		$ID3Filenames &= $AlbumArtFilename & "|"
	ElseIf StringInStr($MIME_Type, "png") Then
		$AlbumArtFilename &= ".png"
		$ID3Filenames &= $AlbumArtFilename & "|"
	Else
		$AlbumArtFilename = "File Type Unknown"
	EndIf

	;Read Picture data to file
	;****************************************************************************************
	$PicFile_h = FileOpen($AlbumArtFilename, 2) ;Open for write and erase existing
	$WriteError = FileWrite($PicFile_h,FileRead($hFile, $LengthToRead))
	FileClose($PicFile_h)
	;****************************************************************************************

	$bReturn = $bText_Encoding & $bMIME_Type
	$bReturn &= $bPicture_Type & $bDescription
	$bReturn &= "|" & $AlbumArtFilename

	Return $bReturn

EndFunc

Func _GetSongLyrics($hFile, $FrameLen)

	Local $LengthToRead = $FrameLen, $LyricsFilename = @ScriptDir & "\" & "SongLyrics.txt", $bReturn
	$ID3Filenames &= $LyricsFilename & "|"

	$bText_Encoding = FileRead($hFile,1)
	$LengthToRead -= 1

	$bLanguage = FileRead($hFile,3)
	$LengthToRead -= 3

	$Content_Descriptor = ""
	$Byte = FileRead($hFile,1)
	$bContent_Descriptor = $Byte
	$LengthToRead -= 1
	While $Byte <> "0x00"
		$Content_Descriptor &= BinaryToString($Byte)
		$Byte = FileRead($hFile, 1)
		$bContent_Descriptor &= $Byte
		$LengthToRead -= 1
	WEnd

	$bLyrics_Text = FileRead($hFile,$LengthToRead)
	$Lyrics_Text = BinaryToString($bLyrics_Text)

	$hLyricFile = FileOpen($LyricsFilename, 2) ;Open for write and erase existing
	FileWrite($hLyricFile, $Lyrics_Text)
	FileClose($hLyricFile)

	$bReturn = $bText_Encoding & $bLanguage & $bContent_Descriptor
	$bReturn &= "|" & $LyricsFilename

	Return $bReturn

EndFunc

Func _ConvertFilterToID3v2_2(ByRef $sFilter)
	
	
	$sFilter = StringReplace($sFilter, "TIT1", "TT1")
	$sFilter = StringReplace($sFilter, "TIT2", "TT2")
	$sFilter = StringReplace($sFilter, "TIT3", "TT3")
	$sFilter = StringReplace($sFilter, "TEXT", "TXT")
	$sFilter = StringReplace($sFilter, "TLAN", "TLA")
	$sFilter = StringReplace($sFilter, "TKEY", "TKE")
	$sFilter = StringReplace($sFilter, "TMED", "TMT")
	$sFilter = StringReplace($sFilter, "TOAL", "TOT")
	$sFilter = StringReplace($sFilter, "TOFN", "TOF")
	$sFilter = StringReplace($sFilter, "TOLY", "TOL")
	$sFilter = StringReplace($sFilter, "TOPE", "TOA")
	$sFilter = StringReplace($sFilter, "TORY", "TOR")
	$sFilter = StringReplace($sFilter, "TPE1", "TP1")
	$sFilter = StringReplace($sFilter, "TPE2", "TP2")
	$sFilter = StringReplace($sFilter, "TPE3", "TP3")
	$sFilter = StringReplace($sFilter, "TPE4", "TP4")
	$sFilter = StringReplace($sFilter, "TPOS", "TPA")
	$sFilter = StringReplace($sFilter, "TALB", "TRK")
	$sFilter = StringReplace($sFilter, "TRCK", "TP2")
	$sFilter = StringReplace($sFilter, "TYER", "TYE")
	$sFilter = StringReplace($sFilter, "COMM", "COM")
	$sFilter = StringReplace($sFilter, "APIC", "PIC")
	$sFilter = StringReplace($sFilter, "USLT", "ULT")
	$sFilter = StringReplace($sFilter, "TSSE", "TSS")
	$sFilter = StringReplace($sFilter, "TENC", "TEN")
	$sFilter = StringReplace($sFilter, "TCOP", "TCR")
	$sFilter = StringReplace($sFilter, "TBPM", "TBP")
	$sFilter = StringReplace($sFilter, "TRDA", "TRD")
	$sFilter = StringReplace($sFilter, "TSIZ", "TSI")
	$sFilter = StringReplace($sFilter, "TSRC", "TRC")
	$sFilter = StringReplace($sFilter, "TCON", "TCO")
	$sFilter = StringReplace($sFilter, "TLEN", "TLE")
	$sFilter = StringReplace($sFilter, "TPUB", "TPB")
	$sFilter = StringReplace($sFilter, "TFLT", "TFT")
	$sFilter = StringReplace($sFilter, "UFID", "UFI")
	$sFilter = StringReplace($sFilter, "TCOM", "TCM")
	$sFilter = StringReplace($sFilter, "WCOM", "WCM")
	$sFilter = StringReplace($sFilter, "WCOP", "WCP")
	$sFilter = StringReplace($sFilter, "WXXX", "WXX")
	$sFilter = StringReplace($sFilter, "WOAR", "WAR")
	$sFilter = StringReplace($sFilter, "WOAS", "WAS")
	$sFilter = StringReplace($sFilter, "WOAF", "WAF")
	$sFilter = StringReplace($sFilter, "WPUB", "WPB")
	$sFilter = StringReplace($sFilter, "PCNT", "CNT")

EndFunc

;Author: YDY (Lazycat)  Thank you for writing all this out
Func _GetGenreByID($iID)
	Local $asGenre = StringSplit("Blues,Classic Rock,Country,Dance,Disco,Funk,Grunge,Hip-Hop," & _
			"Jazz,Metal,New Age, Oldies,Other,Pop,R&B,Rap,Reggae,Rock,Techno,Industrial,Alternative," & _
			"Ska,Death Metal,Pranks,Soundtrack,Euro-Techno,Ambient,Trip-Hop,Vocal,Jazz+Funk,Fusion," & _
			"Trance,Classical,Instrumental,Acid,House,Game,Sound Clip,Gospel,Noise,Alternative Rock," & _
			"Bass,Soul,Punk,Space,Meditative,Instrumental Pop,Instrumental Rock,Ethnic,Gothic,Darkwave," & _
			"Techno-Industrial,Electronic,Pop-Folk,Eurodance,Dream,Southern Rock,Comedy,Cult,Gangsta," & _
			"Top 40,Christian Rap,Pop/Funk,Jungle,Native US,Cabaret,New Wave,Psychadelic,Rave,Showtunes," & _
			"Trailer,Lo-Fi,Tribal,Acid Punk,Acid Jazz,Polka,Retro,Musical,Rock & Roll,Hard Rock,Folk," & _
			"Folk-Rock,National Folk,Swing,Fast Fusion,Bebob,Latin,Revival,Celtic,Bluegrass,Avantgarde," & _
			"Gothic Rock,Progressive Rock,Psychedelic Rock,Symphonic Rock,Slow Rock,Big Band,Chorus," & _
			"Easy Listening,Acoustic,Humour,Speech,Chanson,Opera,Chamber Music,Sonata,Symphony,Booty Bass," & _
			"Primus,Porn Groove,Satire,Slow Jam,Club,Tango,Samba,Folklore,Ballad,Power Ballad,Rhytmic Soul," & _
			"Freestyle,Duet,Punk Rock,Drum Solo,Acapella,Euro-House,Dance Hall,Goa,Drum & Bass,Club-House," & _
			"Hardcore,Terror,Indie,BritPop,Negerpunk,Polsk Punk,Beat,Christian Gangsta,Heavy Metal,Black Metal," & _
			"Crossover,Contemporary C,Christian Rock,Merengue,Salsa,Thrash Metal,Anime,JPop,SynthPop", ",")
    If ($iID >= 0) and ($iID < 148) Then Return $asGenre[$iID + 1]
	Return ("")
EndFunc

Func _GetGenreID($sGrenre)
	Local $asGenre = StringSplit("Blues,Classic Rock,Country,Dance,Disco,Funk,Grunge,Hip-Hop," & _
			"Jazz,Metal,New Age, Oldies,Other,Pop,R&B,Rap,Reggae,Rock,Techno,Industrial,Alternative," & _
			"Ska,Death Metal,Pranks,Soundtrack,Euro-Techno,Ambient,Trip-Hop,Vocal,Jazz+Funk,Fusion," & _
			"Trance,Classical,Instrumental,Acid,House,Game,Sound Clip,Gospel,Noise,Alternative Rock," & _
			"Bass,Soul,Punk,Space,Meditative,Instrumental Pop,Instrumental Rock,Ethnic,Gothic,Darkwave," & _
			"Techno-Industrial,Electronic,Pop-Folk,Eurodance,Dream,Southern Rock,Comedy,Cult,Gangsta," & _
			"Top 40,Christian Rap,Pop/Funk,Jungle,Native US,Cabaret,New Wave,Psychadelic,Rave,Showtunes," & _
			"Trailer,Lo-Fi,Tribal,Acid Punk,Acid Jazz,Polka,Retro,Musical,Rock & Roll,Hard Rock,Folk," & _
			"Folk-Rock,National Folk,Swing,Fast Fusion,Bebob,Latin,Revival,Celtic,Bluegrass,Avantgarde," & _
			"Gothic Rock,Progressive Rock,Psychedelic Rock,Symphonic Rock,Slow Rock,Big Band,Chorus," & _
			"Easy Listening,Acoustic,Humour,Speech,Chanson,Opera,Chamber Music,Sonata,Symphony,Booty Bass," & _
			"Primus,Porn Groove,Satire,Slow Jam,Club,Tango,Samba,Folklore,Ballad,Power Ballad,Rhytmic Soul," & _
			"Freestyle,Duet,Punk Rock,Drum Solo,Acapella,Euro-House,Dance Hall,Goa,Drum & Bass,Club-House," & _
			"Hardcore,Terror,Indie,BritPop,Negerpunk,Polsk Punk,Beat,Christian Gangsta,Heavy Metal,Black Metal," & _
			"Crossover,Contemporary C,Christian Rock,Merengue,Salsa,Thrash Metal,Anime,JPop,SynthPop", ",")

	For $i = 1 to $asGenre[0]
		If $sGrenre == $asGenre[$i] Then
			Return $i - 1
		EndIf
	Next
	
	Return 12 ;Other
	
	
EndFunc


Func _HexToUint32_ID3($HexString4Byte)

	Return Dec(StringLeft($HexString4Byte, 2)) * Dec("FFFFFF") + Dec(StringTrimLeft($HexString4Byte, 2));need this because function Dec is signed

EndFunc

Func _HexToBin_ID3($HexString)
	Local $Bin = ""
	For $i = 1 To StringLen($HexString) Step 1
		$Hex = StringRight(StringLeft($HexString, $i), 1)
		Select
			Case $Hex = "0"
				$Bin = $Bin & "0000"
			Case $Hex = "1"
				$Bin = $Bin & "0001"
			Case $Hex = "2"
				$Bin = $Bin & "0010"
			Case $Hex = "3"
				$Bin = $Bin & "0011"
			Case $Hex = "4"
				$Bin = $Bin & "0100"
			Case $Hex = "5"
				$Bin = $Bin & "0101"
			Case $Hex = "6"
				$Bin = $Bin & "0110"
			Case $Hex = "7"
				$Bin = $Bin & "0111"
			Case $Hex = "8"
				$Bin = $Bin & "1000"
			Case $Hex = "9"
				$Bin = $Bin & "1001"
			Case $Hex = "A"
				$Bin = $Bin & "1010"
			Case $Hex = "B"
				$Bin = $Bin & "1011"
			Case $Hex = "C"
				$Bin = $Bin & "1100"
			Case $Hex = "D"
				$Bin = $Bin & "1101"
			Case $Hex = "E"
				$Bin = $Bin & "1110"
			Case $Hex = "F"
				$Bin = $Bin & "1111"
			Case Else
				SetError(-1)
		EndSelect
	Next
	If @error Then
		Return "ERROR"
	Else
		Return $Bin
	EndIf
EndFunc

Func _BinToHex_ID3($BinString)
	Local $Hex = ""
	If Not IsInt(StringLen($BinString) / 4) Then
		$Num = ((StringLen($BinString) / 4) - Int(StringLen($BinString) / 4)) * 4
;~ 		MsgBox(0, "", $Num)
		For $i = 1 To 4 - $Num Step 1
			$BinString = "0" & $BinString
		Next
	EndIf

	For $i = 4 To StringLen($BinString) Step 4
		$Bin = StringLeft(StringRight($BinString, $i), 4)
		Select
			Case $Bin = "0000"
				$Hex = $Hex & "0"
			Case $Bin = "0001"
				$Hex = $Hex & "1"
			Case $Bin = "0010"
				$Hex = $Hex & "2"
			Case $Bin = "0011"
				$Hex = $Hex & "3"
			Case $Bin = "0100"
				$Hex = $Hex & "4"
			Case $Bin = "0101"
				$Hex = $Hex & "5"
			Case $Bin = "0110"
				$Hex = $Hex & "6"
			Case $Bin = "0111"
				$Hex = $Hex & "7"
			Case $Bin = "1000"
				$Hex = $Hex & "8"
			Case $Bin = "1001"
				$Hex = $Hex & "9"
			Case $Bin = "1010"
				$Hex = $Hex & "A"
			Case $Bin = "1011"
				$Hex = $Hex & "B"
			Case $Bin = "1100"
				$Hex = $Hex & "C"
			Case $Bin = "1101"
				$Hex = $Hex & "D"
			Case $Bin = "1110"
				$Hex = $Hex & "E"
			Case $Bin = "1111"
				$Hex = $Hex & "F"
			Case Else
				SetError(-1)
		EndSelect
	Next
	If @error Then
		Return "ERROR"
	Else
		Return _StringReverse($Hex)
	EndIf
EndFunc
; -----------------------------------------------------------------------------
