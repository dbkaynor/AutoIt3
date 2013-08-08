#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_icon=../icons/HotSun.ico
#AutoIt3Wrapper_outfile=C:\Program Files\AutoIt3\Dougs\DuplicateFinder.exe
#AutoIt3Wrapper_Res_Comment=A Duplicate Finder
#AutoIt3Wrapper_Res_Description=Duplicate Finder
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
;#AutoIt3Wrapper_Run_Obfuscator=nf
;#AutoIt3Wrapper_Run_cvsWrapper=vf
;#Obfuscator_Parameters=/striponly
;#AutoIt3Wrapper_Res_Field=Credits|
;#AutoIt3Wrapper_Run_After=copy "%out%" "..\..\Programs_Updates\AutoIt3Wrapper"
;#NoTrayIcon
;#Tidy_Parameters=/gd /sf
;#Tidy_Parameters=/nsdp /sf1688
;#Tidy_Parameters=/sci=9
;-----------------------------------------------

#cs
    This area is used to store things to do, bugs, and other notes
    
    Fixed:
    
    Todo:
    Use CRC au3
    Run without fsum installed
    
    Combobox RMC box does not work. File move and copy need work
    Don't ask for recycle and so on
    Fix all help calls (F1)
    Run without IrfanView installed
    Allow user to enter path to IrfanView
    Check saved paths
    Verify paths in the results box exist
    Hide main window until move is complete
    Move selected item in ListView (GUICtrlCreateListView) to top of view (search, ???)
    
    autoit3 listview autoscroll
    _GUICtrlListView_EnsureVisible($hWnd, $iIndex[, $fPartialOK = False])
    
#ce

Opt("MustDeclareVars", 1)
Opt("TrayIconDebug", 1)
Opt("TrayAutoPause", 0)
Opt("WinTitleMatchMode", 2)

#include <Array.au3>
#include <ButtonConstants.au3>
#include <Constants.au3>
#include <Date.au3>
#include <EditConstants.au3>
#include <GuiEdit.au3>
#include <GuiComboBox.au3>
#include <GuiComboBoxEx.au3>
#include <GUIConstantsEx.au3>
#include <GUIListBox.au3>
#include <GuiTreeView.au3>
#include <Misc.au3>
#include <String.au3>
#include <StaticConstants.au3>
#include <TreeViewConstants.au3>
#include <WindowsConstants.au3>
#include <Process.au3>
#include <SliderConstants.au3>
#include <ListViewConstants.au3>
#include <GuiListView.au3>
#include <_CRC32.au3>
#include <_DougFunctions.au3>

Global Const $tmp = StringSplit(@ScriptName, ".")
Global Const $ProgramName = $tmp[1]

If _Singleton($ProgramName, 1) = 0 Then
    LogFile(@ScriptLineNumber & " " & $ProgramName & " is already running!", True)
    Exit
EndIf

_Debug("DBGVIEWCLEAR")
DirCreate("AUXFiles")
Global $WorkingFolder = EnvGet("USERPROFILE")
Global $AuxPath = @ScriptDir & "\AUXFiles\"

Global $Log_filename = $AuxPath & $ProgramName & ".log"
Global Const $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")
Global $Project_filename = $AuxPath & $ProgramName & ".prj"
Global $SystemS = $ProgramName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & @OSServicePack & @CRLF & @OSType & @CRLF & @OSArch
Global $Debug = False

Global $IrfanView
Global $DupFinderTempJPG = $AuxPath & "DupFinderTemp.jpg"
Global $FileArray[1] ;This is the listing of all file before any filtering
Global $FilteredArray[1] ;This is the listing of files after filtering
Global $ListResultsArray[1]

Global $HandleArray[1]
Global $ToggleState = False

Const $NameExt = 0
Const $FullPath = 1
Const $Name = 2
Const $Ext = 3
Const $Size = 4
Const $Date = 5
Const $CRC = 6
Const $Attr = 7
Const $Row = 8
Global $DataArray[1]
Global $TimeStamp = 0

Global $hWnd, $iMsg, $iwParam, $ilParam
GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")

FileDelete($Log_filename)

;-----------------------------------------------
; Process the command line
For $x = 1 To $CmdLine[0]
    ConsoleWrite($x & " >> " & $CmdLine[$x] & @CRLF)
    Select
        Case StringInStr($CmdLine[$x], "help") > 0 Or StringInStr($CmdLine[$x], "?") > 0
            Help()
            Exit
        Case StringInStr($CmdLine[$x], "debug") > 0
            $Debug = True
        Case Else
            LogFile(@ScriptLineNumber & " Unknown cmdline option found: >>" & $CmdLine[$x] & "<<", True)
            Exit
    EndSelect
Next
;-----------------------------------------------
; Mainform
Global $MainFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $MainForm = GUICreate("Duplicate finderFileVersion", 850, 820, 10, 10, $MainFormOptions)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlCreateGroup("Buttons", 740, 5, 100, 380);Buttons
GUICtrlSetResizing(-1, $GUI_DOCKALL + $GUI_DOCKLEFT)
Global $ButtonMainStart = GUICtrlCreateButton("Start", 750, 25, 80, 20, $WS_GROUP)
GUICtrlSetResizing(-1, $GUI_DOCKALL + $GUI_DOCKLEFT)
Global $ButtonMainChoseFolders = GUICtrlCreateButton("Chose folders", 750, 50, 80, 20, $WS_GROUP)
GUICtrlSetResizing(-1, $GUI_DOCKALL + $GUI_DOCKLEFT)
Global $ButtonMainLoadProject = GUICtrlCreateButton("Load Project", 750, 75, 80, 20, $WS_GROUP)
GUICtrlSetResizing(-1, $GUI_DOCKALL + $GUI_DOCKLEFT)
Global $ButtonMainSaveProject = GUICtrlCreateButton("Save Project", 750, 100, 80, 20, $WS_GROUP)
GUICtrlSetResizing(-1, $GUI_DOCKALL + $GUI_DOCKLEFT)

Global $ButtonMainEdit = GUICtrlCreateButton("Edit", 750, 125, 80, 20, $WS_GROUP)
GUICtrlSetResizing(-1, $GUI_DOCKALL + $GUI_DOCKLEFT)
Global $ButtonMainAbout = GUICtrlCreateButton("About", 750, 150, 80, 20, $WS_GROUP)
GUICtrlSetResizing(-1, $GUI_DOCKALL + $GUI_DOCKLEFT)
Global $ButtonMainHelp = GUICtrlCreateButton("Help", 750, 175, 80, 20, $WS_GROUP)
GUICtrlSetResizing(-1, $GUI_DOCKALL + $GUI_DOCKLEFT)
Global $ButtonMainClose = GUICtrlCreateButton("Close", 750, 200, 80, 20, $WS_GROUP)
GUICtrlSetResizing(-1, $GUI_DOCKALL + $GUI_DOCKLEFT)
Global $ButtonMainRecycle = GUICtrlCreateButton("Recycle", 750, 225, 80, 20, $WS_GROUP)
GUICtrlSetResizing(-1, $GUI_DOCKALL + $GUI_DOCKLEFT)
Global $ButtonMainSave = GUICtrlCreateButton("Save", 750, 250, 80, 20, $WS_GROUP)
GUICtrlSetResizing(-1, $GUI_DOCKALL + $GUI_DOCKLEFT)

Global $CheckMainAbort = GUICtrlCreateCheckbox("Abort", 775, 290, 80, 20, $WS_GROUP)
GUICtrlSetResizing(-1, $GUI_DOCKALL + $GUI_DOCKLEFT)
GUICtrlSetTip(-1, "Abort the operation")
Global $ButtonMainWorking = GUICtrlCreateButton("", 750, 320, 80, 30, $WS_GROUP)
GUICtrlSetResizing(-1, $GUI_DOCKALL + $GUI_DOCKLEFT)
GUICtrlSetTip(-1, "Progress")

GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateGroup("Filters", 10, 5, 520, 80);filters
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $InputFilter = GUICtrlCreateInput("", 20, 20, 304, 21)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetTip(-1, "Filters string")
Global $RadioInclude = GUICtrlCreateRadio("Include", 330, 25, 60, 20)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $RadioExclude = GUICtrlCreateRadio("Exclude", 400, 25, 60, 20)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $RadioAll = GUICtrlCreateRadio("All", 470, 25, 50, 20)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetState(-1, $GUI_CHECKED)

Global $CheckIncludeFolders = GUICtrlCreateCheckbox("Include folders", 20, 55, 95, 20)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetState(-1, $GUI_UNCHECKED)

Global $InputSelectedProject = GUICtrlCreateInput("project", 120, 55, 400, 20)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "Currently selected project")
GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateGroup("Grouping", 540, 5, 150, 50)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ComboGrouping = GUICtrlCreateCombo("No grouping", 550, 25, 130);, 100, $CBS_AUTOHSCROLL)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetData(-1, "Name & Extension|Full path|Name|Extension|Size|Date|CRC|Attribute", "Name & Extension")
GUICtrlCreateGroup("", -99, -99, 1, 1)

Global $ButtonMainSearchFirst = GUICtrlCreateButton("Search", 10, 90, 50, 20, $WS_GROUP)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $ButtonMainSearchAgain = GUICtrlCreateButton("Again", 70, 90, 50, 20, $WS_GROUP)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $InputSearch = GUICtrlCreateInput("", 130, 90, 180, 20)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetTip(-1, "Search string")

GUICtrlCreateGroup("Status", 10, 113, 720, 280)
GUICtrlSetResizing(-1, $GUI_DOCKALL - $GUI_DOCKWIDTH)
Global $LabelTotal = GUICtrlCreateLabel("Total lines: NA", 20, 140, 150, 20, $SS_SUNKEN)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $LabelFiltered = GUICtrlCreateLabel("Filtered: NA", 200, 140, 150, 20, $SS_SUNKEN)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Global $LabelDisplayed = GUICtrlCreateLabel("Displayed: NA", 420, 140, 150, 20, $SS_SUNKEN)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

Global $ListStatus = GUICtrlCreateList("", 15, 170, 700, 210, BitOR($LBS_NOTIFY, $WS_BORDER, $WS_TABSTOP, $WS_GROUP, $LBS_DISABLENOSCROLL, $WS_VSCROLL, $WS_HSCROLL))
GUICtrlSetData(-1, "")
GUICtrlSetTip(-1, "This is the List of status infomation")
GUICtrlSetResizing(-1, $GUI_DOCKALL - $GUI_DOCKWIDTH)

Global $ListView = GUICtrlCreateListView("Name & Ext|Full path|Name|Ext|Size|Date|CRC|Attr|Row", 15, 400, 825, 400, -1, BitOR($WS_EX_CLIENTEDGE, $LVS_EX_GRIDLINES, $LVS_EX_FULLROWSELECT))
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKBOTTOM)
GUICtrlSetTip(-1, "This is the list view of results")
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $NameExt, 150)
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $FullPath, 200)
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $Name, 75)
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $Ext, 50)
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $Size, 60)
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $Date, 120)
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $CRC, 60)
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $Attr, 50)
GUICtrlSendMsg($ListView, $LVM_SETCOLUMNWIDTH, $Row, 50)

_GUICtrlListView_SetBkColor($ListView, $CLR_WHITE)
_GUICtrlListView_SetTextColor($ListView, $CLR_BLACK)
_GUICtrlListView_SetTextBkColor($ListView, $CLR_Silver)

GUICtrlCreateGroup("", -99, -99, 1, 1)
GUISetState(@SW_SHOW)

;-----------------------------------------------
; Rename, move, copy dialog box ;dbk
Global $RMCForm = GUICreate("Rename, move, copy", 420, 140, 50, 50)
Global $InputRMCInName = GUICtrlCreateInput("???", 10, 10, 400, 20, $es_readonly)
;Global $InputRMCOutName = GUICtrlCreateInput("???", 10, 40, 400, 20)
Global $ComboRMCOutName = GUICtrlCreateCombo("???", 10, 40, 400, 20)

Global $ButtonRMCRename = GUICtrlCreateButton("Rename", 10, 100, 60, 30, $WS_GROUP)
Global $ButtonRMCMove = GUICtrlCreateButton("Move", 10, 100, 60, 30, $WS_GROUP)
Global $ButtonRMCCopy = GUICtrlCreateButton("Copy", 10, 100, 60, 30, $WS_GROUP)
Global $ButtonRMCClose = GUICtrlCreateButton("Close", 200, 100, 60, 30, $WS_GROUP)
Global $ButtonRMCAbout = GUICtrlCreateButton("About", 260, 100, 60, 30, $WS_GROUP)
GUISetState(@SW_HIDE, $RMCForm)

;-----------------------------------------------
; File Single form
Global $SingleForm = GUICreate("Single", 340, 340, 30, 30)
Global $InputSingle = GUICtrlCreateInput("???", 10, 10, 320, 20, $es_readonly)
Global $ButtonSingleShow = GUICtrlCreateButton("Show", 10, 40, 100, 30, $WS_GROUP)
Global $ButtonSingleRecycle = GUICtrlCreateButton("Recycle", 10, 75, 100, 30, $WS_GROUP)
Global $ButtonSingleDelete = GUICtrlCreateButton("Delete", 10, 110, 100, 30, $WS_GROUP)
Global $ButtonSingleRename = GUICtrlCreateButton("Rename", 10, 145, 100, 30, $WS_GROUP)
Global $ButtonSingleMove = GUICtrlCreateButton("Move", 10, 180, 100, 30, $WS_GROUP)
Global $ButtonSingleCopy = GUICtrlCreateButton("Copy", 10, 215, 100, 30, $WS_GROUP)
Global $ButtonSingleClose = GUICtrlCreateButton("Close", 10, 250, 100, 30, $WS_GROUP)
Global $ButtonSingleAbout = GUICtrlCreateButton("About", 10, 285, 100, 30, $WS_GROUP)
Global $ButtonViewRecycleBinS = GUICtrlCreateButton("VRB", 120, 280, 100, 30, $WS_GROUP)
Global $InputS = GUICtrlCreateInput(">>>", 240, 280, 40, 30)
Global $ViewPortS = GUICtrlCreatePic($AuxPath & "Question.jpg", 120, 40, 300, 300, $WS_GROUP)

GUISetState(@SW_HIDE, $SingleForm)
;-----------------------------------------------
; File Multi form
Global $MultiForm = GUICreate("Multi", 1075, 330, 30, 30)
Global $ButtonRecycleSelected = GUICtrlCreateButton("Recycle selected", 10, 10, 100, 30, $WS_GROUP)
Global $CheckDontAsk = GUICtrlCreateCheckbox("Don't ask", 120, 10, 70, 30, $WS_GROUP)
Global $ButtonMultiClose = GUICtrlCreateButton("Close", 210, 10, 50, 30, $WS_GROUP)
Global $ButtonMultiAbout = GUICtrlCreateButton("About", 270, 10, 50, 30, $WS_GROUP)
Global $ButtonViewRecycleBinM = GUICtrlCreateButton("VRB", 330, 10, 50, 30, $WS_GROUP)
Global $LabelIndicator = GUICtrlCreateLabel("NA", 400, 15, 60, 30, BitOR($WS_GROUP, $SS_SUNKEN))

Global $ViewPortM1 = GUICtrlCreatePic($AuxPath & "Question.jpg", 10, 45, 200, 200, $WS_GROUP)
Global $InputM1 = GUICtrlCreateInput("Question.jpg", 10, 250, 200, 50, $ES_MULTILINE)
Global $InputM1a = GUICtrlCreateInput(">>>", 10, 300, 40, 20)
Global $CheckM1 = GUICtrlCreateCheckbox("Select", 60, 300, 70, 30)
Global $ButtonM1 = GUICtrlCreateButton("Go", 60 + 70, 300, 30)

Global $ViewPortM2 = GUICtrlCreatePic($AuxPath & "Question.jpg", 220, 45, 200, 200, $WS_GROUP)
Global $InputM2 = GUICtrlCreateInput("Question.jpg", 220, 250, 200, 50, $ES_MULTILINE)
Global $InputM2a = GUICtrlCreateInput(">>>", 220, 300, 40, 20)
Global $CheckM2 = GUICtrlCreateCheckbox("Select", 270, 300, 70, 30)
Global $ButtonM2 = GUICtrlCreateButton("Go", 270 + 70, 300, 30)

Global $ViewPortM3 = GUICtrlCreatePic($AuxPath & "Question.jpg", 430, 45, 200, 200, $WS_GROUP)
Global $InputM3 = GUICtrlCreateInput("Question.jpg", 430, 250, 200, 50, $ES_MULTILINE)
Global $InputM3a = GUICtrlCreateInput(">>>", 430, 300, 40, 20)
Global $CheckM3 = GUICtrlCreateCheckbox("Select", 480, 300, 70, 30)
Global $ButtonM3 = GUICtrlCreateButton("Go", 480 + 70, 300, 30)

Global $ViewPortM4 = GUICtrlCreatePic($AuxPath & "Question.jpg", 640, 45, 200, 200, $WS_GROUP)
Global $InputM4 = GUICtrlCreateInput("Question.jpg", 640, 250, 200, 50, $ES_MULTILINE)
Global $InputM4a = GUICtrlCreateInput(">>>", 640, 300, 40, 20)
Global $CheckM4 = GUICtrlCreateCheckbox("Select", 690, 300, 70, 30)
Global $ButtonM4 = GUICtrlCreateButton("Go", 690 + 70, 300, 30)

Global $ViewPortM5 = GUICtrlCreatePic($AuxPath & "Question.jpg", 850, 45, 200, 200, $WS_GROUP)
Global $InputM5 = GUICtrlCreateInput("Question.jpg", 850, 250, 200, 50, $ES_MULTILINE)
Global $InputM5a = GUICtrlCreateInput(">>>", 850, 300, 40, 20)
Global $CheckM5 = GUICtrlCreateCheckbox("Select", 900, 300, 70, 30)
Global $ButtonM5 = GUICtrlCreateButton("Go", 900 + 70, 300, 30)

GUISetState(@SW_HIDE, $MultiForm)
;-----------------------------------------------
; Select form
Global $SelectFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $SelectForm = GUICreate("Select folders", 450, 600, 10, 10, $SelectFormOptions)
GUISetFont(10, 400, 0, "Courier New")

Global $ButtonSelectFolders = GUICtrlCreateButton("Select top folder", 15, 15, 140, 20)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKWIDTH + $GUI_DOCKTOP + $GUI_DOCKHEIGHT)

Global $InputSelectFolder = GUICtrlCreateInput($WorkingFolder, 160, 15, 250, 24)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKHEIGHT)

Global $ButtonSelectGetFolders = GUICtrlCreateButton("Get folders", 15, 40, 100, 20)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKWIDTH + $GUI_DOCKTOP + $GUI_DOCKHEIGHT)

Global $ButtonSelectProcess = GUICtrlCreateButton("Process", 125, 40, 70, 20)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKWIDTH + $GUI_DOCKTOP + $GUI_DOCKHEIGHT)

Global $ButtonSelectDone = GUICtrlCreateButton("Done", 205, 40, 50, 20)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKWIDTH + $GUI_DOCKTOP + $GUI_DOCKHEIGHT)

Global $ButtonSelectToggle = GUICtrlCreateButton("Toggle", 265, 40, 70, 20)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKWIDTH + $GUI_DOCKTOP + $GUI_DOCKHEIGHT)

Global $ButtonSelectAbout = GUICtrlCreateButton("About", 345, 40, 50, 20)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKWIDTH + $GUI_DOCKTOP + $GUI_DOCKHEIGHT)

Global $LabelFolderList = GUICtrlCreateLabel("Folder list", 15, 75, 90, 20)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKHEIGHT)
Global $TreeViewFolders = GUICtrlCreateTreeView(10, 100, 400, 200, BitOR($TVS_HASBUTTONS, $TVS_HASLINES, $TVS_LINESATROOT, $TVS_DISABLEDRAGDROP, $TVS_SHOWSELALWAYS, $TVS_CHECKBOXES, $WS_GROUP, $WS_TABSTOP), $WS_EX_CLIENTEDGE)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKWIDTH + $GUI_DOCKTOP) ;+ $GUI_DOCKHEIGHT + $GUI_DOCKBOTTOM)

Global $LabelResults = GUICtrlCreateLabel("Results", 15, 300, 60, 20)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
Global $ListResults = GUICtrlCreateList("", 10, 320, 400, 246)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM)
;-----------------------------------------------

SetDefaults()
LoadProject("Start")
TestForIrfanView()
If $Debug Then Debug()
HotKeySet("{F11}", "GUI_Enable")

;GUISetHelp("notepad c:\config.sys", $MainForm) ; Need a help file to call here
;GUISetHelp("notepad c:\dosvar.bat", $OptionForm) ; Need a help file to call here
;GUISetHelp("notepad c:\autoexec.bat", $SelectFilesForm)

While 1
    Global $nMsg = GUIGetMsg(1)
    Switch $nMsg[0]
        Case $GUI_EVENT_CLOSE ; The following handles the red X on all forms
            Switch $nMsg[1]
                Case $MainForm
                    Exit
                Case $MultiForm
                    GUISetState(@SW_HIDE, $MultiForm)
                Case $SingleForm
                    GUISetState(@SW_HIDE, $SingleForm)
                Case $RMCForm
                    GUISetState(@SW_HIDE, $RMCForm)
                Case $SelectForm
                    GUISetState(@SW_HIDE, $SelectForm)
            EndSwitch

            ;------- Main form
        Case $ButtonMainClose
            Exit
        Case $ComboGrouping
            ConsoleWrite(@ScriptLineNumber & " >>" & GUICtrlRead($ComboGrouping) & "<<" & @CRLF)
        Case $ButtonMainRecycle
            LogFile(@ScriptLineNumber & " Open recycle bin")
            Run("explorer ::{645FF040-5081-101B-9F08-00AA002F954E}")
        Case $ButtonMainSave
            SaveResults()
        Case $ButtonMainStart
            LoadFiles()
        Case $ButtonMainSearchFirst
            Search(True)
        Case $ButtonMainSearchAgain
            Search(False)
        Case $InputSearch
            Search(False)
        Case $ListStatus
            If MsgBox(36 + 8192, "Clear status?", "Are you sure that you want to clear status?" & @CRLF & "(Does not affect the log file.)") = 6 Then
                _GUICtrlListBox_ResetContent($ListStatus)
            EndIf
        Case $InputSearch
        Case $RadioInclude
            ;FilterFiles()
        Case $RadioExclude
            ;FilterFiles()
        Case $RadioAll
            ;FilterFiles()
        Case $ButtonMainSaveProject
            SaveProject()
        Case $ButtonMainLoadProject
            LoadProject("menu")
        Case $ButtonMainEdit
            EditText()
        Case $ButtonMainChoseFolders
            LogFile(@ScriptLineNumber & " ButtonMainChoseFolders  " & $WorkingFolder)
            GUICtrlSetData($InputSelectFolder, $WorkingFolder)
            GUISetState(@SW_SHOW, $SelectForm)
        Case $ButtonMainAbout
            About("Duplicate")
        Case $ButtonMainHelp
            Help()

        Case $ButtonRecycleSelected
            MultiRecycle()
        Case $ButtonM1
            MultiGoButtons('M1')
        Case $ButtonM2
            MultiGoButtons('M2')
        Case $ButtonM3
            MultiGoButtons('M3')
        Case $ButtonM4
            MultiGoButtons('M4')
        Case $ButtonM5
            MultiGoButtons('M5')
        Case $ButtonMultiClose
            GUISetState(@SW_HIDE, $MultiForm)
        Case $ButtonMultiAbout
            About("Multi")
        Case $ButtonViewRecycleBinM
            LogFile(@ScriptLineNumber & " Open recycle bin")
            Run("explorer ::{645FF040-5081-101B-9F08-00AA002F954E}")
            ;------- File Single form
        Case $ButtonSingleRecycle
            SingleRecycle()
        Case $ButtonSingleDelete
            SingleDelete()
        Case $ButtonSingleShow ; Irfanview
            If Not $IrfanView = "XXX" Then
                LogFile(@ScriptLineNumber & " >" & $IrfanView & "<>" & GUICtrlRead($InputSingle) & "<")
                ShellExecuteWait($IrfanView, GUICtrlRead($InputSingle), "", "open")
            EndIf
        Case $ButtonSingleRename
            RenameMoveCopy("R")
        Case $ButtonSingleMove
            RenameMoveCopy("M")
        Case $ButtonSingleCopy
            RenameMoveCopy("C")
        Case $ButtonSingleClose
            GUISetState(@SW_HIDE, $SingleForm)
        Case $ButtonSingleAbout
            About("Single")
        Case $ButtonViewRecycleBinS
            LogFile(@ScriptLineNumber & " Open recycle bin")
            Run("explorer ::{645FF040-5081-101B-9F08-00AA002F954E}")
            ;------- RMC form
        Case $ComboRMCOutName
            _GUICtrlComboBox_AddString($ComboRMCOutName, GUICtrlRead($InputRMCInName))
        Case $InputRMCInName
            _GUICtrlComboBox_SetEditText($ComboRMCOutName, GUICtrlRead($InputRMCInName))
        Case $ComboRMCOutName
            _GUICtrlComboBox_SetEditText($ComboRMCOutName, GUICtrlRead($InputRMCInName))
        Case $ButtonRMCRename
            RMC("Rename")
        Case $ButtonRMCMove
            RMC("Move")
        Case $ButtonRMCCopy
            RMC("Copy")
        Case $ButtonRMCClose
            GUISetState(@SW_HIDE, $RMCForm)
        Case $ButtonRMCAbout
            About("Rename, move, copy"); RMC
        Case $GUI_EVENT_CLOSE
            Exit
        Case $ButtonSelectFolders
            ChoseFolders()
            ;------- select folders form
        Case $ButtonSelectGetFolders
            $WorkingFolder = GUICtrlRead($InputSelectFolder)
            GetFolders()
        Case $ButtonSelectProcess
            Process()
        Case $ButtonSelectToggle
            ToggleChecked()
        Case $ButtonSelectDone
            GUISetState(@SW_HIDE, $SelectForm)
        Case $ButtonSelectAbout
            About("Select")
        Case $InputSelectFolder
            ConsoleWrite(@ScriptLineNumber & " InputSelectFolder" & @CRLF)

    EndSwitch
WEnd
;-----------------------------------------------
; hotkey F11
Func GUI_Enable()
    ConsoleWrite(@ScriptLineNumber & " Hotkey F11" & @CRLF)
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>GUI_Enable
;-----------------------------------------------
;$IrfanView
Func TestForIrfanView()
    Local $tmp = "C:\Program Files\IrfanView\i_view32.exe"
    If FileExists($tmp) Then
        $IrfanView = $tmp
        LogFile(@ScriptLineNumber & " TestForIrfanView found: " & $IrfanView)
        Return
    EndIf

    $tmp = "C:\Program Files (x86)\IrfanView\i_view32.exe"
    If FileExists($tmp) Then
        $IrfanView = $tmp
        LogFile(@ScriptLineNumber & " TestForIrfanView found: " & $IrfanView)
        Return
    EndIf

    If FileExists($IrfanView) Then
        LogFile(@ScriptLineNumber & " TestForIrfanView found: " & $IrfanView)
        Return
    EndIf

    While 1 ; keep trying until a good location is found or cancel is pressed
        $tmp = InputBox("Irfanview was not located", "Enter path to Irfanview and click OK", $IrfanView)
        If @error = 1 Then ExitLoop ; this handles the cancel button
        If FileExists($tmp) Then
            $IrfanView = $tmp
            LogFile(@ScriptLineNumber & " TestForIrfanView found: " & $IrfanView)
            Return
        EndIf
    WEnd

    $IrfanView = "XXX"
    LogFile(@ScriptLineNumber & " Unable to locate IrfanView on this computer")
    MsgBox(16, "IrfanView not found", "Unable to locate IrfanView on this computer")
EndFunc   ;==>TestForIrfanView
;-----------------------------------------------
Func Debug()
    ;   _GUICtrlListBox_ResetContent($ListFolders)
    ;   _GUICtrlListBox_AddString($ListFolders, "C:\Program Files\AutoIt3\Dougs")
    ;   _GUICtrlListBox_AddString($ListFolders, "c:\sdfcvp")
    GUICtrlSetData($InputFilter, "jpg;ini;html;txt")
EndFunc   ;==>Debug
;-----------------------------------------------
Func Search($Reset)
    Static Local $iIndex
    LogFile(@ScriptLineNumber & " Search: " & $Reset & "  " & $iIndex)
    If $Reset = True Then $iIndex = 0
    Local $SearchString = GUICtrlRead($InputSearch)
    For $x = $iIndex To _GUICtrlListView_GetItemCount($ListView)
        Local $RString = _GUICtrlListView_GetItemTextString($ListView, $x)
        If StringInStr($RString, $SearchString) Then
            ConsoleWrite(@ScriptLineNumber & " " & $RString & "  " & $Reset & "  " & $iIndex & @CRLF)
            _GUICtrlListView_SetItemSelected($ListView, $x)
            _GUICtrlListView_EnsureVisible($ListView, $x)
            $iIndex = $x + 1
            ExitLoop
        EndIf
        _GUICtrlListView_SetItemSelected($ListView, 0)
        _GUICtrlListView_EnsureVisible($ListView, 0)
        _GUICtrlListView_EnsureVisible($ListView, 0)
    Next

EndFunc   ;==>Search
;-----------------------------------------------
Func SaveResults()
    Local $FileName = $ProgramName & ".txt"
    LogFile(@ScriptLineNumber & " SaveResults: " & $FileName)
    FileDelete($FileName)

    FileWrite($FileName, "Sort" & @TAB & "Full path" & @TAB & "Name" & @TAB _
             & "Ext" & @TAB & "Size" & @TAB & "Date" & @TAB & "CRC" & @TAB _
             & "Attr" & @TAB & "Row" & @CRLF)
    For $x = 0 To _GUICtrlListView_GetItemCount($ListView) - 1
        Local $RString = _GUICtrlListView_GetItemTextString($ListView, $x)
        ConsoleWrite(@ScriptLineNumber & " " & $RString & @CRLF)
        FileWrite($FileName, StringReplace($RString, '|', @TAB) & @CRLF)
        ;FileWrite($FileName, $RString & @CRLF)
        ConsoleWrite(@ScriptLineNumber & " " & $RString & @CRLF)
    Next
    MsgBox(16, "Results saved", "Results saved as: " & @CRLF & $FileName & @CRLF & "(TAB separated CSV format)")
EndFunc   ;==>SaveResults

;-----------------------------------------------
; This function will verify that the paths exist and will remove duplicates
Func VerifyPaths()
    GuiDisable($GUI_DISABLE)
    LogFile(@ScriptLineNumber & " VerifyPaths")
    For $x = 0 To _GUICtrlListBox_GetCount($ListResults) - 1
        If Not FileExists(_GUICtrlListBox_GetText($ListResults, $x)) Then
            _GUICtrlListBox_DeleteString($ListResults, $x)
        EndIf
    Next
    LogFile(@ScriptLineNumber & " VerifyPaths complete")
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>VerifyPaths
;-----------------------------------------------
;This function will recursively load the files from the folders
Func LoadFiles()
    $TimeStamp = TimerInit()
    GUICtrlSetState($CheckMainAbort, $GUI_UNCHECKED)
    GuiDisable($GUI_DISABLE)
    GUICtrlSetData($ButtonMainWorking, 0)
    LogFile(@ScriptLineNumber & " LoadFiles: " & timer())
    _GUICtrlListView_DeleteAllItems(GUICtrlGetHandle($ListView))

    GUICtrlSetData($LabelFiltered, "Filtered lines: NA")
    GUICtrlSetData($LabelDisplayed, "Displayed lines: NA")
    GUICtrlSetData($LabelTotal, "Total lines: NA")

    If Not IsArray($FileArray) Then Global $FileArray[1]
    If Not IsArray($FilteredArray) Then Global $FilteredArray[1]

    ReDim $ListResultsArray[1]
    ReDim $FileArray[1]
    ReDim $FilteredArray[1]
    Local $tmpArray[1]

    For $x = 0 To _GUICtrlListBox_GetCount($ListResults) - 1
        If CheckAbort(@ScriptLineNumber) <> 0 Then ExitLoop
        ; no double slashes and add slash at tail of path if needed
        Local $tmp = _CleanUpPath(_GUICtrlListBox_GetText($ListResults, $x))

        $tmpArray = _FileListToArrayR($tmp, "*.*", 0, True, 0, "", 1)
        ;_ArrayDisplay($tmpArray, @ScriptLineNumber & "   " & $tmp)

        If $tmpArray = 0 Then
            LogFile(@ScriptLineNumber & " Folder not found: " & $tmp & "  " & timer())
        Else
            LogFile(@ScriptLineNumber & " Folder loaded: " & $tmp & "Files: " & UBound($tmpArray) & "  " & timer())
            For $y = 0 To UBound($tmpArray) - 1
                If CheckAbort(@ScriptLineNumber) <> 0 Then Return
                If StringLen($tmpArray[$y]) > 2 Then _ArrayAdd($FileArray, _AddSlash2PathString($tmp) & $tmpArray[$y])
            Next
        EndIf
    Next
    LogFile(@ScriptLineNumber & " LoadFiles complete  " & Timer())
    If CheckAbort(@ScriptLineNumber) <> 0 Then Return

    _ArrayDelete($FileArray, 0)
    GUICtrlSetData($LabelTotal, "Total lines: " & UBound($FileArray))
    LogFile(@ScriptLineNumber & " Files found: " & UBound($FileArray) & "   " & Timer())

    If UBound($FileArray) < 1 Then
        MsgBox(48, "Aborting!", "No files found")
        GuiDisable($GUI_ENABLE)
        Return
    EndIf

    FilterFiles()
    If CheckAbort(@ScriptLineNumber) <> 0 Then Return
    SplitRawData()
    If CheckAbort(@ScriptLineNumber) <> 0 Then Return


    GUICtrlSetData($LabelFiltered, "Filtered lines: " & UBound($FilteredArray))
    GUICtrlSetData($LabelDisplayed, "Displayed lines: " & _GUICtrlListView_GetItemCount($ListView) - 1)

    GuiDisable($GUI_ENABLE)

    LogFile(@ScriptLineNumber & " Loadfiles complete:  " & Timer())
EndFunc   ;==>LoadFiles

;-----------------------------------------------
Func FilterFiles()
    ;LogFile(@ScriptLineNumber & " FilterFiles:  " & Timer())
    If GUICtrlRead($RadioAll) = $GUI_CHECKED Then LogFile(@ScriptLineNumber & " FilterFiles: All files " & timer())
    If GUICtrlRead($RadioInclude) = $GUI_CHECKED Then LogFile(@ScriptLineNumber & " FilterFiles: Include  " & GUICtrlRead($InputFilter) & " " & timer())
    If GUICtrlRead($RadioExclude) = $GUI_CHECKED Then LogFile(@ScriptLineNumber & " FilterFiles: Exclude  " & GUICtrlRead($InputFilter) & " " & timer())
    If GUICtrlRead($CheckIncludeFolders) = $GUI_UNCHECKED Then
        LogFile(@ScriptLineNumber & " FilterFiles: Do not include folders " & timer())
    Else
        LogFile(@ScriptLineNumber & " FilterFiles: Include folders " & timer())
    EndIf

    ReDim $ListResultsArray[1]
    Local $RawArray[1]
    ReDim $FilteredArray[1]

    If GUICtrlRead($RadioAll) = $GUI_CHECKED Then
        For $A = 0 To UBound($FileArray) - 1
            If StringLen($FileArray[$A]) > 2 Then _ArrayAdd($RawArray, $FileArray[$A])
            If CheckAbort(@ScriptLineNumber) <> 0 Then Return
        Next
    Else
        Local $tmpArray1[1]
        Local $Filters = StringSplit(GUICtrlRead($InputFilter), ";~", 2)
        For $A = 0 To UBound($FileArray) - 1
            For $B = 0 To UBound($Filters) - 1
                If StringRegExp(StringUpper($FileArray[$A]), StringUpper($Filters[$B]) & "\Z", 0) = 1 Then
                    If GUICtrlRead($RadioInclude) = $GUI_CHECKED Then
                        _ArrayAdd($RawArray, $FileArray[$A])
                    ElseIf GUICtrlRead($RadioExclude) = $GUI_CHECKED Then
                        _ArrayAdd($tmpArray1, $FileArray[$A])
                    EndIf
                EndIf
                If CheckAbort(@ScriptLineNumber) <> 0 Then Return
            Next
            If CheckAbort(@ScriptLineNumber) <> 0 Then Return
        Next

        ; This handles the exclude function
        If GUICtrlRead($RadioExclude) = $GUI_CHECKED Then
            Local $tmpArray2[1]
            _ArrayConcatenate($tmpArray2, $FileArray)

            For $A = 0 To UBound($tmpArray1) - 1
                For $B = 0 To UBound($tmpArray2) - 1
                    If StringCompare($tmpArray1[$A], $tmpArray2[$B], 2) = 0 Then
                        $tmpArray2[$B] = ""
                    EndIf
                Next
            Next

            For $B = 0 To UBound($tmpArray2) - 1
                If StringLen($tmpArray2[$B]) > 2 Then _ArrayAdd($RawArray, $tmpArray2[$B])
            Next
        EndIf
    EndIf

    ; remove the folders from the list if needed
    If GUICtrlRead($CheckIncludeFolders) = $GUI_UNCHECKED Then
        For $B = 0 To UBound($RawArray) - 1
            If StringInStr(FileGetAttrib($RawArray[$B]), "D") > 0 Then
                $RawArray[$B] = ""
            EndIf
            If CheckAbort(@ScriptLineNumber) <> 0 Then Return
        Next
    EndIf

    ; if the string is more than 2 characters Multi display it
    For $B = 0 To UBound($RawArray) - 1
        If StringLen($RawArray[$B]) > 2 Then
            _ArrayAdd($FilteredArray, $RawArray[$B])
        EndIf
        If CheckAbort(@ScriptLineNumber) <> 0 Then Return
    Next
    _ArrayDelete($FilteredArray, 0)

EndFunc   ;==>FilterFiles
;-----------------------------------------------
;This function takes the filtered data and processes it into a ~ separated string
Func SplitRawData()
    LogFile(@ScriptLineNumber & " SplitRawData 1:  " & Timer())
    Local $ts
    Local $R
    Local $T
    If Not IsArray($FilteredArray) Then Return
    ;This section splits and then reformats the data
    For $A In $FilteredArray
        Local $B = StringSplit($A, "\")
        Local $C = StringSplit($B[$B[0]], ".")
        If $C[0] = 1 Then _ArrayAdd($C, "")
        Local $fs = CommafyString(String(FileGetSize($A)))
        Local $ft = FileGetTime($A, 0, 0)
        If IsArray($ft) Then
            Local $ft1 = StringFormat("%04d/%02d/%02d %02d:%02d:%02d", $ft[0], $ft[1], $ft[2], $ft[3], $ft[4], $ft[5])

            ; _CRC32($InputData, $type, $CRC32 = -1)
            $ts = $B[$B[0]] & "~" & $A & "~" & $C[1] & "~" & $C[2] & "~" & $fs & "~" & $ft1 & "~" & _CRC32($A, "file") & "~" & FileGetAttrib($A)
            _ArrayAdd($ListResultsArray, $ts)
        EndIf
        If CheckAbort(@ScriptLineNumber) <> 0 Then Return
    Next

    LogFile(@ScriptLineNumber & " SplitRawData 2:  " & Timer())
    ;if $RadioNoGroup is not selected than this section processes and adds the data to the listview

    If GUICtrlRead($ComboGrouping) <> "No grouping" Then
        Local $SplitOffset = -1
        If GUICtrlRead($ComboGrouping) = "Name & Extension" Then
            _GUICtrlListView_SetColumn($ListView, 0, "Name & Extension")
            $SplitOffset = $NameExt
        ElseIf GUICtrlRead($ComboGrouping) = "Full path" Then
            _GUICtrlListView_SetColumn($ListView, 0, "Full Path")
            $SplitOffset = $FullPath
        ElseIf GUICtrlRead($ComboGrouping) = "Name" Then
            _GUICtrlListView_SetColumn($ListView, 0, "Name")
            $SplitOffset = $Name
        ElseIf GUICtrlRead($ComboGrouping) = "Extension" Then
            _GUICtrlListView_SetColumn($ListView, 0, "Extension")
            $SplitOffset = $Ext
        ElseIf GUICtrlRead($ComboGrouping) = "Size" Then
            _GUICtrlListView_SetColumn($ListView, 0, "Size")
            $SplitOffset = $Size
        ElseIf GUICtrlRead($ComboGrouping) = "Date" Then
            _GUICtrlListView_SetColumn($ListView, 0, "Date")
            $SplitOffset = $Date
        ElseIf GUICtrlRead($ComboGrouping) = "CRC" Then
            _GUICtrlListView_SetColumn($ListView, 0, "CRC")
            $SplitOffset = $CRC
        ElseIf GUICtrlRead($ComboGrouping) = "Attribute" Then
            _GUICtrlListView_SetColumn($ListView, 0, "Attribute")
            $SplitOffset = $Attr
        EndIf

        ConsoleWrite(@ScriptLineNumber & " " & $SplitOffset & @CRLF)
        SortDataArray($SplitOffset)

        LogFile(@ScriptLineNumber & " After SortDataArray:  " & Timer())

        _FileWriteFromArray("c:\temp\ResultsArray22.txt", $SplitOffset)
        Local $HashArrayName[1]
        Local $HashArrayCount[1]
        For $A = 0 To UBound($ListResultsArray) - 1
            $B = StringSplit($ListResultsArray[$A], "~")
            If $B[0] <> 8 Then ContinueLoop
            $R = _ArraySearch($HashArrayName, $B[$SplitOffset + 1])

            If $R = -1 And (@error = 6 Or @error = 0) Then ; The data is not found
                _ArrayAdd($HashArrayName, $B[$SplitOffset + 1])
                _ArrayAdd($HashArrayCount, '1')
            ElseIf $R = -1 Then ; A search error occurred
                MsgBox(16, "SplitRawData() search error", _
                        "SplitRawData() search error" & _
                        " R: " & $R & "  Error: " & @error & @CRLF & _
                        $B[$SplitOffset + 1])
                Return
            Else ; The data is found
                $HashArrayCount[$R] += 1
            EndIf
            If CheckAbort(@ScriptLineNumber) <> 0 Then Return
        Next

        LogFile(@ScriptLineNumber & " Now output to Listview (RadioGroup):  " & Timer())

        For $x = 0 To UBound($HashArrayName) - 1
            If $HashArrayCount[$x] > 1 Then
                $T = _GUICtrlListView_AddItem($ListView, $HashArrayName[$x] & "  (" & $HashArrayCount[$x] & ")")
                _GUICtrlListView_AddSubItem($ListView, $T, _GUICtrlListView_GetItemCount($ListView) - 1, $Row)
                For $A = 0 To UBound($ListResultsArray) - 1
                    $B = StringSplit($ListResultsArray[$A], "~")
                    If $B[0] <> 8 Then ContinueLoop
                    ; ConsoleWrite(@ScriptLineNumber & " >>" & $HashArrayName[$x] & "<>" & $B[5] & "<<>>" & $SplitOffset & @CRLF)
                    If StringCompare($HashArrayName[$x], $B[$SplitOffset + 1]) = 0 Then
                        $T = _GUICtrlListView_AddItem($ListView, "")
                        _GUICtrlListView_AddSubItem($ListView, $T, $B[2], $FullPath)
                        _GUICtrlListView_AddSubItem($ListView, $T, $B[3], $Name)
                        _GUICtrlListView_AddSubItem($ListView, $T, $B[4], $Ext)
                        _GUICtrlListView_AddSubItem($ListView, $T, $B[5], $Size)
                        _GUICtrlListView_AddSubItem($ListView, $T, $B[6], $Date)
                        _GUICtrlListView_AddSubItem($ListView, $T, $B[7], $CRC)
                        _GUICtrlListView_AddSubItem($ListView, $T, $B[8], $Attr)
                        _GUICtrlListView_AddSubItem($ListView, $T, _GUICtrlListView_GetItemCount($ListView) - 1, $Row)
                    EndIf
                    If CheckAbort(@ScriptLineNumber) <> 0 Then Return
                Next
            EndIf
        Next
    Else
        ;if $RadioNoGroup is selected than this section processes and adds the data to the listview
        LogFile(@ScriptLineNumber & " Now output to Listview (RadioNoGroup):  " & Timer())
        _GUICtrlListView_SetColumn($ListView, 0, "No grouping")
        For $A = 0 To UBound($ListResultsArray) - 1
            $B = StringSplit($ListResultsArray[$A], "~")
            If $B[0] <> 8 Then ContinueLoop
            $T = _GUICtrlListView_AddItem($ListView, $B[1])
            _GUICtrlListView_AddSubItem($ListView, $T, $B[2], $FullPath)
            _GUICtrlListView_AddSubItem($ListView, $T, $B[3], $Name)
            _GUICtrlListView_AddSubItem($ListView, $T, $B[4], $Ext)
            _GUICtrlListView_AddSubItem($ListView, $T, $B[5], $Size)
            _GUICtrlListView_AddSubItem($ListView, $T, $B[6], $Date)
            _GUICtrlListView_AddSubItem($ListView, $T, $B[7], $CRC)
            _GUICtrlListView_AddSubItem($ListView, $T, $B[8], $Attr)
            _GUICtrlListView_AddSubItem($ListView, $T, _GUICtrlListView_GetItemCount($ListView) - 1, $Row)
            If CheckAbort(@ScriptLineNumber) <> 0 Then Return
        Next
    EndIf
    LogFile(@ScriptLineNumber & " SplitRawData complete:  " & Timer())
EndFunc   ;==>SplitRawData
;-----------------------------------------------
Func SortDataArray($Column)
    LogFile(@ScriptLineNumber & " SortDataArray. Column: " & $Column & "   " & Timer())

    Local $TA1[1]
    Local $TA2[1]
    For $V In $ListResultsArray
        $TA2 = StringSplit($V, "~")
        If $TA2[0] <> 8 Then ContinueLoop
        If $Column = $Size Or $Column = $CRC Then
            _ArrayAdd($TA1, StringFormat("%30s~~~~", $TA2[$Column + 1]) & $V)
        Else
            _ArrayAdd($TA1, $TA2[$Column + 1] & "~~~~" & $V)
        EndIf
    Next

    _ArraySort($TA1)
    LogFile(@ScriptLineNumber & " After ArraySort:  " & Timer())

    ReDim $ListResultsArray[1]

    For $V In $TA1
        $TA2 = StringSplit($V, "~~~~", 1)
        If $TA2[0] <> 2 Then ContinueLoop
        _ArrayAdd($ListResultsArray, $TA2[2])
    Next
    LogFile(@ScriptLineNumber & " After StringSplit:  " & Timer())
    #cs
        ;$ListResultsArray = _ArrayUnique($ListResultsArray)
        ;$ListResultsArray = _ArrayDeleteDupes1($ListResultsArray)
        
        For $i = 0 To UBound($ListResultsArray) - 2
        If StringCompare($ListResultsArray[$i], $ListResultsArray[$i + 1]) = 0 Then $ListResultsArray[$i] = ""
        Next
        LogFile(@ScriptLineNumber & " VerifyFiles: Done unique")
        
        _ArrayDelete($ListResultsArray, 0)
        
        For $STR In $ListResultsArray
        If IsString($STR) = True And StringLen($STR) > 3 Then _GUICtrlListBox_AddString($ListActiveFiles, $STR)
        Next
        
    #ce
    LogFile(@ScriptLineNumber & " After ArrayUnique:  " & Timer())

    _FileWriteFromArray("c:\temp\ResultsArray.txt", $ListResultsArray)

EndFunc   ;==>SortDataArray
;-----------------------------------------------
;This puts commas in the right places in a string and returns the result
Func CommafyString($string)
    If StringLen($string) > 3 Then $string = _StringInsert($string, ",", -3)
    If StringLen($string) > 7 Then $string = _StringInsert($string, ",", -7)
    If StringLen($string) > 11 Then $string = _StringInsert($string, ",", -11)
    Return $string
EndFunc   ;==>CommafyString
;-----------------------------------------------
; Returns 0 if no abort
Func CheckAbort($LineNumber)
    ; $ButtonMainWorking
    Static $cnt
    $cnt += 1
    If Mod($cnt, 100) = 0 Then GUICtrlSetData($ButtonMainWorking, $cnt / 100)
    If GUICtrlRead($CheckMainAbort) = $GUI_CHECKED Then
        LogFile($LineNumber & " CheckAbort true:  " & Timer())
        GuiDisable($GUI_ENABLE)
        Return -1
    EndIf

    Return 0
EndFunc   ;==>CheckAbort
;-----------------------------------------------
Func GuiDisable($choice) ;$GUI_ENABLE $GUI_DISABLE
    For $x = 1 To 100
        GUICtrlSetState($x, $choice)
    Next
    GUICtrlSetState($CheckMainAbort, $GUI_ENABLE)
    GUICtrlSetState($ButtonMainWorking, $GUI_ENABLE)

EndFunc   ;==>GuiDisable
;-----------------------------------------------
Func About(Const $FormID)
    GuiDisable($GUI_DISABLE)

    LogFile(@ScriptLineNumber & " About " & $FormID)
    Local $D = WinGetPos($FormID)
    Local $WinPos
    If IsArray($D) = True Then
        ConsoleWrite(@ScriptLineNumber & " " & $FormID & @CRLF)
        $WinPos = StringFormat("%s" & @CRLF & "WinPOS: %d  %d " & @CRLF & "WinSize: %d %d " & @CRLF & "Desktop: %d %d ", _
                $FormID, $D[0], $D[1], $D[2], $D[3], @DesktopWidth, @DesktopHeight)
    Else
        $WinPos = "About WinPos ERROR. Check the window name: " & $FormID
    EndIf
    LogFile(@CRLF & $SystemS & @CRLF & $WinPos & @CRLF & "Written by Doug Kaynor!", True)
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>About
;-----------------------------------------------
Func LogFile($string, $ShowMSGBox = False)
    $string = StringReplace($string, "-1", "", 1)
    _FileWriteLog($Log_filename, $string)
    _GUICtrlListBox_SetTopIndex($ListStatus, _GUICtrlListBox_AddString($ListStatus, $string))
    _Debug($string, $ShowMSGBox)
EndFunc   ;==>LogFile
;-----------------------------------------------
Func Help()
    LogFile(@ScriptLineNumber & " Help")
    Local $helpstr = 'Startup Project: ' & @CRLF & _
            "help or ?   Display this help file" & @CRLF & _
            "Debug       Secret things" & @CRLF & @CRLF & _
            "Hot Keys: " & @CRLF & _
            "F1  = Help" & @CRLF & _
            "F11 = Unlock GUI"
    LogFile(@ScriptName & @CRLF & $FileVersion & @CRLF & @CRLF & $helpstr, True)
EndFunc   ;==>Help
;-----------------------------------------------
Func SetDefaults()
    GuiDisable($GUI_DISABLE)
    LogFile(@ScriptLineNumber & " SetDefaults")
    $WorkingFolder = EnvGet("USERPROFILE")
    GUICtrlSetState($RadioInclude, $GUI_UNCHECKED)
    GUICtrlSetState($RadioExclude, $GUI_UNCHECKED)
    GUICtrlSetState($RadioAll, $GUI_CHECKED)
    GUICtrlSetState($CheckIncludeFolders, $GUI_UNCHECKED)
    GUICtrlSetData($ComboGrouping, "Name & Extension")
    GUICtrlSetData($InputFilter, "")
    GUICtrlSetData($InputSearch, "")
    GUICtrlSetState($CheckDontAsk, $GUI_UNCHECKED)
    _GUICtrlListBox_ResetContent($ListResults)
    GUICtrlSetData($InputSelectedProject, $Project_filename, "")

    GuiDisable($GUI_ENABLE)
EndFunc   ;==>SetDefaults
;-----------------------------------------------
Func GetWinPos(Const $FormID)
    Local $D = WinGetPos($FormID)
    If IsArray($D) = True Then
        Return StringFormat("%d,%d,%d,%d", $D[0], $D[1], $D[2], $D[3])
    Else
        Return "GetWinPos ERROR, check the window name: " & $FormID
    EndIf
EndFunc   ;==>GetWinPos
;-----------------------------------------------
Func SaveProject()
    GuiDisable($GUI_DISABLE)
    LogFile(@ScriptLineNumber & " SaveProject")
    $Project_filename = FileSaveDialog("Save Project file", $AuxPath, _
            "All Project (*.prj)|DuplicateFinder.prj (DuplicateFinder.prj)|All files (*.*)", 18, $AuxPath & $ProgramName & ".prj")

    Local $File = FileOpen($Project_filename, 2)
    ; Check if file opened for writing OK
    If $File = -1 Then
        LogFile(@ScriptLineNumber & " SaveProject: Unable to open file for writing: " & $Project_filename)
        LogFile("SaveProject: Unable to open file for writing: " & $Project_filename, True)
        GuiDisable($GUI_ENABLE)
        Return
    EndIf

    FileWriteLine($File, "Valid for " & $ProgramName & " Project")
    FileWriteLine($File, "Project file for " & $ProgramName & "  " & _DateTimeFormat(_NowCalc(), 0))
    FileWriteLine($File, "Help: 1 is enabled, 4 is disabled for checkboxes and radio buttons")
    FileWriteLine($File, "RadioInclude:" & GUICtrlRead($RadioInclude))
    FileWriteLine($File, "RadioExclude:" & GUICtrlRead($RadioExclude))
    FileWriteLine($File, "RadioAll:" & GUICtrlRead($RadioAll))
    FileWriteLine($File, "CheckIncludeFolders:" & GUICtrlRead($CheckIncludeFolders))
    FileWriteLine($File, "InputFilter:" & GUICtrlRead($InputFilter))
    FileWriteLine($File, "InputSearch:" & GUICtrlRead($InputSearch))
    FileWriteLine($File, "CheckDontAsk:" & GUICtrlRead($CheckDontAsk))
    FileWriteLine($File, "Grouping:" & GUICtrlRead($ComboGrouping))
    FileWriteLine($File, "Duplicate:" & GetWinPos("Duplicate"))
    FileWriteLine($File, "Single:" & GetWinPos("Single"))
    FileWriteLine($File, "Multi:" & GetWinPos("Multi"))
    FileWriteLine($File, "Select:" & GetWinPos("Select"))
    FileWriteLine($File, "Rename, move, copy:" & GetWinPos("Rename, move, copy"))

    FileWriteLine($File, "Irfanview:" & $IrfanView)
    FileWriteLine($File, "WorkingFolder:" & $WorkingFolder)
    For $x = 0 To _GUICtrlListBox_GetCount($ListResults) - 1
        Local $tmp = _GUICtrlListBox_GetText($ListResults, $x)
        ConsoleWrite(@ScriptLineNumber & " " & $tmp & @CRLF)
        FileWrite($File, "ListFolders:" & $tmp & @CRLF)
    Next

    FileClose($File)
    LogFile(@ScriptLineNumber & " Done SaveProject " & $Project_filename)
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>SaveProject
;-----------------------------------------------
; WinMove ( "title", "text", x, y [, width [, height[, speed]]] )
Func MoveWindow(Const $FormID, Const $MoveString)
    ConsoleWrite(@ScriptLineNumber & " " & $FormID & " " & $MoveString & @CRLF)
    Local $D = StringSplit($MoveString, ",", 2)
    WinMove($FormID, "", $D[0], $D[1], $D[2], $D[3])
EndFunc   ;==>MoveWindow
;-----------------------------------------------
;This opens and loads the Project file
Func LoadProject($type = "start")
    GuiDisable($GUI_DISABLE)
    LogFile(@ScriptLineNumber & " LoadProject: " & $type)

    If StringCompare($type, "menu") = 0 Then
        $Project_filename = FileOpenDialog("Load project file", $AuxPath, _
                "All project files (*.prj)|DuplicateFinder.prj (DuplicateFinder.prj)|All files (*.*)", 18, $AuxPath & $ProgramName & ".prj")
        ;$Project_filename = FileOpenDialog("Load project file", $AuxPath, _
        ;		"All project files (*.prj)|DuplicateFinder.prj (DuplicateFinder.prj)|All files (*.*)", 18)
    EndIf

    Local $File = FileOpen($Project_filename, 0)
    ; Check if file opened for reading OK
    If $File = -1 Then
        LogFile(@ScriptLineNumber & " OpenProject: Unable to open file for reading: " & $Project_filename)
        MsgBox(16, $Project_filename, "OpenProject: Unable to open file for reading: " & @CRLF & $Project_filename)
        GuiDisable($GUI_ENABLE)
        Return
    EndIf
    _GUICtrlListBox_ResetContent($ListResults)

    LogFile(@ScriptLineNumber & " OpenProject: " & $Project_filename)
    ; Read in the first line to verify the file is of the correct type
    If StringCompare(FileReadLine($File, 1), "Valid for " & $ProgramName & " Project") <> 0 Then
        LogFile(@ScriptLineNumber & " Not an Project file for " & $ProgramName)
        LogFile("Not an Project file for " & $ProgramName, True)
        FileClose($File)
        GuiDisable($GUI_ENABLE)
        Return
    EndIf
    ; Read in lines of text until the EOF is reached
    While 1
        Local $LineIn = FileReadLine($File)
        If @error = -1 Then ExitLoop
        If StringInStr($LineIn, ";") = 1 Then ContinueLoop

        If StringInStr($LineIn, "RadioInclude:") Then GUICtrlSetState($RadioInclude, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "RadioExclude:") Then GUICtrlSetState($RadioExclude, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "RadioAll:") Then GUICtrlSetState($RadioAll, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckIncludeFolders:") Then GUICtrlSetState($CheckIncludeFolders, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "InputFilter:") Then GUICtrlSetData($InputFilter, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "InputSearch:") Then GUICtrlSetData($InputSearch, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckDontAsk:") Then GUICtrlSetState($CheckDontAsk, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "Grouping:") Then GUICtrlSetData($ComboGrouping, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "Irfanview:") Then $IrfanView = StringMid($LineIn, StringInStr($LineIn, ":") + 1)
        If StringInStr($LineIn, "Duplicate:") Then MoveWindow("Duplicate", StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "Single:") Then MoveWindow("Single", StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "Multi:") Then MoveWindow("Multi", StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "Select:") Then MoveWindow("Select", StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "Rename, move, copy:") Then MoveWindow("Rename, move, copy", StringMid($LineIn, StringInStr($LineIn, ":") + 1))

        If StringInStr($LineIn, "ListResults:") Then _GUICtrlListBox_AddString($ListResults, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "WorkingFolder:") Then $WorkingFolder = StringMid($LineIn, StringInStr($LineIn, ":") + 1)
        If StringInStr($LineIn, "ListFolders:") Then _GUICtrlListBox_AddString($ListResults, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
    WEnd
    FileClose($File)
    LogFile(@ScriptLineNumber & " Load Project " & $Project_filename)

    GUICtrlSetData($InputSelectedProject, $Project_filename, "")

    GuiDisable($GUI_ENABLE)
EndFunc   ;==>LoadProject
;-----------------------------------------------
Func MultiGoButtons($Button)
    LogFile(@ScriptLineNumber & "  MultiGoButtons " & $Button)
    If $Button = "M1" Then
        GUICtrlSetData($InputSingle, GUICtrlRead($InputM1))
        GUICtrlSetData($InputS, GUICtrlRead($InputM1a))
    EndIf
    If $Button = "M2" Then
        GUICtrlSetData($InputSingle, GUICtrlRead($InputM2))
        GUICtrlSetData($InputS, GUICtrlRead($InputM2a))
    EndIf
    If $Button = "M3" Then
        GUICtrlSetData($InputSingle, GUICtrlRead($InputM3))
        GUICtrlSetData($InputS, GUICtrlRead($InputM3a))
    EndIf
    If $Button = "M4" Then
        GUICtrlSetData($InputSingle, GUICtrlRead($InputM4))
        GUICtrlSetData($InputS, GUICtrlRead($InputM4a))
    EndIf
    If $Button = "M5" Then
        GUICtrlSetData($InputSingle, GUICtrlRead($InputM5))
        GUICtrlSetData($InputS, GUICtrlRead($InputM5a))
    EndIf

    GUICtrlSetPos($ViewPortS, 120, 40, 200, 200)
    GUICtrlSetImage($ViewPortS, $AuxPath & "Question.jpg")
    ;  GUICtrlSetData($InputS, $ItemSelected)

    Local $ArrayTmp = ConvertImage(GUICtrlRead($InputSingle))
    Local $B = _ArraySearch($ArrayTmp, "Image dimensions =", 0, 0, 0, 1, 1, 0)
    Local $C = StringSplit($ArrayTmp[$B], "=P")
    Local $D = StringSplit($C[2], "x")
    GUICtrlSetPos($ViewPortS, 120, 40, $D[1], $D[2])

    GUICtrlSetImage($ViewPortS, $DupFinderTempJPG)
    ; GUISetState(@SW_HIDE, $MultiForm)
    GUISetState(@SW_SHOW, $SingleForm)
    WinActivate("Single")

    ; GUISetState(@SW_SHOW, $SingleForm)
EndFunc   ;==>MultiGoButtons
;-----------------------------------------------
;This function allows the user to edit or view any file, useful for changing the config file
Func EditText($File = "")
    LogFile(@ScriptLineNumber & " Edit " & $File)
    If $File = "" Then $File = FileOpenDialog("View or Edit a file", @ScriptDir, "(Duplicate*.*)|All (*.*)", 1)
    If @error <> 0 Then Return

    If Not FileExists($File) Then
        LogFile(@ScriptLineNumber & " EditText error. " & $File & " does not exist")
        MsgBox(16, "EditText error", "File does not exist" & @CRLF & $File, 2)
        Return
    EndIf

    Const $edit1 = "c:\program files\notepad++\notepad++.exe"
    Const $edit2 = "c:\program files (x86)\notepad++\notepad++.exe"
    Const $edit3 = "notepad.exe"
    Local $editor = ""

    If FileExists($edit1) Then
        $editor = $edit1
    ElseIf FileExists($edit2) Then
        $editor = $edit2
    Else
        $editor = $edit3
    EndIf
    LogFile(@ScriptLineNumber & " Edit: " & $editor & "  " & $File)

    ShellExecute($editor, $File)

EndFunc   ;==>EditText
;-----------------------------------------------
;Converts the image file to $DupFinderTempJPG
; Also returns a array with file info in it
Func ConvertImage($ImageFile)
    ;ConsoleWrite(@ScriptLineNumber & " " & $ImageFile & @CRLF)
    Local $TempFileOut = $AuxPath & "DUPImageTmp.txt"
    FileDelete($TempFileOut)
    Local $FileArray[1]

    ;MsgBox(32, @ScriptLineNumber, "Hello baby  " & $IrfanView)

    If $IrfanView = "XXX" Then
        $DupFinderTempJPG = "Error.jpg"
        Return
    EndIf

    ; first attempt to convert $ImageFile
    Local $ConvertCmd = $ImageFile & " /silent /aspectratio /resize=(200,200) /resample /convert=" & $DupFinderTempJPG
    Local $RunPID = ShellExecuteWait($IrfanView, $ConvertCmd, "", "open", @SW_HIDE) ; 0 is success, 1 is failure
    LogFile(@ScriptLineNumber & " PID:" & $RunPID & "  ImageFile:" & $ImageFile & "  Dupfile:" & $DupFinderTempJPG)

    If $RunPID = 0 Then ; Convert success. This section creates a info file from $DupFinderTempJPG
        Local $tststr = FileGetShortName($DupFinderTempJPG) & " /info=" & $TempFileOut
        Local $Result = ShellExecuteWait($IrfanView, $tststr, "", "open", @SW_HIDE)
        LogFile(@ScriptLineNumber & " ProcessPhotoSize. ShellExecuteWait results: " & $Result & " " & $tststr)
        ;read the file info into an array
        If _FileReadToArray($TempFileOut, $FileArray) = 0 Then
            LogFile(@ScriptLineNumber & " ConvertImage info error" & $TempFileOut & " could not be opened for reading. Error: " & @error)
            MsgBox(16, @ScriptLineNumber & "ConvertImage info erro", $TempFileOut & " could not be opened for reading")
            Return -1
        EndIf
        Return $FileArray
    Else ; Convert error RunPID is not 0
        ConsoleWrite(@ScriptLineNumber & " " & $ImageFile & @CRLF)
        If FileExists($ImageFile) Then
            ConsoleWrite(@ScriptLineNumber & " " & FileCopy($AuxPath & "Error.jpg", $DupFinderTempJPG, 1) & @CRLF)
        ElseIf StringInStr($ImageFile, 'Recycled') <> 0 Then
            ConsoleWrite(@ScriptLineNumber & " " & FileCopy($AuxPath & "recycled.jpg", $DupFinderTempJPG, 1) & @CRLF)
        ElseIf StringInStr($ImageFile, 'Changed') <> 0 Then
            ConsoleWrite(@ScriptLineNumber & " " & FileCopy($AuxPath & "Changed.jpg", $DupFinderTempJPG, 1) & @CRLF)
        ElseIf StringInStr($ImageFile, 'Deleted') <> 0 Then
            ConsoleWrite(@ScriptLineNumber & " " & FileCopy($AuxPath & "Deleted.jpg", $DupFinderTempJPG, 1) & @CRLF)
        Else
            FileCopy($AuxPath & "Question.jpg", $DupFinderTempJPG, 1)
        EndIf
        Return -1
    EndIf

EndFunc   ;==>ConvertImage
;-----------------------------------------------
Func ProcessRow()
    LogFile(@ScriptLineNumber & " ProcessRow")
    ;Set all viewports and strings to default values
    GUICtrlSetState($CheckM1, $GUI_UNCHECKED)
    GUICtrlSetState($CheckM2, $GUI_UNCHECKED)
    GUICtrlSetState($CheckM3, $GUI_UNCHECKED)
    GUICtrlSetState($CheckM4, $GUI_UNCHECKED)
    GUICtrlSetState($CheckM5, $GUI_UNCHECKED)
    GUICtrlSetPos($ViewPortS, 120, 40, 200, 200)
    GUICtrlSetPos($ViewPortM1, 10, 45, 200, 200)
    GUICtrlSetPos($ViewPortM2, 220, 45, 200, 200)
    GUICtrlSetPos($ViewPortM3, 430, 45, 200, 200)
    GUICtrlSetPos($ViewPortM4, 640, 45, 200, 200)
    GUICtrlSetPos($ViewPortM5, 850, 45, 200, 200)
    GUICtrlSetImage($ViewPortS, $AuxPath & "Question.jpg")

    GUICtrlSetData($InputS, "na")
    GUICtrlSetImage($ViewPortM1, $AuxPath & "Question.jpg")
    GUICtrlSetData($InputM1, "Question.jpg")
    GUICtrlSetData($InputM1a, "na")
    GUICtrlSetImage($ViewPortM2, $AuxPath & "Question.jpg")
    GUICtrlSetData($InputM2, "Question.jpg")
    GUICtrlSetData($InputM2a, "na")
    GUICtrlSetImage($ViewPortM3, $AuxPath & "Question.jpg")
    GUICtrlSetData($InputM3, "Question.jpg")
    GUICtrlSetData($InputM3a, "na")
    GUICtrlSetImage($ViewPortM4, $AuxPath & "Question.jpg")
    GUICtrlSetData($InputM4, "Question.jpg")
    GUICtrlSetData($InputM4a, "na")
    GUICtrlSetImage($ViewPortM5, $AuxPath & "Question.jpg")
    GUICtrlSetData($InputM5, "Question.jpg")
    GUICtrlSetData($InputM5a, "na")

    Local $A = StringSplit(_GUICtrlListView_GetItemTextString($ListView, -1), "|")
    Local $ItemSelected = _GUICtrlListView_GetSelectionMark($ListView)
    ConsoleWrite(@ScriptLineNumber & "  >>" & $A[2] & "<<>>" & $ItemSelected & "<<" & @CRLF)
    Local $ArrayTmp[1]
    Local $B, $C, $D

    If $A[1] = "" Then ;Single
        GUICtrlSetData($InputSingle, $A[2], "")
        $ArrayTmp = ConvertImage($A[2])
        If IsArray($ArrayTmp) Then
            $B = _ArraySearch($ArrayTmp, "Image dimensions =", 0, 0, 0, 1, 1, 0)
            $C = StringSplit($ArrayTmp[$B], "=P")
            $D = StringSplit($C[2], "x")
            GUICtrlSetPos($ViewPortS, 120, 40, $D[1], $D[2])
        EndIf
        ConsoleWrite(@ScriptLineNumber & " Single" & $ViewPortS & " " & $DupFinderTempJPG & @CRLF)
        GUICtrlSetImage($ViewPortS, $DupFinderTempJPG)
        GUISetState(@SW_HIDE, $MultiForm)
        GUISetState(@SW_SHOW, $SingleForm)
        WinActivate("Single")
    Else ;Multi
        If GUICtrlRead($ComboGrouping) = "No grouping" Then $ItemSelected -= 1
        Local $Done = False ; M1

        $A = StringSplit(_GUICtrlListView_GetItemTextString($ListView, $ItemSelected + 1), "|")
        ConsoleWrite(@ScriptLineNumber & " M1 " & $A[2] & @CRLF)
        GUICtrlSetData($InputM1a, $ItemSelected + 1, "")
        $ArrayTmp = ConvertImage($A[2])
        If IsArray($ArrayTmp) Then
            $B = _ArraySearch($ArrayTmp, "Image dimensions =", 0, 0, 0, 1, 1, 0)
            $C = StringSplit($ArrayTmp[$B], "=P")
            $D = StringSplit($C[2], "x")
            GUICtrlSetPos($ViewPortM1, 10, 45, $D[1], $D[2])
        EndIf
        GUICtrlSetImage($ViewPortM1, $DupFinderTempJPG)
        GUICtrlSetData($InputM1, $A[2])

        $A = StringSplit(_GUICtrlListView_GetItemTextString($ListView, $ItemSelected + 2), "|")
        ConsoleWrite(@ScriptLineNumber & " M2 " & $A[2] & @CRLF)
        If ($A[2]) = '' Then $Done = True; M2
        If Not $Done Then
            GUICtrlSetData($InputM2a, $ItemSelected + 2)
            $ArrayTmp = ConvertImage($A[2])
            If IsArray($ArrayTmp) Then
                $B = _ArraySearch($ArrayTmp, "Image dimensions =", 0, 0, 0, 1, 1, 0)
                $C = StringSplit($ArrayTmp[$B], "=P")
                $D = StringSplit($C[2], "x")
                GUICtrlSetPos($ViewPortM2, 220, 45, $D[1], $D[2])
            EndIf
            GUICtrlSetImage($ViewPortM2, $DupFinderTempJPG)
            GUICtrlSetData($InputM2, $A[2])
        EndIf

        $A = StringSplit(_GUICtrlListView_GetItemTextString($ListView, $ItemSelected + 3), "|")
        ConsoleWrite(@ScriptLineNumber & " M3 " & $A[2] & @CRLF)
        If ($A[2]) = '' Then $Done = True; M3
        If Not $Done Then
            GUICtrlSetData($InputM3a, $ItemSelected + 3)
            $ArrayTmp = ConvertImage($A[2])
            If IsArray($ArrayTmp) Then
                $B = _ArraySearch($ArrayTmp, "Image dimensions =", 0, 0, 0, 1, 1, 0)
                $C = StringSplit($ArrayTmp[$B], "=P")
                $D = StringSplit($C[2], "x")
                GUICtrlSetPos($ViewPortM3, 430, 45, $D[1], $D[2])
            EndIf
            GUICtrlSetImage($ViewPortM3, $DupFinderTempJPG)
            GUICtrlSetData($InputM3, $A[2])
        EndIf

        $A = StringSplit(_GUICtrlListView_GetItemTextString($ListView, $ItemSelected + 4), "|")
        ConsoleWrite(@ScriptLineNumber & " M4 " & $A[2] & @CRLF)
        If ($A[2]) = '' Then $Done = True; M4
        If Not $Done Then
            GUICtrlSetData($InputM4a, $ItemSelected + 4)
            $ArrayTmp = ConvertImage($A[2])
            If IsArray($ArrayTmp) Then
                $B = _ArraySearch($ArrayTmp, "Image dimensions =", 0, 0, 0, 1, 1, 0)
                $C = StringSplit($ArrayTmp[$B], "=P")
                $D = StringSplit($C[2], "x")
                GUICtrlSetPos($ViewPortM4, 640, 45, $D[1], $D[2])
            EndIf
            GUICtrlSetImage($ViewPortM4, $DupFinderTempJPG)
            GUICtrlSetData($InputM4, $A[2])
        EndIf

        $A = StringSplit(_GUICtrlListView_GetItemTextString($ListView, $ItemSelected + 5), "|")
        ConsoleWrite(@ScriptLineNumber & " M5 " & $A[2] & @CRLF)
        If ($A[2]) = '' Then $Done = True; M5
        If Not $Done Then
            GUICtrlSetData($InputM5a, $ItemSelected + 5)
            $ArrayTmp = ConvertImage($A[2])
            If IsArray($ArrayTmp) Then
                $B = _ArraySearch($ArrayTmp, "Image dimensions =", 0, 0, 0, 1, 1, 0)
                $C = StringSplit($ArrayTmp[$B], "=P")
                $D = StringSplit($C[2], "x")
                GUICtrlSetPos($ViewPortM5, 850, 45, $D[1], $D[2])
            EndIf
            GUICtrlSetImage($ViewPortM5, $DupFinderTempJPG)
            GUICtrlSetData($InputM5, $A[2])
        EndIf

        If ($A[2]) = '' Then $Done = True
        $A = StringSplit(_GUICtrlListView_GetItemTextString($ListView, $ItemSelected + 6), "|")
        If $Done = False And FileExists($A[2]) Then
            GUICtrlSetData($LabelIndicator, "MORE")
            GUICtrlSetBkColor($LabelIndicator, $COLOR_RED)
        Else
            GUICtrlSetData($LabelIndicator, "ALL")
            GUICtrlSetBkColor($LabelIndicator, 0x00ff00)
        EndIf

        GUISetState(@SW_HIDE, $SingleForm)
        GUISetState(@SW_SHOW, $MultiForm)
        WinActivate("Multi")

    EndIf

EndFunc   ;==>ProcessRow
;-----------------------------------------------
Func RenameMoveCopy($type);dbk  RMC
    LogFile(@ScriptLineNumber & " RenameMoveCopy " & $type)
    GUICtrlSetData($InputRMCInName, GUICtrlRead($InputSingle))

    _GUICtrlComboBox_AddString($ComboRMCOutName, GUICtrlRead($InputRMCInName))
    _GUICtrlComboBox_SetEditText($ComboRMCOutName, GUICtrlRead($InputRMCInName))

    GUISetState(@SW_SHOW, $RMCForm)

    GUISetState(@SW_SHOW, $RMCForm)
    GUICtrlSetState($ButtonRMCCopy, $GUI_HIDE)
    GUICtrlSetState($ButtonRMCMove, $GUI_HIDE)
    GUICtrlSetState($ButtonRMCRename, $GUI_HIDE)
    If StringCompare($type, "C") = 0 Then GUICtrlSetState($ButtonRMCCopy, $GUI_SHOW)
    If StringCompare($type, "M") = 0 Then GUICtrlSetState($ButtonRMCMove, $GUI_SHOW)
    If StringCompare($type, "R") = 0 Then GUICtrlSetState($ButtonRMCRename, $GUI_SHOW)
EndFunc   ;==>RenameMoveCopy
;-----------------------------------------------
; This is where the actions are actually performed
Func RMC($type)
    LogFile(@ScriptLineNumber & " RMC " & $type)
    If MsgBox(4 + 32 + 256, "Rename, Move, Copy " & $type, $type & @CRLF & GUICtrlRead($InputRMCInName) & @CRLF & _
            " to" & @CRLF & _
            _GUICtrlComboBox_GetEditText($ComboRMCOutName) & @CRLF & @CRLF & _
            "Are you sure?") = 6 Then ; 6 = yes,  7 = no
        Local $Res
        If StringCompare($type, "Copy") = 0 Then
            $Res = FileCopy(GUICtrlRead($InputRMCInName), _GUICtrlComboBox_GetEditText($ComboRMCOutName))
        ElseIf StringCompare($type, "Move") = 0 Then
            $Res = FileMove(GUICtrlRead($InputRMCInName), _GUICtrlComboBox_GetEditText($ComboRMCOutName))
        ElseIf StringCompare($type, "Rename") = 0 Then
            $Res = FileMove(GUICtrlRead($InputRMCInName), _GUICtrlComboBox_GetEditText($ComboRMCOutName))
        Else
            MsgBox(16, "Error", "Should not be able to see this")
        EndIf

        If $Res = 0 Then
            MsgBox(16, "Operation failed " & $type, $type & " not successfull" & @CRLF _
                     & "Check that source and destination names are valid")
        Else
            If StringCompare($type, "Copy") <> 0 Then
                _GUICtrlListView_AddSubItem($ListView, GUICtrlRead($InputS), "Changed", 01)
                GUICtrlSetData($InputSingle, "Changed --> " & GUICtrlRead($InputRMCInName))
                GUICtrlSetPos($ViewPortS, 120, 40, 200, 200)
                GUICtrlSetImage($ViewPortS, $AuxPath & "Changed.jpg")
            EndIf
        EndIf
    EndIf
    GUISetState(@SW_HIDE, $RMCForm)

EndFunc   ;==>RMC
;-----------------------------------------------
Func SingleRecycle()
    LogFile(@ScriptLineNumber & " SingleRecycle")
    Global $R = MsgBox(4, "Recycle single", "Are you sure that you want to recycle this item?" & @CRLF & "(Move to recycle bin.)")
    If $R = 6 Then ; yes
        $R = GUICtrlRead($InputSingle)
        If FileExists($R) Then
            If FileRecycle($R) = 1 Then
                _GUICtrlListView_AddSubItem($ListView, GUICtrlRead($InputS), "Recycled", 01)
                GUICtrlSetData($InputSingle, "Recycled --> " & $R)
                GUICtrlSetPos($ViewPortS, 120, 40, 200, 200)
                GUICtrlSetImage($ViewPortS, $AuxPath & "Recycled.jpg")
            EndIf
        EndIf
    EndIf
    GUISetState(@SW_HIDE, $RMCForm)
EndFunc   ;==>SingleRecycle
;-----------------------------------------------
Func SingleDelete()
    LogFile(@ScriptLineNumber & " SingleDelete")
    Global $R = MsgBox(4, "Delete single", "Are you sure that you want to delete this item?" & @CRLF & "(This is permanent!)")
    If $R = 6 Then ; yes
        $R = GUICtrlRead($InputSingle)
        If FileExists($R) Then
            If FileDelete($R) = 1 Then
                _GUICtrlListView_AddSubItem($ListView, GUICtrlRead($InputS), "Deleted", 01)
                GUICtrlSetData($InputSingle, "Deleted --> " & $R)
                GUICtrlSetPos($ViewPortS, 120, 40, 200, 200)
                GUICtrlSetImage($ViewPortS, $AuxPath & "Deleted.jpg")
            EndIf
        EndIf
    EndIf
    GUISetState(@SW_HIDE, $RMCForm)
EndFunc   ;==>SingleDelete
;-----------------------------------------------
Func MultiRecycle()
    LogFile(@ScriptLineNumber & " MultiRecycle")
    Local $R = 6
    If GUICtrlRead($CheckDontAsk) = $GUI_UNCHECKED Then
        $R = MsgBox(4, "Multiple recycle", "Are you sure that you want to recycle these items?" & @CRLF & "(Move to recycle bin.)")
    EndIf
    If $R = 6 Then ; yes
        $R = GUICtrlRead($InputM1)
        If FileExists($R) And GUICtrlRead($CheckM1) = $GUI_CHECKED Then
            If FileRecycle($R) = 1 Then
                _GUICtrlListView_AddSubItem($ListView, GUICtrlRead($InputM1a), "Recycled", 01)
                GUICtrlSetData($InputM1, "Recycled --> " & $R)
                GUICtrlSetPos($ViewPortM1, 10, 45, 200, 200)
                GUICtrlSetImage($ViewPortM1, $AuxPath & "Recycled.jpg")
            EndIf
        EndIf
        $R = GUICtrlRead($InputM2)
        If FileExists($R) And GUICtrlRead($CheckM2) = $GUI_CHECKED Then
            If FileRecycle($R) = 1 Then
                _GUICtrlListView_AddSubItem($ListView, GUICtrlRead($InputM2a), "Recycled", 01)
                GUICtrlSetData($InputM2, "Recycled --> " & $R)
                GUICtrlSetPos($ViewPortM2, 220, 40, 200, 200)
                GUICtrlSetImage($ViewPortM2, $AuxPath & "Recycled.jpg")
            EndIf
        EndIf
        $R = GUICtrlRead($InputM3)
        If FileExists($R) And GUICtrlRead($CheckM3) = $GUI_CHECKED Then
            If FileRecycle($R) = 1 Then
                _GUICtrlListView_AddSubItem($ListView, GUICtrlRead($InputM3a), "Recycled", 01)
                GUICtrlSetData($InputM3, "Recycled --> " & $R)
                GUICtrlSetPos($ViewPortM3, 430, 40, 200, 200)
                GUICtrlSetImage($ViewPortM3, $AuxPath & "Recycled.jpg")
            EndIf
        EndIf
        $R = GUICtrlRead($InputM4)
        If FileExists($R) And GUICtrlRead($CheckM4) = $GUI_CHECKED Then
            _GUICtrlListView_AddSubItem($ListView, GUICtrlRead($InputM4a), "Recycled", 01)
            If FileRecycle($R) = 1 Then
                GUICtrlSetData($InputM4, "Recycled --> " & $R)
                GUICtrlSetPos($ViewPortM4, 640, 40, 200, 200)
                GUICtrlSetImage($ViewPortM4, $AuxPath & "Recycled.jpg")
            EndIf
        EndIf
        $R = GUICtrlRead($InputM5)
        If FileExists($R) And GUICtrlRead($CheckM5) = $GUI_CHECKED Then
            If FileRecycle($R) = 1 Then
                _GUICtrlListView_AddSubItem($ListView, GUICtrlRead($InputM5a), "Recycled", 1)
                GUICtrlSetData($InputM5, "Recycled --> " & $R)
                GUICtrlSetPos($ViewPortM5, 850, 40, 200, 200)
                GUICtrlSetImage($ViewPortM5, $AuxPath & "Recycled.jpg")
            EndIf
        EndIf
    EndIf
EndFunc   ;==>MultiRecycle

;-----------------------------------------------
;Listview column sort and double click
Func WM_NOTIFY($hWnd, $iMsg, $iwParam, $ilParam)
    #forceref $hWnd, $iMsg, $iwParam, $ilParam
    Local $hWndFrom, $iCode, $tNMHDR

    $tNMHDR = DllStructCreate($tagNMHDR, $ilParam)
    $hWndFrom = HWnd(DllStructGetData($tNMHDR, "hWndFrom"))
    $iCode = DllStructGetData($tNMHDR, "Code")
    If @error Then
        ConsoleWrite(@ScriptLineNumber & " WM_NOTIFY error  " & @error & @CRLF)
        Return
    EndIf

    Local $tInfo
    Local $ColumnIndex

    ;_GUICtrlListView_GetSelectedIndices($ListView)
    ;_GUICtrlListView_GetItemSelected($ListView,5)

    Switch $hWndFrom
        Case GUICtrlGetHandle($ListView)
            Switch $iCode
                Case $LVN_COLUMNCLICK
                    $tInfo = DllStructCreate($tagNMLISTVIEW, $ilParam)
                    $ColumnIndex = DllStructGetData($tInfo, "SubItem")
                    _Debug(@ScriptLineNumber & "  " & $ColumnIndex)
                    _ListView_Sort($ColumnIndex)
                Case $NM_CLICK
                    ProcessRow()
                Case Else
                    $tInfo = DllStructCreate($tagNMLISTVIEW, $ilParam)
                    $ColumnIndex = DllStructGetData($tInfo, "SubItem")
                    Local $Result = DllStructGetData($tInfo, 3)
                    If $Result = -114 Then

                        _Debug(@ScriptLineNumber & "  " & $iCode & "  " & $Result)
                    EndIf
            EndSwitch
    EndSwitch

    Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_NOTIFY

;-----------------------------------------------
;Listview column sort part 2
Func _ListView_Sort($cIndex = 0)
    If GUICtrlRead($ComboGrouping) <> "No grouping" Then Return

    Local $iColumnsCount, $iDimension, $iItemsCount, $aItemsTemp, $aItemsText, $iCurPos, $iImgSummand, $i, $j
    $iColumnsCount = _GUICtrlListView_GetColumnCount($ListView)
    $iDimension = $iColumnsCount * 2
    $iItemsCount = _GUICtrlListView_GetItemCount($ListView)

    If $iItemsCount < 1 Then Return

    Local $aItemsTemp[1][$iDimension]

    For $i = 0 To $iItemsCount - 1
        $aItemsTemp[0][0] += 1
        ReDim $aItemsTemp[$aItemsTemp[0][0] + 1][$iDimension]
        $aItemsText = _GUICtrlListView_GetItemTextArray($ListView, $i)
        $iImgSummand = $aItemsText[0] - 1
        For $j = 1 To $aItemsText[0]
            $aItemsTemp[$aItemsTemp[0][0]][$j - 1] = $aItemsText[$j]
            $aItemsTemp[$aItemsTemp[0][0]][$j + $iImgSummand] = _GUICtrlListView_GetItemImage($ListView, $i, $j - 1)
        Next
    Next

    $iCurPos = $aItemsTemp[1][$cIndex]
    _ArraySort($aItemsTemp, 0, 1, 0, $cIndex)
    If StringInStr($iCurPos, $aItemsTemp[1][$cIndex]) Then _ArraySort($aItemsTemp, 1, 1, 0, $cIndex)
    For $i = 1 To $aItemsTemp[0][0]
        For $j = 1 To $iColumnsCount
            _GUICtrlListView_SetItemText($ListView, $i - 1, $aItemsTemp[$i][$j - 1], $j - 1)
            _GUICtrlListView_SetItemImage($ListView, $i - 1, $aItemsTemp[$i][$j + $iImgSummand], $j - 1)
        Next
    Next
EndFunc   ;==>_ListView_Sort
;-----------------------------------------------
Func ChoseFolders()
    LogFile(@ScriptLineNumber & " ChoseFolders")
    Local $tmp = FileSelectFolder("Chose a folder", "", 7, $WorkingFolder)
    If $tmp = "" Then Return
    $WorkingFolder = $tmp
    GUICtrlSetData($InputSelectFolder, $tmp)
EndFunc   ;==>ChoseFolders
;-----------------------------------------------
Func GetFolders()
    LogFile(@ScriptLineNumber & " GetFolders")
    GuiDisable($GUI_DISABLE)
    _GUICtrlTreeView_DeleteAll($TreeViewFolders)
    _GUICtrlListBox_ResetContent($ListResults)
    ReDim $HandleArray[1]

    Local $FolderArray = _FileListToArrayR($WorkingFolder, "*", 2, 1, 0, "", 1)
    _ArrayDelete($FolderArray, 0);
    _ArrayUnique($FolderArray)
    _ArraySort($FolderArray)

    If IsArray($FolderArray) Then
        For $x In $FolderArray
            _ArrayAdd($HandleArray, _GUICtrlTreeView_Add($TreeViewFolders, 0, $x))
        Next
        _ArrayDelete($HandleArray, 0)
    Else
        MsgBox(16, "Path not found", "Path not found:" & @CRLF & $WorkingFolder)
    EndIf
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>GetFolders
;-----------------------------------------------
Func ToggleChecked()
    $ToggleState = Not $ToggleState
    For $x = 0 To UBound($HandleArray) - 1
        _GUICtrlTreeView_SetChecked($TreeViewFolders, $HandleArray[$x], $ToggleState)
    Next
EndFunc   ;==>ToggleChecked
;-----------------------------------------------\
Func Process()
    Local $ArrayTmp[1]
    _GUICtrlListBox_ResetContent($ListResults)
    For $x = 0 To UBound($HandleArray) - 1
        Local $AA = _GUICtrlTreeView_GetChecked($TreeViewFolders, $HandleArray[$x])
        Local $BB = _GUICtrlTreeView_GetText($TreeViewFolders, $HandleArray[$x])
        If $AA Then _ArrayAdd($ArrayTmp, $BB)
    Next

    _ArrayDelete($ArrayTmp, 0)
    CleanPaths($ArrayTmp)

    If IsArray($ArrayTmp) Then
        For $x In $ArrayTmp
            _GUICtrlListBox_AddString($ListResults, $WorkingFolder & "\" & $x)
        Next
    EndIf
EndFunc   ;==>Process
;-----------------------------------------------
Func CleanPaths(ByRef $Array)
    For $x = 0 To UBound($Array) - 1
        For $y = 0 To UBound($Array) - 1
            If StringInStr($Array[$y], $Array[$x]) <> 0 And _
                    StringCompare($Array[$y], $Array[$x]) <> 0 Then
                $Array[$y] = ""
            EndIf
        Next
    Next
    _RemoveBlankLines($Array)
EndFunc   ;==>CleanPaths

;-----------------------------------------------
Func Timer()
    Return Int(TimerDiff($TimeStamp) / 100) / 10 & " seconds"
EndFunc   ;==>Timer
;-----------------------------------------------
