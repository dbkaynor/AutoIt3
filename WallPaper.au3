#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_icon=../icons/BrickWall.ico
#AutoIt3Wrapper_outfile=WallPaper.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Comment=A wallpaper changer
#AutoIt3Wrapper_Res_Description=Wallpaper changer
#AutoIt3Wrapper_Res_Fileversion=1.0.0.39
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=Y
#AutoIt3Wrapper_Res_ProductVersion=666
#AutoIt3Wrapper_Res_LegalCopyright=Copyright 2011 Douglas B Kaynor
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
;#AutoIt3Wrapper_Run_Obfuscator=n
;#AutoIt3Wrapper_Run_cvsWrapper=v
;#Obfuscator_Parameters=/striponly
;#AutoIt3Wrapper_Res_Field=Credits|
;#AutoIt3Wrapper_Run_After=copy "%out%" "..\..\Programs_Updates\AutoIt3Wrapper"
;#NoTrayIcon
;#Tidy_Parameters=/gd /sf
;#Tidy_Parameters=/nsdp /sf
;#Tidy_Parameters=/sci=9
#cs
    This area is used to store things todo, bugs, and other notes
    
    Fixed:
    
    Todo:
    Verify file should sync file counts
    Invalid files type cause odd behavior
    Fix all help calls (F1)
    Hide window before movewin
    Get pictures from remote site(s)
    Filter unwanted file types (ini and so on)
#CE

Opt("MustDeclareVars", 1)
Opt("WinTitleMatchMode", 2)
Opt("GUIResizeMode", 0)

#include <Array.au3>
#include <ButtonConstants.au3>
#include <Color.au3>
#include <ComboConstants.au3>
#include <Constants.au3>
#include <Date.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiEdit.au3>
#include <GUIListBox.au3>
#include <GuiTreeView.au3>
#include <ListBoxConstants.au3>
#include <Misc.au3>
#include <SliderConstants.au3>
#include <StaticConstants.au3>
#include <String.au3>
#include <TreeViewConstants.au3>
#include <WindowsConstants.au3>
#include <_DougFunctions.au3>

_Debug("DBGVIEWCLEAR")
DirCreate("AUXFiles")

Global Const $tmp = StringSplit(@ScriptName, ".")
Global Const $ProgramName = $tmp[1]

;Global $AuxPath = @ScriptDir & "\AUXFiles\"
Global Const $FileVersion = "  Ver: " & FileGetVersion(@AutoItExe, "Fileversion")
Global Const $i_view32_ini = $AuxPath & "i_view32.ini"

Global $Option_filename = $AuxPath & "WallPaper-" & @ComputerName & ".opt"
Global $Picture_filename = $AuxPath & "WallPaper-" & @ComputerName & ".lst"
Global Const $Used_filename = $AuxPath & "WallPaper-" & @ComputerName & ".usd"
Global Const $Log_filename = $AuxPath & "WallPaper-" & @ComputerName & ".log"
Global $Filter_filename = $AuxPath & "WallPaper-" & @ComputerName & ".flt"
Global $WallpaperTempJPG = $AuxPath & "WallpaperTemp.jpg"
Global Const $WallpaperJPG = $AuxPath & "Wallpaper.jpg"
Global Const $WallpaperMainInfo = $AuxPath & "WallpaperM.inf"
Global Const $WallpaperSelectInfo = $AuxPath & "WallpaperS.inf"
Global Const $WallpaperTempInfo = $AuxPath & "WallpaperT.inf"
Global Const $FileHelp = $AuxPath & "WallpaperToDoList.txt"

;Global $ResultLocation = ""
Global $WorkingFolder = "JUNK"
Global $IrfanView = "junk"
Global $SystemS = $ProgramName & @CRLF & $FileVersion & @CRLF & @OSVersion & @CRLF & @OSServicePack & @CRLF & @OSType & @CRLF & @OSArch

Global $FontName = "Courier New"
Global $FontColorFG = 0x00FFFF
Global $FontColorFG_REF = 16776960
Global $FontColorBG = 0
Global $FontPointSize = 12
Global $FontWeight = 700
Global $FontItalic = 0
Global $FontUnderline = 0
Global $FontStrikethru = 0

Global $SavedTime = -9
Global $CurrentTime = 0
Global $TimeDiff = 0
Global $Running = False
Global $Pause = False
Global $Debug = False
Global $Hide = False
Global $CurrentFile = 'Unknown'

Global $FolderHandleArray[1]
Global $FileHandleArray[1]

If _Singleton($ProgramName, 1) = 0 Then
    LogFile(@ScriptLineNumber & " " & $ProgramName & " is already running!", True)
    Exit
EndIf

FileDelete($WallpaperTempJPG)
FileDelete($WallpaperMainInfo)
FileDelete($WallpaperSelectInfo)

Opt("TrayMenuMode", 3)
Opt("TrayOnEventMode", 1)
;Opt("TrayIconDebug", 1)
TraySetIcon("..\icons\BrickWall.ico")
Global $TrayChangeWallpaper = TrayCreateItem("Change wallpaper")
TrayItemSetOnEvent(-1, "ChangeWallpaper")
Global $TrayStartRunning = TrayCreateItem("Start running")
TrayItemSetOnEvent(-1, "StartRunning")
Global $TrayTogglePause = TrayCreateItem("Toggle pause")
TrayItemSetOnEvent(-1, "TogglePause")
TrayCreateItem("")
Global $TrayShowMainForm = TrayCreateItem("Show main form")
TrayItemSetOnEvent(-1, "ShowMainForm")
Global $TrayShowSelectForm = TrayCreateItem("Show select form")
TrayItemSetOnEvent(-1, "ShowSelectForm")
Global $TrayShowOptionsForm = TrayCreateItem("Show options form")
TrayItemSetOnEvent(-1, "ShowOptionsForm")
Global $TrayHideAllForms = TrayCreateItem("Hide all forms")
TrayItemSetOnEvent(-1, "HideAllForms")
TrayCreateItem("")
Global $TrayShowInfo = TrayCreateItem("Show main info")
TrayItemSetOnEvent(-1, "ShowMainInfo")
Global $TrayEditCurrentFile = TrayCreateItem("Edit current file")
TrayItemSetOnEvent(-1, "EditCurrentFile")
Global $TrayGoToCurrentFolder = TrayCreateItem("Go to current folder")
TrayItemSetOnEvent(-1, "TaskGoToCurrentFolder")
Global $TrayGoToEXEFolder = TrayCreateItem("Go to EXE")
TrayItemSetOnEvent(-1, "TaskGoToEXEFolder")
Global $TrayReloadCurrentFile = TrayCreateItem("Reload current file")
TrayItemSetOnEvent(-1, "ReloadCurrentFile")
TrayCreateItem("")
Global $TrayGUIEnable = TrayCreateItem("GUI Enable")
TrayItemSetOnEvent(-1, "GUI_Enable")
Global $TrayGUIFind = TrayCreateItem("Find GUI")
TrayItemSetOnEvent(-1, "FindGUI")
Global $exit = TrayCreateItem("Exit")
TrayItemSetOnEvent(-1, "ExitEvent")

FileInstall(".\AUXFiles\Wallpaper.jpg", $WallpaperJPG)
FileInstall(".\AUXFiles\i_view32.ini", $i_view32_ini)

LogFile(@ScriptLineNumber & " Command line arguments: " & $CmdLine[0] & "  " & $CmdLineRaw)

For $x = 1 To $CmdLine[0]
    ConsoleWrite($x & " >> " & $CmdLine[$x] & @CRLF)
    Select
        Case StringInStr($CmdLine[$x], "help") > 0 Or StringInStr($CmdLine[$x], "?") > 0
            Help("cmdline")
            Exit
        Case StringInStr($CmdLine[$x], "debug") > 0
            $Debug = True
        Case StringInStr($CmdLine[$x], "run") > 0
            $Running = True
        Case StringInStr($CmdLine[$x], "hide") > 0
            $Hide = True
        Case Else
            LogFile(@ScriptLineNumber & " Unknown cmdline option found: >>" & $CmdLine[$x] & "<<", True)
            Exit
    EndSelect
Next

Global $temp ; this is a junk varible for general purpose use

; Main form
Global $MainForm = GUICreate($ProgramName & "  " & @ComputerName & "  " & $FileVersion, 520, 160, 10, 10)
GUISetFont(10, 400, 0, "Courier New")
Global $ButtonChangeWallpaper = GUICtrlCreateButton("Change wallpaper", 10, 10, 160, 50)
GUICtrlSetTip(-1, "Change the current wallpaper")
Global $ButtonOptions = GUICtrlCreateButton("Options", 170, 10, 130, 25)
GUICtrlSetTip(-1, "Go to the options menu")
Global $ButtonSelectFilesMain = GUICtrlCreateButton("Select files", 170, 40, 130, 25)
GUICtrlSetTip(-1, "Go to the file select menu")
Global $ButtonStart = GUICtrlCreateButton("Start", 10, 70, 60, 25)
GUICtrlSetTip(-1, "Start the automatic wallpaper changer")
Global $ButtonStop = GUICtrlCreateButton("Stop", 70, 70, 60, 25)
GUICtrlSetTip(-1, "Stop the automatic wallpaper changer")
Global $ButtonHide = GUICtrlCreateButton("Hide", 130, 70, 60, 25)
GUICtrlSetTip(-1, "Hide forms, move to tray")

Global $ButtonPause = GUICtrlCreateButton("Pause", 190, 70, 60, 25)
GUICtrlSetTip(-1, "Pause operations")

Global $ButtonAboutMain = GUICtrlCreateButton("About", 300, 10, 60, 25)
GUICtrlSetTip(-1, "About the program and some Debug stuff")
Global $ButtonHelpMain = GUICtrlCreateButton("Help", 300, 40, 60, 25)
GUICtrlSetTip(-1, "Display help information")

Global $ButtonInfoMain = GUICtrlCreateButton("Picture info", 365, 10, 145, 25)
GUICtrlSetTip(-1, "Show file information")
Global $ButtonEditMain = GUICtrlCreateButton("Edit picture", 365, 40, 145, 25)
GUICtrlSetTip(-1, "Edit current file")
Global $ButtonReloadMain = GUICtrlCreateButton("Reload picture", 365, 70, 145, 25)
GUICtrlSetTip(-1, "Reload current file")
Global $ButtonFolderMain = GUICtrlCreateButton("Go to folder", 365, 100, 145, 25)
GUICtrlSetTip(-1, "Go to folder ")

Global $ButtonExitMain = GUICtrlCreateButton("Exit", 365, 130, 145, 25)
GUICtrlSetTip(-1, "Exit the program")

Global $LabelCountMain = GUICtrlCreateLabel("Count", 260, 70, 100, 20, $SS_SUNKEN)
GUICtrlSetTip(-1, "Wallpaper timer display")
Global $EditStatusMain = GUICtrlCreateEdit("Status", 10, 110, 345, 40, $WS_HSCROLL)
GUICtrlSetTip(-1, "Name of current wallpaper")
GUICtrlSetData($EditStatusMain, $ProgramName & " " & @ComputerName & " " & $FileVersion)
; "text", left, top, width, height

; Options form
#region ### START Koda GUI section ### Form=C:\Program Files\AutoIt3\Dougs\OptionForm.kxf
Global $OptionForm = GUICreate("Select options", 600, 400, 10, 10)
Global $ButtonDoneOption = GUICtrlCreateButton("Done", 464, 10, 100, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Close option form and return to main form")
Global $ButtonInstallAutostart = GUICtrlCreateButton("Install Autostart", 464, 40, 100, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Install the program to load on startup")
Global $ButtonUninstallAutostart = GUICtrlCreateButton("Uninstall Autostart", 464, 70, 100, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Uninstall the program from load on startup")
Global $ButtonSetDefaults = GUICtrlCreateButton("Restore defaults", 464, 100, 100, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Set all options to factory default")
Global $ButtonSaveOptions = GUICtrlCreateButton("Save options", 464, 130, 100, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Save current options to an opt file")
Global $ButtonLoadOptions = GUICtrlCreateButton("Load options", 464, 160, 100, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Load options from an opt file")
Global $ButtonHelpOptions = GUICtrlCreateButton("Help", 464, 190, 100, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Show program help")
Global $ButtonAboutOptions = GUICtrlCreateButton("About", 464, 220, 100, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Show program about info")
Global $ButtonEditOptions = GUICtrlCreateButton("Edit text", 464, 250, 100, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Edit a text file")

Global $InputIrfanviewFolder = GUICtrlCreateInput("poo", 10, 370, 240, 20, $ES_AUTOHSCROLL)
GUICtrlSetTip(-1, "Edit the path to Irfanview")
Global $ButtonIrfanviewFolder = GUICtrlCreateButton("Save", 260, 370, 40, 20)
GUICtrlSetTip(-1, "Save the path to Irfanview")

Global $ButtonHideAllForms = GUICtrlCreateButton("Hide all forms", 464, 280, 100, 25, $WS_GROUP)
GUICtrlSetTip(-1, "Hide all forms, Use F9 to restore")

Global $CheckboxCheckForUsed = GUICtrlCreateCheckbox("Check for already used", 310, 290, 150, 20)

; "text", left, top, width, height
Global $LabelMinPictFileSize = GUICtrlCreateLabel("Min picture file size (Kb)", 310, 320, 120, 20, $SS_SUNKEN)
Global $SliderMinPictFileSize = GUICtrlCreateSlider(435, 320, 80, 25, BitOR($TBS_AUTOTICKS, $TBS_NOTICKS))
GUICtrlSetLimit(-1, 100, 1)
GUICtrlSetData(-1, 5)
Global $LabelMinPictFileSizeValue = GUICtrlCreateLabel("", 520, 320, 40, 20, $SS_SUNKEN)

Global $LabelMaxLogFileSize = GUICtrlCreateLabel("Max log file size (Kb)", 310, 350, 120, 20, $SS_SUNKEN)
Global $SliderMaxLogFileSize = GUICtrlCreateSlider(435, 350, 80, 25, BitOR($TBS_AUTOTICKS, $TBS_NOTICKS))
GUICtrlSetLimit(-1, 100, 10)
GUICtrlSetData(-1, 5)
Global $LabelMaxLogFileSizeValue = GUICtrlCreateLabel("", 520, 350, 40, 20, $SS_SUNKEN)

Global $GroupNameDisplayOptions = GUICtrlCreateGroup("Name display options", 10, 10, 290, 180)
Global $CheckboxDisplayName = GUICtrlCreateCheckbox("Display name", 27, 35, 105, 20)
Global $CheckboxTransparent = GUICtrlCreateCheckbox("Transparent", 27, 65, 90, 20)
Global $GroupPosition = GUICtrlCreateGroup("Position", 27, 100, 90, 65)
Global $RadioTop = GUICtrlCreateRadio("Top", 37, 115, 65, 20)
Global $RadioBottom = GUICtrlCreateRadio("Bottom", 37, 135, 65, 20)
GUICtrlCreateGroup("", -99, -99, 1, 1)

Global $LabelFontSample = GUICtrlCreateLabel("Font Sample", 130, 30, 160, 90, $SS_SUNKEN)
Global $ButtonSetFont = GUICtrlCreateButton("Set font", 130, 130, 160, 25, $WS_GROUP)
Global $ButtonTextBackground = GUICtrlCreateButton("Text Background", 130, 160, 160, 25, $WS_GROUP)

GUICtrlCreateGroup("", -99, -99, 1, 1)
Global $GroupScalingOptions = GUICtrlCreateGroup("Scaling Options", 310, 10, 145, 125)
Global $CheckboxAspect = GUICtrlCreateCheckbox("Maintain aspect", 325, 30, 100, 20)
Global $CheckboxShrinkToFit = GUICtrlCreateCheckbox("Shrink to fit", 325, 50, 80, 20)
Global $CheckboxEnlargeToFit = GUICtrlCreateCheckbox("Enlarge to fit", 325, 70, 80, 20)

;"text", left, top, width, height
Global $LabelFitPercent = GUICtrlCreateLabel("Fit Percent", 320, 100, 55, 20, $SS_SUNKEN)
Global $SliderFitPercent = GUICtrlCreateSlider(375, 100, 50, 25, BitOR($TBS_AUTOTICKS, $TBS_NOTICKS))
GUICtrlSetLimit(-1, 110, 10)
GUICtrlSetData(-1, 10)
Global $LabelFitPercentValue = GUICtrlCreateLabel("", 430, 100, 20, 20, $SS_SUNKEN)

GUICtrlCreateGroup("", -99, -99, 1, 1)
Global $GroupChangeOptions = GUICtrlCreateGroup("Change options (minutes)", 10, 200, 290, 160); 10, 10, 289, 180
Global $RadioEveryMinutes = GUICtrlCreateRadio("Every Minutes", 40, 220, 100, 33)
Global $SliderEveryMinutes = GUICtrlCreateSlider(135, 220, 85, 25, BitOR($TBS_AUTOTICKS, $TBS_NOTICKS))
GUICtrlSetLimit(-1, 60, 1)
GUICtrlSetData(-1, 1)
Global $LabelEveryMinutes = GUICtrlCreateLabel("EveryMinutes", 225, 220, 48, 20, $SS_SUNKEN)
Global $RadioEveryFixed = GUICtrlCreateRadio("Every Fixed", 40, 250, 73, 33)
Global $ComboEveryFixed = GUICtrlCreateCombo("1 Minute", 144, 250, 105, 25)
GUICtrlSetData(-1, "1 minute|5 minutes|10 minutes|15 minutes|30 Minutes|1 hour|12 hours|1 day")
Global $RadioRandom = GUICtrlCreateRadio("Random", 40, 280, 60, 25)
Global $SliderMin = GUICtrlCreateSlider(75, 320, 80, 25, BitOR($TBS_AUTOTICKS, $TBS_NOTICKS))
GUICtrlSetLimit(-1, 30, 1)
GUICtrlSetData(-1, 1)
Global $LabelMin = GUICtrlCreateLabel("Min", 50, 320, 25, 20, $SS_SUNKEN)
Global $LabelMinV = GUICtrlCreateLabel("MinV", 120, 295, 30, 20, $SS_SUNKEN)
Global $LabelMax = GUICtrlCreateLabel("Max", 160, 320, 25, 20, $SS_SUNKEN)
Global $SliderMax = GUICtrlCreateSlider(200, 320, 80, 25, BitOR($TBS_AUTOTICKS, $TBS_NOTICKS))
GUICtrlSetLimit(-1, 600, 31)
GUICtrlSetData(-1, 31)
Global $LabelMaxV = GUICtrlCreateLabel("MaxV", 230, 295, 30, 20, $SS_SUNKEN)
Global $LabelValue = GUICtrlCreateLabel("V", 170, 295, 25, 20, $SS_SUNKEN)
GUICtrlCreateGroup("", -99, -99, 1, 1)
Global $GroupModes = GUICtrlCreateGroup("Placement options", 310, 140, 145, 120)
Global $RadioCentered = GUICtrlCreateRadio("Centered", 320, 160, 81, 25)
Global $RadioTiled = GUICtrlCreateRadio("Tiled", 320, 180, 81, 25)
Global $RadioStretched = GUICtrlCreateRadio("Stretched", 320, 200, 81, 25)
Global $RadioStretchedProportional = GUICtrlCreateRadio("Stretched-proportional", 320, 220, 120, 25)
GUICtrlCreateGroup("", -99, -99, 1, 1)

; Select folders form
Global $SelectFoldersFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $SelectFoldersForm = GUICreate("Select folders", 705, 600, 10, 10);, $SelectFoldersFormOptions)
GUISetFont(10, 400, 0, "Courier New")

Global $ButtonSelectTopFolder = GUICtrlCreateButton("Select top folder", 15, 15, 250, 20)
Global $InputWorkingFolder = GUICtrlCreateInput("POO", 270, 15, 410, 20)
Global $ButtonGetFolders = GUICtrlCreateButton("Get folders", 15, 40, 150, 20)
Global $ButtonSelectFolderProcess = GUICtrlCreateButton("Process", 170, 40, 110, 20)
Global $ButtonSelectFolderToggle = GUICtrlCreateButton("Toggle", 285, 40, 110, 20)
Global $ButtonSelectFolderDone = GUICtrlCreateButton("Done", 400, 40, 70, 20)
Global $ButtonSelectFolderAbout = GUICtrlCreateButton("About", 475, 40, 75, 20)
Global $CheckBoxSelectFolderRecursive = GUICtrlCreateCheckbox("Recursive", 555, 40, 140, 20)
GUICtrlCreateLabel("Search", 10, 70, 80, 15)
Global $InputBoxSearchFolders = GUICtrlCreateInput("", 90, 70, 140, 25)
Global $ButtonSearchFolders = GUICtrlCreateButton("Search", 245, 70, 180, 25)
Global $ButtonSearchFoldersAgain = GUICtrlCreateButton("Search again", 450, 70, 200, 25)
Global $ButtonMarkFoundFolders = GUICtrlCreateButton("Mark found", 235, 100, 200, 25)
Global $ButtonToggleAllFolders = GUICtrlCreateButton("Toggle all", 450, 100, 200, 25)
GUICtrlCreateLabel("Folder list", 15, 100, 95, 20, $SS_SUNKEN)

;dbk
Global $TreeViewSelectFolders = GUICtrlCreateTreeView(10, 130, 665, 200, BitOR($TVS_HASBUTTONS, $TVS_HASLINES, $TVS_LINESATROOT, $TVS_DISABLEDRAGDROP, $TVS_SHOWSELALWAYS, $TVS_CHECKBOXES, $WS_GROUP, $WS_TABSTOP), $WS_EX_CLIENTEDGE)
Global $ListSelectFolderResults = GUICtrlCreateList("", 10, 340, 665, 246, BitOR($LBS_DISABLENOSCROLL, $LBS_SORT, $WS_BORDER, $WS_HSCROLL, $WS_VSCROLL))

; Select Files form
;Global $SelectFilesFormOptions = BitOR($WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_POPUP, $WS_POPUPWINDOW, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS)
Global $SelectFilesForm = GUICreate("Select files", 950, 610, 10, 10) ;, $SelectFilesFormOptions)
;_GUICtrlMenu_EnableMenuItem($SelectFilesForm, $SC_close, $MF_grayed)
GUISetFont(10, 400, 0, "Courier New")

Global $CheckSkipSmallFiles = GUICtrlCreateCheckbox("Skip small files", 10, 10, 160, 20)
Global $CheckAppendFiles = GUICtrlCreateCheckbox("Append new files", 180, 10, 160, 20)
GUICtrlCreateLabel("Files in list:", 360, 10, 120, 15)

Global $LabelActiveCount = GUICtrlCreateLabel("", 480, 10, 50, 15)

;dbk Files search
GUICtrlCreateLabel("Search", 10, 35, 70, 15)
Global $InputBoxSearchFiles = GUICtrlCreateInput("", 80, 30, 140, 25)
Global $ButtonSearchFiles = GUICtrlCreateButton("Search", 220, 30, 70, 25)
Global $ButtonSearchFilesAgain = GUICtrlCreateButton("Search again", 300, 30, 120, 25)
Global $ButtonMarkFoundFiles = GUICtrlCreateButton("Mark found", 430, 30, 110, 25)
Global $ButtonToggleAllFiles = GUICtrlCreateButton("Toggle all", 550, 30, 110, 25)

Global $TreeViewActiveFiles = GUICtrlCreateTreeView(10, 60, 680, 480, BitOR($TVS_HASBUTTONS, $TVS_HASLINES, $TVS_LINESATROOT, $TVS_DISABLEDRAGDROP, $TVS_SHOWSELALWAYS, $TVS_CHECKBOXES, $WS_GROUP, $WS_TABSTOP), $WS_EX_CLIENTEDGE)
GUICtrlSetTip(-1, "list active files")

Global $EditStatusSelect = GUICtrlCreateEdit("Status", 10, 550, 770, 50, $WS_HSCROLL)
GUICtrlSetTip(-1, "Name of current wallpaper")
Global $ButtonInfoSelect = GUICtrlCreateButton("Preview info", 800, 560, 130, 25)
GUICtrlSetTip(-1, "Show file information")
Global $ButtonSelectFolder = GUICtrlCreateButton("Select folder", 700, 10, 150, 25)

Global $ButtonGetFiles = GUICtrlCreateButton("Get files", 700, 40, 150, 25)
Global $ButtonRemoveChecked = GUICtrlCreateButton("Remove checked", 700, 70, 150, 25)
Global $ButtonRemoveAllFiles = GUICtrlCreateButton("Remove all", 700, 130, 150, 25)
Global $ButtonVerifyFiles = GUICtrlCreateButton("Verify files", 700, 160, 150, 25)
Global $ButtonEditPicture = GUICtrlCreateButton("Edit picture", 700, 190, 150, 25)
Global $ButtonOpenFolder = GUICtrlCreateButton("Go to folder", 700, 220, 150, 25)
Global $ButtonFilterFileList = GUICtrlCreateButton("Filter File List", 700, 250, 150, 25)
Global $ButtonGotoEXE = GUICtrlCreateButton("Go to EXE", 850, 250, 90, 25)
Global $ButtonDoneSelect = GUICtrlCreateButton("Done", 850, 10, 90, 25)
Global $ButtonSaveList = GUICtrlCreateButton("Save", 850, 40, 90, 25)
Global $ButtonLoadList = GUICtrlCreateButton("Load", 850, 70, 90, 25)
Global $ButtonEditList = GUICtrlCreateButton("Edit text", 850, 100, 90, 25)
Global $ButtonCancelSelect = GUICtrlCreateButton("Cancel", 850, 130, 90, 25)
Global $ButtonAboutSelect = GUICtrlCreateButton("About", 850, 160, 90, 25)
Global $ButtonHelpSelect = GUICtrlCreateButton("Help", 850, 190, 90, 25)
Global $ButtonPreviewSelect = GUICtrlCreateButton("Preview", 850, 220, 90, 25)
Global $ButtonSetAsWallpaper = GUICtrlCreateButton("Set as wallpaper", 700, 290, 200, 25)

LogFile(@ScriptLineNumber & " " & "-----------------" & $ProgramName & " Started -----------------")
If FileExists($WallpaperJPG) = False Then
    LogFile(@ScriptLineNumber & " " & $WallpaperJPG & " does not exist")
    MsgBox(16, $WallpaperJPG, "File does not exist" & @CRLF & $WallpaperJPG)
EndIf
;left, top, width, height
Global $PicPreview = GUICtrlCreatePic($WallpaperJPG, 700, 340, 200, 200)

;GUISetHelp("notepad c:\config.sys", $MainForm) ; Need a help file to call here
;GUISetHelp("notepad c:\Fvar.bat", $OptionForm) ; Need a help file to call here
;GUISetHelp("notepad c:\autoexec.bat", $SelectFilesForm)

;If Not $Hide Then SplashImageOn("Wallpaper is starting. Please wait.", $AuxPath & "Wallpaper.jpg", -1, -1, -1, -1, 18)

SetDefaults()
LoadOptions("Start")
If Not FileExists($WorkingFolder) Then $WorkingFolder = _AddSlash2PathString(EnvGet("USERPROFILE"))
GUICtrlSetData($InputWorkingFolder, $WorkingFolder)
GUICtrlSetData($InputIrfanviewFolder, $IrfanView)
TestForIrfanView()
TestForRequiredFiles()

LoadList("Start")
VerifyFileCounts()
If $Hide = True Then
    HideAllForms()
Else
    ShowMainForm()
EndIf

Const $edit1 = "c:\program files\notepad++\notepad++.exe"
Const $edit2 = "c:\program files (x86)\notepad++\notepad++.exe"
Const $edit3 = "notepad.exe"
Global $editor = ""

If FileExists($edit1) Then
    $editor = $edit1
ElseIf FileExists($edit2) Then
    $editor = $edit2
Else
    $editor = $edit3
EndIf
LogFile(@ScriptLineNumber & " Text editor: " & $editor)

GUISetHelp($editor & " " & $FileHelp, $MainForm)

;HotKeySet("%{f11}", "ChangeWallpaper")
;HotKeySet("%E", "GUI_Enable")

;SplashOff()

While 1
    Global $nMsg = GUIGetMsg(1)
    Switch $nMsg[0]
        Case $GUI_EVENT_CLOSE
            Exit
            ; Main form
        Case $GUI_EVENT_RESIZED

        Case $ButtonChangeWallpaper
            ChangeWallpaper()
        Case $ButtonStart
            StartRunning()
        Case $ButtonStop
            $Running = False
        Case $ButtonHide
            HideAllForms()
        Case $ButtonPause
            TogglePause()
        Case $ButtonOptions
            ShowOptionsForm()
        Case $ButtonSelectFilesMain
            ShowSelectForm()
        Case $ButtonHelpMain
            Help($ProgramName)

        Case $ButtonAboutMain
            About($ProgramName)
        Case $ButtonExitMain
            Exit

            ; set options
        Case $RadioEveryMinutes
            HandleChangeOptions("EveryMinutes")
        Case $SliderEveryMinutes
            GUICtrlSetData($LabelEveryMinutes, GUICtrlRead($SliderEveryMinutes))
            HandleChangeOptions("EveryMinutes")
        Case $RadioEveryFixed
            HandleChangeOptions("EveryFixed")
        Case $ComboEveryFixed
            HandleChangeOptions("EveryFixed")
        Case $RadioRandom
            HandleChangeOptions("Random")
        Case $ButtonSetDefaults
            SetDefaults()
        Case $ButtonSaveOptions
            SaveOptions()
        Case $ButtonLoadOptions
            LoadOptions("menu")
        Case $ButtonEditOptions
            EditText()
        Case $CheckboxDisplayName
            HandleDisplayName()

        Case $ButtonTextBackground
            $FontColorBG = _ChooseColor(2)
            HandleDisplayName()
        Case $ButtonSetFont
            Global $FontInfo = _ChooseFont($FontName, $FontPointSize, $FontColorFG_REF, $FontWeight, $FontItalic, $FontUnderline, $FontStrikethru)
            If IsArray($FontInfo) Then
                ;_ArrayDisplay($FontInfo, @ScriptLineNumber)
                $FontName = $FontInfo[2]
                $FontPointSize = $FontInfo[3]
                $FontWeight = $FontInfo[4]
                $FontColorFG = $FontInfo[7]
                $FontColorFG_REF = $FontInfo[5]
                $FontItalic = BitAND($FontInfo[1], 2)
                $FontUnderline = BitAND($FontInfo[1], 4)
                $FontStrikethru = BitAND($FontInfo[1], 8)
                HandleDisplayName()
            EndIf

        Case $SliderMinPictFileSize
            GUICtrlSetData($LabelMinPictFileSizeValue, GUICtrlRead($SliderMinPictFileSize))

        Case $SliderMaxLogFileSize
            GUICtrlSetData($LabelMaxLogFileSizeValue, GUICtrlRead($SliderMaxLogFileSize))

        Case $CheckboxTransparent
            HandleDisplayName()
        Case $RadioBottom
            HandleDisplayName()
        Case $RadioTop
            HandleDisplayName()
        Case $CheckboxEnlargeToFit
            HandleSizetoFit()
        Case $CheckboxShrinkToFit
            HandleSizetoFit()

        Case $ButtonPreviewSelect
            ShowPreview()
        Case $TreeViewActiveFiles
            ShowPreview()
        Case $SliderFitPercent
            GUICtrlSetData($LabelFitPercentValue, GUICtrlRead($SliderFitPercent))
        Case $SliderMin
            GUICtrlSetData($LabelMinV, GUICtrlRead($SliderMin))
        Case $SliderMax
            GUICtrlSetData($LabelMaxV, GUICtrlRead($SliderMax))
        Case $ButtonHelpOptions
            Help("Select options")
        Case $ButtonAboutOptions
            About("Select options")

        Case $InputIrfanviewFolder
            $IrfanView = GUICtrlRead($InputIrfanviewFolder)
            ConsoleWrite(@ScriptLineNumber & " " & $IrfanView & @CRLF)
        Case $ButtonIrfanviewFolder
            $IrfanView = GUICtrlRead($InputIrfanviewFolder)
            If FileExists($IrfanView) Then
                MsgBox(48, "Success", "Irfanview path updated" & @CRLF & $IrfanView)
            Else
                MsgBox(48, "Fail", "Irfanview path is not valid" & @CRLF & $IrfanView)
            EndIf

        Case $ButtonDoneOption
            GUISetState(@SW_HIDE, $OptionForm)
            GUISetState(@SW_SHOW, $MainForm)
            GUISetState(@SW_HIDE, $SelectFilesForm)
            GUISetState(@SW_HIDE, $SelectFoldersForm)

        Case $ButtonHideAllForms
            HideAllForms()

            ; select folders
        Case $ButtonSelectTopFolder
            ChoseFolders()
        Case $ButtonGetFolders
            $WorkingFolder = _AddSlash2PathString(GUICtrlRead($InputWorkingFolder))
            GetFolders()
        Case $ButtonSelectFolderProcess
            ProcessFolders()
        Case $ButtonSelectFolderToggle
            ToggleFoldersChecked()
        Case $ButtonSelectFolderDone
            GUISetState(@SW_HIDE, $OptionForm)
            GUISetState(@SW_HIDE, $MainForm)
            GUISetState(@SW_SHOW, $SelectFilesForm)
            GUISetState(@SW_HIDE, $SelectFoldersForm)
        Case $ButtonSelectFolderAbout
            About("Select folders")
        Case $InputBoxSearchFolders
            SearchFolders(True)
        Case $ButtonSearchFolders
            SearchFolders(True)
        Case $ButtonSearchFoldersAgain
            SearchFolders(False)
        Case $ButtonMarkFoundFolders
            MarkAllFolderFound()
        Case $ButtonToggleAllFolders
            ToggleAllFoldersChecks()

        Case $ButtonSelectFolder
            ; SelectFolder()
            GUISetState(@SW_HIDE, $OptionForm)
            GUISetState(@SW_HIDE, $MainForm)
            GUISetState(@SW_HIDE, $SelectFilesForm)
            GUISetState(@SW_SHOW, $SelectFoldersForm)
        Case $ButtonGetFiles
            GetFiles()
        Case $ButtonRemoveChecked
            RemoveChecked()
        Case $ButtonOpenFolder
            GoToCurrentFolder("select")
        Case $ButtonFilterFileList
            FilterFileList()
        Case $ButtonGotoEXE
            GotoEXEFolder()
        Case $ButtonRemoveAllFiles
            RemoveAll()
        Case $ButtonVerifyFiles
            VerifyFiles()
        Case $ButtonEditPicture
            EditPicture()
        Case $ButtonSearchFiles
            SearchFiles(True)
        Case $ButtonSearchFilesAgain
            SearchFiles(False)
        Case $InputBoxSearchFiles
            SearchFiles(True)
        Case $ButtonMarkFoundFiles
            MarkAllFilesFound()
        Case $ButtonToggleAllFiles
            ToggleAllFilesChecks()

        Case $ButtonInfoMain
            LogFile(@ScriptLineNumber & " EditStatusMain")
            ClipPut(GUICtrlRead($EditStatusMain))
            EditText($WallpaperMainInfo)
        Case $ButtonEditMain
            EditCurrentFile()
        Case $ButtonReloadMain
            ReloadCurrentFile()
        Case $ButtonFolderMain
            GoToCurrentFolder("main")

        Case $ButtonInfoSelect
            LogFile(@ScriptLineNumber & " LabelStatusSelect")
            ShowPreview()
            ClipPut(GUICtrlRead($EditStatusSelect))
            EditText($WallpaperSelectInfo)
        Case $PicPreview
            LogFile(@ScriptLineNumber & " PicPreview")
            ShowPreview()
            ;ClipPut(GUICtrlRead($EditStatusSelect))
            ;EditText($WallpaperSelectInfo)
        Case $ButtonDoneSelect
            GUISetState(@SW_HIDE, $OptionForm)
            GUISetState(@SW_SHOW, $MainForm)
            GUISetState(@SW_HIDE, $SelectFilesForm)
            GUISetState(@SW_HIDE, $SelectFoldersForm)
        Case $ButtonSaveList
            SaveList()
        Case $ButtonLoadList
            LoadList("menu")
        Case $ButtonEditList
            EditText()
        Case $ButtonCancelSelect
            GUISetState(@SW_SHOW, $OptionForm)
            GUISetState(@SW_HIDE, $MainForm)
            GUISetState(@SW_HIDE, $SelectFilesForm)
            GUISetState(@SW_HIDE, $SelectFoldersForm)
        Case $ButtonAboutSelect
            About("Select files")
        Case $ButtonHelpSelect
            Help("Select files")
        Case $ButtonSetAsWallpaper
            SetAsWallpaper()

        Case $ButtonInstallAutostart
            FileCreateShortcut(@ScriptDir & "\Wallpaper.a3x", _
                    @StartupCommonDir & "\Wallpaper.lnk", _
                    @ScriptDir, _
                    "run hide", _
                    "Wallpaper changer", _
                    @ScriptDir & "\..\icons\Brickwall.ico") ; ../icons/BrickWall.ico

        Case $ButtonUninstallAutostart
            ConsoleWrite(@ScriptLineNumber & " StartupCommonDir   " & @StartupCommonDir & @CRLF)
            FileDelete(@StartupCommonDir & "\Wallpaper.lnk")
    EndSwitch

    If $Running Then CheckChangeCounter()
WEnd
;-----------------------------------------------
Func StartRunning()
    If Not FileExists($IrfanView) Then
        $Running = False
        LogFile(@ScriptLineNumber & " StartRunning: " & $IrfanView & " does not exist", True)
        Return
    EndIf
    $Running = True
    $Pause = False
    ChangeWallpaper()
    SetUpDelayTime()
EndFunc   ;==>StartRunning
;-----------------------------------------------
Func FindGUI()
    WinMove("WallPaper ", "", 10, 10)
    WinMove("Select files", "", 10, 10)
    WinMove("Select folders", "", 10, 10)
    WinMove("Select options", "", 10, 10)
EndFunc   ;==>FindGUI
;-----------------------------------------------
Func TogglePause()
    $Pause = Not $Pause
    LogFile(@ScriptLineNumber & " Pause: " & $Pause)
EndFunc   ;==>TogglePause
;-----------------------------------------------
; This function returns true if the screen saver is running
Func TestforScreenSaver()
    Local $var = RegRead("HKEY_CURRENT_USER\Control Panel\Desktop", "SCRNSAVE.EXE")
    Local $temp = StringSplit($var, "\")
    Local $PName = $temp[$temp[0]]
    If ProcessExists($PName) Then Return True
    Return False
EndFunc   ;==>TestforScreenSaver
;-----------------------------------------------
Func CheckChangeCounter()
    If Not FileExists($IrfanView) Then
        $Running = False
        LogFile(@ScriptLineNumber & " CheckChangeCounter: " & $IrfanView & " does not exist", True)
        Return
    EndIf

    If $CurrentFile = "Unknown" Then
        ConsoleWrite(@ScriptLineNumber & " " & $CurrentFile & @CRLF)
        ChangeWallpaper()
        SetUpDelayTime()
    EndIf

    $CurrentTime = TimerDiff($SavedTime) / 1000 ; seconds

    ;This slows down the test attempts
    If Mod(Int($CurrentTime * 400), 10) <> 0 Then Return

    Local $TS
    If StringCompare($CurrentFile, "Unknown") = 0 Then
        $TS = $CurrentFile
    Else
        Local $TA = StringSplit($CurrentFile, "\")
        $TS = ".....\" & $TA[UBound($TA) - 2] & "\" & $TA[UBound($TA) - 1]
    EndIf

    If $CurrentTime > $TimeDiff * 60 Then
        SetUpDelayTime()
        Local $T = TestforScreenSaver()
        If $Pause Or $T Then
            LogFile(@ScriptLineNumber & " Pause: " & $Pause & "    TestforScreenSaver: " & $T)
            Return
        EndIf
        ;ConsoleWrite(@ScriptLineNumber & " " & $CurrentTime & " " & $TimeDiff * 60 & @CRLF)
        ChangeWallpaper()
    EndIf

    If $Pause Then
        GUICtrlSetData($LabelCountMain, "Paused")
        TraySetToolTip(StringFormat("%s  Paused", $TS & @CRLF))
    Else
        GUICtrlSetData($LabelCountMain, StringFormat("%d %d", $CurrentTime, $TimeDiff * 60))
        TraySetToolTip(StringFormat("%s  %d  %d", $TS & @CRLF, $CurrentTime, $TimeDiff * 60))
    EndIf
EndFunc   ;==>CheckChangeCounter
;-----------------------------------------------
Func SetUpDelayTime()
    $SavedTime = TimerInit()

    If GUICtrlRead($RadioEveryMinutes) = $GUI_CHECKED Then
        $TimeDiff = GUICtrlRead($LabelEveryMinutes)
    ElseIf GUICtrlRead($RadioEveryFixed) = $GUI_CHECKED Then
        Local $tmp = StringSplit(GUICtrlRead($ComboEveryFixed), " ", 2)
        If StringInStr($tmp[1], "minute") <> 0 Then
            $TimeDiff = $tmp[0]
        ElseIf StringInStr($tmp[1], "hour") <> 0 Then
            $TimeDiff = $tmp[0] * 60
        ElseIf StringInStr($tmp[1], "day") <> 0 Then
            $TimeDiff = $tmp[0] * 60 * 24
        EndIf
    ElseIf GUICtrlRead($RadioRandom) = $GUI_CHECKED Then
        $TimeDiff = Random(GUICtrlRead($LabelMinV), GUICtrlRead($LabelMaxV), 1)
        GUICtrlSetData($LabelValue, $TimeDiff)
    Else
        LogFile(@ScriptLineNumber & " We should not have gotten here: SetUpDelayTime", True)
    EndIf
EndFunc   ;==>SetUpDelayTime
;-----------------------------------------------
Func RemoveAll()
    GuiDisable($GUI_DISABLE)
    LogFile(@ScriptLineNumber & @ScriptLineNumber & "RemoveAll")
    _GUICtrlTreeView_DeleteAll($TreeViewActiveFiles)
    ReDim $FileHandleArray[1]

    GUICtrlSetData($LabelActiveCount, _GUICtrlTreeView_GetCount($TreeViewSelectFolders))

    GuiDisable($GUI_ENABLE)
EndFunc   ;==>RemoveAll
;-----------------------------------------------
Func RemoveChecked()
    GuiDisable($GUI_DISABLE)
    Local $x = 0
    While $x < UBound($FileHandleArray)
        If _GUICtrlTreeView_GetChecked($TreeViewActiveFiles, $FileHandleArray[$x]) Then
            _GUICtrlTreeView_Delete($TreeViewActiveFiles, $FileHandleArray[$x])
            _ArrayDelete($FileHandleArray, $x)
        Else
            $x = $x + 1
        EndIf
    WEnd

    VerifyFileCounts()
    GUICtrlSetData($LabelActiveCount, _GUICtrlTreeView_GetCount($TreeViewActiveFiles))
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>RemoveChecked
;-----------------------------------------------
Func EditPicture()
    GuiDisable($GUI_DISABLE)
    LogFile(@ScriptLineNumber & " EditPicture")
    Local $x = _GUICtrlTreeView_GetSelection($TreeViewActiveFiles)
    Local $File = _GUICtrlTreeView_GetText($TreeViewActiveFiles, $x)
    ConsoleWrite(@ScriptLineNumber & " " & $x & "   " & $File & @CRLF)
    If FileExists($IrfanView) Then
        LogFile(@ScriptLineNumber & " ShellExecuteWait results: " & ShellExecuteWait($IrfanView, $File, "", "open", @SW_SHOW))
    Else
        LogFile(@ScriptLineNumber & " EditCurrentFile: " & $IrfanView & " does not exist", True)
    EndIf
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>EditPicture
;-----------------------------------------------
Func VerifyFiles()
    GuiDisable($GUI_DISABLE)
    If Not $Hide Then SplashImageOn("Wallpaper is starting. Please wait.", $AuxPath & "Wallpaper.jpg", -1, -1, -1, -1, 18)
    LogFile(@ScriptLineNumber & " VerifyFiles: begin files exist")

    Local $ArrayFailed[1]
    Local $ArrayGood[1]
    If _GUICtrlTreeView_GetCount($TreeViewActiveFiles) = 0 Then
        MsgBox(64, "VerifyFiles error", "No files to verify")
        GuiDisable($GUI_ENABLE)
        Return
    EndIf

    For $x In $FileHandleArray
        Local $temp = _GUICtrlTreeView_GetText($TreeViewActiveFiles, $x)
        ;  ConsoleWrite(@ScriptLineNumber & " " & $x & " >" & $temp & "< " & @CRLF)
        If FileExists($temp) = 0 Then
            _ArrayAdd($ArrayFailed, $temp)
        Else
            _ArrayAdd($ArrayGood, $temp)
        EndIf
    Next

    LogFile(@ScriptLineNumber & " VerifyFiles: files exist complete")

    _ArraySort($ArrayGood)
    LogFile(@ScriptLineNumber & " VerifyFiles: sort complete")

    _GUICtrlTreeView_DeleteAll($TreeViewActiveFiles)

    ReDim $FileHandleArray[1]
    For $i = 0 To UBound($ArrayGood) - 2
        If StringCompare($ArrayGood[$i], $ArrayGood[$i + 1]) = 0 Then $ArrayGood[$i] = ""
    Next
    LogFile(@ScriptLineNumber & " VerifyFiles: Done unique")
    _ArrayDelete($ArrayGood, 0)

    For $STR In $ArrayGood
        If IsString($STR) = True And StringLen($STR) > 3 Then
            _ArrayAdd($FileHandleArray, _GUICtrlTreeView_Add($TreeViewActiveFiles, 0, $STR))
        EndIf
    Next

    _ArrayDelete($FileHandleArray, 0)

    VerifyFileCounts()

    GUICtrlSetData($LabelActiveCount, _GUICtrlTreeView_GetCount($TreeViewActiveFiles))
    LogFile(@ScriptLineNumber & " VerifyFiles: complete")
    SplashOff()
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>VerifyFiles
;-----------------------------------------------
Func VerifyFileCounts()
    LogFile(@ScriptLineNumber & " VerifyFilesCounts")
    Local $a = _GUICtrlTreeView_GetCount($TreeViewActiveFiles)
    Local $b = UBound($FileHandleArray)
    ConsoleWrite(@ScriptLineNumber & " File count verify: " & $a & " " & $b & @CRLF)
    If _GUICtrlTreeView_GetCount($TreeViewActiveFiles) <> UBound($FileHandleArray) Then
        ;MsgBox(48, "File list count mismatch", "TreeViewActiveFiles: " & $a & @CRLF & "FileHandleArray: " & $b, 2)
        LogFile(@ScriptLineNumber & "File list count mismatch: TreeViewActiveFiles:" & $a & "  FileHandleArray:" & $b)
    EndIf
EndFunc   ;==>VerifyFileCounts
;-----------------------------------------------
; This function returns true is the file has already been displayed, false if not
Func CheckForUsed($File)
    If GUICtrlRead($CheckboxCheckForUsed) = $GUI_UNCHECKED Then
        LogFile(@ScriptLineNumber & " CheckForUsed skipped " & $File)
        Return False
    EndIf

    Local $UsedArray
    _FileReadToArray($Used_filename, $UsedArray)
    Local $count = 0
    While 1
        If $count > UBound($UsedArray) - 1 Then ExitLoop
        If StringInStr($UsedArray[$count], $File) <> 0 Then
            LogFile(@ScriptLineNumber & " CheckForUsed TRUE " & $File & " has already been displayed")
            Return True
        EndIf
        $count += 1
    WEnd
    LogFile(@ScriptLineNumber & " CheckForUsed FALSE " & $File & " has not been displayed")
    Return False
EndFunc   ;==>CheckForUsed
;-----------------------------------------------
;Uses a random value to select a JPG
;Checks to verify if it has been displayed
;Tries all files before giving up
Func GetAFileToDisplay()
    LogFile(@ScriptLineNumber & " GetAFileToDisplay")
    Local Const $MaxTries = 50
    Local $count = 0

    ; try to get a random file first $Maxtries times
    Do
        Local $x = Random(0, UBound($FileHandleArray) - 1, 1)
        $CurrentFile = _GUICtrlTreeView_GetText($TreeViewSelectFolders, $FileHandleArray[$x])
        ConsoleWrite(@ScriptLineNumber & " " & $x & " >" & $CurrentFile & "< " & $count & @CRLF)
        $count += 1
        If CheckForused($CurrentFile) = False Then Return $CurrentFile
    Until $count > $MaxTries

    ; now try all files in the list
    For $x In $FileHandleArray
        $CurrentFile = _GUICtrlTreeView_GetText($TreeViewSelectFolders, $x)
        ConsoleWrite(@ScriptLineNumber & " " & $x & " >" & $CurrentFile & "< " & @CRLF)
        If CheckForUsed($CurrentFile) = False Then Return $CurrentFile
    Next

    FileDelete($Used_filename)
    LogFile(@ScriptLineNumber & " " & $Used_filename & ": used list cleared")

    Return $CurrentFile
EndFunc   ;==>GetAFileToDisplay
;-----------------------------------------------
Func ChangeWallpaper()
    If Not FileExists($IrfanView) Then
        LogFile(@ScriptLineNumber & " ChangeWallpaper: " & $IrfanView & " does not exist", True)
        Return
    EndIf

    GuiDisable($GUI_DISABLE)
    LogFile(@ScriptLineNumber & " ChangeWallpaper")
    LogFileLimit()
    VerifyFileCounts()

    Local $count = 0
    Do
        $count += 1
        $CurrentFile = GetAFileToDisplay()
    Until ProcessWallpaper($CurrentFile) Or $count > 10

    LogFile(@ScriptLineNumber & " ChangeWallpaper " & $CurrentFile)
    $SavedTime = TimerInit()
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>ChangeWallpaper
;-----------------------------------------------
Func SetAsWallpaper()
    If Not FileExists($IrfanView) Then
        LogFile(@ScriptLineNumber & " SetAsWallpaper: " & $IrfanView & " does not exist", True)
        Return
    EndIf

    GuiDisable($GUI_DISABLE)
    LogFile(@ScriptLineNumber & " SetAsWallpaper")
    $SavedTime = TimerInit()

    Local $x = _GUICtrlTreeView_GetSelection($TreeViewActiveFiles)
    $CurrentFile = _GUICtrlTreeView_GetText($TreeViewActiveFiles, $x)

    ProcessWallpaper($CurrentFile)
    GUICtrlSetData($EditStatusSelect, $CurrentFile)

    GuiDisable($GUI_ENABLE)
EndFunc   ;==>SetAsWallpaper
;-----------------------------------------------
; Checks the dimensions of picture file
; Pass 2 file name in and creates a text file of data
; $FileSource & $FileTemp are JPG's and FileOutput is a text file
; Returns -1 if the both file dimensions are reduced
; Returns +1 if the both file dimensions are enlarged
; Returns 0 if neither of the above is true

Func CheckPictureInfo($FileSource, $FileTemp, $FileOutput)
    LogFile(@ScriptLineNumber & " CheckPictureInfo " & $FileSource & " " & $FileTemp & " " & $FileOutput)
    If Not FileExists($IrfanView) Then Return
    Local $A1 = ProcessPictureSize($FileSource)
    Local $A2 = ProcessPictureSize($FileTemp)
    _ArrayConcatenate($A1, $A2)
    _FileWriteFromArray($FileOutput, $A1)

    ;$PictureDemensions
    Local $R = _ArraySearch($A1, "Image dimensions =", 0, 0, 0, 1)
    ConsoleWrite(@ScriptLineNumber & " Source: " & $A1[$R] & @CRLF)
    $R = StringSplit($A1[$R], " ", 2)
    ;_ArrayDisplay($R, @ScriptLineNumber)

    Local $S = _ArraySearch($A2, "Image dimensions =", 0, 0, 0, 1)
    ConsoleWrite(@ScriptLineNumber & " Temp: " & $A2[$S] & @CRLF)
    $S = StringSplit($A2[$S], " ", 2)
    ;_ArrayDisplay($S, @ScriptLineNumber)

    ConsoleWrite(@ScriptLineNumber & " " & _SystemLocalTime() & $FileOutput & @CRLF)
    Local $PercentString = StringFormat(" %2.2f  %2.2f", Number($S[3]) / Number($R[3]), Number($S[5]) / Number($R[5]))
    Local $DataString = $R[3] & ">" & $S[3] & "  " & $R[5] & ">" & $S[5]
    If (Number($R[3]) > Number($S[3])) And (Number($R[5]) > Number($S[5])) Then
        ConsoleWrite(@ScriptLineNumber & " File dimensions are both reduced " & $PercentString & "  " & $DataString & @CRLF)
        FileWrite($FileOutput, "File dimensions are both reduced  " & $PercentString & "  " & $DataString)
        Return "reduced " & $PercentString
    ElseIf (Number($R[3]) < Number($S[3])) And (Number($R[5]) < Number($S[5])) Then
        ConsoleWrite(@ScriptLineNumber & " File dimensions are both enlarged " & $PercentString & "  " & $DataString & @CRLF)
        FileWrite($FileOutput, "File dimensions are both enlarged  " & $PercentString & "  " & $DataString)
        Return "enlarged " & $PercentString
    Else
        ConsoleWrite(@ScriptLineNumber & " File dimensions are undetermined " & $PercentString & "  " & $DataString & @CRLF)
        FileWrite($FileOutput, "File dimensions are undetermined " & $PercentString & "  " & $DataString)
        Return "undetermined " & $PercentString
    EndIf
EndFunc   ;==>CheckPictureInfo
;-----------------------------------------------
;Takes a Picture file name and path and returns an array of data
Func ProcessPictureSize($TestFile)
    LogFile(@ScriptLineNumber & " ProcessPictureSize " & $TestFile)
    If Not FileExists($IrfanView) Then Return

    Local $FileArray[1]
    Local $TempFileOut = $AuxPath & "PPSTemp.txt"


    ;this section creates a info file from a Picture file
    If Not FileExists($TestFile) Then
        LogFile(@ScriptLineNumber & " ProcessPictureSize error: " & $TestFile & " does not exist " & @error)
        MsgBox(16, @ScriptLineNumber & " ProcessPictureSize error", $TestFile & " does not exist")
        Return -1
    Else
        Local $tststr = FileGetShortName($TestFile) & " /info=" & $TempFileOut


        Local $Result = ShellExecuteWait($IrfanView, $tststr, "", "open", @SW_HIDE)
        LogFile(@ScriptLineNumber & " ProcessPictureSize. ShellExecuteWait results: " & $Result & " " & $tststr)
    EndIf

    ;read the file info into an array
    If _FileReadToArray($TempFileOut, $FileArray) = 0 Then
        LogFile(@ScriptLineNumber & " ProcessPictureSize error " & $TempFileOut & " could not be opened for reading. Error: " & @error)
        MsgBox(16, @ScriptLineNumber & " ProcessPictureSize error", $TempFileOut & " could not be opened for reading")
        Return -1
    EndIf

    _ArrayDelete($FileArray, 0)
    _ArrayAdd($FileArray, $TestFile)
    _ArrayAdd($FileArray, "---------------------------------------------")
    Return $FileArray
EndFunc   ;==>ProcessPictureSize
;-----------------------------------------------
Func ShowPreview()
    If Not FileExists($IrfanView) Then
        LogFile(@ScriptLineNumber & " ShowPreview: " & $IrfanView & " does not exist", True)
        Return
    EndIf

    GuiDisable($GUI_DISABLE)
    Local $RunPID

    LogFile(@ScriptLineNumber & " ShowPreview")
    GUICtrlSetData($EditStatusSelect, "Working....")

    Local $x = _GUICtrlTreeView_GetSelection($TreeViewActiveFiles)
    Local $File = _GUICtrlTreeView_GetText($TreeViewActiveFiles, $x)

    ConsoleWrite(@ScriptLineNumber & " " & $File & @CRLF)

    GUICtrlSetData($EditStatusSelect, $File)

    If Not FileExists($File) Then
        MsgBox(48, "Invalid file", "File does not exist" & @CRLF & $File)
        GuiDisable($GUI_ENABLE)
        Return
    EndIf

    Local $AspectRatio
    If GUICtrlRead($CheckboxAspect) = $GUI_CHECKED Then $AspectRatio = " /aspectratio "
    Local $ConvertCmd = $File & $AspectRatio & " /resize=(200,200) /resample /convert=" & $WallpaperTempJPG
    $RunPID = ShellExecuteWait($IrfanView, $ConvertCmd, "", "open", @SW_HIDE)
    LogFile(@ScriptLineNumber & " ShowPreview. ShellExecuteWait results: " & $RunPID & " " & $ConvertCmd)
    LogFile(@ScriptLineNumber & " ShowPreview ShellExecuteWait results: " & $RunPID & "   " & $ConvertCmd)
    LogFile(@ScriptLineNumber & " GUICtrlSetImage results: " & GUICtrlSetImage($PicPreview, $WallpaperTempJPG))

    ;This line gets the combined data for both source and temp files
    CheckPictureInfo(GUICtrlRead($EditStatusSelect), $WallpaperTempJPG, $WallpaperSelectInfo)

    ;Now we do the temp file for aspect ratio correction if needed
    If GUICtrlRead($CheckboxAspect) = $GUI_CHECKED Then
        Local $a = ProcessPictureSize($WallpaperTempJPG)
        Local $b = _ArraySearch($a, "Image dimensions =", 0, 0, 0, 1, 1, 0)
        ConsoleWrite(@ScriptLineNumber & " " & $b & " " & @error & @CRLF)
        Local $C = StringSplit($a[$b], "=P")
        Local $d = StringSplit($C[2], "x")

        Local $JA = ControlGetPos("", "", $PicPreview)
        GUICtrlSetPos($PicPreview, $JA[0], $JA[1], $d[1], $d[2])

    EndIf

    GuiDisable($GUI_ENABLE)
EndFunc   ;==>ShowPreview
;-----------------------------------------------
;This fuction converts the source JPG to $WallpaperTempJPG in the correct size
Func ConvertFile($File)
    LogFile(@ScriptLineNumber & " ConvertFile: " & $File)
    FileWriteLine($Used_filename, _SystemLocalTime() & " ~ " & $File)
    If Not FileExists($File) Then
        LogFile(@ScriptLineNumber & "  ConvertFile. File does not exist: " & $File)
        MsgBox(16, @ScriptLineNumber & " ConvertFile", $File & " does not exist")
        Return
    EndIf

    Local $AspectRatio = ""
    Local $Advanced = ""
    If GUICtrlRead($CheckboxDisplayName) = $GUI_CHECKED Then $Advanced = " /advancedbatch /ini=" & '"' & $AuxPath & '"'
    If GUICtrlRead($CheckboxAspect) = $GUI_CHECKED Then $AspectRatio = " /aspectratio "
    Local $Percent = GUICtrlRead($LabelFitPercentValue) / 100
    Local $Size = StringFormat(" /resize=(%d,%d) ", @DesktopWidth * $Percent, @DesktopHeight * $Percent)

    Local $ConvertCmd = FileGetShortName($File) & $Size & $AspectRatio & " " & $Advanced & " /resample /convert=" & $WallpaperTempJPG
    ;Local $STR = FileGetShortName($File) & $AspectRatio & " " & $Advanced & " /resample /convert=" & $WallpaperTempJPG

    If FileExists($IrfanView) Then
        Local $RunPID = ShellExecuteWait($IrfanView, $ConvertCmd, "", "open", @SW_HIDE)
        LogFile(@ScriptLineNumber & " ConvertFile. ShellExecuteWait results: " & $RunPID & " " & $ConvertCmd)
    Else
        LogFile(@ScriptLineNumber & " ConvertFile: " & $IrfanView & " does not exist", True)
        Return
    EndIf

    ;Local $Result = ShellExecuteWait($Infraview, FileGetShortName($File) & $Size & $AspectRatio & $Advanced & " /resample /convert=" & $WallpaperTempJPG, "", "open")
    LogFile(@ScriptLineNumber & " " & $ConvertCmd)
    GUICtrlSetData($EditStatusMain, $File)
    LogFile(@ScriptLineNumber & " ConvertFile. ShellExecuteWait results: " & $RunPID)
EndFunc   ;==>ConvertFile
;-----------------------------------------------
; $File is a JPG to display
; Returns true if file is diplayed and false if not
Func ProcessWallpaper($File)
    LogFile(@ScriptLineNumber & " ProcessWallpaper: " & $File)
    If $File = '' Then
        LogFile(@ScriptLineNumber & " ProcessWallpaper. Filename is blank")
        MsgBox(48, @ScriptLineNumber & " ProcessWallpaper", "Filename is blank")
        $Running = False
        $Hide = False
        Return -1
    EndIf

    If FileExists($File) = False Then
        LogFile(@ScriptLineNumber & "ProcessWallpaper. File does not exist" & $File)
        If MsgBox(49, @ScriptLineNumber & " ProcessWallpaper", "File does not exist: " & $File, 5) = 2 Then Exit
        Return
    EndIf
    GUICtrlSetData($EditStatusMain, "Working....")

    Local $WallSetting
    If GUICtrlRead($RadioCentered) = $GUI_CHECKED Then
        $WallSetting = 0
    ElseIf GUICtrlRead($RadioTiled) = $GUI_CHECKED Then
        $WallSetting = 1
    ElseIf GUICtrlRead($RadioStretched) = $GUI_CHECKED Then
        $WallSetting = 2
    ElseIf GUICtrlRead($RadioStretchedProportional) = $GUI_CHECKED Then
        $WallSetting = 1
    Else
        Exit
    EndIf

    HandleDisplayName($File)

    ;Convert the file to the correct size and check the resulting size
    ConvertFile($File)
    ;Now get the resulting status
    Local $F = CheckPictureInfo(GUICtrlRead($EditStatusMain), $WallpaperTempJPG, $WallpaperMainInfo)

    Local $WorkingFile
    ConsoleWrite(@ScriptLineNumber & " " & $F & @CRLF)
    If (StringInStr($F, "enlarged") <> 0 And GUICtrlRead($CheckboxEnlargeToFit) = $GUI_UNCHECKED) Or _
            (StringInStr($F, "reduced") <> 0 And GUICtrlRead($CheckboxShrinkToFit) = $GUI_UNCHECKED) Then
        $WorkingFile = $File
    Else
        $WorkingFile = $WallpaperTempJPG
    EndIf

    Local $Advanced = ""
    If GUICtrlRead($CheckboxDisplayName) = $GUI_CHECKED Then
        $Advanced = " /advancedbatch /ini=" & '"' & $AuxPath & '"'
        Local $STR = FileGetShortName($WorkingFile) & $Advanced & " /resample /convert=" & $WallpaperTempJPG
        Local $Result = ShellExecuteWait($IrfanView, $STR, "", "open", @SW_HIDE)
        ConsoleWrite(@ScriptLineNumber & " ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ " & $Result & @CRLF)
    EndIf

    ChangeDestopWallpaper($WallpaperTempJPG, $WallSetting)

    LogFile(@ScriptLineNumber & " ProcessWallpaper complete")
    Return True
EndFunc   ;==>ProcessWallpaper
;-----------------------------------------------
;$Option_filename
;$Picture_filename
Func TestForRequiredFiles()
    Return
    If Not FileExists($Option_filename) Then
        MsgBox(48, "Required file not found ", $Option_filename)
        $Hide = False
        $Running = False
    EndIf
    If Not FileExists($Picture_filename) Then
        MsgBox(48, "Required file not found ", $Picture_filename)
        $Hide = False
        $Running = False
    EndIf

    ;Abort if no valid pictures are found DBK\
    Local $TA
    Local $Found = False
    _FileReadToArray($Picture_filename, $TA)
    For $x In $TA
        If FileExists($x) Then
            $Found = True
            ExitLoop
        EndIf
    Next

    If Not $Found Then
        MsgBox(48, "No valid picture files found", "No valid picture files found in " & @CRLF & $Picture_filename)
        $Hide = False
        $Running = False
    EndIf

EndFunc   ;==>TestForRequiredFiles
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


    Local $msg = "Unable to locate IrfanView on this computer" & @CRLF & $ProgramName & " can not run correctly" & @CRLF & _
            $IrfanView & @CRLF & "\Program Files" & @CRLF & "\Program Files (x86)" & @CRLF & _
            "You can use the options window to fix this."

    GUICtrlSetData($InputIrfanviewFolder, $IrfanView)

    LogFile(@ScriptLineNumber & " " & $msg)
    MsgBox(16, "IrfanView not found", $msg)

EndFunc   ;==>TestForIrfanView

;-----------------------------------------------5
Func HandleSizetoFit()
    If GUICtrlRead($CheckboxShrinkToFit) = $GUI_CHECKED Or _
            GUICtrlRead($CheckboxEnlargeToFit) = $GUI_CHECKED Then
        GUICtrlSetState($SliderFitPercent, $GUI_ENABLE)
        GUICtrlSetState($LabelFitPercentValue, $GUI_ENABLE)
        GUICtrlSetState($CheckboxAspect, $GUI_ENABLE)
    Else
        GUICtrlSetState($SliderFitPercent, $GUI_DISABLE)
        GUICtrlSetState($LabelFitPercentValue, $GUI_DISABLE)
        GUICtrlSetState($CheckboxAspect, $GUI_DISABLE)
    EndIf
EndFunc   ;==>HandleSizetoFit
;-----------------------------------------------
Func HandleDisplayName($FileName = "")
    Local $setting
    If GUICtrlRead($CheckboxDisplayName) = $GUI_CHECKED Then
        $setting = $GUI_ENABLE
    Else
        $setting = $GUI_DISABLE
    EndIf
    GUICtrlSetState($CheckboxTransparent, $setting)
    GUICtrlSetState($RadioTop, $setting)
    GUICtrlSetState($RadioBottom, $setting)
    GUICtrlSetState($ButtonTextBackground, $setting)
    GUICtrlSetState($ButtonSetFont, $setting)
    GUICtrlSetState($LabelFontSample, $setting)

    GUICtrlSetFont($LabelFontSample, $FontPointSize, $FontWeight, $FontItalic + $FontUnderline + $FontStrikethru, $FontName)
    GUICtrlSetColor($LabelFontSample, $FontColorFG)
    GUICtrlSetBkColor($LabelFontSample, $FontColorBG)

    ;Now the ini file needs to be fixed to reflect the options
    Local $ArrayINI
    If _FileReadToArray($i_view32_ini, $ArrayINI) = 0 Then ; read the ini data in
        LogFile(@ScriptLineNumber & "i_view32.ini could not be opened for reading :" & $ArrayINI)
        MsgBox(16, "i_view32.ini file error", "i_view32.ini could not be opened for reading:" & @CRLF & $ArrayINI & @CRLF & @error)
        Return
    EndIf

    For $i = 0 To UBound($ArrayINI) - 1
        Local $T = ''
        If $FontItalic = 2 Then
            $T = "|255|"
        Else
            $T = "|0|"
        EndIf
        If $FontUnderline = 4 Then
            $T = $T & "1|"
        Else
            $T = $T & "0|"
        EndIf
        If $FontStrikethru = 8 Then
            $T = $T & "1|"
        Else
            $T = $T & "0|"
        EndIf

        If StringInStr($ArrayINI[$i], "FontParam=") <> 0 Then
            $ArrayINI[$i] = "FontParam=-" & $FontPointSize & "|0|0|0|" & $FontWeight & $T & "|3|2|1|20|"
            _Debug(@ScriptLineNumber & " " & _SystemLocalTime() & $ArrayINI[$i])
        ElseIf StringInStr($ArrayINI[$i], "AddText=") = 1 Then ; AddText=$D$F
            Local $R = "AddText=" & $FileName
            _Debug(@ScriptLineNumber & " " & $R)
            $ArrayINI[$i] = $R
            _Debug(@ScriptLineNumber & " " & _SystemLocalTime() & $ArrayINI[$i])
        ElseIf StringInStr($ArrayINI[$i], "Font=") = 1 Then
            $ArrayINI[$i] = "Font=" & $FontName
        ElseIf StringInStr($ArrayINI[$i], "FontColor=") = 1 Then
            $ArrayINI[$i] = "FontColor=" & $FontColorFG_REF
        ElseIf StringInStr($ArrayINI[$i], "TxtBgkr=") = 1 Then
            ;ConsoleWrite(@ScriptLineNumber & " Font stuff: " & _ColorGetBlue($FontColorBG) & " " & _ColorGetGreen($FontColorBG) & " " & _ColorGetRed($FontColorBG) & @CRLF)
            Local $aiInput[3] = [_ColorGetBlue($FontColorBG), _ColorGetGreen($FontColorBG), _ColorGetRed($FontColorBG)]
            $ArrayINI[$i] = "TxtBgkr=" & _ColorSetRGB($aiInput)
        ElseIf StringInStr($ArrayINI[$i], "TranspText=") = 1 Then
            If GUICtrlRead($CheckboxTransparent) = $GUI_CHECKED Then
                $ArrayINI[$i] = "TranspText=1"
            Else
                $ArrayINI[$i] = "TranspText=0"
            EndIf
        ElseIf StringInStr($ArrayINI[$i], "Corner=") = 1 Then
            If GUICtrlRead($RadioTop) = $GUI_CHECKED Then
                $ArrayINI[$i] = "Corner=0"
            ElseIf GUICtrlRead($RadioBottom) = $GUI_CHECKED Then
                $ArrayINI[$i] = "Corner=2"
            Else
                MsgBox(16, "Illegal value", "Position")
            EndIf
        EndIf
    Next


    _ArrayDelete($ArrayINI, 0)
    If _FileWriteFromArray($i_view32_ini, $ArrayINI) = 0 Then ; write the ini data out
        LogFile(@ScriptLineNumber & "File could not be opened for writing" & "  " & $ArrayINI & "  " & @error)
        MsgBox(16, "Ini file error", "File could not be opened for writing" & @CRLF & $ArrayINI & @CRLF & @error)
        Return
    EndIf

EndFunc   ;==>HandleDisplayName
;-----------------------------------------------
;"EveryMinutes"  "EveryFixed"  "Random"
Func HandleChangeOptions($type)
    GUICtrlSetState($SliderEveryMinutes, $GUI_DISABLE)
    GUICtrlSetState($LabelEveryMinutes, $GUI_DISABLE)
    GUICtrlSetState($ComboEveryFixed, $GUI_DISABLE)
    GUICtrlSetState($LabelMin, $GUI_DISABLE)
    GUICtrlSetState($LabelMinV, $GUI_DISABLE)
    GUICtrlSetState($LabelMax, $GUI_DISABLE)
    GUICtrlSetState($LabelMaxV, $GUI_DISABLE)
    GUICtrlSetState($SliderMin, $GUI_DISABLE)
    GUICtrlSetState($SliderMax, $GUI_DISABLE)
    GUICtrlSetState($LabelValue, $GUI_DISABLE)
    If StringCompare($type, "EveryMinutes") = 0 Then
        GUICtrlSetState($SliderEveryMinutes, $GUI_ENABLE)
        GUICtrlSetState($LabelEveryMinutes, $GUI_ENABLE)
    ElseIf StringCompare($type, "EveryFixed") = 0 Then
        GUICtrlSetState($ComboEveryFixed, $GUI_ENABLE)
    ElseIf StringCompare($type, "Random") = 0 Then
        GUICtrlSetState($LabelMin, $GUI_ENABLE)
        GUICtrlSetState($LabelMinV, $GUI_ENABLE)
        GUICtrlSetState($LabelMax, $GUI_ENABLE)
        GUICtrlSetState($LabelMaxV, $GUI_ENABLE)
        GUICtrlSetState($SliderMin, $GUI_ENABLE)
        GUICtrlSetState($SliderMax, $GUI_ENABLE)
        GUICtrlSetState($LabelValue, $GUI_ENABLE)
    Else
        LogFile("How did we get here? HandleChangeOptions:" & $type, True)
    EndIf
    SetUpDelayTime()
EndFunc   ;==>HandleChangeOptions
;-----------------------------------------------
Func ToggleAllFoldersChecks()
    GuiDisable($GUI_DISABLE)
    Static $ToggleFolderState = True
    LogFile(@ScriptLineNumber & " ToggleAllFoldersChecks " & $ToggleFolderState)
    ;clears all check marks for folderlist
    For $x = 0 To UBound($FolderHandleArray) - 1
        _GUICtrlTreeView_SetChecked($TreeViewSelectFolders, $FolderHandleArray[$x], $ToggleFolderState)
    Next
    $ToggleFolderState = Not $ToggleFolderState
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>ToggleAllFoldersChecks
;-----------------------------------------------
Func MarkAllFolderFound()
    LogFile(@ScriptLineNumber & " MarkAllFoldersFound")
    GuiDisable($GUI_DISABLE)
    For $x = 0 To UBound($FolderHandleArray) - 1
        Local $TS = _GUICtrlTreeView_GetText($TreeViewSelectFolders, $FolderHandleArray[$x])
        If StringInStr($TS, GUICtrlRead($InputBoxSearchFolders)) > 0 Then
            ConsoleWrite(@ScriptLineNumber & " " & GUICtrlRead($InputBoxSearchFolders) & " " & $x & "  " & $FolderHandleArray[$x] & @CRLF)
            _GUICtrlTreeView_SelectItem($TreeViewSelectFolders, $FolderHandleArray[$x])
            _GUICtrlTreeView_SetChecked($TreeViewSelectFolders, $FolderHandleArray[$x])
        EndIf
    Next
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>MarkAllFolderFound

;-----------------------------------------------
Func SearchFolders($Reset)
    Static $SearchFoldersPosition
    LogFile(@ScriptLineNumber & " SearchFolders " & $Reset)
    GuiDisable($GUI_DISABLE)
    If $Reset Then
        $SearchFoldersPosition = 0
    Else
        $SearchFoldersPosition = $SearchFoldersPosition + 1
    EndIf

    For $x = $SearchFoldersPosition To UBound($FolderHandleArray) - 1
        Local $TS = _GUICtrlTreeView_GetText($TreeViewSelectFolders, $FolderHandleArray[$x])
        If StringInStr($TS, GUICtrlRead($InputBoxSearchFolders)) > 0 Then
            _GUICtrlTreeView_SelectItem($TreeViewSelectFolders, $FolderHandleArray[$x])
            _GUICtrlTreeView_SetChecked($TreeViewSelectFolders, $FolderHandleArray[$x])
            $SearchFoldersPosition = $x
            ExitLoop
        EndIf
    Next
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>SearchFolders
;-----------------------------------------------
Func ToggleAllFilesChecks()
    GuiDisable($GUI_DISABLE)
    Static $ToggleFilesState = True
    LogFile(@ScriptLineNumber & " ToggleAllFilesChecks " & $ToggleFilesState & "  " & UBound($FileHandleArray))
    ;clears all check marks for file list
    For $x = 0 To UBound($FileHandleArray) - 1
        _GUICtrlTreeView_SetChecked($TreeViewActiveFiles, $FileHandleArray[$x], $ToggleFilesState)
    Next
    $ToggleFilesState = Not $ToggleFilesState
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>ToggleAllFilesChecks
;-----------------------------------------------
Func MarkAllFilesFound()
    LogFile(@ScriptLineNumber & " MarkAllFilesFound")
    GuiDisable($GUI_DISABLE)
    For $x = 0 To UBound($FileHandleArray) - 1
        Local $TS = _GUICtrlTreeView_GetText($TreeViewActiveFiles, $FileHandleArray[$x])
        If StringInStr($TS, GUICtrlRead($InputBoxSearchFiles)) > 0 Then
            ConsoleWrite(@ScriptLineNumber & " " & GUICtrlRead($InputBoxSearchFiles) & " " & $x & "  " & $FileHandleArray[$x] & @CRLF)
            _GUICtrlTreeView_SelectItem($TreeViewActiveFiles, $FileHandleArray[$x])
            _GUICtrlTreeView_SetChecked($TreeViewActiveFiles, $FileHandleArray[$x])
        EndIf
    Next
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>MarkAllFilesFound

;-----------------------------------------------
Func SearchFiles($Reset)
    Static $SearchFilesPosition

    LogFile(@ScriptLineNumber & " SearchFiles " & $Reset)
    GuiDisable($GUI_DISABLE)
    If $Reset Then
        $SearchFilesPosition = 0
    Else
        $SearchFilesPosition = $SearchFilesPosition + 1
    EndIf

    For $x = $SearchFilesPosition To UBound($FileHandleArray) - 1
        Local $TS = _GUICtrlTreeView_GetText($TreeViewActiveFiles, $FileHandleArray[$x])
        If StringInStr($TS, GUICtrlRead($InputBoxSearchFiles)) > 0 Then
            _GUICtrlTreeView_SelectItem($TreeViewActiveFiles, $FileHandleArray[$x])
            _GUICtrlTreeView_SetChecked($TreeViewActiveFiles, $FileHandleArray[$x])
            $SearchFilesPosition = $x
            ExitLoop
        EndIf
    Next
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>SearchFiles
;-----------------------------------------------
;This function allows the user to edit or view any file, useful for changing the config file
Func EditText($File = "")
    LogFile(@ScriptLineNumber & " Edit " & $File)
    If $File = "" Then $File = FileOpenDialog("View or Edit a file", @ScriptDir, "Wallpaper (Wallpaper*.*)|All (*.*)", 1)
    If @error <> 0 Then Return

    If Not FileExists($File) Then
        LogFile(@ScriptLineNumber & " EditText error. " & $File & " does not exist")
        MsgBox(16, "EditText error", "File does not exist" & @CRLF & $File, 2)
        Return
    EndIf

    ShellExecute($editor, $File, @SW_HIDE)
EndFunc   ;==>EditText
;-----------------------------------------------
Func SaveList()
    GuiDisable($GUI_DISABLE)

    LogFile(@ScriptLineNumber & " SaveList")
    $Picture_filename = FileSaveDialog("Save list file", $AuxPath, _
            $ProgramName & "Lists (*.lst)|Wallpaper (Wallpaper.*)|All files (*.*)", 18, "WallPaper-" & @ComputerName & ".lst")

    Local $File = FileOpen($Picture_filename, 2)
    ; Check if file opened for writing OK
    If $File = -1 Then
        LogFile("SaveList: Unable to open file for writing: " & $Picture_filename, True)
        LogFile(@ScriptLineNumber & " SaveList: Unable to open file for writing: " & $Picture_filename)
        GuiDisable($GUI_ENABLE)
        Return
    EndIf
    LogFile(@ScriptLineNumber & " SaveList  " & $Picture_filename)
    If Not $Hide Then SplashImageOn("Wallpaper is starting. Please wait.", $AuxPath & "Wallpaper.jpg", -1, -1, -1, -1, 18)
    FileWriteLine($File, "Valid for Wallpaper list")

    For $x = 0 To UBound($FileHandleArray) - 1
        Local $TFN = _GUICtrlTreeView_GetText($TreeViewActiveFiles, $FileHandleArray[$x])
        If StringLen($TFN) > 3 Then FileWriteLine($File, $TFN)
    Next

    FileClose($File)
    SplashOff()
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>SaveList
;-----------------------------------------------
Func LoadList($type = "start")
    GuiDisable($GUI_DISABLE)
    LogFile(@ScriptLineNumber & " LoadList " & $type)

    If StringCompare($type, "menu") = 0 Then
        $Picture_filename = FileOpenDialog("Load options file", $AuxPath, _
                $ProgramName & "Lists (*.lst)|Wallpaper (Wallpaper.*)|All files (*.*)", 18, $AuxPath & "WallPaper-" & @ComputerName & ".opt")
    EndIf

    If Not $Hide Then SplashImageOn("Wallpaper is starting. Please wait.", $AuxPath & "Wallpaper.jpg", -1, -1, -1, -1, 18)
    Local $TempArray
    Local $Result = _FileReadToArray($Picture_filename, $TempArray)

    ; Check if file opened for reading OK
    If $Result <> 1 Then
        LogFile(@ScriptLineNumber & " LoadList: _FileReadToArray failed: " & $Picture_filename)
        MsgBox(16, $Picture_filename, "File does not exist" & @CRLF & $Picture_filename)
        SplashOff()
        GuiDisable($GUI_ENABLE)
        Return
    EndIf
    LogFile(@ScriptLineNumber & " Done reading file " & $Picture_filename)

    ; Read in the first line to verify the file is of the correct type
    If StringCompare($TempArray[1], "Valid for WallPaper list") <> 0 Then
        LogFile("Not an valid list file for WallPaper", True)
        LogFile(@ScriptLineNumber & " Not an vaild list file for WallPaper")
        SplashOff()
        GuiDisable($GUI_ENABLE)
        Return
    EndIf
    _GUICtrlTreeView_DeleteAll($TreeViewActiveFiles)

    ;_ArrayDisplay($TempArray, @ScriptLineNumber)


    _ArraySort($TempArray) ; sort the entries
    LogFile(@ScriptLineNumber & " Done with arraysort")

    For $i = 0 To UBound($TempArray) - 2 ; remove non-unique entries
        If StringCompare($TempArray[$i], $TempArray[$i + 1]) = 0 Then $TempArray[$i] = ""
    Next
    LogFile(@ScriptLineNumber & " Done with unique")

    ReDim $FileHandleArray[1]
    For $i In $TempArray ; if not a string of character 2 is not : then ignore the line
        If StringInStr($i, ":") = 2 And VerifyFileTypes($i) >= 0 Then
            _ArrayAdd($FileHandleArray, _GUICtrlTreeView_Add($TreeViewActiveFiles, 0, $i))
        EndIf
    Next
    _ArrayDelete($FileHandleArray, 0)
    SplashOff()
    LogFile(@ScriptLineNumber & " LoadList done. Files in list: " & _GUICtrlTreeView_GetCount($TreeViewActiveFiles) & "  " & $Picture_filename)
    GUICtrlSetData($LabelActiveCount, _GUICtrlTreeView_GetCount($TreeViewActiveFiles))
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>LoadList
;-----------------------------------------------
Func VerifyFileTypes($File)
    ;This is an array of files types that can be converted to wallpaper
    Static $FileTypes[31] = ['.JPG', '.JPEG', '.BMP', '.GIF', '.TIF', '.TIFF', '.TTF', '.TXT', _
            '.INI', '.HTML', '.HTML', '.RAW', '.PNG', '.PNG934', '.PCX', _
            '.PHOTOCD', '.ICO', '.ECW', '.EMF', '.FSH', '.JLS', '.JP2', '.JNG', '.JPM', '.PCX', _
            '.PBM', '.PGM', '.PNG', '.PPM', '.RAW', '.TGA']
    Local $szDrive, $szDir, $szFName, $szExt
    _PathSplit($File, $szDrive, $szDir, $szFName, $szExt)
    ;  ConsoleWrite(@ScriptLineNumber & " " & $szExt & " " & _ArraySearch($FileTypes, StringUpper($szExt)) & @CRLF)
    Return _ArraySearch($FileTypes, StringUpper($szExt))
EndFunc   ;==>VerifyFileTypes
;-----------------------------------------------
Func SetDefaults()
    GuiDisable($GUI_DISABLE)
    LogFile(@ScriptLineNumber & " SetDefaults")

    WinMove("WallPaper", "", 10, 10)
    WinMove("Select files", "", 10, 10)
    WinMove("Select folders", "", 10, 10)
    WinMove("Select options", "", 10, 10)
    $WorkingFolder = _AddSlash2PathString(EnvGet("USERPROFILE"))
    GUICtrlSetData($InputWorkingFolder, $WorkingFolder)
    $FontName = "Courier New"
    $FontColorFG = 0x00FFFF
    $FontColorFG_REF = 16776960
    $FontColorBG = 0
    $FontPointSize = 12
    $FontWeight = 700
    $FontItalic = 0
    $FontUnderline = 0
    $FontStrikethru = 0

    ; select files page
    ;GUICtrlSetData($LabelNewCount, 0)
    GUICtrlSetData($LabelActiveCount, 0)
    GUICtrlSetState($CheckBoxSelectFolderRecursive, $GUI_CHECKED)
    GUICtrlSetState($CheckSkipSmallFiles, $GUI_CHECKED)
    GUICtrlSetState($CheckAppendFiles, $GUI_UNCHECKED)

    ; Options page
    GUICtrlSetState($CheckboxDisplayName, $GUI_CHECKED)
    GUICtrlSetState($CheckboxTransparent, $GUI_CHECKED)
    GUICtrlSetState($CheckboxCheckForUsed, $GUI_CHECKED)
    GUICtrlSetData($SliderMinPictFileSize, 10)
    GUICtrlSetData($LabelMinPictFileSizeValue, GUICtrlRead($SliderMinPictFileSize))
    GUICtrlSetData($SliderMaxLogFileSize, 50)
    GUICtrlSetData($LabelMaxLogFileSizeValue, GUICtrlRead($SliderMaxLogFileSize))

    ;	$FontColor = 16777215
    ;	$TxtBgkr = 0
    GUICtrlSetState($RadioTop, $GUI_CHECKED)
    GUICtrlSetState($RadioBottom, $GUI_UNCHECKED)
    GUICtrlSetState($RadioEveryMinutes, $GUI_CHECKED)
    GUICtrlSetData($SliderEveryMinutes, 10)
    GUICtrlSetData($LabelEveryMinutes, GUICtrlRead($SliderEveryMinutes))
    GUICtrlSetState($RadioEveryFixed, $GUI_UNCHECKED)
    GUICtrlSetData($ComboEveryFixed, "10 minutes")
    GUICtrlSetState($RadioRandom, $GUI_UNCHECKED)

    GUICtrlSetData($SliderMin, 8)
    GUICtrlSetData($SliderMax, 100)
    GUICtrlSetData($LabelMinV, GUICtrlRead($SliderMin))
    GUICtrlSetData($LabelMaxV, GUICtrlRead($SliderMax))
    GUICtrlSetData($LabelValue, Random(GUICtrlRead($LabelMinV), GUICtrlRead($LabelMaxV), 1))

    GUICtrlSetState($CheckboxAspect, $GUI_CHECKED)
    GUICtrlSetState($CheckboxShrinkToFit, $GUI_CHECKED)
    GUICtrlSetState($CheckboxEnlargeToFit, $GUI_CHECKED)
    GUICtrlSetData($SliderFitPercent, 75)
    GUICtrlSetData($LabelFitPercentValue, GUICtrlRead($SliderFitPercent))

    GUICtrlSetState($RadioCentered, $GUI_CHECKED)
    GUICtrlSetState($RadioTiled, $GUI_UNCHECKED)
    GUICtrlSetState($RadioStretched, $GUI_UNCHECKED)
    GUICtrlSetState($RadioStretchedProportional, $GUI_UNCHECKED)
    ;GUICtrlSetData($LabelNewCount, 0)
    GUICtrlSetData($LabelActiveCount, 0)
    HandleDisplayName()
    HandleChangeOptions("EveryMinutes")
    HandleSizetoFit()
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>SetDefaults
;-----------------------------------------------
Func SaveOptions()
    GuiDisable($GUI_DISABLE)
    LogFile(@ScriptLineNumber & " SaveOptions")
    $Option_filename = FileSaveDialog("Save options file", $AuxPath, _
            $ProgramName & "Options (*.opt)|Wallpaper (Wallpaper.*)|All files (*.*)", 18, "WallPaper-" & @ComputerName & ".opt")

    Local $File = FileOpen($Option_filename, 2)
    ; Check if file opened for writing OK
    If $File = -1 Then
        LogFile(@ScriptLineNumber & " SaveOptions: Unable to open file for writing: " & $Option_filename)
        LogFile("SaveOptions: Unable to open file for writing: " & $Option_filename, True)
        GuiDisable($GUI_ENABLE)
        Return
    EndIf

    FileWriteLine($File, "Valid for " & $ProgramName & " options")
    FileWriteLine($File, "Options file for " & $Option_filename & "  " & _DateTimeFormat(_NowCalc(), 0))
    FileWriteLine($File, "Help 1 is enabled, 4 is disabled for checkboxes")
    FileWriteLine($File, "Irfanview:" & $IrfanView)
    FileWriteLine($File, "WorkingFolder:" & $WorkingFolder)
    FileWriteLine($File, "CheckBoxSelectFolderRecursive:" & GUICtrlRead($CheckBoxSelectFolderRecursive))
    FileWriteLine($File, "CheckSkipSmallFiles:" & GUICtrlRead($CheckSkipSmallFiles))
    FileWriteLine($File, "CheckAppendFiles:" & GUICtrlRead($CheckAppendFiles))
    FileWriteLine($File, "SliderMinPictFileSize:" & GUICtrlRead($SliderMinPictFileSize))
    FileWriteLine($File, "LabelMinPictFileSizeValue:" & GUICtrlRead($LabelMinPictFileSizeValue))
    FileWriteLine($File, "SliderMaxLogFileSize:" & GUICtrlRead($SliderMaxLogFileSize))
    FileWriteLine($File, "LabelMaxLogFileSizeValue:" & GUICtrlRead($LabelMaxLogFileSizeValue))
    ; Options page
    FileWriteLine($File, "CheckboxDisplayName:" & GUICtrlRead($CheckboxDisplayName))
    FileWriteLine($File, "CheckboxTransparent:" & GUICtrlRead($CheckboxTransparent))
    FileWriteLine($File, "CheckboxCheckForUsed:" & GUICtrlRead($CheckboxCheckForUsed))
    FileWriteLine($File, "RadioTop:" & GUICtrlRead($RadioTop))
    FileWriteLine($File, "RadioBottom:" & GUICtrlRead($RadioBottom))

    FileWriteLine($File, "RadioEveryMinutes:" & GUICtrlRead($RadioEveryMinutes))
    FileWriteLine($File, "SliderEveryMinutes:" & GUICtrlRead($SliderEveryMinutes))
    FileWriteLine($File, "LabelEveryMinutes:" & GUICtrlRead($LabelEveryMinutes))
    FileWriteLine($File, "RadioEveryFixed:" & GUICtrlRead($RadioEveryFixed))
    FileWriteLine($File, "ComboEveryFixed:" & GUICtrlRead($ComboEveryFixed))
    FileWriteLine($File, "RadioRandom:" & GUICtrlRead($RadioRandom))
    FileWriteLine($File, "SliderMin:" & GUICtrlRead($SliderMin))
    FileWriteLine($File, "SliderMax:" & GUICtrlRead($SliderMax))
    FileWriteLine($File, "LabelMinV:" & GUICtrlRead($SliderMin))
    FileWriteLine($File, "LabelMaxV:" & GUICtrlRead($SliderMax))
    FileWriteLine($File, "LabelValue:" & GUICtrlRead($LabelValue))
    FileWriteLine($File, "CheckboxAspect:" & GUICtrlRead($CheckboxAspect))
    FileWriteLine($File, "CheckboxShrinkToFit:" & GUICtrlRead($CheckboxShrinkToFit))
    FileWriteLine($File, "CheckboxEnlargeToFit:" & GUICtrlRead($CheckboxEnlargeToFit))
    FileWriteLine($File, "SliderFitPercent:" & GUICtrlRead($SliderFitPercent))
    FileWriteLine($File, "LabelFitPercent:" & GUICtrlRead($LabelFitPercentValue))
    FileWriteLine($File, "RadioCentered:" & GUICtrlRead($RadioCentered))
    FileWriteLine($File, "RadioTiled:" & GUICtrlRead($RadioTiled))
    FileWriteLine($File, "RadioStretched:" & GUICtrlRead($RadioStretched))
    FileWriteLine($File, "RadioStretchedProportional:" & GUICtrlRead($RadioStretchedProportional))

    FileWriteLine($File, "FontName:" & $FontName)
    FileWriteLine($File, "FontColorFG:" & $FontColorFG)
    FileWriteLine($File, "FontColorFG_REF:" & $FontColorFG_REF)
    FileWriteLine($File, "FontColorBG:" & $FontColorBG)
    FileWriteLine($File, "FontPointSize:" & $FontPointSize)
    FileWriteLine($File, "FontWeight:" & $FontWeight)
    FileWriteLine($File, "FontItalic:" & $FontItalic)
    FileWriteLine($File, "FontUnderline:" & $FontUnderline)
    FileWriteLine($File, "FontStrikethru:" & $FontStrikethru)

    Local $F = WinGetPos("WallPaper", "")
    FileWriteLine($File, "WallPaperWinpos:" & $F[0] & " " & $F[1] & " " & $F[2] & " " & $F[3])

    $F = WinGetPos("Select files", "")
    FileWriteLine($File, "SelectFilesWinpos:" & $F[0] & " " & $F[1] & " " & $F[2] & " " & $F[3])

    $F = WinGetPos("Select folders", "")
    FileWriteLine($File, "SelectFoldersWinpos:" & $F[0] & " " & $F[1] & " " & $F[2] & " " & $F[3])

    $F = WinGetPos("Select options", "")
    FileWriteLine($File, "SelectOptionsWinpos:" & $F[0] & " " & $F[1] & " " & $F[2] & " " & $F[3])

    FileClose($File)
    LogFile(@ScriptLineNumber & " Done SaveOptions")
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>SaveOptions
;-----------------------------------------------
;This loads the options file
Func LoadOptions($type = "start")
    GuiDisable($GUI_DISABLE)
    LogFile(@ScriptLineNumber & " LoadOptions  " & $type)

    If StringCompare($type, "menu") = 0 Then
        $Option_filename = FileOpenDialog("Load options file", $AuxPath, _
                $ProgramName & "Options (*.opt)|Wallpaper (Wallpaper.*)|All files (*.*)", 18, "WallPaper-" & @ComputerName & ".opt")
    EndIf

    Local $File = FileOpen($Option_filename, 0)
    ; Check if file opened for reading OK
    If $File = -1 Then
        LogFile(@ScriptLineNumber & " LoadOptions: Unable to open file for reading: " & $Option_filename)
        MsgBox(16, $Option_filename, "LoadOptions: Unable to open file for reading: " & @CRLF & $Option_filename)
        GuiDisable($GUI_ENABLE)
        Return
    EndIf

    LogFile(@ScriptLineNumber & " LoadOptions " & $Option_filename)
    ; Read in the first line to verify the file is of the correct type
    If StringCompare(FileReadLine($File, 1), "Valid for WallPaper options") <> 0 Then
        LogFile(@ScriptLineNumber & " Not an options file for WallPaper")
        LogFile("Not an options file for WallPaper ", True)
        FileClose($File)
        GuiDisable($GUI_ENABLE)
        Return
    EndIf
    ; Read in lines of text until the EOF is reached
    While 1
        Local $LineIn = FileReadLine($File)
        If @error = -1 Then ExitLoop
        If StringInStr($LineIn, ";") = 1 Then ContinueLoop

        If StringInStr($LineIn, "WallPaperWinpos:") Then
            Local $F = StringMid($LineIn, StringInStr($LineIn, ":") + 1)
            $F = StringSplit($F, " ", 2)
            WinMove("WallPaper", "", $F[0], $F[1], $F[2], $F[3])
        EndIf

        If StringInStr($LineIn, "SelectFilesWinpos:") Then
            $F = StringMid($LineIn, StringInStr($LineIn, ":") + 1)
            $F = StringSplit($F, " ", 2)
            WinMove("Select files", "", $F[0], $F[1], $F[2], $F[3])
        EndIf

        If StringInStr($LineIn, "SelectFoldersWinpos:") Then
            $F = StringMid($LineIn, StringInStr($LineIn, ":") + 1)
            $F = StringSplit($F, " ", 2)
            WinMove("Select folders", "", $F[0], $F[1], $F[2], $F[3])
        EndIf

        If StringInStr($LineIn, "SelectOptionsWinpos:") Then
            $F = StringMid($LineIn, StringInStr($LineIn, ":") + 1)
            $F = StringSplit($F, " ", 2)
            WinMove("Select options", "", $F[0], $F[1], $F[2], $F[3])
        EndIf

        If StringInStr($LineIn, "Irfanview:") Then $IrfanView = StringMid($LineIn, StringInStr($LineIn, ":") + 1)
        If StringInStr($LineIn, "WorkingFolder:") Then $WorkingFolder = StringMid($LineIn, StringInStr($LineIn, ":") + 1)
        ; select folder page
        If StringInStr($LineIn, "CheckBoxSelectFolderRecursive:") Then GUICtrlSetState($CheckBoxSelectFolderRecursive, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckSkipSmallFiles:") Then GUICtrlSetState($CheckSkipSmallFiles, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckAppendFiles:") Then GUICtrlSetState($CheckAppendFiles, StringMid($LineIn, StringInStr($LineIn, ":") + 1))

        If StringInStr($LineIn, "SliderMinPictFileSize:") Then GUICtrlSetData($SliderMinPictFileSize, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "LabelMinPictFileSizeValue:") Then GUICtrlSetData($LabelMinPictFileSizeValue, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "SliderMaxLogFileSize:") Then GUICtrlSetData($SliderMaxLogFileSize, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "LabelMaxLogFileSizeValue:") Then GUICtrlSetData($LabelMaxLogFileSizeValue, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckboxDisplayName:") Then GUICtrlSetState($CheckboxDisplayName, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckboxTransparent:") Then GUICtrlSetState($CheckboxTransparent, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "RadioTop:") Then GUICtrlSetState($RadioTop, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "RadioBottom:") Then GUICtrlSetState($RadioBottom, StringMid($LineIn, StringInStr($LineIn, ":") + 1))

        If StringInStr($LineIn, "RadioEveryMinutes:") Then GUICtrlSetState($RadioEveryMinutes, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "SliderEveryMinutes:") Then GUICtrlSetData($SliderEveryMinutes, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "LabelEveryMinutes:") Then GUICtrlSetData($LabelEveryMinutes, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "RadioEveryFixed:") Then GUICtrlSetState($RadioEveryFixed, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "ComboEveryFixed:") Then GUICtrlSetData($ComboEveryFixed, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "RadioRandom:") Then GUICtrlSetState($RadioRandom, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "SliderMin:") Then GUICtrlSetData($SliderMin, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "SliderMax:") Then GUICtrlSetData($SliderMax, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "LabelMinV:") Then GUICtrlSetData($LabelMinV, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "LabelMaxV:") Then GUICtrlSetData($LabelMaxV, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "LabelValue:") Then GUICtrlSetData($LabelValue, StringMid($LineIn, StringInStr($LineIn, ":") + 1))

        If StringInStr($LineIn, "CheckboxAspect:") Then GUICtrlSetState($CheckboxAspect, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckboxShrinkToFit:") Then GUICtrlSetState($CheckboxShrinkToFit, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckboxEnlargeToFit:") Then GUICtrlSetState($CheckboxEnlargeToFit, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "SliderFitPercent:") Then GUICtrlSetData($SliderFitPercent, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "LabelFitPercent:") Then GUICtrlSetData($LabelFitPercentValue, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "RadioCentered:") Then GUICtrlSetState($RadioCentered, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "RadioTiled:") Then GUICtrlSetState($RadioTiled, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "RadioStretched:") Then GUICtrlSetState($RadioStretched, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "RadioStretchedProportional:") Then GUICtrlSetState($RadioStretchedProportional, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckboxTransparent:") Then GUICtrlSetState($CheckboxTransparent, StringMid($LineIn, StringInStr($LineIn, ":") + 1))
        If StringInStr($LineIn, "CheckboxCheckForUsed:") Then GUICtrlSetState($CheckboxCheckForUsed, StringMid($LineIn, StringInStr($LineIn, ":") + 1))

        If StringInStr($LineIn, "FontName:") Then $FontName = StringMid($LineIn, StringInStr($LineIn, ":") + 1)
        If StringInStr($LineIn, "FontColorFG:") Then $FontColorFG = StringMid($LineIn, StringInStr($LineIn, ":") + 1)
        If StringInStr($LineIn, "FontColorFG_REF:") Then $FontColorFG_REF = StringMid($LineIn, StringInStr($LineIn, ":") + 1)
        If StringInStr($LineIn, "FontColorBG:") Then $FontColorBG = StringMid($LineIn, StringInStr($LineIn, ":") + 1)
        If StringInStr($LineIn, "FontPointSize:") Then $FontPointSize = StringMid($LineIn, StringInStr($LineIn, ":") + 1)
        If StringInStr($LineIn, "FontWeight:") Then $FontWeight = StringMid($LineIn, StringInStr($LineIn, ":") + 1)
        If StringInStr($LineIn, "FontItalic:") Then $FontItalic = StringMid($LineIn, StringInStr($LineIn, ":") + 1)
        If StringInStr($LineIn, "FontUnderline:") Then $FontUnderline = StringMid($LineIn, StringInStr($LineIn, ":") + 1)
        If StringInStr($LineIn, "FontStrikethru:") Then $FontStrikethru = StringMid($LineIn, StringInStr($LineIn, ":") + 1)

    WEnd
    FileClose($File)

    If Not FileExists($WorkingFolder) Then
        If MsgBox(1 + 16, "Working Folder error", $WorkingFolder & @CRLF & "does not exist" & @CRLF & "OK to continue, Cancel to quit") = 2 Then Exit
    EndIf

    HandleDisplayName()
    HandleChangeOptions("EveryMinutes")
    HandleSizetoFit()
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>LoadOptions
;-----------------------------------------------
Func About(Const $FormID)
    GuiDisable($GUI_DISABLE)
    LogFile(@ScriptLineNumber & " About " & $FormID)
    Local $d = WinGetPos($FormID)
    Local $WinPos
    If IsArray($d) = True Then
        ConsoleWrite(@ScriptLineNumber & $FormID & @CRLF)
        $WinPos = StringFormat("%s" & @CRLF & "WinPOS: %d  %d " & @CRLF & "WinSize: %d %d " & @CRLF & "Desktop: %d %d ", _
                $FormID, $d[0], $d[1], $d[2], $d[3], @DesktopWidth, @DesktopHeight)
    Else
        $WinPos = ">>>About ERROR, Check the window name<<<"
    EndIf
    LogFile(@CRLF & $SystemS & @CRLF & $WinPos & @CRLF & "Written by Doug Kaynor!", True)
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>About
;-----------------------------------------------
Func LogFile($string, $ShowMSGBox = False)
    $string = StringReplace($string, "-1", "", 1)
    _FileWriteLog($Log_filename, $string)
    If $ShowMSGBox Then
        MsgBox(16, 'Debug message', $string)
        _Debug($string, $ShowMSGBox)
    EndIf
EndFunc   ;==>LogFile
;-----------------------------------------------
Func Help(Const $FormID)
    GuiDisable($GUI_DISABLE)
    LogFile(@ScriptLineNumber & " Help " & $FormID)
    Local $helpstr = 'Startup options: ' & @CRLF & _
            "help or ?   Display this help file" & @CRLF & _
            "Run         Start changing wallpaper on startup" & @CRLF & _
            "Hide        Hide the GUI on startup" & @CRLF & _
            "Debug       Secret things" & @CRLF & @CRLF & _
            "Hot Keys: " & @CRLF & _
            "F1  = Help" & @CRLF & _
            "F10 = Change wallpaper" & @CRLF & _
            "F11 = Unlock GUI"
    LogFile(@ScriptName & @CRLF & $FileVersion & @CRLF & @CRLF & $helpstr, True)
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>Help
;-----------------------------------------------
Func ShowMainInfo()
    LogFile(@ScriptLineNumber & " EditStatusMain")
    ClipPut(GUICtrlRead($EditStatusMain))
    EditText($WallpaperMainInfo)
EndFunc   ;==>ShowMainInfo
;-----------------------------------------------
Func EditCurrentFile()
    GuiDisable($GUI_DISABLE)
    Local $File = GUICtrlRead($EditStatusMain)
    ConsoleWrite(@ScriptLineNumber & " Editcurrentfile " & $File & @CRLF)
    LogFile(@ScriptLineNumber & " Editcurrentfile  " & $File)
    If FileExists($IrfanView) Then
        LogFile(@ScriptLineNumber & " ShellExecuteWait results: " & ShellExecuteWait($IrfanView, $File, "", "open", @SW_SHOW))
    Else
        LogFile(@ScriptLineNumber & " EditCurrentFile: " & $IrfanView & " does not exist", True)
    EndIf
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>EditCurrentFile
;-----------------------------------------------
;This will remove all matching files from the list of pictures to display
Func FilterFileList()
    LogFile(@ScriptLineNumber & " FilterFileList")
    If Not FileExists($Filter_filename) Then
        LogFile(@ScriptLineNumber & " FilterFileList: " & $Filter_filename & " does not exist", True)
        Return
    Else
        ; Check if filter file opened for reading OK
        Local $FiltersA
        If Not _FileReadToArray($Filter_filename, $FiltersA) Then
            LogFile(@ScriptLineNumber & " FilterFileList: Unable to open filter file for reading: " & $Filter_filename & ' ' & @error, True)
            GuiDisable($GUI_ENABLE)
            Return
        EndIf

        ; Read in the first line to verify the file is of the correct type
        Local $TstString = "Valid for WallPaper filters"
        ConsoleWrite(@ScriptLineNumber & " " & $TstString & @CRLF)
        ConsoleWrite(@ScriptLineNumber & " " & $FiltersA[1] & @CRLF)
        If Not StringCompare($FiltersA[1], $TstString) = 0 Then
            LogFile(@ScriptLineNumber & " Not a valid filter file for WallPaper" & @CRLF & $TstString & @CRLF & $FiltersA[1], True)
            GuiDisable($GUI_ENABLE)
            Return
        EndIf

        ; Check if picture file opened for reading OK
        Local $PicturesA
        If Not _FileReadToArray($Picture_filename, $PicturesA) Then
            LogFile(@ScriptLineNumber & " FilterFileList: Unable to open picture file for reading: " & $Picture_filename + ' ' & @error, True)
            GuiDisable($GUI_ENABLE)
            Return
        EndIf

        ConsoleWrite(@ScriptLineNumber & " FiltersA:" & $FiltersA[0] & @CRLF)
        ConsoleWrite(@ScriptLineNumber & " $PicturesA:" & $PicturesA[0] & @CRLF)

        For $FilterData In $FiltersA
            For $PictureName In $PicturesA
                ConsoleWrite(@ScriptLineNumber & " " & $FilterData & '   ' & $PictureName & @CRLF)
                ;StringInStr
                ;StringRegExp
            Next
        Next

        LogFile(@ScriptLineNumber & " FilterFileList: Both files opened OK", True)



    EndIf

EndFunc   ;==>FilterFileList
;-----------------------------------------------
Func TaskGoToEXEFolder()
    GoToExeFolder()
EndFunc   ;==>TaskGoToEXEFolder
;-----------------------------------------------
Func GoToEXEFolder()
    ;LogFile(@ScriptLineNumber & " GoToEXEFolder  " & $pathx)
    ;LogFile(@ScriptLineNumber & " ShellExecuteWait results: " & ShellExecuteWait("explorer.exe", $pathx, "", "open", @SW_SHOW))
    GuiDisable($GUI_ENABLE)
    ConsoleWrite(@ScriptLineNumber & "  " & @ScriptDir & @CRLF)
    LogFile(@ScriptLineNumber & " ShellExecuteWait results: " & ShellExecuteWait("explorer.exe", @ScriptDir, "", "open", @SW_SHOW))
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>GoToEXEFolder
;-----------------------------------------------
Func TaskGoToCurrentFolder()
    GoToCurrentFolder("main")

EndFunc   ;==>TaskGoToCurrentFolder
;-----------------------------------------------
Func GoToCurrentFolder($type)
    GuiDisable($GUI_DISABLE)

    ConsoleWrite(@ScriptLineNumber & " GoToCurrentFolder: " & $type & @CRLF)
    Local $File
    Select
        Case $type = "main"
            $File = GUICtrlRead($EditStatusMain)
        Case $type = "select"
            Local $x = _GUICtrlTreeView_GetSelection($TreeViewActiveFiles)
            $File = _GUICtrlTreeView_GetText($TreeViewActiveFiles, $x)
    EndSelect

    If Not FileExists($File) Then
        MsgBox(16, "Path not found", "path not found:" & @CRLF & $File)
        Return
    EndIf

    Local $pathx
    Local $path = StringSplit($File, "\") ; filesplit
    For $x = 1 To $path[0] - 1
        $pathx &= $path[$x] & "\"
    Next

    ConsoleWrite(@ScriptLineNumber & " pathx: " & $pathx & @CRLF)

    LogFile(@ScriptLineNumber & " GoToCurrentFolder  " & $pathx)
    LogFile(@ScriptLineNumber & " ShellExecuteWait results: " & ShellExecuteWait("explorer.exe", $pathx, "", "open", @SW_SHOW))
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>GoToCurrentFolder
;-----------------------------------------------
Func ReloadCurrentFile()
    GuiDisable($GUI_DISABLE)
    Local $File = GUICtrlRead($EditStatusMain)
    LogFile(@ScriptLineNumber & " ReloadCurrentFile " & $File)
    ConsoleWrite(@ScriptLineNumber & " ReloadCurrentFile " & $File & @CRLF)
    ProcessWallpaper($File)
    ;ConvertFile($File)
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>ReloadCurrentFile
;-----------------------------------------------
Func ShowMainForm()
    ;MoveWindows()
    GUISetState(@SW_HIDE, $OptionForm)
    GUISetState(@SW_SHOW, $MainForm)
    GUISetState(@SW_HIDE, $SelectFilesForm)
    GUISetState(@SW_HIDE, $SelectFoldersForm)
EndFunc   ;==>ShowMainForm
;-----------------------------------------------
Func ShowSelectForm()
    ; MoveWindows()
    GUISetState(@SW_HIDE, $OptionForm)
    GUISetState(@SW_HIDE, $MainForm)
    GUISetState(@SW_SHOW, $SelectFilesForm)
    GUISetState(@SW_HIDE, $SelectFoldersForm)
EndFunc   ;==>ShowSelectForm
;-----------------------------------------------
Func ShowOptionsForm()
    ; MoveWindows()
    GUISetState(@SW_SHOW, $OptionForm)
    GUISetState(@SW_HIDE, $MainForm)
    GUISetState(@SW_HIDE, $SelectFilesForm)
    GUISetState(@SW_HIDE, $SelectFoldersForm)
EndFunc   ;==>ShowOptionsForm
;-----------------------------------------------
Func HideAllForms()
    GUISetState(@SW_HIDE, $OptionForm)
    GUISetState(@SW_HIDE, $MainForm)
    GUISetState(@SW_HIDE, $SelectFilesForm)
    GUISetState(@SW_HIDE, $SelectFoldersForm)
EndFunc   ;==>HideAllForms
;-----------------------------------------------
Func LogFileLimit()
    LogFile(@ScriptLineNumber & " LogFileLimit")
    Const $MAXFS = 30000
    Local $AA[1]
    LogFile(@ScriptLineNumber & " Current size:" & FileGetSize($Log_filename) & " Max size:" & $MAXFS)
    If FileGetSize($Log_filename) > $MAXFS Then
        _FileReadToArray($Log_filename, $AA)
        FileDelete($Log_filename)
        For $x = UBound($AA) / 2 To UBound($AA) - 1
            FileWriteLine($Log_filename, $AA[$x])
        Next
    EndIf
    LogFile(@ScriptLineNumber & " LogFileLimit completed. Current size:" & FileGetSize($Log_filename))
EndFunc   ;==>LogFileLimit
;-----------------------------------------------
Func ExitEvent()
    Exit
EndFunc   ;==>ExitEvent
;-----------------------------------------------
; hotkey F11
Func GUI_Enable()
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>GUI_Enable
;-----------------------------------------------
Func GuiDisable($choice) ;$GUI_ENABLE $GUI_DISABLE
    For $x = 1 To 200
        GUICtrlSetState($x, $choice)
    Next
EndFunc   ;==>GuiDisable
;-----------------------------------------------
Func ChangeDestopWallpaper($sImg, $iStyle = 0)
    Local $sKey = "HKEY_CURRENT_USER\Control Panel\Desktop"
    Local $SPI_SETDESKWALLPAPER = 20
    Local $SPIF_UPDATEINIFILE = 1
    Local $SPIF_SENDCHANGE = 2
    If $iStyle = 1 Then
        RegWrite($sKey, "TileWallPaper", "REG_SZ", 1)
        RegWrite($sKey, "WallpaperStyle", "REG_SZ", 0)
    Else
        RegWrite($sKey, "TileWallPaper", "REG_SZ", 0)
        RegWrite($sKey, "WallpaperStyle", "REG_SZ", $iStyle)
    EndIf

    DllCall("user32.dll", "int", "SystemParametersInfo", _
            "int", $SPI_SETDESKWALLPAPER, _
            "int", 0, _
            "str", $sImg, _
            "int", BitOR($SPIF_UPDATEINIFILE, $SPIF_SENDCHANGE))

    Return 0
EndFunc   ;==>ChangeDestopWallpaper
;-----------------------------------------------
;The following section is for select folders
;-----------------------------------------------
Func ChoseFolders()
    Local $tmp = FileSelectFolder("Chose a folder", "", 7, $WorkingFolder)
    If $tmp = "" Then Return
    $WorkingFolder = _AddSlash2PathString($tmp)
    GUICtrlSetData($InputWorkingFolder, $WorkingFolder)
EndFunc   ;==>ChoseFolders
;-----------------------------------------------
Func GetFolders()
    _GUICtrlTreeView_DeleteAll($TreeViewSelectFolders)
    _GUICtrlListBox_ResetContent($ListSelectFolderResults)

    GuiDisable($GUI_DISABLE)
    ReDim $FolderHandleArray[1]
    Local $FolderArray[1]

    ;Func _FileListToArrayR($sPath, $sFilter = "*", $iFlag = 0, $iRecurse = 0, $iBaseDir = 1, $sExclude = "", $i_Options = 1)

    If GUICtrlRead($CheckBoxSelectFolderRecursive) = $GUI_CHECKED Then
        ConsoleWrite(@ScriptLineNumber & " #1 Recursive " & $WorkingFolder & @CRLF)
        $FolderArray = _FileListToArrayR($WorkingFolder, "*", 2, 1, 1, "", 1)
    Else
        ConsoleWrite(@ScriptLineNumber & " #2 Not Recursive " & $WorkingFolder & @CRLF)
        $FolderArray = _FileListToArrayR($WorkingFolder, "*", 2, 0, 1, "", 1)
        ; $FolderArray = _FileListToArray($WorkingFolder, "*")
    EndIf
    _ArrayDelete($FolderArray, 0)
    _ArrayUnique($FolderArray)
    _ArraySort($FolderArray)

    _ArrayInsert($FolderArray, 0, $WorkingFolder)

    If IsArray($FolderArray) Then
        For $x In $FolderArray
            ;_GUICtrlTreeView_Add($TreeViewFolders, 0, $X)
            _ArrayAdd($FolderHandleArray, _GUICtrlTreeView_Add($TreeViewSelectFolders, 0, $x))
        Next
        _ArrayDelete($FolderHandleArray, 0)
    Else
        MsgBox(16, "Path not found", "Path not found:" & @CRLF & $WorkingFolder)
    EndIf
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>GetFolders
;-----------------------------------------------
Func ToggleFoldersChecked()
    ;_ArrayDisplay($FolderHandleArray)
    Static $ToggleFolderState
    $ToggleFolderState = Not $ToggleFolderState
    For $x = 0 To UBound($FolderHandleArray) - 1
        _GUICtrlTreeView_SetChecked($TreeViewSelectFolders, $FolderHandleArray[$x], $ToggleFolderState)
    Next
EndFunc   ;==>ToggleFoldersChecked
;-----------------------------------------------\
Func ProcessFolders()
    GuiDisable($GUI_DISABLE)
    Local $ArrayTmp[1]

    If Not IsArray($FolderHandleArray) Then
        MsgBox(48, "No folders", "No folders to process")
        GuiDisable($GUI_ENABLE)
        Return
    EndIf

    ;_ArrayDisplay($FolderHandleArray, @ScriptLineNumber)
    _GUICtrlListBox_ResetContent($ListSelectFolderResults)
    For $x = 0 To UBound($FolderHandleArray) - 1
        Local $AA = _GUICtrlTreeView_GetChecked($TreeViewSelectFolders, $FolderHandleArray[$x])
        Local $BB = _GUICtrlTreeView_GetText($TreeViewSelectFolders, $FolderHandleArray[$x])
        If $AA Then _ArrayAdd($ArrayTmp, $BB)
    Next

    _ArrayDelete($ArrayTmp, 0)

    If GUICtrlRead($CheckBoxSelectFolderRecursive) = $GUI_CHECKED Then CleanPaths($ArrayTmp)

    ConsoleWrite(@ScriptLineNumber & " " & UBound($ArrayTmp) & @CRLF)
    If Not IsArray($ArrayTmp) Then
        MsgBox(48, "No folders", "No folders to process")
        GuiDisable($GUI_ENABLE)
        Return
    EndIf

    For $x In $ArrayTmp
        _GUICtrlListBox_AddString($ListSelectFolderResults, $x)
    Next
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>ProcessFolders
;-----------------------------------------------
;Remove any sub-paths if a parent is already included
Func CleanPaths(ByRef $Array)
    For $x = 0 To UBound($Array) - 1
        For $Y = 0 To UBound($Array) - 1
            If StringInStr($Array[$Y], $Array[$x]) <> 0 And _
                    StringCompare($Array[$Y], $Array[$x]) <> 0 Then
                $Array[$Y] = ""
            EndIf
        Next
    Next
    _RemoveBlankLines($Array)
EndFunc   ;==>CleanPaths
;-----------------------------------------------
Func GetFiles() ; dbk
    GuiDisable($GUI_DISABLE)
    LogFile(@ScriptLineNumber & " GetFiles started")
    SplashImageOn("Working", $AuxPath & "Wallpaper.jpg", -1, -1, -1, -1, 18)
    Local $FileArray

    If GUICtrlRead($CheckAppendFiles) = $GUI_UNCHECKED Then
        _GUICtrlTreeView_DeleteAll($TreeViewActiveFiles)
        ReDim $FileHandleArray[1]
    EndIf
    GUICtrlSetData($LabelActiveCount, 0)

    For $x = 0 To _GUICtrlListBox_GetCount($ListSelectFolderResults) - 1
        Local $File = _GUICtrlListBox_GetText($ListSelectFolderResults, $x)

        ;$FolderArray = _FileListToArray R($sPath, $sFilter = "*", $iFlag = 0, $iRecurse = 0, $iBaseDir = 1, $sExclude = "", $i_Options = 1)

        If GUICtrlRead($CheckBoxSelectFolderRecursive) = $GUI_CHECKED Then
            ConsoleWrite(@ScriptLineNumber & " #1 Recursive " & $File & @CRLF)
            $FileArray = _FileListToArrayR($File, "*", 1, 1, 1, "ini", 1)
            _ArrayDelete($FileArray, 0)
            _ArrayInsert($FileArray, 0, $File)

            If IsArray($FileArray) Then
                For $i In $FileArray
                    If StringInStr($i, ":") = 2 And VerifyFileTypes($i) >= 0 Then
                        _ArrayAdd($FileHandleArray, _GUICtrlTreeView_Add($TreeViewActiveFiles, 0, $i))
                    EndIf
                Next
            EndIf
        Else
            ConsoleWrite(@ScriptLineNumber & " #2 Not Recursive " & $File & @CRLF)
            $FileArray = _FileListToArrayR($File, "*", 1, 0, 1, "ini", 1)
            _ArrayDelete($FileArray, 0)
            _ArrayInsert($FileArray, 0, $File)
            If IsArray($FileArray) Then
                For $i In $FileArray
                    If StringInStr($i, ":") = 2 And VerifyFileTypes($i) >= 0 Then
                        _ArrayAdd($FileHandleArray, _GUICtrlTreeView_Add($TreeViewActiveFiles, 0, $i))
                    EndIf
                Next
            EndIf
        EndIf
    Next

    GUICtrlSetData($LabelActiveCount, _GUICtrlTreeView_GetCount($TreeViewActiveFiles))

    ;_ArrayDisplay($FileHandleArray, @ScriptLineNumber)
    SplashOff()
    LogFile(@ScriptLineNumber & " GetFiles completed")
    GuiDisable($GUI_ENABLE)
EndFunc   ;==>GetFiles
;-----------------------------------------------


