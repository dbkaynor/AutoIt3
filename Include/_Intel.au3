#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Run_Tidy=y
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****

#cs  This is a list of functions in this UDF
	Func OIDs()
	Func WireShark()
	Func EventMonitor()
	Func EventViewer()
	Func ResourceMonitor()
#ce

#include-once
Opt("MustDeclareVars", 1) ; 0=no, 1=require pre-declare

;#include <Array.au3>
;#include <Date.au3>
;#include <File.au3>

;-----------------------------------------------
Func OIDs()
	Local $OIDsString = 'c:\4.01-with EEE\bin\oids3.exe'
	If Not FileExists($OIDsString) Then

		$OIDsString = 'c:\2.60\oids.exe'
		If Not FileExists($OIDsString) Then
			If MsgBox(20, 'Error launching OIDs', 'Error launching OIDs:' & @CRLF & 'Should I install it?') = 6 Then
				If DirCopy(@ScriptDir & '\AUXFiles\Utils\OIDs-O-Matic\4.01-with EEE', 'c:\4.01-with EEE', 1) <> 1 Then
					MsgBox(16, 'Copy failed', '4.01-with EEE')
				EndIf
				If DirCopy(@ScriptDir & '\AUXFiles\Utils\OIDs-O-Matic\2.60', 'c:\2.60', 1) <> 1 Then
					MsgBox(16, 'Copy failed', '2.60')
				EndIf
				MsgBox(48, 'Files copied.', '4.01-with EEE and 2.60 versions copied')
				Return
			EndIf
			Return
		EndIf
	EndIf
	ShellExecute($OIDsString)
EndFunc   ;==>OIDs
;-----------------------------------------------
Func WireShark()
	Local $WireSharkString
	Local $WireSharkInstallString

	Switch @OSArch
		Case 'x86'
			$WireSharkString = 'C:\Program Files (x86)\wireshark\wireshark.exe'
			$WireSharkInstallString = @ScriptDir & '\AUXFiles\Utils\wireshark-win32-1.65.exe'

		Case 'x64'
			$WireSharkString = 'C:\Program Files\wireshark\wireshark.exe'
			$WireSharkInstallString = @ScriptDir & '\AUXFiles\Utils\wireshark-win64-1.65.exe'
		Case Else
			MsgBox(16, 'Wireshark case error', 'How did we get here? ' & @OSArch)
	EndSwitch

	If Not FileExists($WireSharkString) Then
		If MsgBox(20, 'Error launching Wireshark', 'Error launching Wireshark:' & @CRLF & 'Should I install it?') = 6 Then
			If ShellExecute($WireSharkInstallString) <> 1 Then
				MsgBox(16, 'Wire Shark Install failed', 'Wire Shark Install failed')
				Return
			EndIf
			Return
		EndIf
	EndIf
	ShellExecute($WireSharkString)
EndFunc   ;==>WireShark
;-----------------------------------------------
Func EventMonitor()
	Local $EventMonitorString = @ScriptDir & "\eventmon.exe"
	If Not FileExists($EventMonitorString) Then
		MsgBox(16, 'Error launching EventMonitor', 'Error launching EventMonitor:' & @CRLF & $EventMonitorString)
	EndIf
	ShellExecuteWait($EventMonitorString)
EndFunc   ;==>EventMonitor
;-----------------------------------------------
Func NICViewer()
	Local $NICViewerString = @ScriptDir & "\nicviewer.exe"
	If Not FileExists($NICViewerString) Then
		MsgBox(16, 'Error launching NICViewer', 'Error launching EventMonitor:' & @CRLF & $NICViewerString)
	EndIf
	ShellExecuteWait($NICViewerString)
EndFunc   ;==>NICViewer
;-----------------------------------------------
Func EventViewer()
	ShellExecuteWait("eventvwr.msc", "/s")
EndFunc   ;==>EventViewer
;-----------------------------------------------
Func ResourceMonitor()
	ShellExecute("resmon.exe")
EndFunc   ;==>ResourceMonitor
;-----------------------------------------------
Func PerformanceMonitor()
	ShellExecute("perfmon.exe")
EndFunc   ;==>PerformanceMonitor
;-----------------------------------------------
Func DeviceManager()
	EnvSet('devmgr_show_nonpresent_devices', '1')
	EnvUpdate()
	ShellExecute("devmgmt.msc")
EndFunc   ;==>DeviceManager
;-----------------------------------------------
Func GroupPolicyEditor()
	ShellExecute("gpedit.msc")
EndFunc   ;==>GroupPolicyEditor
;-----------------------------------------------
Func SecurityPolicyEditor()
	;set devmgr_show_nonpresent_devices=1
	ShellExecute("secpol.msc")
EndFunc   ;==>SecurityPolicyEditor
;-----------------------------------------------
Func Services()
	ShellExecute("services.msc")
EndFunc   ;==>Services
;-----------------------------------------------
Func NetworkSharing()
	ShellExecute("explorer.exe", "shell:::{8E908FC9-BECC-40f6-915B-F4CA0E70D03D}")
EndFunc   ;==>NetworkSharing
;-----------------------------------------------
Func NetworkControlPanel()
	ShellExecute('ncpa.cpl')
EndFunc   ;==>NetworkControlPanel
;-----------------------------------------------
Func FirewallControlPanel()
	ShellExecute('firewall.cpl')
EndFunc   ;==>FirewallControlPanel
;-----------------------------------------------

#cs
	
	This is pre-built menu code
	Global $ToolsMenu = GUICtrlCreateMenu('Tools')
	Global $OIDsItem = GUICtrlCreateMenuItem('OIDs', $ToolsMenu)
	Global $WireSharkItem = GUICtrlCreateMenuItem('WireShark', $ToolsMenu)
	Global $EventMonitorItem = GUICtrlCreateMenuItem('Event Monitor', $ToolsMenu)
	Global $EventViewerItem = GUICtrlCreateMenuItem('Event Viewer', $ToolsMenu)
	Global $PerformanceMonitorItem = GUICtrlCreateMenuItem('Performance Monitor', $ToolsMenu)
	Global $ResourceMonitorItem = GUICtrlCreateMenuItem('Resource Monitor', $ToolsMenu)
	Global $DeviceManagerItem = GUICtrlCreateMenuItem('Device Manager', $ToolsMenu)
	Global $GroupPolicyEditorItem = GUICtrlCreateMenuItem('Group Policy Editor', $ToolsMenu)
	Global $ServicesItem = GUICtrlCreateMenuItem('Services', $ToolsMenu)
	
	
	
	;--- Menus
	Case $OIDsItem
	LaunchOIDs()
	Case $WireSharkItem
	LaunchWireShark()
	Case $EventMonitorItem
	LaunchEventMonitor()
	Case $EventViewerItem
	LaunchEventviewer()
	Case $ResourceMonitorItem
	ResourceMonitor()
	Case $PerformanceMonitorItem
	PerformanceMonitor()
	Case $DeviceManagerItem
	DeviceManager()
	Case $GroupPolicyEditorItem
	GroupPolicyEditor()
	Case $ServicesItem
	Services()
#ce




