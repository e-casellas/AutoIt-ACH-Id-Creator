#include <IE.au3>
#include <Excel.au3>
#include <MsgBoxConstants.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <GUIConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include "MetroGUI-UDF\MetroGUI_UDF.au3"
#include "MetroGUI-UDF\_GUIDisable.au3"
Opt("WinTitleMatchMode", 2)

;Variables
Global $oIE, $oExcel, $sAccountNum, $sCompanyId, $sCompanyName, $sEntryDesc

;ESC key will stop this bot
HotKeySet("{ESC}", "_Exit")
Func _Exit()
	Exit
EndFunc   ;==>_Exit

#Region GUI
_Metro_EnableHighDPIScaling()
_SetTheme("LightGray")
$GUIThemeColor = 0xeff4ff
$Form1 = _Metro_CreateGUI("ACH ID Tracker Creator", 310, 175, -1, -1, True)
$ButtonBKColor = 0x603cba
$Button1 = _Metro_CreateButtonEx2("Start", 210, 50, 80, 30)
$Button2 = _Metro_CreateButtonEx2("Stop", 210, 100, 80, 30)
$Label1 = GUICtrlCreateLabel("ACH ID Tracker Creator", 50, 5, 123, 17)
GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
GUICtrlSetColor(-1, 0x603cba)
$Label2 = GUICtrlCreateLabel("", 50, 70, 105, 40)
GUICtrlSetFont(-1, 11, 400, 0, "Segoe UI")
GUICtrlSetColor(-1, 0xee1111)
GUICtrlSetState(-1, $GUI_HIDE)
$Control_Buttons = _Metro_AddControlButtons(True, False, True, False, True) ;CloseBtn = True, MaximizeBtn = True, MinimizeBtn = True, FullscreenBtn = True, MenuBtn = True
$GUI_CLOSE_BUTTON = $Control_Buttons[0]
$GUI_MINIMIZE_BUTTON = $Control_Buttons[3]
$GUI_MENU_BUTTON = $Control_Buttons[6]
GUISetState(@SW_SHOW)
#EndRegion GUI

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE, $GUI_CLOSE_BUTTON
			ExitLoop
			Exit
		Case $Form1
		Case $GUI_MINIMIZE_BUTTON
			GUISetState(@SW_MINIMIZE, $Form1)
		Case $GUI_MENU_BUTTON
			Local $MenuButtonsArray[2] = ["About", "Exit"]
			Local $MenuSelect = _Metro_MenuStart($Form1, 50, $MenuButtonsArray)
			Switch $MenuSelect ;Above function returns the index number of the selected button from the provided buttons array.
				Case "0"
				Case "1"
					_Metro_GUIDelete($Form1)
					Exit
			EndSwitch
		Case $Button1
			CheckMenu()
			OpenExcelInput()
			AddLoop()
			_Excel_Close($oExcel)
		Case $Button2
			Exit
		Case $Label1
	EndSwitch
WEnd

Func CheckExists()
	If WinExists("FIS ACH") = 0 Then
		MsgBox($MB_SYSTEMMODAL, "ACH ID Creator", "ACH Tracker window not found. Please make sure you are logged in and try again.")
		Exit
	EndIf
EndFunc   ;==>CheckExists

Func CheckMenu()
	$oIE = _IEAttach("FIS ACH")
	_IEGetObjByName($oIE, "inp1_searchByCompanyId")
	If @error Then
		MsgBox($MB_ICONWARNING, "ACH ID Creator", "Please open menu SETUP>ACH PROFILE>CREATE/UPDATE ACH before running this bot.")
		Exit
	EndIf
EndFunc   ;==>CheckMenu

Func OpenExcelInput()
	$oExcel = _Excel_Open(False, False, False, False, True)
	$sWorkbook = @ScriptDir & "\input.xlsx"
	Global $oWorkbook = _Excel_BookOpen($oExcel, $sWorkbook)
	Global $lLastRow = $oWorkbook.ActiveSheet.UsedRange.Rows.Count
EndFunc   ;==>OpenExcelInput

Func AddLoop()
	For $i = 2 To $lLastRow Step 1
		WinActivate("FIS ACH")
		$oIE = _IEAttach("FIS ACH")
		GUICtrlSetData($Label2, "Working on cell       " & $i & " out of " & $lLastRow)
		GUICtrlSetState($Label2, $GUI_SHOW)
		$sCompanyId = _Excel_RangeRead($oWorkbook, "Sheet1", "A" & $i, 1)
		$sCompanyName = _Excel_RangeRead($oWorkbook, "Sheet1", "B" & $i, 1)
		$sAccountNum = _Excel_RangeRead($oWorkbook, "Sheet1", "C" & $i, 1)
		$sEntryDesc = _Excel_RangeRead($oWorkbook, "Sheet1", "D" & $i, 1)

		Write($oIE, "inp1_searchByCompanyId", $sCompanyId)
		Click($oIE, "selectButton")
		While 1
			If WinExists("Profile Search Results") Or WinExists("Create ACH Profile") Then
				ExitLoop
				Sleep(100)
			EndIf
		WEnd
		If WinExists("Profile Search Results") Then
			$oIE = _IEAttach("Profile Search Results")
			Sleep(500)
			Click($oIE, "createNew")
		Else
			$oIE = _IEAttach("Create ACH Profile")
		EndIf
		Do
			$oIE = _IEAttach("Create ACH Profile")
		Until Not @error
		Do
			_IEGetObjById($oIE, "COMPANY_NAME")
		Until Not @error
		Write($oIE, "COMPANY_NAME", $sCompanyName)
		Write($oIE, "POINT", "ORG493AA ")
		Write($oIE, "ABA", "123456789")
		Write($oIE, "ENTRY_DESC", $sEntryDesc)
		Write($oIE, "BATCH", "U")
		Write($oIE, "OFFSET_TIMING", "S")
		Write($oIE, "SETTLE_TYPE", "U")
		Write($oIE, "OFFSET_ENTRY", "Checking")
		Write($oIE, "OFFSET_ABA", "123456789")
		Write($oIE, "OFFSET_ACCT", $sAccountNum)
		Write($oIE, "RETURN_ACCT", $sAccountNum)
		Write($oIE, "ORIGINATION_ACCT", $sAccountNum)
		Write($oIE, "FULL_COMPANY_NAME", $sCompanyName)
		Write($oIE, "CONTACT_NAME", "CMOU")
		Write($oIE, "phoneAreaCode_CONTACT_PHONE", "888")
		Write($oIE, "phonePrefix_CONTACT_PHONE", "555")
		Write($oIE, "phoneSuffix_CONTACT_PHONE", "7777")
		Click($oIE, "apply")
		_IELoadWait($oIE)
		Sleep(200)
		Print()
	Next
EndFunc   ;==>AddLoop

Func Print()
	_IEAction($oIE, "print")
	WinWait("Print")
	WinActivate("Print")
	ControlClick("Print", "", "Button10")
	Sleep(500)
	ControlClick("Print", "", "Button13")
	Sleep(500)
	Do
		WinActivate("Review ACH Profile")
	Until Not @error
	Sleep(500)
	Send("^w")
	Sleep(500)
EndFunc   ;==>Print

#Region MyFunctions ===================================================================
Func Click($Tab, $ObjIdOrName)
	$Obj = _IEGetObjByName($Tab, $ObjIdOrName)
	If @error Then $oObj = _IEGetObjById($Tab, $ObjIdOrName)
	_IEAction($Obj, "click")
EndFunc   ;==>Click

Func Write($Tab, $ObjIdOrName, $Text)
	$Obj = _IEGetObjByName($Tab, $ObjIdOrName)
	If @error Then $oObj = _IEGetObjById($Tab, $ObjIdOrName)
	_IEFormElementSetValue($Obj, $Text)
EndFunc   ;==>Write
#EndRegion MyFunctions ===================================================================

Exit
