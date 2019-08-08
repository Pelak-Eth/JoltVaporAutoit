#include <Excel.au3>
#include <IE.au3>
#include <File.au3>
#include <String.au3>
#include <Array.au3>
#include <Date.au3>
#include <MsgBoxConstants.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <date.au3>
#include <Misc.au3>
#Include <Restart.au3>

;##########################
;Allow only one instance ##
;##########################
If _Singleton("Print Labels On Screen", 1) = 0 Then
	MsgBox(4096, "ALREADY RUNNING!", "CLOSING DUPLICATE INSTANCE", 3)
    Exit
EndIf

;#####################################################################################################
;Global Variables																					##
Global $sFilePath = @ScriptDir & "\Label Spreadsheet CSV.csv" ;								        ##
;Global $oExcel = _ExcelBookOpen($sFilePath1) ;														##
Global $response; tnaah

;Function for running other Autoit scripts from here
Func _RunAU3($sFilePath, $sWorkingDir = "", $iShowFlag = @SW_SHOW, $iOptFlag = 0)
    Return Run('"' & @AutoItExe & '" /AutoIt3ExecuteScript "' & $sFilePath & '"', $sWorkingDir, $iShowFlag, $iOptFlag)
EndFunc

#Region ### START Koda GUI section ### Form=
$Form1_1 = GUICreate("Form1", ((@DesktopWidth/2) - 350), 142, (@DesktopWidth/2)-46, (@DesktopHeight - 142), $WS_POPUP, BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW))
GUISetBkColor(0x000000)

$Print = GUICtrlCreateButton("LABEL PRINTING", 0, 0, ((@DesktopWidth/2) - 350), 160)
GUICtrlSetFont(-1, 32, 400, 0, "skrunch")
GUICtrlSetColor(-1, 0xC8C8C8)
GUICtrlSetBkColor(-1, 0x000000)

;~ Dim $Form1_AccelTable[1][2] = [["^!p", $Print]]
;~ GUISetAccelerators($Form1_AccelTable)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

;MAIN CONTROL LOOP GUI
While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit

		Case $Print
			If WinExists("P-touch Editor") Then
				WinClose("P-touch Editor")
			EndIf
			;MsgBox(4096, "Hit Print", "")
			While 1
			getLabels()
			Sleep(1000)
				If WinExists("P-touch Editor") Then
					Sleep(2000)
					SplashTextOn("Title", "PRINTING LABELS NOW", -1, 60, -1, -1, 1, "", 24, 700)
					ControlClick("P-touch Editor", "", "[ID:5351]", "left")
				    winwait("Print");wait for print dialog window
					ControlClick("Print", "", "[ID:5202]", "left") ;print all sheets radio box
					;MsgBox(4096, "WAIT HERE", "", 3000)
					ControlClick("Print", "", "[ID:1]", "left") ; click print
					;ControlClick("P-touch Editor", "", "[ID:5350]", "left", 1); sends command to print
					;MsgBox(4096, "Print Commanda sent", "")
					Sleep(2000)
					SplashOff()

				Else
					ExitLoop
				EndIf
			Sleep(2000)
			;WinClose("P-touch Editor")
			If ProcessExists("Ptedit51.exe") Then ; Check if P-Touch is running
				;MsgBox($MB_SYSTEMMODAL, "", "P-Touch is running")
				ProcessClose("Ptedit51.exe")
			EndIf
				ExitLoop
			WEnd

        Case $GUI_EVENT_SECONDARYUP
            $aCInfo = GUIGetCursorInfo($Form1_1)
            If $aCInfo[4] = $Print Then
                If MsgBox(36, 'Restarting...', 'Press Yes to restart this script.') = 6 Then
					_ScriptRestart()
				EndIf
            EndIf
	EndSwitch
WEnd

Func getLabels()
	local $flavor[0]
	local $nicNumber[0]

	#Region GET DATE IN NICE READABLE FORMAT  XX-XX-XXXX
	$timeSys = _Date_Time_GetSystemTime()
	$unformattedTime = _Date_Time_SystemTimeToDateTimeStr($timeSys)
	$slashTime = StringRegExpReplace($unformattedTime, "/", "-")
	$timefull = StringLeft($slashTime, 10)
	Local $time = StringRegExpReplace($timefull, "20(?=\d\d)", "")
	;MsgBox(4096, "Date", $time)
	#EndRegion

	If Not WinExists("Touch Order") Then
		$message = "WRONG PAGE.  DO YOU WANT TO PRINT A CUSTOM LABEL?"
		Call("pMsgBox", $message); Yes --> $response = 1 ,, No --> $response = 0 ,, $message = string text
			If $response = 1 Then
				_RunAU3("Custom Label Print GUI.au3", @ScriptDir)
				SplashTextOn("Title", "LOADING....  PLEASE WAIT", -1, 60, -1, -1, 1, "", 24, 700)
				Sleep(2000)
				SplashOff()
				Return
			ElseIf $response = 0 Then
				Return
			EndIf
	EndIf

	WinWait("Touch Order")
	Local $oIE1 = _IEAttach("Touch Order")
	_IEAction($oIE1, "refresh")
	If @error <> 0 Then
		;do nothing
	EndIf
	_IELoadWait($oIE1, 500, 30000)
		If @error Then
			MsgBox(4096, "ERROR: ", @error & @extended)
		EndIf
	Local $oIE = _IEAttach("Touch Order")
	Local $oTable = _IETableGetCollection($oIE, 0)
		If @error = 7 Then
			$message = "NO ORDERS.  DO YOU WANT TO PRINT A CUSTOM LABEL?"
			Call("pMsgBox", $message); Yes --> $response = 1 ,, No --> $response = 0 ,, $message = string text
				If $response = 1 Then
					_RunAU3("Custom Label Print GUI.au3", @ScriptDir)
					SplashTextOn("Title", "LOADING....  PLEASE WAIT", -1, 60, -1, -1, 1, "", 24, 700)
					Sleep(2000)
					SplashOff()
					Return
				ElseIf $response = 0 Then
					Return
				EndIf
			Return
		Else
			SplashTextOn("Title", "STARTING TO PRINT NOW....", -1, 60, -1, -1, 1, "", 24, 700)
			Local $orderArray = _IETableWriteToArray($oTable)
			;_ArrayDisplay($orderArray, "order array = ")
			Local $hFileOpen = FileOpen($sFilePath, $FO_OVERWRITE)
			Sleep(1500)
					If (UBound($orderArray, 2) - 1) >= 9 Then
						$numberOfLabels = 9
					Else
						$numberOfLabels = (UBound($orderArray, 2) - 1)
					 EndIf
					 $csvData = "";define string
				For $i = 1 To $numberOfLabels
					;MsgBox(4096, "Array Size = ", $LabelNum)
					$flavorText = $orderArray[1][$i] ;flavor name
					$nicData = $orderArray[5][$i]; nic starts at [5]
					$nicNumberArray = _StringBetween($nicData, '"', '"')
					;_ArrayDisplay($nicNumberArray, "Nic Data = ")
					$nicText = $nicNumberArray[0];write nic
					$vgPG = $orderArray[6][$i]; read VG / PG
					$vgPGArray = _StringBetween($vgPG, '"', '"')
					;_ArrayDisplay($vgPGArray, "VG / PG array = ")
					$vgText = $vgPGArray[0] ;write VG
					$mentholText = $orderArray[3][$i]; write menthol
					$sizeText = $orderArray[2][$i]
					$partData = $flavorText&","&$nicText&","&$time&","&$vgText&","&$mentholText&","&$sizeText&"," ; combine label data to single string
				    $csvData &= $partData ; combine all label data
				    ;MsgBox(4096, "text = ", $csvData)
			   Next
				    FileWriteLine($hFileOpen, "flavor1,nic1,date1,vg1,menthol1,size1,flavor2,nic2,date2,vg2,menthol2,size2,flavor3,nic3,date3,vg3,menthol3,size3,flavor4,nic4,date4,vg4,menthol4,size4,flavor5,nic5,date5,vg5,menthol5,size5,flavor6,nic6,date6,vg6,menthol6,size6,flavor7,nic7,date7,vg7,menthol7,size7,flavor8,nic8,date8,vg8,menthol8,size8,flavor9,nic9,date9,vg9,menthol9,size9") ; CSV name headers
				    FileWriteLine($hFileOpen, $csvData)
					SplashOff()
					FileClose($hFileOpen)
					FileFlush($hFileOpen)
					sleep(1000)
					If $numberOfLabels >= 9 Then
						MsgBox(4096, "PRINT MSG", "ONLY PRINTING 9 LABELS", 5)
						ShellExecute(@ScriptDir&"\9dymo.lbx", "", "", "", @SW_HIDE)
					ElseIf $sizeText == "250 mL" Then
						ShellExecute(@ScriptDir&"\250label.lbx", "", "", "open", @SW_HIDE)
					Else
						ShellExecute(@ScriptDir&"\"&$numberOfLabels&"dymo.lbx", "", "", "open", @SW_HIDE)
					EndIf
		EndIf
	EndFunc

Func pMsgBox($message)
	#Region ### START Koda GUI section ### Form=C:\Users\windows\Documents\Autoit\PMsgBox.kxf
	$PMSGBOX = GUICreate("PMsgBox", 597, 210, (@DesktopWidth/2)-299, 217, $WS_POPUP, BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW))
	GUISetBkColor(0x000000)

	$LabelMessage = GUICtrlCreateLabel($message, 17, 8, 571, 28)
	GUICtrlSetFont(-1, 16, 400, 0, "MS Sans Serif")
	GUICtrlSetColor(-1, 0xC8C8C8)
	GUICtrlSetBkColor(-1, 0x000000)

	$Yes = GUICtrlCreateButton("YES", 1, 48, 257, 161)
	GUICtrlSetFont(-1, 72, 400, 2, "skrunch")
	GUICtrlSetColor(-1, 0xC8C8C8)
	GUICtrlSetBkColor(-1, 0x000000)

	$No = GUICtrlCreateButton("NO", 328, 48, 265, 161)
	GUICtrlSetFont(-1, 72, 400, 2, "skrunch")
	GUICtrlSetColor(-1, 0xC8C8C8)
	GUICtrlSetBkColor(-1, 0x000000)

	GUISetState(@SW_SHOW)

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				Exit

			Case $Yes
				$response = 1
				GUIDelete()
				Return $response

			Case $No
				$response = 0
				GUIDelete()
				Return $response

		EndSwitch
	WEnd
	#EndRegion ### END Koda GUI section #############################################
EndFunc