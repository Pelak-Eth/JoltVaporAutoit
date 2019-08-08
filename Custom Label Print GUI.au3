#include <Excel.au3>
#include <File.au3>
#include <String.au3>
#include <GUIConstantsEx.au3>
#include <Array.au3>
#include <GuiComboBox.au3>
#include <MsgBoxConstants.au3>
#include <WindowsConstants.au3>
#include <ButtonConstants.au3>
#include <Misc.au3>
#include <SQLite.au3>
#include <SQLite.dll.au3>
#include <GuiComboBox.au3>
#include <date.au3>

;##########################
;Allow only one instance ##
;##########################
If _Singleton("Custom Label Print GUI", 1) = 0 Then
	MsgBox(4096, "ALREADY RUNNING!", "CLOSING DUPLICATE INSTANCE", 3)
    Exit
EndIf

Global $rQuantity = 1

Func _JoltStyle()
	GUICtrlSetColor(-1, 0xC8C8C8)
	GUICtrlSetBkColor(-1, 0x000000)
EndFunc

Call("sql_Lite")
Func sql_Lite()
_SQLite_Startup()
If @error Then
    MsgBox($MB_SYSTEMMODAL, "SQLite Error", "SQLite3.dll Can't be Loaded!")
    Exit -1
EndIf
ConsoleWrite("_SQLite_LibVersion=" & _SQLite_LibVersion() & @CRLF)
Local $sDbName = @ScriptDir & "\JoltRecipesDB"
Local $hDskDb = _SQLite_Open($sDbName) ; Open a permanent disk database
If @error Then
    MsgBox($MB_SYSTEMMODAL, "SQLite Error", "Can't open or create a permanent Database!")
    Exit -1
EndIf

; Query
Local $aArray1, $iRows, $iColumns, $iRval
Global $aResult = _SQLite_GetTable2d(-1, "SELECT * FROM Recipes order by Recipe asc;", $aArray1, $iRows, $iColumns)
;_ArrayDisplay($aArray1, "Array = ")
If $aResult = $SQLITE_OK Then
    ;_SQLite_Display2DResult($aResult)
Else
    MsgBox($MB_SYSTEMMODAL, "SQLite Error: " & $iRval, _SQLite_ErrMsg())
	Exit
EndIf
Global $aArray = $aArray1

EndFunc

#Region ### START Koda GUI section ### Form=C:\Users\windows\Documents\Autoit\Custom Label Print GUI.kxf
$CustomLabelPrint = GUICreate("CustomLabelGUI", 1030, 489, -5, 176, $WS_BORDER, BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW)) ;$WS_POPUP, BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW))
GUISetBkColor(0x000000)

$hCombo = GUICtrlCreateCombo("", 16, 296, 377, 45, $CBS_DROPDOWN)
GUICtrlSetFont(-1, 24, 800, 0, "MS Sans Serif")
_JoltStyle()
	;Populate Recipe Combobox
	Pop_Flavor($hCombo)

$NicCombo = GUICtrlCreateCombo("0", 435, 295, 185, 45, $CBS_DROPDOWN)
GUICtrlSetData(-1, "1|2|3|4|5|6|7|8|9|10|11|12|13|14|15|16|17|18|19|20|21|22|23|24|25|26|27|28|29|30|31|32|33|34|35|36", "0")
GUICtrlSetFont(-1, 24, 800, 0, "MS Sans Serif")
_JoltStyle()

$Label2 = GUICtrlCreateLabel("NIC LEVEL", 442, 256, 165, 37)
GUICtrlSetFont(-1, 24, 400, 2, "skrunch")
_JoltStyle()

$vgCombo = GUICtrlCreateCombo("40", 435, 195, 185, 45, $CBS_DROPDOWNLIST)
GUICtrlSetData(-1, "MAX|50|60|70|80|90", "40")
GUICtrlSetFont(-1, 24, 800, 0, "MS Sans Serif")
_JoltStyle()

$Label3 = GUICtrlCreateLabel("VG PERCENT", 435, 159, 193, 37)
GUICtrlSetFont(-1, 24, 400, 2, "skrunch")
_JoltStyle()

$b10mL = GUICtrlCreateButton("PRINT 10mL BOTTLE", 662, 84, 345, 53)
GUICtrlSetFont(-1, 24, 400, 2, "skrunch")
_JoltStyle()

$b30mL = GUICtrlCreateButton("PRINT 30mL BOTTLE", 662, 153, 345, 53)
GUICtrlSetFont(-1, 24, 400, 2, "skrunch")
_JoltStyle()

$b50mL = GUICtrlCreateButton("PRINT 50mL BOTTLE", 662, 222, 345, 53)
GUICtrlSetFont(-1, 24, 400, 2, "skrunch")
_JoltStyle()

$b100mL = GUICtrlCreateButton("PRINT 100mL BOTTLE", 662, 291, 345, 53)
GUICtrlSetFont(-1, 24, 400, 2, "skrunch")
_JoltStyle()

$Label1 = GUICtrlCreateLabel("TYPE OR SELECT RECIPE", 16, 256, 380, 37)
GUICtrlSetFont(-1, 24, 400, 2, "skrunch")
_JoltStyle()

$Exit = GUICtrlCreateButton("CLOSE WINDOW", 0, 0, 1017, 73)
GUICtrlSetFont(-1, 36, 400, 2, "skrunch")
_JoltStyle()

$MentholCombo = GUICtrlCreateCombo("None", 107, 195, 201, 45, $CBS_DROPDOWNLIST)
GUICtrlSetData(-1, "LIGHT|MEDIUM|HEAVY|SUPER")
GUICtrlSetFont(-1, 24, 800, 0, "MS Sans Serif")
_JoltStyle()

$MentholLabel = GUICtrlCreateLabel("ADD MENTHOL", 107, 159, 204, 37)
GUICtrlSetFont(-1, 24, 400, 2, "skrunch")
_JoltStyle()

$QuantityCombo = GUICtrlCreateCombo("1", 436, 110, 185, 45, $CBS_DROPDOWN)
GUICtrlSetData(-1, "5|10|20")
GUICtrlSetFont(-1, 24, 800, 0, "MS Sans Serif")
_JoltStyle()

$QuantityLabel = GUICtrlCreateLabel("QUANTITY", 449, 72, 158, 37)
GUICtrlSetFont(-1, 24, 400, 2, "skrunch")
_JoltStyle()

$VendorLabel = GUICtrlCreateLabel("VENDOR", 140, 72, 121, 37)
GUICtrlSetFont(-1, 24, 400, 2, "skrunch")
_JoltStyle()

$VendorCombo = GUICtrlCreateCombo("JOLT VAPOR", 15, 109, 377, 45, $CBS_DROPDOWN)
GUICtrlSetData(-1, "DIGITAL VAPOR DEN", "JOLT VAPOR")
GUICtrlSetFont(-1, 24, 400, 2, "skrunch")
_JoltStyle()

$PplasticLabel = GUICtrlCreateButton("PRINT PLASTIC PREMIX LABEL", 20, 409, 313, 53)
GUICtrlSetFont(-1, 16, 400, 2, "skrunch")
_JoltStyle()

$PglassLabel = GUICtrlCreateButton("PRINT GLASS PREMIX LABEL", 354, 408, 313, 53)
GUICtrlSetFont(-1, 16, 400, 2, "skrunch")
_JoltStyle()

$PconcLabel = GUICtrlCreateButton("PRINT CONCENTRATE LABEL", 692, 407, 313, 53)
GUICtrlSetFont(-1, 16, 400, 2, "skrunch")
_JoltStyle()

$Label7 = GUICtrlCreateLabel("BROTHER WHITE LABEL PRINTING", 248, 360, 516, 37)
GUICtrlSetFont(-1, 24, 400, 0, "skrunch")
_JoltStyle()

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()

	;Reregister Window Statuses to detect text, etc
	GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")

	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			_SQLite_Close()
			_SQLite_Shutdown()
			Exit

		Case $Exit
			_SQLite_Close()
			_SQLite_Shutdown()
			Exit

		Case $PglassLabel
			$concTotal = 0; tnaah declare
			$pgTotal = 0
			$vgTotal = 0
			$rName = GUICtrlRead($hCombo)
			$rQuantity = GUICtrlRead($QuantityCombo)

			Local $iIndex = _ArraySearch($aArray, $rName, 0, 0, 0, 0); get Recipe index from recipe array
				If @error <> 0 Then
					MsgBox(4096, "NOT FOUND", "TRY AGAIN")
					Exit
				EndIf

		#Region; Calculate Conc Volume
			For $i = 1 To 9 Step 1
				$aIngredient = $aArray[$iIndex][$i]
					If $aIngredient = "" Then
						ExitLoop
					EndIf
				$sArray = StringSplit($aIngredient, "|")
				;_ArrayAdd($arrayOfIngredients, $sArray[1])
				$singleIngredientVolume = Round(($sArray[3] * 50), 1);50 multiplier so total is 250mL
				$concTotal = $singleIngredientVolume + $concTotal
				;MsgBox(4096, "CONC TOTAL = ", "CONC = " & $concTotal)

		#EndRegion

		#Region; Add Ingredients to Array and Nicely Format them
			Next
			;calculate PG and VG totals
			$pgTotal = (250 * 0.6) - $concTotal
			$vgTotal = 100

			;add nice text
			$pgTotal = $pgTotal & "mL PG"
			$vgTotal = $vgTotal & "mL VG"
			$concTotal = $concTotal & "mL of Premix"

			;CSV 'Database' stuff
			local $sFilePath = @ScriptDir & "\Brother Conc Printing CSV.csv" ;
			Local $hFileOpen = FileOpen($sFilePath, $FO_OVERWRITE)
			$csvData = $rname&","&$concTotal&","&$pgTotal&","&$vgTotal
			FileWriteLine($hFileOpen, "recipeName,concTotal,pgTotal,vgTotal") ; CSV name headers
			FileWriteLine($hFileOpen, $csvData)
			FileClose($hFileOpen)
			FileFlush($hFileOpen)
			sleep(500)
			MsgBox(4096, "READY TO Execute", "CONC= " & $concTotal & ", PG= " & $pgTotal)
			ShellExecute(@ScriptDir & "\Glass Concentrate Premix Labels 0.94 inch.lbx", "", "", "", @SW_HIDE)
			Sleep(500)
			;MsgBox(4096, "PASSED IT", "BLOCKED?")

			If WinExists("P-touch Editor") Then
				Sleep(3000)
				;MsgBox(4096, "Quantity", "Quantity", 30)
				ControlSetText("P-touch Editor", "", "[ID:5148]", $rQuantity)
				SplashTextOn("Title", "PRINTING LABELS NOW", -1, 60, -1, -1, 1, "", 24, 700)
				;ControlClick("P-touch Editor", "", "[ID:5350]", "left", 1); sends command to print
				ControlClick("P-touch Editor", "", "[ID:5351]", "left")
				winwait("Print");wait for print dialog window
				;ControlClick("Print", "", "[ID:5202]", "left") ;print all sheets radio box
				ControlClick("Print", "", "[ID:1]", "left") ; click print
				;MsgBox(4096, "Print Commanda sent", "")
				Sleep(2000)
				SplashOff()
			Else
				;MsgBox(4096, "WTF!", "Window Not Found!")
			EndIf
			If ProcessExists("Ptedit52.exe") Then ; Check if P-Touch is running process is running.
				;MsgBox($MB_SYSTEMMODAL, "", "P-Touch is running")
				ProcessClose("Ptedit52.exe")
			EndIf

		Case $PplasticLabel
			$concTotal = 0; tnaah declare
			$pgTotal = 0
			$vgTotal = 0
			$rName = GUICtrlRead($hCombo)
			$rQuantity = GUICtrlRead($QuantityCombo)

			;CSV 'Database' stuff
			local $sFilePath = @ScriptDir & "\Brother Printing CSV.csv" ;
			Local $hFileOpen = FileOpen($sFilePath, $FO_OVERWRITE)
			$csvData = $rname&","&$concTotal&","&$pgTotal&","&$vgTotal
			FileWriteLine($hFileOpen, "recipeName,concTotal,pgTotal,vgTotal") ; CSV name headers
			FileWriteLine($hFileOpen, $csvData)
			FileClose($hFileOpen)
			FileFlush($hFileOpen)
			sleep(500)

			ShellExecute(@ScriptDir & "\Plastic Premix Labels 0.94 inch.lbx", "", "", "", @SW_HIDE)

			sleep(500)

			If WinExists("P-touch Editor") Then
				Sleep(2000)
				ControlSetText("P-touch Editor", "", "[ID:5148]", $rQuantity)
				SplashTextOn("Title", "PRINTING LABELS NOW", -1, 60, -1, -1, 1, "", 24, 700)
				;ControlClick("P-touch Editor", "", "[ID:5350]", "left", 1); sends command to print
				ControlClick("P-touch Editor", "", "[ID:5351]", "left")
				winwait("Print");wait for print dialog window
				ControlClick("Print", "", "[ID:5202]", "left") ;print all sheets radio box
				ControlClick("Print", "", "[ID:1]", "left") ; click print
				;MsgBox(4096, "Print Commanda sent", "")
				Sleep(2000)
				SplashOff()
			EndIf
			If ProcessExists("Ptedit52.exe") Then ; Check if P-Touch is running process is running.
				;MsgBox($MB_SYSTEMMODAL, "", "P-Touch is running")
				ProcessClose("Ptedit52.exe")
			EndIf

		Case $PconcLabel
			$concTotal = 0; tnaah declare
			$pgTotal = 0
			$vgTotal = 0
			$rName = GUICtrlRead($hCombo)
			$rQuantity = GUICtrlRead($QuantityCombo)

			;CSV 'Database' stuff
			local $sFilePath = @ScriptDir & "\Brother Printing CSV.csv" ;
			Local $hFileOpen = FileOpen($sFilePath, $FO_OVERWRITE)
			$csvData = $rname&","&$concTotal&","&$pgTotal&","&$vgTotal
			FileWriteLine($hFileOpen, "recipeName,concTotal,pgTotal,vgTotal") ; CSV name headers
			FileWriteLine($hFileOpen, $csvData)
			FileClose($hFileOpen)
			FileFlush($hFileOpen)
			sleep(500)

			ShellExecute(@ScriptDir & "\Concentrate Single.lbx", "", "", "", @SW_HIDE)

			sleep(500)

			If WinExists("P-touch Editor") Then
				Sleep(2000)
				ControlSetText("P-touch Editor", "", "[ID:5148]", $rQuantity)
				SplashTextOn("Title", "PRINTING LABELS NOW", -1, 60, -1, -1, 1, "", 24, 700)
				;ControlClick("P-touch Editor", "", "[ID:5350]", "left", 1); sends command to print
				ControlClick("P-touch Editor", "", "[ID:5351]", "left")
				winwait("Print");wait for print dialog window
				ControlClick("Print", "", "[ID:5202]", "left") ;print all sheets radio box
				ControlClick("Print", "", "[ID:1]", "left") ; click print
				;MsgBox(4096, "Print Commanda sent", "")
				Sleep(2000)
				SplashOff()
			EndIf
			If ProcessExists("Ptedit52.exe") Then ; Check if P-Touch is running process is running.
				;MsgBox($MB_SYSTEMMODAL, "", "P-Touch is running")
				ProcessClose("Ptedit52.exe")
			EndIf

		Case $b10mL
			$rName = GUICtrlRead($hCombo)
			$rSize = "10mL"
			$rNic = GUICtrlRead($NicCombo)
			$rVG = GUICtrlRead($vgCombo)
			;$rPG = (100 - $rVG)
			$rMenthol = GUICtrlRead($MentholCombo)
			$rQuantity = GUICtrlRead($QuantityCombo)
			;Call("Validate_Int", $rQuantity)
			$rVendor = GUICtrlRead($VendorCombo)
			Call("Print_Label", $rName, $rSize, $rNic, $rVG, $rMenthol, $rQuantity, $rVendor)

			;Clear Data from inputs
			;GUICtrlSetData($hCombo, "")

			;Re Populate Comboboxes
			Pop_Flavor($hCombo)

		;Set Focus on First Flavor Combobox
		WinActivate("CustomLabelGUI")
		_WinAPI_SetFocus(ControlGetHandle("CustomLabelGUI", "", $hCombo))

		Case $b30mL
			$rName = GUICtrlRead($hCombo)
			$rSize = "30mL"
			$rNic = GUICtrlRead($NicCombo)
			$rVG = GUICtrlRead($vgCombo)
			;$rPG = (100 - $rVG)
			$rMenthol = GUICtrlRead($MentholCombo)
			$rQuantity = GUICtrlRead($QuantityCombo)
			;Call("Validate_Int", $rQuantity)
			;MsgBox(4096, "$rQuantity", "$rQuantity = " & $rQuantity)
			$rVendor = GUICtrlRead($VendorCombo)
			Call("Print_Label", $rName, $rSize, $rNic, $rVG, $rMenthol, $rQuantity, $rVendor)
			;Clear Data from inputs
			;GUICtrlSetData($hCombo, "")

			;Re Populate Comboboxes
			Pop_Flavor($hCombo)

		;Set Focus on First Flavor Combobox
		WinActivate("CustomLabelGUI")
		_WinAPI_SetFocus(ControlGetHandle("CustomLabelGUI", "", $hCombo))

		Case $b50mL
			$rName = GUICtrlRead($hCombo)
			$rSize = "50mL"
			$rNic = GUICtrlRead($NicCombo)
			$rVG = GUICtrlRead($vgCombo)
			;$rPG = (100 - $rVG)
			$rMenthol = GUICtrlRead($MentholCombo)
			$rQuantity = GUICtrlRead($QuantityCombo)
			;Call("Validate_Int", $rQuantity)
			$rVendor = GUICtrlRead($VendorCombo)
			Call("Print_Label", $rName, $rSize, $rNic, $rVG, $rMenthol, $rQuantity, $rVendor)
			;Clear Data from inputs
			;GUICtrlSetData($hCombo, "")

			;Re Populate Comboboxes
			Pop_Flavor($hCombo)

		;Set Focus on First Flavor Combobox
		WinActivate("CustomLabelGUI")
		_WinAPI_SetFocus(ControlGetHandle("CustomLabelGUI", "", $hCombo))

		Case $b100mL
			$rName = GUICtrlRead($hCombo)
			$rSize = "100mL"
			$rNic = GUICtrlRead($NicCombo)
			$rVG = GUICtrlRead($vgCombo)
			;$rPG = (100 - $rVG)
			$rMenthol = GUICtrlRead($MentholCombo)
			$rQuantity = GUICtrlRead($QuantityCombo)
			;Call("Validate_Int", $rQuantity)
			$rVendor = GUICtrlRead($VendorCombo)
			Call("Print_Label", $rName, $rSize, $rNic, $rVG, $rMenthol, $rQuantity, $rVendor)
			;Clear Data from inputs
			;GUICtrlSetData($hCombo, "")

			;Re Populate Comboboxes
			Pop_Flavor($hCombo)

		;Set Focus on First Flavor Combobox
		WinActivate("CustomLabelGUI")
		_WinAPI_SetFocus(ControlGetHandle("CustomLabelGUI", "", $hCombo))
	EndSwitch

	;Must Register Window Status to detect text, etc
 	GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")
WEnd
#EndRegion ### END Koda GUI section ######################################################

;#####################################################################################
;AUTOCOMPLETE NEEDED FUNCTIONS
Func _Edit_Changed($hCombo)
    _GUICtrlComboBox_AutoComplete($hCombo)
EndFunc   ;==>_Edit_Changed

Func WM_COMMAND($hWnd, $iMsg, $iwParam, $ilParam)

    #forceref $hWnd, $iMsg, $ilParam

    $iIDFrom = BitAND($iwParam, 0xFFFF) ; Low Word
    $iCode = BitShift($iwParam, 16) ; Hi Word
    If $iCode = $CBN_EDITCHANGE Then
        Switch $iIDFrom
            Case $hCombo
                _Edit_Changed($hCombo)
        EndSwitch
    EndIf
    Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_COMMAND
;#####################################################################################
;END AUTOCOMPLETE NEEDED FUNCTIONS

Func Pop_Flavor($hCombo)
	;Global $aArray = _ExcelReadSheetToArray($oExcel, 1, 1, 0, 8) ;Starting on the 1st Row
	;_ArrayDisplay($aArray, "Array = ")
    ; Add Flavor Array Data
    _GUICtrlComboBox_BeginUpdate($hCombo)
	For $i = 1 to (UBound($aArray)-1) step 1
		$data1 = $aArray[$i][0]
		;MsgBox(4096, "DATA = ", $data1)
		_GUICtrlComboBox_AddString($hCombo, $data1)
	Next
    _GUICtrlComboBox_EndUpdate($hCombo)
EndFunc

Func Print_Label($rName, $rSize, $rNic, $rVG, $rMenthol, $rQuantity, $rVendor)
	#Region GET DATE IN NICE WINDOWS FORMAT  XX-XX-XXXX
	$timeSys = _Date_Time_GetSystemTime()
	$unformattedTime = _Date_Time_SystemTimeToDateTimeStr($timeSys)
	$slashTime = StringRegExpReplace($unformattedTime, "/", "-")
	$timefull = StringLeft($slashTime, 10)
	$time = StringRegExpReplace($timefull, "20(?=\d\d)", "")
	;MsgBox(4096, "Date", $time)
	#EndRegion

	;Close P-Touch Editor if open to avoid write lock on csv file
	If ProcessExists("Ptedit52.exe") Then ; Check if P-Touch is running
	 ;MsgBox($MB_SYSTEMMODAL, "", "P-Touch is running")
	 ProcessClose("Ptedit52.exe")
	EndIf

    local $sFilePath = @ScriptDir & "\Label Spreadsheet CSV.csv" ;
	Local $hFileOpen = FileOpen($sFilePath, $FO_OVERWRITE)
	$csvData = $rname&","&$rNic&","&$time&","&$rVG&","&$rMenthol&","&$rSize
    FileWriteLine($hFileOpen, "flavor1,nic1,date1,vg1,menthol1,size1,flavor2,nic2,date2,vg2,menthol2,size2,flavor3,nic3,date3,vg3,menthol3,size3,flavor4,nic4,date4,vg4,menthol4,size4,flavor5,nic5,date5,vg5,menthol5,size5,flavor6,nic6,date6,vg6,menthol6,size6,flavor7,nic7,date7,vg7,menthol7,size7,flavor8,nic8,date8,vg8,menthol8,size8,flavor9,nic9,date9,vg9,menthol9,size9") ; CSV name headers
    FileWriteLine($hFileOpen, $csvData)
    FileClose($hFileOpen)
    FileFlush($hFileOpen)
	sleep(1000)

#Region; CHECK VENDOR AND OPEN CORRECT FILE
;MsgBox(4096, "Vendor = ", "Vendor is:  " & $rVendor)
	If $rVendor == "DIGITAL VAPOR DEN" Then
	   ;MsgBox(4096, "Before", "")
	   ShellExecute(@ScriptDir&"\Digital Vapor Den Label.lbx", "", "", "", @SW_HIDE)
	   ;MsgBox(4096, "After", "")
	   sleep(500)
	Else
	   ShellExecute(@ScriptDir & "\1dymo.lbx", "", "", "", @SW_HIDE)
	   sleep(500)
	EndIf
#EndRegion

		If WinExists("P-touch Editor") Then
			Sleep(3000)
			ControlSetText("P-touch Editor", "", "[ID:5148]", $rQuantity)
			SplashTextOn("Title", "PRINTING LABELS NOW", -1, 60, -1, -1, 1, "", 24, 700)
			;ControlClick("P-touch Editor", "", "[ID:5350]", "left", 1); sends command to print
			ControlClick("P-touch Editor", "", "[ID:5351]", "left")
			winwait("Print");wait for print dialog window
			ControlClick("Print", "", "[ID:5202]", "left") ;print all sheets radio box
			ControlClick("Print", "", "[ID:1]", "left") ; click print
			;MsgBox(4096, "Print Commanda sent", "")
			Sleep(2000)
			SplashOff()
		Else
			MsgBox(4096, "WTF!", "Not Catching Window")
		EndIf
		If ProcessExists("Ptedit52.exe") Then ; Check if P-Touch is running process is running.
			;MsgBox($MB_SYSTEMMODAL, "", "P-Touch is running")
			ProcessClose("Ptedit52.exe")
		EndIf
EndFunc

Func Validate_Int($value)
	#forceref $value
	If StringRegExp($value, "\d{1,2}") Then
		;do nothing
		;MsgBox(4096, "GOOD", "NO ERROR WITH: " & $rQuantity)
		Return
	Else
		MsgBox(4096, "USER ERROR!", "NOT A VALID NUMBER: " & $value & " SET TO 1", 3)
		Global $rQuantity = 1; hardcode overwrite variable
		Return
	EndIf
EndFunc
