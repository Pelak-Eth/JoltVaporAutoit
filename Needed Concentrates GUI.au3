#include <Excel.au3>
#include <File.au3>
#include <String.au3>
#include <Array.au3>
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
#include <Date.au3>

Call("sql_Lite"); Open DB, Write DB to Array
Func sql_Lite()
_SQLite_Startup()
If @error Then
    MsgBox($MB_SYSTEMMODAL, "SQLite Error", "SQLite3.dll Can't be Loaded!")
    Exit -1
EndIf
ConsoleWrite("_SQLite_LibVersion=" & _SQLite_LibVersion() & @CRLF)
Local $sDbName = @ScriptDir & "\JoltRecipesDB"
Global $hDskDb = _SQLite_Open($sDbName) ; Open a permanent disk database
If @error Then
    MsgBox($MB_SYSTEMMODAL, "SQLite Error", "Can't open or create a permanent Database!")
    Exit -1
EndIf

; Query
Local $aArray1, $iRows, $iColumns, $iRval
Local $aResult = _SQLite_GetTable2d(-1, "SELECT * FROM Recipes order by Recipe asc;", $aArray1, $iRows, $iColumns); Write DB to 2D Array
If $aResult = $SQLITE_OK Then
    ;_SQLite_Display2DResult($aResult)
Else
    MsgBox($MB_SYSTEMMODAL, "SQLite Error: " & $iRval, _SQLite_ErrMsg())
	Exit
EndIf

Local $aArray2, $iRows1, $iColumns1, $iRval1
Local $aResult2 = _SQLite_GetTable2d(-1, "SELECT * FROM Concentrates order by Concentrate asc;", $aArray2, $iRows1, $iColumns1)

Global $aArray = $aArray1
Global $cArray = $aArray2
;_ArrayDisplay($cArray, "Array = "); COL[0] = Concentrate, COL[1] = Vendor, COL[2] = Volume, COL[3] = Min, COL[4] = Max,  ALL DATA STARTS AT ROW[1]
EndFunc

Func _JoltStyle()
	GUICtrlSetColor(-1, 0xC8C8C8)
	GUICtrlSetBkColor(-1, 0x000000)
EndFunc

#Region ###
$Form1 = GUICreate("CONCENTRATES NEEDED GUI", 615, 566, 1301, 240, $WS_SYSMENU, BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW))
GUISetBkColor(0x000000)

$Flavor1 = GUICtrlCreateLabel("SELECT RECIPE", 232, 48, 147, 24)
GUICtrlSetFont(-1, 14, 400, 2, "skrunch")
_JoltStyle()
$hCombo = GUICtrlCreateCombo("", 120, 80, 345, 25)
_JoltStyle()

$mBulkLabel = GUICtrlCreateLabel("How Many mL of Recipe", 200, 120, 448, 24)
GUICtrlSetFont(-1, 14, 400, 2, "skrunch")
_JoltStyle()

$bMultiplier = GUICtrlCreateInput("250", 232, 152, 129, 46)
GUICtrlSetFont(-1, 28, 400, 2, "skrunch")
_JoltStyle()

$addConc = GUICtrlCreateButton("ADD CONCENTRATES", 8, 240, 193, 129)
GUICtrlSetFont(-1, 14, 400, 2, "skrunch")
_JoltStyle()

$export = GUICtrlCreateButton("EXPORT TO EXCEL", 212, 240, 193, 129)
GUICtrlSetFont(-1, 16, 400, 2, "skrunch")
_JoltStyle()

$calcAll = GUICtrlCreateButton("CALCULATE ALL", 416, 240, 193, 129)
GUICtrlSetFont(-1, 14, 400, 2, "skrunch")
_JoltStyle()

$Label4 = GUICtrlCreateLabel("", 99, 147, 4, 4)
GUICtrlSetFont(-1, 14, 400, 2, "skrunch")
_JoltStyle()

$ViewArray = GUICtrlCreateButton("DISPLAY RESULTS", 102, 384, 193, 129)
GUICtrlSetFont(-1, 14, 400, 2, "skrunch")
_JoltStyle()

$OrderTPA = GUICtrlCreateButton("ORDER FROM TPA", 319, 384, 193, 129)
GUICtrlSetFont(-1, 14, 400, 2, "skrunch")
_JoltStyle()

    ; Add Flavor Array Data
	Pop_Flavor($hCombo)

;Set Focus on First Flavor Combobox
_WinAPI_SetFocus(ControlGetHandle("CONCENTRATES NEEDED GUI", "", $hCombo))

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
While 1
	$nMsg = GUIGetMsg()
		;Reregister Window Statuses to detect text, etc
		GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			_SQLite_Close()
			_SQLite_Shutdown()
			Exit

		Case $addConc
 			$rName = GUICtrlRead($hCombo)
			$bMult = (GUICtrlRead($bMultiplier) / 5)

			Call("getRecipe", $rName, $bMult, $cArray)

			;Clear Data from inputs
			GUICtrlSetData($hCombo, "")

			;Re Populate Comboboxes
			Pop_Flavor($hCombo)

		;Set Focus on First Flavor Combobox
		_WinAPI_SetFocus(ControlGetHandle("CONCENTRATES NEEDED GUI", "", $hCombo))

		Case $export
			#Region GET DATE IN NICE WINDOWS FORMAT  XX-XX-XXXX
			$timeSys = _Date_Time_GetSystemTime()
			$unformattedTime = _Date_Time_SystemTimeToDateTimeStr($timeSys)
			$slashTime = StringRegExpReplace($unformattedTime, "/", "-")
			$time = StringLeft($slashTime, 10)
			#EndRegion

			For $a = 1 to ((UBound($cArray)) - 1)
				$mlVolume = $cArray[$a][2]
				$ozVolume = Ceiling($mlVolume / 29.5735)
				$cArray[$a][3] = $ozVolume
			Next
				_ArrayDisplay($cArray, "TOTAL CONCENTRATES NEEDED")
			Local $oExcel = _ExcelBookNew() ;Create new book, make it visible

			_ExcelWriteSheetFromArray($oExcel, $cArray, 1, 1, 0, 0) ; Write Array to Excel

			MsgBox($MB_SYSTEMMODAL, "Exiting", "Press OK to Save File and Exit"); Popup Confirmation of Saving
			_ExcelBookSaveAs($oExcel, @DesktopDir & "\Concentrates Needed For "&$time&".xls", "xls", 0, 1) ; Save to Desktop with Date Stamp
			_ExcelBookClose($oExcel) ;Close Excel
		Case $calcAll
			Local $sArray[8] = ["", "", "", "", "", "", "", ""]; redeclare as blank
			Local $totalVolumeArray[9] = [0,0,0,0,0,0,0,0,0]; redeclare as blank
			;Iterate through Array of Recipes
			For $i = 3 to ((UBound($aArray)) - 1)
				For $a = 1 to 8 step 1
					$fullString = $aArray[$i][$a]
					If $fullString == "" Then
						ExitLoop
					EndIf
					$sArray = StringSplit($fullString, "|")
					$iName = $sArray[1]
					$bMult = GUICtrlRead($bMultiplier)
						If UBound($sArray) <> 4 Then
							MsgBox(4096, "ERROR!", "ERROR WITH: " & $fullString & "Array Line: " & $i)
							;_ArrayDisplay($aArray, "aArray = ")
							ExitLoop
						Else
							$finalVolume = ($sArray[3] * $bMult)
							$totalVolumeArray[$a] = $finalVolume
							Call("addConcentrate", $iName, $finalVolume, $cArray)
						EndIf
				Next
;~ 				;total up concentrates
;~ 				Local $totalVolume = 0; redeclare as 0
;~ 				Local $ingQuantity = 0; redeclare as 0
;~ 				Local $newbMult = 0; redeclare as 0
;~ 					For $val = 1 to 8 step 1
;~ 						$totalVolume = $totalVolumeArray[$val] + $totalVolume
;~ 							If not $totalVolume == 0 Then
;~ 								$ingQuantity = $ingQuantity + 1
;~ 							EndIf
;~ 					Next
;~ 						If $totalVolume < 420 Then
;~ 							MsgBox(4096, "TOTAL VOLUME", "TOTAL VOL = " & $totalVolume &"mL", 2)
;~ 							$newbMult = Ceiling(((420-$totalVolume)/$ingQuantity))
;~ 							MsgBox(4096, "NEW MULT", "NEW MULT = " & $newbMult)
;~ 							For $a = 1 to 8 step 1
;~ 								$fullString = $aArray[$i][$a]
;~ 								If $fullString == "" Then
;~ 									ExitLoop
;~ 								EndIf
;~ 								$sArray = StringSplit($fullString, "|")
;~ 								$iName = $sArray[1]
;~ 								$finalVolume = ($sArray[3] * $newbMult)
;~ 								Call("addConcentrate", $iName, $finalVolume, $cArray)
;~ 							Next
;~ 								MsgBox(4096, "NEW TOTAL", "TOTAL VOL = " & $totalVolume)
;~ 						EndIf
			Next
				;Convert mL to Ounces
				For $v = 1 to ((UBound($cArray)) - 1)
					$mlVolume = $cArray[$v][2]
					$ozVolume = Ceiling($mlVolume / 29.5735)
					$cArray[$v][3] = $ozVolume
				Next
					_ArrayDisplay($cArray, "TOTAL CONCENTRATES NEEDED")

		Case $ViewArray
				_ArrayDisplay($cArray, "cArray = ")

		Case $OrderTPA
			MsgBox(4096, "TODO", "ORDER FROM TPA")

	EndSwitch

	;Must Register Window Status to detect text, etc
 	GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")

WEnd
WEnd

#Region;   AUTOCOMPLETE NEEDED FUNCTIONS
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
#EndRegion;   END AUTOCOMPLETE NEEDED FUNCTIONS

Func Pop_Flavor($hCombo)
    ; Add Flavors from Array to Combobox
    _GUICtrlComboBox_BeginUpdate($hCombo)
	For $i = 3 to (UBound($aArray)-1) step 1
		$data1 = $aArray[$i][0]
		_GUICtrlComboBox_AddString($hCombo, $data1)
	Next
    _GUICtrlComboBox_EndUpdate($hCombo)
EndFunc

Func getRecipe($rName, $bMult, ByRef $cArray)
$sizeDataVolume = ($bMult * 5)
$fVG = ($sizeDataVolume * 0.4)
$iPG = ($sizeDataVolume * 0.6)
$sPG = 0
local $arrayOfIngredients[0]
local $arrayofVendors[0]
local $arrayOfVolumes[0]
;## Set ingredient label text to blank if not used
Local $Ingredient[8] = ["", "", "", "", "", "", ""]
Local $iIndex = _ArraySearch($aArray, $rName)
If @error Then
    MsgBox($MB_SYSTEMMODAL, "Not Found", '"' & $rName & '" was not found in the array.', 3)
	Return
Else
	For $i = 1 To 8 Step 1
		$aIngredient = $aArray[$iIndex][$i]
			If $aIngredient = "" Then
				Return
			EndIf
		$sArray = StringSplit($aIngredient, "|")
		_ArrayAdd($arrayOfIngredients, $sArray[1])
		If UBound($sArray) <> 4 Then
			MsgBox(4096, "ERROR!", "ERROR WITH: " & $aIngredient & "UB= " & UBound($sArray))
		Else
			$finalVolume = ($sArray[3] * $bMult)
		EndIf
		_ArrayAdd($arrayOfVolumes, $finalVolume)
		_ArrayAdd($arrayofVendors, $sArray[2])
		$sPG = $finalVolume + $sPG
		$iName = $sArray[1]
		Call("addConcentrate", $iName, $finalVolume, $cArray)
	Next
		_ArrayDisplay($cArray, "CURRENT CONCENTRATES NEEDED = ")
EndIf
Return
EndFunc

Func addConcentrate($iName, $finalVolume, ByRef $cArray)
	$cIndex = _ArraySearch($cArray, $iName);Get Array Index of Concentrate
		If @error <> 0 Then
			MsgBox(4096, "NOT FOUND", "Conc: " & $iName)
			Return
		EndIf
			;Add Volumes Needed back to Array
			$currentVolume = $cArray[$cIndex][2]
			$newVolume = $finalVolume + $currentVolume
			$cArray[$cIndex][2] = $newVolume
	Return $cArray
EndFunc
