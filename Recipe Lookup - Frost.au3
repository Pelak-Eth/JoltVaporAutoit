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
#include <Crypt.au3>
#include <WinAPIMisc.au3>
#include <Misc.au3>
#include <SQLite.au3>
#include <SQLite.dll.au3>
#include <GuiComboBox.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <StaticConstants.au3>
#Include <Restart.au3>

;##########################
;Allow only one instance ##
;##########################
If _Singleton("Recipe Lookup", 1) = 0 Then
	MsgBox(4096, "ALREADY RUNNING!", "CLOSING DUPLICATE INSTANCE", 3)
    Exit
EndIf

Global $ingredient1, $ingredient2, $ingredient3, $ingredient4, $ingredient5, $ingredient6, $ingredient7, $Recipe1, $Recipe2, $AddRecipe, $SConcentrate, $response, $NewRecipeGUI
Global $sSearch = ""
Global $concVendor = ""
Global $fCombo = ""
$recipeReturn = 0
Global $mentholVolume = 0
Global $bulkMix
Global $stop = 0

Func _JoltStyle()
	GUICtrlSetColor(-1, 0xC8C8C8)
	GUICtrlSetBkColor(-1, 0x000000)
EndFunc

Call("sql_Lite"); Open DB, Write DB to Array
Func sql_Lite()
   local $sqlDll = @ScriptDir & "\sqlite3.dll"
	_SQLite_Startup($sqlDll, False, 1)
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

Func sqlQueryToArray()
	; Query
Local $aArray1, $iRows, $iColumns, $iRval
Local $aResult = _SQLite_GetTable2d($hDskDb, "SELECT * FROM Recipes order by Recipe asc;", $aArray1, $iRows, $iColumns); Write DB to 2D Array
If $aResult = $SQLITE_OK Then
    ;_SQLite_Display2DResult($aResult)
Else
    MsgBox($MB_SYSTEMMODAL, "SQLite Error: " & $iRval, _SQLite_ErrMsg())
	Exit
EndIf

Local $aArray2, $iRows1, $iColumns1, $iRval1
Local $aResult2 = _SQLite_GetTable2d($hDskDb, "SELECT * FROM Concentrates order by Concentrate asc;", $aArray2, $iRows1, $iColumns1)

Global $aArray = $aArray1
Global $cArray = $aArray2
EndFunc

;Function for running other Autoit scripts from here
Func _RunAU3($sFilePath, $sWorkingDir = @ScriptDir, $iShowFlag = @SW_SHOW, $iOptFlag = 0)
    Return Run('"' & @AutoItExe & '" /AutoIt3ExecuteScript "' & $sFilePath & '"', $sWorkingDir, $iShowFlag, $iOptFlag)
EndFunc

#Region ### START Koda GUI section ### Form=
$RecipeLookup = GUICreate("Recipe Lookup", 300, 142, (@DesktopWidth - 400), (@DesktopHeight - 142), $WS_POPUP, BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW))
GUISetBkColor(0x000000)

$Lookup = GUICtrlCreateButton("Recipe Stuff", 0, 0, ((@DesktopWidth/2) - 650), 160);(@DesktopHeight - 142))
GUICtrlSetFont(-1, 32, 400, 0, "skrunch")
_JoltStyle()

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

#Region ;Main Control Loop Start
While 1
While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $Lookup
			sqlQueryToArray()
			Call("recipeLookup")
			Global $stop = 0; set back to 0, allow lookups again
			ExitLoop

        Case $GUI_EVENT_SECONDARYUP
            $aCInfo = GUIGetCursorInfo($RecipeLookup)
            If $aCInfo[4] = $Lookup Then
                If MsgBox(36, 'Restarting...', 'Press Yes to restart this script.') = 6 Then
					_ScriptRestart()
				EndIf
            EndIf

		Case $GUI_EVENT_CLOSE
			_SQLite_Close()
			_SQLite_Shutdown()
			Exit
	EndSwitch
WEnd
WEnd
#EndRegion ;Main Control Loop End

Func recipeLookup()
If not WinExists("Touch Order") Then
		Call("recipeChoiceGUI", $message)
Else

Global $vgPercentArray[10]
WinWait("Touch Order")
Local $oIE = _IEAttach("Touch Order")
Local $oTable = _IETableGetCollection($oIE, 0)
Local $isData = _IETableWriteToArray($oTable)

EndIf
#Region; IF NOTHING TO PRINT, THEN PROMPT
	If UBound($isData, 2) = 1 Then
			_RunAU3("Bulk Mix GUI.au3")
			Return
	Else
		Local $orderArray = _IETableWriteToArray($oTable)
			If Not IsArray($orderArray) Then
				MsgBox(4096, "ERROR!", "Not Reading Array from IE!")
				Exit
			EndIf
	EndIf
#EndRegion

If @error = 0 Then
;_ArrayDisplay($orderArray, "order array = "); [1][1+]=Recipe Name, [2][1+]=Size mL, [3][1+]= Menthol(None), [4][1+]= "INT" Menthol Volume, [5][1+]=(NIC)"INT" mg Add: 5.12 mL, [6][1+]=(VG / PG)"int" / "int", [7][1+]=(premix) Float mL
For $a = 1 to (UBound($orderArray, 2)-1) step 1
	Global $fBoost = 1
	Global $mentholVolume = $orderArray[4][$a]
	Global $sSearch = $orderArray[1][$a]; [1][x] x= order number, iterate 1 to array length (-1)
	$sizeDataFull = $orderArray[2][$a]; size multiplier
	Global $sizeDataVolume = StringTrimRight($sizeDataFull, 3)
	$nicNumberArray = $orderArray[5][$a]
	$nicNumberArray2 = _StringBetween($nicNumberArray, 'Add: ', 'ml')
	$nicNumber = $nicNumberArray2[0]
	;MsgBox(4096, "Nic = ", $nicNumber)
	Global $vgPercent = $orderArray[6][$a]; Read VG

		If $sizeDataVolume = 5 Then
			Global $sizeData = 1

		ElseIf $sizeDataVolume = 6 Then
			Global $sizeData = 1.2

		ElseIf $sizeDataVolume = 10 Then
			Global $sizeData = 2

		ElseIf $sizeDataVolume = 11 Then
			Global $sizeData = 2.2

		Elseif $sizeDataVolume = 30 Then
			Global $sizeData = 6

		Elseif $sizeDataVolume = 32 Then
			Global $sizeData = 6.4

		Elseif $sizeDataVolume = 50 Then
			Global $sizeData = 10

		Elseif $sizeDataVolume = 100 Then
			Global $sizeData = 20

		Else
			MsgBox(4096, "SIZE ERROR", "WHAT THE FUCK SIZE IS THIS?")

		EndIf
		If $sSearch = "" Then
			ExitLoop
		Else
			If StringRight($sSearch, 4) == "XTRA" Then
				;MsgBox(4096, "XTRA Flavor Full String = ", $sSearch)
				$tSearch = StringTrimRight($sSearch, 5)
				;MsgBox(4096, "XTRA FLAVOR", "Flavor = " & $tSearch)
				Global $sSearch = $tSearch
				Global $fBoost = 1.7
			EndIf

			If $stop <> 0 Then
				Return
			EndIf

			Call(getRecipe, $sSearch, $sizeData, $vgPercent, $sizeDataVolume, $fBoost, $nicNumber, $mentholVolume)

		EndIf
	Next
	;done
	Return
EndIf
EndFunc

Func getRecipe($sSearch, $sizeData, $vgPercent, $sizeDataVolume, $fBoost, $nicNumber, $mentholVolume)
;Declare arrays so its not angry
local $arrayOfIngredients[0]
local $arrayofVendors[0]
local $arrayOfVolumes[0]

;## Set values to blank if not used
Local $Ingredient[9] = ["", "", "", "", "", "", "", "", ""]
Local $NicInfo = ""
Local $nicVolume = 0
Local $mentholInfo = ""

Local $iIndex = _ArraySearch($aArray, $sSearch, 0, 0, 0, 0); get Recipe index from recipe array
If @error <> 0 Then

		$message = "NOT FOUND.  DO YOU WANT TO ADD RECIPE?"
		Call("pMsgBox", $message); Yes --> $response = 1 ,, No --> $response = 0 ,, $message = string text

		If $response = 1 Then; yes
			IniWrite(@ScriptDir & "\RecipesINI.ini", "RECIPEINFO", "recipename", $sSearch)
			_RunAU3("New Recipe GUI.au3")
			_SQLite_Close()
			Call("sql_Lite"); refresh db
			Return

		ElseIf $response = 0 Then; no
			Return

		ElseIf $response = 2 Then; cancel
			Global $stop = 1
			Return

		EndIf
	Return

Else
	$message = string("SHOW RECIPE FOR: " & $sSearch & "?")
	Call("pMsgBox", $message); Yes --> $response = 1 ,, No --> $response = 0 ,, $message = string text
	If $response = 1 Then; yes


	If $mentholVolume <> 0 Then
		$mentholInfo = string("Menthol to Add: " & $mentholVolume & "mL")
	Else
		$mentholInfo = ""
	EndIf
	;Calculate and set Nicotine Volume
	$nicVolume = $nicNumber ;no calc, just match POS value
	;Round(($nicNumber * $sizeDataVolume) / 100, 2)
	$NicInfo = "Nicotine to Add: " & $nicVolume & "mL" & "     " & $mentholInfo

	;Premixed Volume
	$preMixVolume = ($sizeDataVolume - $nicVolume - $mentholVolume)

#Region; Calculate Ingredient Volumes
	For $i = 1 To 9 Step 1
		$aIngredient = $aArray[$iIndex][$i]
			If $aIngredient = "" Then
				ExitLoop
			EndIf
		$sArray = StringSplit($aIngredient, "|")
		_ArrayAdd($arrayOfIngredients, $sArray[1])
		$finalVolume = Round(($sArray[3] * $sizeData * $fBoost), 1)
		_ArrayAdd($arrayOfVolumes, $finalVolume)
		_ArrayAdd($arrayofVendors, $sArray[2])
#EndRegion

#Region; Add Ingredients to Array and Nicely Format them
	Next
		$tVolume = 0
		For $f = 0 to (UBound($arrayOfVolumes)-1)
			$tVolume = Round($arrayOfVolumes[$f] + $tVolume, 1)
			_ArrayInsert($Ingredient, $f, $Ingredient[$f] & "INGREDIENT: " & $arrayofVendors[$f] & "  " & $arrayOfIngredients[$f] & "  Add:  " & $arrayOfVolumes[$f] & "mL")
				If $f < (UBound($arrayOfVolumes)-1) Then
					$Ingredient[$f] = $Ingredient[$f] & ""&@CRLF&""
				EndIf
#EndRegion
		Next
#Region; SPECIAL INSTRUCTIONS, MIX TWO RECIPES OR ADD SOMETHING TO RECIPE
			If ($aArray[$iIndex][9]) == "" Then
				$specialInstructions = ""

			Else
				$instructArray = StringSplit(($aArray[$iIndex][9]), "|")
				If $instructArray[1] == "Add" Then; [1]=Flag(Add or Percent), [2]=Add Premixed, [3]=Concentrate, [4]=Vendor, [5]=volume
					$instructVolume = Round(($instructArray[5] * $sizeData * $fBoost), 1)
					$recipeVolume = Round($sizeDataVolume - $mentholVolume - $nicVolume - $instructVolume, 1)
					$specialInstructions = "OR: " & $recipeVolume & "mL of: " & $instructArray[2] & " " & $instructArray[4] & " " & $instructArray[3] & " Add: " & $instructVolume & " mL"

				ElseIf $instructArray[1] == "Percent" Then; [1]=Flag(Add or Percent), [2]=Premixed Recipe1, [3]=Percent1, [4]=Premixed Recipe2, [5]=Percent2
					$percent1Vol = ($instructArray[3] * $sizeDataVolume) - (0.5 * ($nicVolume + $mentholVolume))
					$percent2Vol = ($instructArray[5] * $sizeDataVolume) - (0.5 * ($nicVolume + $mentholVolume))
					$specialInstructions = "OR: " & "Add: " & $percent1Vol & " mL of " & $instructArray[2] & "  Add: " & $percent2Vol & " mL of " & $instructArray[4]

				Else
					$specialInstructions = ""
				EndIf
			EndIf
#EndRegion

#Region; VG AND PG STUFF
			If $vgPercent == "MAX" Then
				$totalVG = Round(($sizeDataVolume - $tVolume - $mentholVolume - $nicVolume), 1)
				$vgInfo = "FILL REMAINDER WITH VG, WHICH IS: " & $totalVG & "mL OF VG"

			ElseIf $vgPercent <> 40 Then
				;MsgBox(4096, "VG NOT 40", "VG= " & $vgPercent)
				$vgVolume = Round((($vgPercent / 100) * $sizeDataVolume) - ($nicVolume/2), 1)
				$pgVolume = Round((((100 - $vgPercent)/100) * $sizeDataVolume - ($nicVolume/2) - $tVolume), 1); $pgVolume = ($sizeDataVolume - $vgVolume - ($nicVolume/2) - $tVolume)
				$vgInfo = "VG TO ADD: " & $vgVolume & "mL      " & "PG TO ADD: " & $pgVolume & "mL"

			Else
				;MsgBox(4096, "VG IS 40", "IS 40")
				$vgVolume = Round((($vgPercent/100)*$sizeDataVolume) - ($nicVolume/2), 1)
				$pgVolume = Round((((100 - $vgPercent)/100) * $sizeDataVolume - ($nicVolume/2) - $tVolume - $mentholVolume), 1)
				;MsgBox(4096, "Values = ", "Total= " & $sizeDataVolume & "VG= " & $vgVolume & "Nic=" & ($nicVolume/2) & "ConcTotal = " & $tVolume)
				$vgInfo = "VG TO ADD: " & $vgVolume & "mL      " & "PG TO ADD: " & $pgVolume & "mL"

			EndIf
#EndRegion

			;Check if bulk mixed
			If $aArray[$iIndex][10] == "yes" Then
				$bulkMix = "YES"
			ElseIf $aArray[$iIndex][10] == "no" Then
				$bulkMix = "NO"
			Else
				$bulkMix = "ERROR"
			EndIf

			;SET FLAVOR BOOST
			If $fBoost <> 1 Then
				$sSearch = $sSearch & " XTRA"
			;Check if Glass bulk mixed
			If $aArray[$iIndex][11] == "yes" Then
				$bulkMix = "GLASS"
			ElseIf $aArray[$iIndex][11] == "no" Then
				$bulkMix = "NO"
			Else
				$bulkMix = "ERROR"
			EndIf
			EndIf

;MsgBox(4096, "PG", "PG = " & $pgVolume)
			#Region ### START Koda GUI section ### Form=C:\Users\windows\Documents\Autoit\Recipe Lookup GUI.kxf
			$ShowRecipe = GUICreate("ShowRecipe", 1028, 600, (@DesktopWidth/2)-514, -8,  $WS_POPUP, BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW))
			GUISetBkColor(0x000000)

			$CloseRecipe = GUICtrlCreateButton("Close Recipe View", 0, 488, 1025, 145)
			GUICtrlSetFont(-1, 48, 400, 2, "skrunch")
			_JoltStyle()

			$RecipeName = GUICtrlCreateLabel("BULK=" & $bulkMix & "   " & $sSearch & " " & $sizeDataVolume & "mL", 50, 8, (@DesktopWidth-100), 41)
			GUICtrlSetFont(-1, 36, 400, 2, "skrunch")
			_JoltStyle()

			$IngredientLabel0 = GUICtrlCreateLabel($Ingredient[0], 95, 52, 900, 41)
			GUICtrlSetFont(-1, 22, 800, 0, "MS Sans Serif")
			_JoltStyle()

			$IngredientLabel1 = GUICtrlCreateLabel($Ingredient[1], 95, 83, 900, 41)
			GUICtrlSetFont(-1, 22, 800, 0, "MS Sans Serif")
			_JoltStyle()

			$IngredientLabel2 = GUICtrlCreateLabel($Ingredient[2], 95, 114, 900, 41)
			GUICtrlSetFont(-1, 22, 800, 0, "MS Sans Serif")
			_JoltStyle()

			$IngredientLabel3 = GUICtrlCreateLabel($Ingredient[3], 95, 145, 900, 41)
			GUICtrlSetFont(-1, 22, 800, 0, "MS Sans Serif")
			_JoltStyle()

			$IngredientLabel4 = GUICtrlCreateLabel($Ingredient[4], 95, 175, 900, 41)
			GUICtrlSetFont(-1, 22, 800, 0, "MS Sans Serif")
			_JoltStyle()

			$IngredientLabel5 = GUICtrlCreateLabel($Ingredient[5], 95, 206, 900, 41)
			GUICtrlSetFont(-1, 22, 800, 0, "MS Sans Serif")
			_JoltStyle()

			$IngredientLabel6 = GUICtrlCreateLabel($Ingredient[6], 95, 237, 900, 41)
			GUICtrlSetFont(-1, 22, 800, 0, "MS Sans Serif")
			_JoltStyle()

			$IngredientLabel7 = GUICtrlCreateLabel($Ingredient[7], 88, 268, 900, 33)
			GUICtrlSetFont(-1, 22, 800, 0, "MS Sans Serif")
			_JoltStyle()

			$NicLabel = GUICtrlCreateLabel($NicInfo, 90, 294, 900, 33)
			GUICtrlSetFont(-1, 28, 800, 0, "MS Sans Serif")
			_JoltStyle()

			$Special = GUICtrlCreateLabel($specialInstructions, 10, 339, @DesktopWidth, 41)
			GUICtrlSetFont(-1, 22, 800, 0, "MS Sans Serif")
			_JoltStyle()

			$VGLabel = GUICtrlCreateLabel($vgInfo, 95, 386, 900, 41)
			GUICtrlSetFont(-1, 28, 800, 0, "MS Sans Serif")
			_JoltStyle()

			$TotalVolume = GUICtrlCreateLabel("Total Concentrate = " & $tVolume & " mL" & "   OR Add Premixed " & $preMixVolume & "mL", 95, 438, 900, 41)
			GUICtrlSetFont(-1, 28, 800, 0, "MS Sans Serif")
			_JoltStyle()

			GUISetState(@SW_SHOW)
			#EndRegion ### END Koda GUI section ###

#Region; Recipe Lookup Control Loop
			While 1
				$nMsg = GUIGetMsg()
				Switch $nMsg
					Case $GUI_EVENT_CLOSE
						GUIDelete($ShowRecipe)
						Return
					Case $CloseRecipe
						GUIDelete($ShowRecipe)
						Return
				EndSwitch
			WEnd
#EndRegion

	ElseIf $response = 0 Then; no
		Return

	Elseif $response = 2 Then;cancel
		Global $stop = 1
		Return
	EndIf
EndIf
Return
EndFunc

;#####################################################################################
;AUTOCOMPLETE NEEDED FUNCTIONS
Func _Edit_Changed($ingredient)
    _GUICtrlComboBox_AutoComplete($ingredient)
EndFunc   ;==>_Edit_Changed

Func WM_COMMAND($hWnd, $iMsg, $iwParam, $ilParam)

    #forceref $hWnd, $iMsg, $ilParam

    $iIDFrom = BitAND($iwParam, 0xFFFF) ; Low Word
    $iCode = BitShift($iwParam, 16) ; Hi Word
    If $iCode = $CBN_EDITCHANGE Then
        Switch $iIDFrom
            Case $ingredient1
                _Edit_Changed($ingredient1)
            Case $ingredient2
                _Edit_Changed($ingredient2)
            Case $ingredient3
                _Edit_Changed($ingredient3)
            Case $ingredient4
                _Edit_Changed($ingredient4)
            Case $ingredient5
                _Edit_Changed($ingredient5)
            Case $ingredient6
                _Edit_Changed($ingredient6)
            Case $ingredient7
                _Edit_Changed($ingredient7)
			Case $Recipe1
				_Edit_Changed($Recipe1)
			Case $Recipe2
				_Edit_Changed($Recipe2)
			Case $AddRecipe
				_Edit_Changed($AddRecipe)
			Case $SConcentrate
				_Edit_Changed($SConcentrate)
        EndSwitch
    EndIf
    Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_COMMAND
;#####################################################################################
;END AUTOCOMPLETE NEEDED FUNCTIONS

Func Pop_Ingredient($Ingredient)
    ; Add Flavor Array Data
    _GUICtrlComboBox_BeginUpdate($ingredient)
	For $i = 2 to (UBound($cArray) - 1) step 1
		$data1 = $cArray[$i][0]
		_GUICtrlComboBox_AddString($ingredient, $data1)
	Next
    _GUICtrlComboBox_EndUpdate($ingredient)
EndFunc

Func get_Vendor($rSConcentrate)
	$concIndex = _ArraySearch($cArray, $rSConcentrate)
		If @error <> 0 Then
			$concVendor = "GUESS"
		Else
			$concVendor = $cArray[$concIndex][1]
			;MsgBox(4096, "Vendor = ", $concVendor)
		EndIf
	Return $concVendor
EndFunc

Func Pop_Recipe($Recipe)
    ; Add Flavors from Array to Combobox
    _GUICtrlComboBox_BeginUpdate($Recipe)
	For $i = 3 to (UBound($aArray)-1) step 1
		$data1 = $aArray[$i][0]
		_GUICtrlComboBox_AddString($Recipe, $data1)
	Next
    _GUICtrlComboBox_EndUpdate($Recipe)
EndFunc

Func Validate_Conc($cFlavor)
	$concIndex = _ArraySearch($cArray, $cFlavor)
		If @error <> 0 Then
			;MsgBox(4096, "NOT FOUND", $cFlavor & " is not defined!")
			_RunAU3("Add New Concentrate.au3", @ScriptDir)
		Else
			$concIndex = _ArraySearch($cArray, $cFlavor)
				If @error <> 0 Then
					MsgBox(4096, "REFRESHING", "REFRESHING TO CHECK FOR NEW CONCENTRATES", 2)
					_SQLite_Close()
					Call("sql_Lite"); refresh db
					;MsgBox(4096, "ERROR!", $cFlavor & " Not Found in Concentrates")
					Return

				Else
					$concVendor = $cArray[$concIndex][1]
					;MsgBox(4096, "Vendor = ", $concVendor)
					Return $concVendor

				EndIf
			EndIf
	Return
EndFunc


Func pMsgBox($message)
#Region ### START Koda GUI section ### Form=C:\Users\windows\Documents\Autoit\PMsgBox.kxf
$PMSGBOX = GUICreate("PMsgBox", 597, 325, (@DesktopWidth/2)-299, 217, $WS_POPUP, BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW))
GUISetBkColor(0x000000)

$LabelMessage = GUICtrlCreateLabel($message, 17, 8, 571, 28)
GUICtrlSetFont(-1, 16, 400, 0, "MS Sans Serif")
_JoltStyle()

$Yes = GUICtrlCreateButton("YES", 1, 48, 257, 161)
GUICtrlSetFont(-1, 72, 400, 2, "skrunch")
_JoltStyle()

$No = GUICtrlCreateButton("NO", 328, 48, 265, 161)
GUICtrlSetFont(-1, 72, 400, 2, "skrunch")
_JoltStyle()

$Cancel = GUICtrlCreateButton("CANCEL LOOKUPS", 8, 224, 577, 89)
GUICtrlSetFont(-1, 36, 800, 0, "skrunch")
_JoltStyle()

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

		Case $Cancel
			$response = 2
			GUIDelete()
			Return $response

	EndSwitch
WEnd
#EndRegion ### END Koda GUI section #############################################
EndFunc

Func recipeChoiceGUI($message)
#Region ### START Koda GUI section ### Form=C:\Users\windows\Documents\Autoit\Recipe Choice GUI.kxf
$recipeChoiceGUI = GUICreate("Recipe Choice GUI", 1030, 327, (@DesktopWidth/2)-515, 220, $WS_POPUP, BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW))
GUISetBkColor(0x000000)

$labelMessage = GUICtrlCreateLabel($message, 17, 8, 955, 28)
GUICtrlSetFont(-1, 16, 400, 0, "MS Sans Serif")
_JoltStyle()

$bulkMix = GUICtrlCreateButton("Bulk Mix", 44, 48, 257, 161)
GUICtrlSetFont(-1, 36, 400, 2, "skrunch")
_JoltStyle()

$newRecipe = GUICtrlCreateButton("NEW RECIPE", 371, 48, 265, 161)
GUICtrlSetFont(-1, 28, 400, 2, "skrunch")
_JoltStyle()

$newConcentrate = GUICtrlCreateButton("New Concentrate", 697, 48, 265, 161)
GUICtrlSetFont(-1, 20, 400, 2, "skrunch")
_JoltStyle()

$Exit = GUICtrlCreateButton("CLOSE", 10, 247, 993, 73)
GUICtrlSetFont(-1, 36, 400, 2, "skrunch")
_JoltStyle()

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			$response = -1
			GUIDelete()
			Return $response

		Case $Exit
			$response = -1
			GUIDelete()
			Return $response

		Case $bulkMix
			$response = 1
			GUIDelete()
			Return $response

		Case $newRecipe
			$response = 2
			GUIDelete()
			Return $response

		Case $newConcentrate
			$response = 3
			GUIDelete()
			Return $response

	EndSwitch
WEnd
EndFunc