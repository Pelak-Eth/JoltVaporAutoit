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
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <StaticConstants.au3>
#include <SQLite.au3>
#include <SQLite.dll.au3>
#include <Misc.au3>

;##########################
;Allow only one instance ##
;##########################
If _Singleton("New Recipe GUI", 1) = 0 Then
	MsgBox(4096, "ALREADY RUNNING!", "CLOSING DUPLICATE INSTANCE", 3)
    Exit
EndIf

Global $sSearch = IniRead(@ScriptDir & "\RecipesINI.ini", "RECIPEINFO", "recipename", "")
Global $concVendor = ""
Global $fCombo = ""

Func _JoltStyle()
	GUICtrlSetColor(-1, 0xC8C8C8)
	GUICtrlSetBkColor(-1, 0x000000)
EndFunc

;Function for running other Autoit scripts from here
Func _RunAU3($sFilePath, $sWorkingDir = @ScriptDir, $iShowFlag = @SW_SHOW, $iOptFlag = 0)
    Return Run('"' & @AutoItExe & '" /AutoIt3ExecuteScript "' & $sFilePath & '"', $sWorkingDir, $iShowFlag, $iOptFlag)
EndFunc

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

Func sqlQueryToArray()
	; Query
Local $aArray1, $iRows, $iColumns, $iRval
Local $aResult = _SQLite_GetTable2d($hDskDb, "SELECT * FROM Recipes order by Recipe asc;", $aArray1, $iRows, $iColumns); Write DB to 2D Array
;~ If $aResult = $SQLITE_OK Then
;~     ;_SQLite_Display2DResult($aResult)
;~ Else
;~     MsgBox($MB_SYSTEMMODAL, "SQLite Error: " & $iRval, _SQLite_ErrMsg())
;~ 	Exit
;~ EndIf

Local $aArray2, $iRows1, $iColumns1, $iRval1
Local $aResult2 = _SQLite_GetTable2d($hDskDb, "SELECT * FROM Concentrates order by Concentrate asc;", $aArray2, $iRows1, $iColumns1)

Global $aArray = $aArray1
Global $cArray = $aArray2
EndFunc

#Region ### START Koda GUI section ### Form=C:\Users\Jolt\Documents\Autoit\newrecipebottlesize.kxf
$PickSize = GUICreate("PickSize", 914, 322, ((@DesktopWidth/2)-457), 360, $WS_POPUP, BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW))
GUISetBkColor(0x000000)

$c10ml = GUICtrlCreateButton("10 mL", 16, 99, 201, 105)
GUICtrlSetFont(-1, 36, 800, 0, "skrunch")
_JoltStyle()

$c30ml = GUICtrlCreateButton("30 mL", 244, 99, 201, 105)
GUICtrlSetFont(-1, 36, 800, 0, "skrunch")
_JoltStyle()

$c50ml = GUICtrlCreateButton("50 mL", 473, 99, 201, 105)
GUICtrlSetFont(-1, 36, 800, 0, "skrunch")
_JoltStyle()

$c100ml = GUICtrlCreateButton("100 mL", 701, 99, 201, 105)
GUICtrlSetFont(-1, 36, 800, 0, "skrunch")
_JoltStyle()

$Label1 = GUICtrlCreateLabel("WHICH SIZE BOTTLE DID YOU MAKE IT WITH?", 24, 24, 864, 46)
GUICtrlSetFont(-1, 30, 800, 0, "skrunch")
_JoltStyle()

$close = GUICtrlCreateButton("CANCEL / CLOSE", 24, 232, 865, 73)
GUICtrlSetFont(-1, 32, 800, 0, "skrunch")
_JoltStyle()

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###


While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit

		Case $c10ml
			Global $sizeMult = 2
			GUIDelete()
			ExitLoop
		Case $c30ml
			Global $sizeMult = 6
			GUIDelete()
			ExitLoop
		Case $c50ml
			Global $sizeMult = 10
			GUIDelete()
			ExitLoop
		Case $c100ml
			Global $sizeMult = 20
			GUIDelete()
			ExitLoop
		Case $close
			_SQLite_Close()
			_SQLite_Shutdown()
			;Call("sql_Lite"); Open DB, Write DB to Array
			Exit
	EndSwitch
WEnd

#Region ### START Koda GUI section ### Form=c:\users\windows\documents\autoit\new recipe gui.kxf
			$NewRecipeGUI = GUICreate("ADD NEW RECIPE GUI", 1266, 473, ((@DesktopWidth/2)-633), ((@DesktopHeight/2)-236))
			GUISetBkColor(0x000000)

			$Title = GUICtrlCreateLabel("NEW RECIPE - TYPE NUMBERS CAREFULLY", 350, 0, 700, 37)
			GUICtrlSetFont(-1, 24, 400, 2, "skrunch")
			_JoltStyle()

			$RecipeName = GUICtrlCreateInput($sSearch, 368, 48, 529, 32)
			GUICtrlSetFont(-1, 14, 400, 0, "MS Sans Serif")
			_JoltStyle()

			$Ingredient1 = GUICtrlCreateCombo("", 24, 128, 153, 25)
			$Volume1 = GUICtrlCreateInput("", 22, 216, 153, 21)
			_JoltStyle()

			$Ingredient2 = GUICtrlCreateCombo("", 203, 128, 153, 25)
			$Volume2 = GUICtrlCreateInput("", 201, 216, 153, 21)
			_JoltStyle()

			$Ingredient3 = GUICtrlCreateCombo("", 381, 128, 153, 25)
			$Volume3 = GUICtrlCreateInput("", 379, 216, 153, 21)
			_JoltStyle()

			$Ingredient4 = GUICtrlCreateCombo("", 560, 128, 153, 25)
			$Volume4 = GUICtrlCreateInput("", 558, 216, 153, 21)
			_JoltStyle()

			$Ingredient5 = GUICtrlCreateCombo("", 739, 128, 153, 25)
			$Volume5 = GUICtrlCreateInput("", 737, 216, 153, 21)
			_JoltStyle()

			$Ingredient6 = GUICtrlCreateCombo("", 917, 128, 153, 25)
			$Volume6 = GUICtrlCreateInput("", 915, 216, 153, 21)
			_JoltStyle()

			$Ingredient7 = GUICtrlCreateCombo("", 1096, 128, 153, 25)
			$Volume7 = GUICtrlCreateInput("", 1094, 216, 153, 21)
			_JoltStyle()

			$LabelName = GUICtrlCreateLabel("ENTER RECIPE NAME =", 112, 52, 245, 26)
			GUICtrlSetFont(-1, 16, 400, 2, "skrunch")
			_JoltStyle()

			$VolumeLabel1 = GUICtrlCreateLabel("Enter mL Volume", 21, 192, 143, 20)
			GUICtrlSetFont(-1, 12, 400, 2, "skrunch")
			_JoltStyle()

			$Label1 = GUICtrlCreateLabel("Enter mL Volume", 201, 192, 143, 20)
			GUICtrlSetFont(-1, 12, 400, 2, "skrunch")
			_JoltStyle()

			$Label2 = GUICtrlCreateLabel("Enter mL Volume", 381, 192, 143, 20)
			GUICtrlSetFont(-1, 12, 400, 2, "skrunch")
			_JoltStyle()

			$Label3 = GUICtrlCreateLabel("Enter mL Volume", 561, 192, 143, 20)
			GUICtrlSetFont(-1, 12, 400, 2, "skrunch")
			_JoltStyle()

			$Label4 = GUICtrlCreateLabel("Enter mL Volume", 741, 192, 143, 20)
			GUICtrlSetFont(-1, 12, 400, 2, "skrunch")
			_JoltStyle()

			$Label5 = GUICtrlCreateLabel("Enter mL Volume", 921, 192, 143, 20)
			GUICtrlSetFont(-1, 12, 400, 2, "skrunch")
			_JoltStyle()

			$Label6 = GUICtrlCreateLabel("Enter mL Volume", 1101, 192, 143, 20)
			GUICtrlSetFont(-1, 12, 400, 2, "skrunch")
			_JoltStyle()

			$Label7 = GUICtrlCreateLabel("Enter Flavor", 35, 104, 114, 20)
			GUICtrlSetFont(-1, 12, 400, 2, "skrunch")
			_JoltStyle()

			$Label8 = GUICtrlCreateLabel("Enter Flavor", 215, 104, 114, 20)
			GUICtrlSetFont(-1, 12, 400, 2, "skrunch")
			_JoltStyle()

			$Label9 = GUICtrlCreateLabel("Enter Flavor", 395, 104, 114, 20)
			GUICtrlSetFont(-1, 12, 400, 2, "skrunch")
			_JoltStyle()

			$Label10 = GUICtrlCreateLabel("Enter Flavor", 575, 104, 114, 20)
			GUICtrlSetFont(-1, 12, 400, 2, "skrunch")
			_JoltStyle()

			$Label11 = GUICtrlCreateLabel("Enter Flavor", 755, 104, 114, 20)
			GUICtrlSetFont(-1, 12, 400, 2, "skrunch")
			_JoltStyle()

			$Label12 = GUICtrlCreateLabel("Enter Flavor", 935, 104, 114, 20)
			GUICtrlSetFont(-1, 12, 400, 2, "skrunch")
			_JoltStyle()

			$Label13 = GUICtrlCreateLabel("Enter Flavor", 1115, 104, 114, 20)
			GUICtrlSetFont(-1, 12, 400, 2, "skrunch")
			_JoltStyle()

			$Label14 = GUICtrlCreateLabel("COMBINED TWO FLAVORS?", 152, 272, 237, 24)
			GUICtrlSetFont(-1, 14, 400, 2, "skrunch")
			_JoltStyle()

			$Label15 = GUICtrlCreateLabel("ADDED 1 THING TO RECIPE?", 806, 272, 248, 24)
			GUICtrlSetFont(-1, 14, 400, 2, "skrunch")
			_JoltStyle()

			$Percent1 = GUICtrlCreateCombo("", 79, 310, 89, 25, $CBS_DROPDOWNLIST)
			GUICtrlSetData(-1, "70|60|50|40|30|20")
			_JoltStyle()

			$Percent2 = GUICtrlCreateCombo("", 80, 381, 89, 25, $CBS_DROPDOWNLIST)
			GUICtrlSetData(-1, "70|60|50|40|30|20")
			_JoltStyle()

;~ 			$Percent3 = GUICtrlCreateCombo("", 80, 381, 89, 25, $CBS_DROPDOWNLIST)
;~ 			GUICtrlSetData(-1, "70|60|50|40|30|20")
;~ 			_JoltStyle()

			$Recipe1 = GUICtrlCreateCombo("", 307, 310, 153, 25)
			_JoltStyle()

			$Recipe2 = GUICtrlCreateCombo("", 307, 381, 153, 25)
			_JoltStyle()

			$Label16 = GUICtrlCreateLabel("Percent of", 180, 310, 113, 24)
			GUICtrlSetFont(-1, 14, 400, 2, "skrunch")
			_JoltStyle()

			$Label17 = GUICtrlCreateLabel("Percent of", 182, 381, 113, 24)
			GUICtrlSetFont(-1, 14, 400, 2, "skrunch")
			_JoltStyle()

			$AddRecipe = GUICtrlCreateCombo("", 1073, 335, 153, 25)
			_JoltStyle()

			$SConcentrate = GUICtrlCreateCombo("", 883, 335, 153, 25)
			_JoltStyle()

			$SVolume = GUICtrlCreateCombo("", 659, 335, 153, 21, $CBS_DROPDOWNLIST)
			GUICtrlSetData(-1, "0.02|0.05|0.1|0.15|0.2|0.25|0.3|0.35|0.4|0.5|0.6|0.7|0.8|0.9|1.0|1.1|1.2|1.3|1.4|1.5|1.6")
			_JoltStyle()

			$Label18 = GUICtrlCreateLabel("TO", 1043, 335, 28, 24)
			GUICtrlSetFont(-1, 14, 400, 2, "skrunch")
			_JoltStyle()

			$Label19 = GUICtrlCreateLabel("ADD", 619, 335, 33, 24)
			GUICtrlSetFont(-1, 14, 400, 2, "skrunch")
			_JoltStyle()

			$Label20 = GUICtrlCreateLabel("mL of", 819, 335, 57, 24)
			GUICtrlSetFont(-1, 14, 400, 2, "skrunch")
			_JoltStyle()

			$Label21 = GUICtrlCreateLabel("VOLUME", 691, 311, 75, 24)
			GUICtrlSetFont(-1, 14, 400, 2, "skrunch")
			_JoltStyle()

			$Label22 = GUICtrlCreateLabel("CONCENTRATE", 891, 311, 131, 24)
			GUICtrlSetFont(-1, 14, 400, 2, "skrunch")
			_JoltStyle()

			$Label23 = GUICtrlCreateLabel("RECIPE", 1115, 311, 72, 24)
			GUICtrlSetFont(-1, 14, 400, 2, "skrunch")
			_JoltStyle()

			$Label24 = GUICtrlCreateLabel("PERCENT", 83, 343, 85, 24)
			GUICtrlSetFont(-1, 14, 400, 2, "skrunch")
			_JoltStyle()

			$Label25 = GUICtrlCreateLabel("RECIPE", 347, 343, 72, 24)
			GUICtrlSetFont(-1, 14, 400, 2, "skrunch")
			_JoltStyle()

			$Submit = GUICtrlCreateButton("SUBMIT NEW RECIPE", 312, 416, 657, 49)
			GUICtrlSetFont(-1, 16, 400, 2, "skrunch")
			_JoltStyle()

			;Set Focus on First Flavor Combobox
			_WinAPI_SetFocus(ControlGetHandle("NEW RECIPE GUI", "", $RecipeName))

			;Populate Comboboxes
			Pop_Ingredient($ingredient1)
			Pop_Ingredient($ingredient2)
			Pop_Ingredient($ingredient3)
			Pop_Ingredient($ingredient4)
			Pop_Ingredient($ingredient5)
			Pop_Ingredient($ingredient6)
			Pop_Ingredient($ingredient7)
			Pop_Ingredient($SConcentrate)
			Pop_Recipe($Recipe1)
			Pop_Recipe($Recipe2)
			Pop_Recipe($AddRecipe)

			GUISetState(@SW_SHOW)

			While 1
				$nMsg = GUIGetMsg()
					;Reregister Window Statuses to detect text, etc
					GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")
				Switch $nMsg
					Case $GUI_EVENT_CLOSE
						ExitLoop

					Case $Submit
							$RName = GUICtrlRead($RecipeName)

						;Check if recipe exists
						Local $iIndex = _ArraySearch($aArray, $RName, 0, 0, 0, 0)
							If @error <> 0 Then
								;exit if, continue
							Else
								MsgBox(4096, "Already Exists", $RName & " has already been added.  Press OK to exit")
								_SQLite_Close()
								_SQLite_Shutdown()
								Exit
							EndIf

							$cFlavor1 = GUICtrlRead($ingredient1)
								If $cFlavor1 == "" Then
									;MsgBox(4096, "BLANK", "FLAVOR IS: " & $cFlavor1)
									$RFlavor1 = ""
								Else
									$cVolume1 = (GUICtrlRead($Volume1) / $sizeMult)
									Validate_Conc($cFlavor1)
									$RFlavor1 = $cFlavor1&"|"&$concVendor&"|"&$cVolume1
								EndIf

							$cFlavor2 = GUICtrlRead($ingredient2)
								If $cFlavor2 == "" Then
									$RFlavor2 = ""
								Else
									$cVolume2 = (GUICtrlRead($Volume2) / $sizeMult)
									Validate_Conc($cFlavor2)
									;get_Vendor($cFlavor2)
									$RFlavor2 = $cFlavor2&"|"&$concVendor&"|"&$cVolume2
								EndIf

							$cFlavor3 = GUICtrlRead($ingredient3)
								If $cFlavor3 == "" Then
									$RFlavor3 = ""
								Else
									$cVolume3 = (GUICtrlRead($Volume3) / $sizeMult)
									Validate_Conc($cFlavor3)
									;get_Vendor($cFlavor3)
									$RFlavor3 = $cFlavor3&"|"&$concVendor&"|"&$cVolume3
								EndIf

							$cFlavor4 = GUICtrlRead($ingredient4)
								If $cFlavor4 == "" Then
									$RFlavor4 = ""
								Else
									$cVolume4 = (GUICtrlRead($Volume4) / $sizeMult)
									Validate_Conc($cFlavor4)
									;get_Vendor($cFlavor4)
									$RFlavor4 = $cFlavor4&"|"&$concVendor&"|"&$cVolume4
								EndIf

							$cFlavor5 = GUICtrlRead($ingredient5)
								If $cFlavor5 == "" Then
									$RFlavor5 = ""
								Else
									$cVolume5 = (GUICtrlRead($Volume5) / $sizeMult)
									Validate_Conc($cFlavor5)
									;get_Vendor($cFlavor5)
									$RFlavor5 = $cFlavor5&"|"&$concVendor&"|"&$cVolume5
								EndIf

							$cFlavor6 = GUICtrlRead($ingredient6)
								If $cFlavor6 == "" Then
									$RFlavor6 = ""
								Else
									$cVolume6 = (GUICtrlRead($Volume6) / $sizeMult)
									Validate_Conc($cFlavor6)
									;get_Vendor($cFlavor6)
									$RFlavor6 = $cFlavor6&"|"&$concVendor&"|"&$cVolume6
								EndIf

							$cFlavor7 = GUICtrlRead($ingredient7)
								If $cFlavor7 == "" Then
									$RFlavor7 = ""
								Else
									$cVolume7 = (GUICtrlRead($Volume7) / $sizeMult)
									Validate_Conc($cFlavor7)
									;get_Vendor($cFlavor7)
									$RFlavor7 = $cFlavor7&"|"&$concVendor&"|"&$cVolume7
								EndIf

							$fCombo = ""
							;Combination Recipe Info ex: Percent|Melon Balla|0.6|Blue Racer|0.4  or   Add|Lemonade Plus|Apple|TPA|1.0
							$rPercent1 = GUICtrlRead($Percent1)
							If $rPercent1 <> "" Then
								$rRecipe1 = GUICtrlRead($Recipe1)
								$rRecipe2 = GUICtrlRead($Recipe2)
								$rPercent1 = GUICtrlRead($Percent1) / 100
								$rPercent2 = GUICtrlRead($Percent2) / 100
								$fCombo = "Percent|"&$rRecipe1&"|"&$rPercent1&"|"&$rRecipe2&"|"&$rPercent2
							Else
								;MsgBox(4096, "BLANK", "Recipe = " & $rRecipe1)
							EndIf

							$rAddRecipe = GUICtrlRead($AddRecipe)
							If $rAddRecipe <> "" And $rPercent1 == "" Then
								$rAddRecipe = GUICtrlRead($AddRecipe)
								$rSConcentrate = GUICtrlRead($SConcentrate)
								get_Vendor($rSConcentrate)
								$rVendor = $concVendor
								;MsgBox(4096, "VENDOR", "VENDOR = " & $rVendor)
								$rSVolume = (GUICtrlRead($SVolume) / $sizeMult)
								$fCombo = "Add|" & $rAddRecipe & " Plus " & "|" & $rSConcentrate&"|"&$rVendor&"|"&$rSVolume
							Else
							EndIf

						;INSERT DATA INTO DATABASE
						;Recipe	Ingredient1	Ingredient2	Ingredient3	Ingredient4	Ingredient5	Ingredient6	Ingredient7	Ingredient8	Combination
						$InsertSQL = "INSERT INTO Recipes (Recipe,Ingredient1,Ingredient2,Ingredient3,Ingredient4,Ingredient5,Ingredient6,Ingredient7,Combination) VALUES (" & _SQLite_FastEscape($RName) &","&_SQLite_FastEscape($RFlavor1) &","& _SQLite_FastEscape($RFlavor2) &","& _SQLite_FastEscape($RFlavor3) &","& _SQLite_FastEscape($RFlavor4) &","& _SQLite_FastEscape($RFlavor5) &","& _SQLite_FastEscape($RFlavor6) &","& _SQLite_FastEscape($RFlavor7) &","& _SQLite_FastEscape($fCombo)&")"
							If Not _SQLite_Exec(-1, $InsertSQL) = $SQLITE_OK Then
								MsgBox(16, "SQLite Error", _SQLite_ErrMsg())
							Else
								_SQLite_Exec(-1, $InsertSQL)
							EndIf

						; Delete GUI
						IniWrite(@ScriptDir & "\RecipesINI.ini", "RECIPEINFO", "recipename", "")
						GUIDelete($NewRecipeGUI)
						_SQLite_Close()
						_SQLite_Shutdown()
						;Call("sql_Lite"); Open DB, Write DB to Array
						Exit

					EndSwitch

				;Must Register Window Status to detect text, etc
				GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")

			WEnd
#EndRegion ### END Koda GUI section #########################################

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
	For $i = 1 to (UBound($cArray) - 1) step 1
		$data1 = $cArray[$i][0]
		_GUICtrlComboBox_AddString($ingredient, $data1)
	Next
    _GUICtrlComboBox_EndUpdate($ingredient)
EndFunc

Func get_Vendor($rSConcentrate)
	$concIndex = _ArraySearch($cArray, $rSConcentrate)
		If @error <> 0 Then
			$concVendor = "FIXME"
		Else
			$concVendor = $cArray[$concIndex][1]
			;MsgBox(4096, "Vendor = ", $concVendor)
		EndIf
	Return $concVendor
EndFunc

Func Pop_Recipe($Recipe)
    ; Add Flavors from Array to Combobox
    _GUICtrlComboBox_BeginUpdate($Recipe)
	For $i = 1 to (UBound($aArray)-1) step 1
		$data1 = $aArray[$i][0]
		_GUICtrlComboBox_AddString($Recipe, $data1)
	Next
    _GUICtrlComboBox_EndUpdate($Recipe)
EndFunc

Func Validate_Conc($cFlavor)
	$concIndex = _ArraySearch($cArray, $cFlavor)
		If @error <> 0 Then
			IniWrite(@ScriptDir & "\RecipesINI.ini", "RECIPEINFO", "concentrate", $cFlavor)
			$concVendor = "FIXME"
			_RunAU3("Add New Concentrate.au3", @ScriptDir)
			Return $concVendor
		Else
			$concIndex = _ArraySearch($cArray, $cFlavor)
				If @error <> 0 Then
					MsgBox(4096, "ERROR!", $cFlavor & " Not Found in Concentrates")
					Return
				Else
					$concVendor = $cArray[$concIndex][1]
					;MsgBox(4096, "Vendor = ", $concVendor)
					Return $concVendor
				EndIf
			EndIf
EndFunc