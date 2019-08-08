#include <Excel.au3>
#include <File.au3>
#include <String.au3>
#include <Array.au3>
#include <Date.au3>
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
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <StaticConstants.au3>
;#####################################################################################################
;Global Variables																					##
;Global $sFilePath1 = @ScriptDir & "\Vape Recipes Autoit.xlsx";										##
;Global $oExcel = _ExcelBookOpen($sFilePath1)													;   ##
;#####################################################################################################

;##########################
;Allow only one instance ##
;##########################
If _Singleton("Bulk Mix GUI", 1) = 0 Then
	MsgBox(4096, "ALREADY RUNNING!", "CLOSING DUPLICATE INSTANCE", 3)
    Exit
EndIf

Func _JoltStyle()
	GUICtrlSetColor(-1, 0xC8C8C8)
	GUICtrlSetBkColor(-1, 0x000000)
EndFunc

Global $bulkMix
;$rName = "" ;default blank value

Call("sql_Lite")
Func sql_Lite()
   local $sqlDll = @ScriptDir & "\sqlite3.dll"
	_SQLite_Startup($sqlDll, False, 1)
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

#Region MIX GUI
$MixGUI = GUICreate("MixGUI", 1028, 380, 918, 550, $WS_BORDER, BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW));$WS_POPUP, BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW))
GUISetBkColor(0x000000)

$RecipeLabel = GUICtrlCreateLabel("START TYPING / SELECT", 24, 256, 380, 37)
GUICtrlSetFont(-1, 24, 400, 2, "skrunch")
_JoltStyle()

$hCombo = GUICtrlCreateCombo("", 24, 294, 377, 45)
GUICtrlSetFont(-1, 24, 800, 0, "MS Sans Serif")
_JoltStyle()

$VGLabel = GUICtrlCreateLabel("VG %", 27, 190, 66, 34)
GUICtrlSetFont(-1, 22, 400, 2, "skrunch")
_JoltStyle()

$VGCombo = GUICtrlCreateCombo("", 97, 185, 73, 45, $CBS_DROPDOWNLIST)
GUICtrlSetData($VGCombo, "100|50|60|70|80|90|40", "40")
GUICtrlSetFont(-1, 22, 800, 0, "MS Sans Serif")
_JoltStyle()

$MentholLabel = GUICtrlCreateLabel("MENTHOL", 294, 86, 95, 24)
GUICtrlSetFont(-1, 16, 400, 2, "skrunch")
_JoltStyle()

$Menthol = GUICtrlCreateCombo("", 292, 113, 105, 45, $CBS_DROPDOWNLIST)
GUICtrlSetData($Menthol, "LIGHT|MEDIUM|HEAVY|SUPER|NONE", "NONE")
GUICtrlSetFont(-1, 20, 800, 0, "MS Sans Serif")
_JoltStyle()

$PremixBottle = GUICtrlCreateButton("PREMIX BOTTLE", 440, 272, 288, 69)
GUICtrlSetFont(-1, 24, 400, 2, "skrunch")
_JoltStyle()

$SaltNicLabel = GUICtrlCreateLabel("SALT:", 220, 190, 80, 34)
GUICtrlSetFont(-1, 22, 400, 2, "skrunch")
_JoltStyle()

$saltNic = GUICtrlCreateCombo("", 311, 185, 80, 45, $CBS_DROPDOWNLIST)
GUICtrlSetData(-1, "YES|NO", "NO")
GUICtrlSetFont(-1, 22, 800, 0, "skrunch")
_JoltStyle()

$FlavorShotLabel = GUICtrlCreateLabel("Flavor shot?", 13, 86, 114, 20)
GUICtrlSetFont(-1, 12, 400, 2, "skrunch")
_JoltStyle()

$FlavorShot = GUICtrlCreateCombo("", 27, 113, 80, 45, $CBS_DROPDOWNLIST)
GUICtrlSetData(-1, "YES|NO", "NO")
GUICtrlSetFont(-1, 22, 800, 0, "skrunch")
_JoltStyle()

$NicLabel = GUICtrlCreateLabel("Nicotine?", 144, 86, 121, 24)
GUICtrlSetFont(-1, 18, 400, 2, "skrunch")
_JoltStyle()

$NicCombo = GUICtrlCreateCombo("", 144, 113, 121, 45, $CBS_DROPDOWN)
GUICtrlSetData(-1, "0|1|2|3|4|5|6|7|8|9|10|11|12|13|14|15|16|17|18|19|20|21|22|23|24|25|26|27|28|29|30|40|50", "0")
GUICtrlSetFont(-1, 24, 800, 0, "MS Sans Serif")
_JoltStyle()

$b10mL = GUICtrlCreateButton("10mL BOTTLE", 744, 84, 265, 53)
GUICtrlSetFont(-1, 24, 400, 2, "skrunch")
_JoltStyle()

$b30mL = GUICtrlCreateButton("30mL BOTTLE", 744, 154, 265, 53)
GUICtrlSetFont(-1, 24, 400, 2, "skrunch")
_JoltStyle()

$b50mL = GUICtrlCreateButton("50mL BOTTLE", 744, 224, 265, 53)
GUICtrlSetFont(-1, 24, 400, 2, "skrunch")
_JoltStyle()

$b100mL = GUICtrlCreateButton("100mL BOTTLE", 744, 294, 265, 53)
GUICtrlSetFont(-1, 24, 400, 2, "skrunch")
_JoltStyle()

$MultiplierLabel = GUICtrlCreateLabel("Mixed Vol mL", 440, 94, 195, 24)
GUICtrlSetFont(-1, 14, 400, 2, "skrunch")
_JoltStyle()

$Multiplier = GUICtrlCreateInput("220", 440, 128, 137, 45)
GUICtrlSetFont(-1, 24, 400, 0, "MS Sans Serif")
_JoltStyle()

$CustomMult = GUICtrlCreateButton("VOLUME", 438, 184, 137, 69)
GUICtrlSetFont(-1, 24, 400, 2, "skrunch")
_JoltStyle()

$CustomVol = GUICtrlCreateLabel("Conc Vol mL", 583, 94, 153, 24)
GUICtrlSetFont(-1, 14, 400, 2, "skrunch")
_JoltStyle()

$VolMultiplier = GUICtrlCreateInput("420", 590, 128, 137, 45)
GUICtrlSetFont(-1, 24, 400, 0, "MS Sans Serif")
_JoltStyle()

$ConcMult = GUICtrlCreateButton("CONC", 590, 184, 137, 69)
GUICtrlSetFont(-1, 24, 400, 2, "skrunch")
_JoltStyle()

$Exit = GUICtrlCreateButton("CLOSE WINDOW", 0, 0, 1017, 73)
GUICtrlSetFont(-1, 24, 400, 2, "skrunch")
_JoltStyle()

; Add Flavor Array Data
Pop_Flavor($hCombo)

;Set Focus on Flavor Combobox
WinActivate("MixGUI")
_WinAPI_SetFocus(ControlGetHandle("MixGUI", "", $hCombo))

GUISetState(@SW_SHOW)

#EndRegion MIX GUI

#Region MIX GUI CONTROLS
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

		Case $Exit
			_SQLite_Close()
			_SQLite_Shutdown()
			Exit

		Case $PremixBottle
			$rName = GUICtrlRead($hCombo)
			If $rName == "" Then
				MsgBox(4096, "DUMBASS ALERT", "YOU DID NOT SELECT A RECIPE", 5)
				ExitLoop
			EndIf
			$bMult = 50
			$rNic = GUICtrlRead($NicCombo)
			$rVolume = 250
			$rBoost = GUICtrlRead($FlavorShot)
			$rVG = GUICtrlRead($VGCombo)
			$rMenthol = GUICtrlRead($Menthol)
			$rSalt = GUICtrlRead($saltNic)

			Call("getRecipe", $rName, $bMult, $rNic, $rVolume, $rBoost, $rVG, $rMenthol, $rSalt)

			;Clear Data from inputs
			GUICtrlSetData($hCombo, "")

			;Re Populate Comboboxes
			Pop_Flavor($hCombo)

			;Set Focus on Flavor Combobox
			WinActivate("MixGUI")
			_WinAPI_SetFocus(ControlGetHandle("MixGUI", "", $hCombo))

		Case $b10mL
			$rName = GUICtrlRead($hCombo)
			If $rName == "" Then
				MsgBox(4096, "DUMBASS ALERT", "YOU DID NOT SELECT A RECIPE", 5)
				ExitLoop
			EndIf
			$bMult = 2
			$rNic = GUICtrlRead($NicCombo)
			$rVolume = 10
			$rBoost = GUICtrlRead($FlavorShot)
			$rVG = GUICtrlRead($VGCombo)
			$rMenthol = GUICtrlRead($Menthol)
			$rSalt = GUICtrlRead($saltNic)

			Call("getRecipe", $rName, $bMult, $rNic, $rVolume, $rBoost, $rVG, $rMenthol, $rSalt)

			;Clear Data from inputs
			GUICtrlSetData($hCombo, "")

			;Re Populate Comboboxes
			Pop_Flavor($hCombo)

			;Set Focus on Flavor Combobox
			WinActivate("MixGUI")
			_WinAPI_SetFocus(ControlGetHandle("MixGUI", "", $hCombo))

		Case $b30mL
			$rName = GUICtrlRead($hCombo)
			If $rName == "" Then
				MsgBox(4096, "DUMBASS ALERT", "YOU DID NOT SELECT A RECIPE", 5)
				ExitLoop
			EndIf
			$bMult = 6
			$rNic = GUICtrlRead($NicCombo)
			$rVolume = 30
			$rBoost = GUICtrlRead($FlavorShot)
			$rVG = GUICtrlRead($VGCombo)
			$rMenthol = GUICtrlRead($Menthol)
			$rSalt = GUICtrlRead($saltNic)

			Call("getRecipe", $rName, $bMult, $rNic, $rVolume, $rBoost, $rVG, $rMenthol, $rSalt)

			;Clear Data from inputs
			;GUICtrlSetData($hCombo, "")
			GUICtrlSetData($Menthol, "NONE")
			GUICtrlSetData($VGCombo, "40")
			GUICtrlSetData($FlavorShot, "NO")
			GUICtrlSetData($saltNic, "NO")
			GUICtrlSetData($NicCombo, "")

			;Re Populate Comboboxes
			;Pop_Flavor($hCombo)

			;Set Focus on Flavor Combobox
			WinActivate("MixGUI")
			_WinAPI_SetFocus(ControlGetHandle("MixGUI", "", $hCombo))

		Case $b50mL
			$rName = GUICtrlRead($hCombo)
			If $rName == "" Then
				MsgBox(4096, "DUMBASS ALERT", "YOU DID NOT SELECT A RECIPE", 5)
				ExitLoop
			EndIf
			$bMult = 10
			$rNic = GUICtrlRead($NicCombo)
			$rVolume = 50
			$rBoost = GUICtrlRead($FlavorShot)
			$rVG = GUICtrlRead($VGCombo)
			$rMenthol = GUICtrlRead($Menthol)
			$rSalt = GUICtrlRead($saltNic)

			Call("getRecipe", $rName, $bMult, $rNic, $rVolume, $rBoost, $rVG, $rMenthol, $rSalt)

			;Clear Data from inputs
			;GUICtrlSetData($hCombo, "")
			GUICtrlSetData($Menthol, "NONE")
			GUICtrlSetData($VGCombo, "40")
			GUICtrlSetData($FlavorShot, "NO")
			GUICtrlSetData($saltNic, "NO")
			GUICtrlSetData($NicCombo, "")

			;Re Populate Comboboxes
			;Pop_Flavor($hCombo)

			;Set Focus on Flavor Combobox
			WinActivate("MixGUI")
			_WinAPI_SetFocus(ControlGetHandle("MixGUI", "", $hCombo))

		Case $b100mL
			$rName = GUICtrlRead($hCombo)
			If $rName == "" Then
				MsgBox(4096, "DUMBASS ALERT", "YOU DID NOT SELECT A RECIPE", 5)
				ExitLoop
			EndIf
			$bMult = 20
			$rNic = GUICtrlRead($NicCombo)
			$rVolume = 100
			$rBoost = GUICtrlRead($FlavorShot)
			$rVG = GUICtrlRead($VGCombo)
			$rMenthol = GUICtrlRead($Menthol)
			$rSalt = GUICtrlRead($saltNic)

			Call("getRecipe", $rName, $bMult, $rNic, $rVolume, $rBoost, $rVG, $rMenthol, $rSalt)

			;Clear Data from inputs
			;GUICtrlSetData($hCombo, "")
			GUICtrlSetData($Menthol, "NONE")
			GUICtrlSetData($VGCombo, "40")
			GUICtrlSetData($FlavorShot, "NO")
			GUICtrlSetData($saltNic, "NO")
			GUICtrlSetData($NicCombo, "")

			;Re Populate Comboboxes
			;Pop_Flavor($hCombo)

			;Set Focus on Flavor Combobox
			WinActivate("MixGUI")
			_WinAPI_SetFocus(ControlGetHandle("MixGUI", "", $hCombo))

		Case $CustomMult
			$rName = GUICtrlRead($hCombo)
			If $rName == "" Then
				MsgBox(4096, "DUMBASS ALERT", "YOU DID NOT SELECT A RECIPE", 5)
				ExitLoop
			EndIf
			$prebMult = GUICtrlRead($Multiplier)
			$bMult = ($prebMult / 5)
			$rNic = GUICtrlRead($NicCombo)
			$rVolume = $bMult * 5
			$rBoost = GUICtrlRead($FlavorShot)
			$rVG = GUICtrlRead($VGCombo)
			$rMenthol = GUICtrlRead($Menthol)
			$rSalt = GUICtrlRead($saltNic)

			Call("getRecipe", $rName, $bMult, $rNic, $rVolume, $rBoost, $rVG, $rMenthol, $rSalt)

			;Clear Data from inputs
			;GUICtrlSetData($hCombo, "")

			;Re Populate Comboboxes
			Pop_Flavor($hCombo)

			;Set Focus on Flavor Combobox
			WinActivate("MixGUI")
			_WinAPI_SetFocus(ControlGetHandle("MixGUI", "", $hCombo))

		Case $ConcMult
			$rVolume = GUICtrlRead($VolMultiplier)
			$rName = GUICtrlRead($hCombo)
			If $rName == "" Then
				MsgBox(4096, "DUMBASS ALERT", "YOU DID NOT SELECT A RECIPE", 5)
				ExitLoop
			EndIf
			$rNic = GUICtrlRead($NicCombo)
			$rBoost = GUICtrlRead($FlavorShot)
			$rVG = GUICtrlRead($VGCombo)
			$rMenthol = GUICtrlRead($Menthol)
			$rSalt = GUICtrlRead($saltNic)

			Call("specificVolume", $rVolume, $rName, $rNic, $rBoost, $rVG, $rMenthol, $rSalt)

			;Clear Data from inputs
			;GUICtrlSetData($hCombo, "")

			;Re Populate Comboboxes
			Pop_Flavor($hCombo)

			;Set Focus on Flavor Combobox
			WinActivate("MixGUI")
			_WinAPI_SetFocus(ControlGetHandle("MixGUI", "", $hCombo))

	EndSwitch

	;Must Register Window Status to detect text, etc
 	GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")
WEnd
WEnd
#EndRegion MIX GUI CONTROLS


Func specificVolume($rVolume, $rName, $rNic, $rBoost, $rVG, $rMenthol, $rSalt)
;define arrays to hold data
local $arrayOfIngredients[0]
local $arrayOfIngredients1[0]
local $arrayofVendors[0]
local $arrayOfVolumes[0]
local $arrayOfVolumes1[0]

;## Set ingredient label text to blank if not used
Local $Ingredient[10] = ["", "", "", "", "", "", "", "", "", ""]

	;Find Recipe Index
	Local $iIndex = _ArraySearch($aArray, $rName, 0, 0, 0, 0)

	;Make sure recipe is found
	If @error Then
    MsgBox($MB_SYSTEMMODAL, "Not Found", '"' & $rName & '" was not found in the array.', 3)
	Return ;error so return to GUI
	Else
	EndIf

	;Flavor Boost Check
	If $rBoost == "YES" Then
		$fBoost = 1.7
		$rName = ($rName & " XTRA")

		;Check if Glass bulk mixed
		If $aArray[$iIndex][11] == "yes" Then
			$bulkMix = "GLASS"
		ElseIf $aArray[$iIndex][11] == "no" Then
			$bulkMix = "NO"
		Else
			$bulkMix = "ERROR"
		EndIf
	Else
		$fBoost = 1
	EndIf

	;name before adding stuff
	$rJustName = $rName

	$rName = $rVolume & " mL - " & $rName & " CONCENTRATE"

	;Calculate bmult when total concentrate = $rVolume
	For $i = 1 To 8 Step 1
	$aIngredient = $aArray[$iIndex][$i]
		If $aIngredient = "" Then
			ExitLoop
		EndIf
	$sArray = StringSplit($aIngredient, "|")
	_ArrayAdd($arrayOfIngredients1, $sArray[1])
	$multVolume = ($sArray[3] * $fBoost)
	_ArrayAdd($arrayOfVolumes1, $multVolume)
	Next
	;MsgBox(4096, "LOOP 1", "Clear")
		$totalVolume1 = 0
		For $j = 0 to (UBound($arrayOfVolumes1)-1)
			$totalVolume1 = $arrayOfVolumes1[$j] + $totalVolume1
		Next
			;MsgBox(4096, "Total", "Total Vol = " & $totalVolume1)
			$bMult = Round(($rVolume / $totalVolume1), 3)
			;MsgBox(4096, "MULTIPLIER", "bMult = " & $bMult)

			;bMult is defined, so calculate these now
			$rVolume = $bMult * 5; total volume size

			;Calculate and set Nicotine Volume
			$rNicVol = Round(($rNic * $rVolume) / 100, 2)

		For $i = 1 To 8 Step 1
	$aIngredient = $aArray[$iIndex][$i]
		If $aIngredient = "" Then
			ExitLoop
		EndIf
	$sArray = StringSplit($aIngredient, "|")
	_ArrayAdd($arrayOfIngredients, $sArray[1])
	$finalVolume = Round(($sArray[3] * $bMult * $fBoost), 1)
	_ArrayAdd($arrayOfVolumes, $finalVolume)
	_ArrayAdd($arrayofVendors, $sArray[2])
	;$sPG = $finalVolume + $sPG
		Next
			;MsgBox(4096, "LOOP 3", "Clear")
		$tVolume = 0
;~ 		$fPG = ($iPG - $sPG)
			For $f = 0 to (UBound($arrayOfVolumes)-1)
				$tVolume = $arrayOfVolumes[$f] + $tVolume
				_ArrayInsert($Ingredient, $f, $Ingredient[$f] & "INGREDIENT: " & $arrayofVendors[$f] & "  " & $arrayOfIngredients[$f] & "  Add:  " & $arrayOfVolumes[$f] & "mL")
					If $f < (UBound($arrayOfVolumes)-1) Then
						$Ingredient[$f] = $Ingredient[$f] & ""&@CRLF&""
					EndIf
			Next
				;MsgBox(4096, "LOOP 4", "Clear")
				;_ArrayDisplay($Ingredient, "Ingredient Array")

				If ($aArray[$iIndex][9]) == "" Then
					$specialInstructions = ""
				Else
					$instructArray = StringSplit(($aArray[$iIndex][9]), "|")
					If $instructArray[1] == "Add" Then; [1]=Flag(Add or Percent), [2]=Add Premixed, [3]=Concentrate, [4]=Vendor, [5]=volume
						$instructVolume = Round(($instructArray[5] * $bMult * $fBoost), 1)
						$recipeVolume = Round($rVolume - $rNicVol - $instructVolume, 1)
						$specialInstructions = "OR: " & $recipeVolume & "mL of: " & $instructArray[2] & " " & $instructArray[4] & " " & $instructArray[3] & " Add: " & $instructVolume & " mL"
					ElseIf $instructArray[1] == "Percent" Then; [1]=Flag(Add or Percent), [2]=Premixed Recipe1, [3]=Percent1, [4]=Premixed Recipe2, [5]=Percent2
						$percent1Vol = ($instructArray[3] * $rVolume) - (0.5 * ($rNicVol)); + $mentholVolume))
						$percent2Vol = ($instructArray[5] * $rVolume) - (0.5 * ($rNicVol)); + $mentholVolume))
						$specialInstructions = "OR: " & "Add: " & $percent1Vol & " mL of " & $instructArray[2] & "  Add: " & $percent2Vol & " mL of " & $instructArray[4]
					Else
						$specialInstructions = ""
					EndIf
				 EndIf

			   ;MENTHOL CALCULATIONS
			   If $rMenthol = "NONE" Then
				  $mVolume = 0
			   ElseIf $rMenthol = "LIGHT" Then
				  $mVolume = 0.2 * $bMult
			   ElseIf $rMenthol = "MEDIUM" Then
				  $mVolume = 0.4 * $bMult
			   ElseIf $rMenthol = "HEAVY" Then
				  $mVolume = 0.8 * $bMult
			   ElseIf $rMenthol = "SUPER" Then
				  $mVolume = 1.6 * $bMult
			   EndIf

			   ;Premixed Volume
				$preMixVolume = ($rVolume - $rNicVol - $mVolume)

				If $rVG == 100 Then
					$vgVolume = Round((($rVG / 100) * $rVolume) - $rNicVol - $tVolume - $mVolume, 1)
					$vgInfo = "FILL REMAINDER WITH VG." & "  VG TO ADD: " & $vgVolume & "mL"
				ElseIf $rVG <> 40 Then
					;MsgBox(4096, "VG 1 IS: ", $rVG)
					$vgVolume = Round((($rVG / 100) * $rVolume) - ($rNicVol/2), 1)
					$pgVolume = Round((((100 - $rVG)/100) * $rVolume - ($rNicVol/2) - $tVolume - $mVolume), 1); $pgVolume = ($rVolume - $vgVolume - ($rNic/2) - $tVolume)
					$vgInfo = "VG TO ADD: " & $vgVolume & "mL      " & "PG TO ADD: " & $pgVolume & "mL"
				Else
					;MsgBox(4096, "VG 2 IS: ", $rVG)
					$vgVolume = Round((($rVG/100)*$rVolume) - ($rNicVol/2), 1)
					$pgVolume = Round((((100 - $rVG)/100) * $rVolume - ($rNicVol/2) - $tVolume - $mVolume), 1)
					;MsgBox(4096, "Values = ", "Total= " & $rVolume & "VG= " & $vgVolume & "Nic=" & ($rNic/2) & "ConcTotal = " & $tVolume)
					$vgInfo = "VG TO ADD: " & $vgVolume & "mL      " & "PG TO ADD: " & $pgVolume & "mL"
				EndIf

			DisplayMixData($rName, $rJustName, $rNic, $rNicVol, $rVG, $rMenthol, $mVolume, $rVolume, $Ingredient, $specialInstructions, $vgInfo, $tVolume, $preMixVolume)
	Return
EndFunc

Func getInfoPrint($rVolume, $rJustName, $rNic, $rBoost, $rVG, $rMenthol, $rQuantity)
   	#Region GET DATE IN NICE WINDOWS FORMAT  XX-XX-XXXX
	$timeSys = _Date_Time_GetSystemTime()
	$unformattedTime = _Date_Time_SystemTimeToDateTimeStr($timeSys)
	$slashTime = StringRegExpReplace($unformattedTime, "/", "-")
	$timefull = StringLeft($slashTime, 10)
	Local $time = StringRegExpReplace($timefull, "20(?=\d\d)", "")
	;MsgBox(4096, "Date", $time)
	#EndRegion

   If $rVG = "100" Then
	  $rVG = "MAX"
   EndIf
   ;close P-Touch Editor before attempting to write CSV to avoid write lock
   If WinExists("P-touch Editor") Then
	   WinClose("P-touch Editor")
   EndIf
	If ProcessExists("Ptedit52.exe") Then ; Check if P-Touch is running
	 ;MsgBox($MB_SYSTEMMODAL, "", "P-Touch is running")
	 ProcessClose("Ptedit52.exe")
	EndIf

   $csvData = $rName&","&$rNic&","&$time&","&$rVG&","&$rMenthol&","&$rVolume&"mL"&","

   Local $sFilePath = @ScriptDir & "\Label Spreadsheet CSV.csv"
   Local $hFileOpen = FileOpen($sFilePath, $FO_OVERWRITE)
   FileWriteLine($hFileOpen, "flavor1,nic1,date1,vg1,menthol1,size1,flavor2,nic2,date2,vg2,menthol2,size2,flavor3,nic3,date3,vg3,menthol3,size3,flavor4,nic4,date4,vg4,menthol4,size4,flavor5,nic5,date5,vg5,menthol5,size5,flavor6,nic6,date6,vg6,menthol6,size6,flavor7,nic7,date7,vg7,menthol7,size7,flavor8,nic8,date8,vg8,menthol8,size8,flavor9,nic9,date9,vg9,menthol9,size9") ; CSV name headers
   FileWriteLine($hFileOpen, $csvData)
   FileClose($hFileOpen)
   FileFlush($hFileOpen)

   ;MsgBox(4096, "Hit Print", "")

   ShellExecute(@ScriptDir&"\1dymo.lbx", "", "", "", @SW_HIDE)

	WinWait("P-touch Editor", "")
		Sleep(3000)
		SplashTextOn("Title", "PRINTING LABELS NOW", -1, 60, -1, -1, 1, "", 24, 700)
		ControlSetText("P-touch Editor", "", "[ID:5148]", $rQuantity)
		Sleep(500)
		ControlClick("P-touch Editor", "", "[ID:5351]", "left")
		winwait("Print");wait for print dialog window
		ControlClick("Print", "", "[ID:5202]", "left") ;print all sheets radio box
		;MsgBox(4096, "WAIT HERE", "", 3000)
		ControlClick("Print", "", "[ID:1]", "left") ; click print
		;ControlClick("P-touch Editor", "", "[ID:5350]", "left", 1); sends command to print
		;MsgBox(4096, "Print Commanda sent", "")
		Sleep(2000)
		SplashOff()
		 Sleep(2000)
		 ;WinClose("P-touch Editor")
		 If ProcessExists("Ptedit52.exe") Then ; Check if P-Touch is running
			 ;MsgBox($MB_SYSTEMMODAL, "", "P-Touch is running")
			 ProcessClose("Ptedit52.exe")
		  EndIf
		  Return

EndFunc
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

Func getRecipe($rName, $bMult, $rNic, $rVolume, $rBoost, $rVG, $rMenthol, $rSalt)

local $arrayOfIngredients[0]
local $arrayofVendors[0]
local $arrayOfVolumes[0]

;## Set ingredient label text to blank if not used
Local $Ingredient[10] = ["", "", "", "", "", "", "", "", "", ""]
Local $iIndex = _ArraySearch($aArray, $rName, 0, 0, 0, 0)

;Make sure recipe is found
If @error Then
MsgBox($MB_SYSTEMMODAL, "Not Found", '"' & $rName & '" was not found in the array.', 3)
Return ;error so return to GUI
Else
EndIf

;Calculate and set Nicotine Volume
$rNicVol = Round(($rNic * $rVolume) / 100, 2)

;Salt Check
If $rSalt == "YES" Then
	$rNicVol = $rNicVol / 2.5
Else
	;nothing
EndIf

If @error Then
    MsgBox($MB_SYSTEMMODAL, "Not Found", '"' & $rName & '" was not found in the array.', 3)
	Return
Else
EndIf

	;Flavor Boost Check
If $rBoost == "YES" Then
	$fBoost = 1.7
	$rName = ($rName & " XTRA")
	;Check if Glass bulk mixed
	If $aArray[$iIndex][11] == "yes" Then
		$bulkMix = "GLASS"
	ElseIf $aArray[$iIndex][11] == "no" Then
		$bulkMix = "NO"
	Else
		$bulkMix = "ERROR"
	EndIf
Else
	$fBoost = 1
EndIf

;name before adding stuff
$rJustName = $rName

;Add Size to Name Text
$rName = $rVolume & " mL - " & $rName

    ;MsgBox($MB_SYSTEMMODAL, "Found", '"' & $rName & " Recipe:", 1)
	For $i = 1 To 10 Step 1
		$aIngredient = $aArray[$iIndex][$i]
			If $aIngredient = "" Then
				ExitLoop
			EndIf
		$sArray = StringSplit($aIngredient, "|")
		_ArrayAdd($arrayOfIngredients, $sArray[1])
		$finalVolume = Round(($sArray[3] * $bMult * $fBoost), 1)
		_ArrayAdd($arrayOfVolumes, $finalVolume)
		_ArrayAdd($arrayofVendors, $sArray[2])
		;$sPG = $finalVolume + $sPG
	Next
		$tVolume = 0
;~ 		$fPG = ($iPG - $sPG)
		For $f = 0 to (UBound($arrayOfVolumes)-1)
			$tVolume = $arrayOfVolumes[$f] + $tVolume
			_ArrayInsert($Ingredient, $f, $Ingredient[$f] & "INGREDIENT: " & $arrayofVendors[$f] & "  " & $arrayOfIngredients[$f] & "  Add:  " & $arrayOfVolumes[$f] & "mL")
				If $f < (UBound($arrayOfVolumes)-1) Then
					$Ingredient[$f] = $Ingredient[$f] & ""&@CRLF&""
				EndIf
		Next
			;_ArrayDisplay($Ingredient, "Ingredient Array")

			If ($aArray[$iIndex][9]) == "" Then
				$specialInstructions = ""
			Else
				$instructArray = StringSplit(($aArray[$iIndex][9]), "|")
				If $instructArray[1] == "Add" Then; [1]=Flag(Add or Percent), [2]=Add Premixed, [3]=Concentrate, [4]=Vendor, [5]=volume
					$instructVolume = Round(($instructArray[5] * $bMult * $fBoost), 1)
					$recipeVolume = Round($rVolume - $rNicVol - $instructVolume, 1)
					$specialInstructions = "OR: " & $recipeVolume & "mL of: " & $instructArray[2] & " " & $instructArray[4] & " " & $instructArray[3] & " Add: " & $instructVolume & " mL"
				ElseIf $instructArray[1] == "Percent" Then; [1]=Flag(Add or Percent), [2]=Premixed Recipe1, [3]=Percent1, [4]=Premixed Recipe2, [5]=Percent2
					$percent1Vol = ($instructArray[3] * $rVolume) - (0.5 * ($rNicVol)); + $mentholVolume))
					$percent2Vol = ($instructArray[5] * $rVolume) - (0.5 * ($rNicVol)); + $mentholVolume))
					$specialInstructions = "OR: " & "Add: " & $percent1Vol & " mL of " & $instructArray[2] & "  Add: " & $percent2Vol & " mL of " & $instructArray[4]
				Else
					$specialInstructions = ""
				EndIf
			EndIf

			;MENTHOL CALCULATIONS
			If $rMenthol = "NONE" Then
			   $mVolume = 0
			ElseIf $rMenthol = "LIGHT" Then
			   $mVolume = 0.2 * $bMult
			ElseIf $rMenthol = "MEDIUM" Then
			   $mVolume = 0.4 * $bMult
			ElseIf $rMenthol = "HEAVY" Then
			   $mVolume = 0.8 * $bMult
			ElseIf $rMenthol = "SUPER" Then
			   $mVolume = 1.6 * $bMult
			Else
			   $mVolume = 0
			EndIf

			;Premixed Volume
			$preMixVolume = ($rVolume - $rNicVol - $mVolume)

			If $rVG == 100 Then
				$vgVolume = Round((($rVG / 100) * $rVolume) - $rNicVol - $tVolume - $mVolume, 1)
				$vgInfo = "FILL REMAINDER WITH VG." & "  VG TO ADD: " & $vgVolume & "mL"
			ElseIf $rVG <> 40 Then
				;MsgBox(4096, "VG 1 IS: ", $rVG)
				$vgVolume = Round((($rVG / 100) * $rVolume) - ($rNicVol/2), 1)
				$pgVolume = Round((((100 - $rVG)/100) * $rVolume - ($rNicVol/2) - $tVolume - $mVolume), 1); $pgVolume = ($rVolume - $vgVolume - ($rNic/2) - $tVolume)
				$vgInfo = "VG TO ADD: " & $vgVolume & "mL      " & "PG TO ADD: " & $pgVolume & "mL"
			Else
				;MsgBox(4096, "VG 2 IS: ", $rVG)
				$vgVolume = Round((($rVG/100)*$rVolume) - ($rNicVol/2), 1)
				$pgVolume = Round((((100 - $rVG)/100) * $rVolume - ($rNicVol/2) - $tVolume - $mVolume), 1)
				;MsgBox(4096, "Values = ", "Total= " & $rVolume & "VG= " & $vgVolume & "Nic=" & ($rNic/2) & "ConcTotal = " & $tVolume)
				$vgInfo = "VG TO ADD: " & $vgVolume & "mL      " & "PG TO ADD: " & $pgVolume & "mL"
			EndIf

			;Call("DisplayMixData", $rName, $rNic, $rVG, $rMenthol, $rVolume, $Ingredient, $specialInstructions, $vgInfo, $tVolume, $preMixVolume)
			DisplayMixData($rName, $rJustName, $rNic, $rNicVol, $rVG, $rMenthol, $mVolume, $rVolume, $Ingredient, $specialInstructions, $vgInfo, $tVolume, $preMixVolume)

Return
EndFunc

Func DisplayMixData($rName, $rJustName, $rNic, $rNicVol, $rVG, $rMenthol, $mVolume, $rVolume, $Ingredient, $specialInstructions, $vgInfo, $tVolume, $preMixVolume)
	$MixDataGUI = GUICreate("RECIPE INFO", 1041, 911, 374, 0, $WS_BORDER, BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW))
	GUISetBkColor(0x000000)

	$QtyLabel = GUICtrlCreateLabel("QUANTITY", 892, 768, 142, 36)
	GUICtrlSetFont(-1, 20, 400, 0, "MS Sans Serif")
	_JoltStyle()

	$PrintQty = GUICtrlCreateCombo("", 896, 816, 137, 80, $CBS_DROPDOWN)
	GUICtrlSetData($PrintQty, "1|2|3|4|5|6|7|8|9|10", "1")
	GUICtrlSetFont(-1, 24, 400, 0, "MS Sans Serif")
	_JoltStyle()

	$Print = GUICtrlCreateButton("PRINT", 892, 3, 145, 759)
	GUICtrlSetFont(-1, 30, 400, 2, "skrunch")
	_JoltStyle()

	$RecipeName = GUICtrlCreateLabel($rName, 8, 8, 880, 41)
	GUICtrlSetFont(-1, 34, 800, 0, "Skrunch")
	_JoltStyle()

	$NicLabel = GUICtrlCreateLabel("NICOTINE -- ADD: " & $rNicVol & " mL", 8, 64, 880, 41)
	GUICtrlSetFont(-1, 24, 400, 0, "MS Sans Serif")
	_JoltStyle()

	$MentholLabel = GUICtrlCreateLabel("MENTHOL -- ADD: " & $mVolume & " mL", 8, 114, 880, 41)
	GUICtrlSetFont(-1, 24, 400, 0, "MS Sans Serif")
	_JoltStyle()

	$Ingredient1 = GUICtrlCreateLabel($Ingredient[0], 88, 168, 800, 41)
	GUICtrlSetFont(-1, 22, 800, 0, "MS Sans Serif")
	_JoltStyle()
	$Ingredient2 = GUICtrlCreateLabel($Ingredient[1], 88, 212, 800, 41)
	GUICtrlSetFont(-1, 22, 800, 0, "MS Sans Serif")
	_JoltStyle()
	$Ingredient3 = GUICtrlCreateLabel($Ingredient[2], 88, 256, 800, 41)
	GUICtrlSetFont(-1, 22, 800, 0, "MS Sans Serif")
	_JoltStyle()
	$Ingredient4 = GUICtrlCreateLabel($Ingredient[3], 88, 300, 800, 41)
	GUICtrlSetFont(-1, 22, 800, 0, "MS Sans Serif")
	_JoltStyle()
	$Ingredient5 = GUICtrlCreateLabel($Ingredient[4], 88, 344, 800, 41)
	GUICtrlSetFont(-1, 22, 800, 0, "MS Sans Serif")
	_JoltStyle()
	$Ingredient6 = GUICtrlCreateLabel($Ingredient[5], 88, 388, 800, 41)
	GUICtrlSetFont(-1, 22, 800, 0, "MS Sans Serif")
	_JoltStyle()
	$Ingredient7 = GUICtrlCreateLabel($Ingredient[6], 88, 432, 800, 41)
	GUICtrlSetFont(-1, 22, 800, 0, "MS Sans Serif")
	_JoltStyle()
	$Ingredient8 = GUICtrlCreateLabel($Ingredient[7], 88, 476, 800, 41)
	GUICtrlSetFont(-1, 22, 800, 0, "MS Sans Serif")
	_JoltStyle()
	$Ingredient9 = GUICtrlCreateLabel($Ingredient[8], 88, 520, 800, 41)
	GUICtrlSetFont(-1, 22, 800, 0, "MS Sans Serif")
	_JoltStyle()
	$Ingredient10 = GUICtrlCreateLabel($Ingredient[9], 88, 564, 800, 41)
	GUICtrlSetFont(-1, 22, 800, 0, "MS Sans Serif")
	_JoltStyle()

	$Special = GUICtrlCreateLabel($specialInstructions, 8, 608, 880, 41)
	GUICtrlSetFont(-1, 22, 800, 0, "MS Sans Serif")
	_JoltStyle()

	$VGLabel = GUICtrlCreateLabel($vgInfo, 88, 660, 785, 41)
	GUICtrlSetFont(-1, 24, 400, 0, "MS Sans Serif")
	_JoltStyle()

	$TotalVolume = GUICtrlCreateLabel("Total Concentrate = " & $tVolume & " mL" & "   OR Add Premixed " & $preMixVolume & "mL", 88, 710, 800, 41)
	GUICtrlSetFont(-1, 24, 400, 0, "MS Sans Serif")
	_JoltStyle()

	$CloseRecipe = GUICtrlCreateButton("Close Recipe View", 0, 760, 889, 145)
	GUICtrlSetFont(-1, 48, 400, 2, "skrunch")
	_JoltStyle()

	GUISetState(@SW_SHOW, $MixDataGUI)

		While 1
			$nMsg = GUIGetMsg()
			Switch $nMsg
				Case $GUI_EVENT_CLOSE
					GUIDelete($MixDataGUI)
					Return
				Case $CloseRecipe
					GUIDelete($MixDataGUI)
					Return
				Case $Print
					$rQuantity = GUICtrlRead($PrintQty)
					;MsgBox(4096, "QTY", "QTY= " & $rQuantity)
					Call("getInfoPrint", $rVolume, $rJustName, $rNic, $rBoost, $rVG, $rMenthol, $rQuantity)
			EndSwitch
		WEnd
	Return
EndFunc
