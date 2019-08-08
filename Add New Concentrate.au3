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
If _Singleton("Add New Concentrate", 1) = 0 Then
	MsgBox(4096, "ALREADY RUNNING!", "CLOSING DUPLICATE INSTANCE", 3)
    Exit
EndIf

Local $concVendor = ""
Local $fCombo = ""
Local $concentrate = IniRead(@ScriptDir & "\RecipesINI.ini", "RECIPEINFO", "concentrate", "")

Func _JoltStyle()
	GUICtrlSetColor(-1, 0xC8C8C8)
	GUICtrlSetBkColor(-1, 0x000000)
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
Local $aResult = _SQLite_GetTable2d(-1, "SELECT * FROM Recipes;", $aArray1, $iRows, $iColumns); Write DB to 2D Array
If $aResult = $SQLITE_OK Then
    ;_SQLite_Display2DResult($aResult)
Else
    MsgBox($MB_SYSTEMMODAL, "SQLite Error: " & $iRval, _SQLite_ErrMsg())
	Exit
EndIf

Local $aArray2, $iRows1, $iColumns1, $iRval1
Local $aResult2 = _SQLite_GetTable2d(-1, "SELECT * FROM Concentrates;", $aArray2, $iRows1, $iColumns1)

Global $aArray = $aArray1
Global $cArray = $aArray2
;_ArrayDisplay($cArray, "Array = "); COL[0] = Concentrate, COL[1] = Vendor, COL[2] = Volume, COL[3] = Min, COL[4] = Max,  ALL DATA STARTS AT ROW[1]
EndFunc

Func sqlQueryToArray()
; Query DB and Write to Array
Local $aArray1, $iRows, $iColumns, $iRval
Local $aResult = _SQLite_GetTable2d($hDskDb, "SELECT * FROM Recipes order by Recipe asc;", $aArray1, $iRows, $iColumns); Write DB to 2D Array

Local $aArray2, $iRows1, $iColumns1, $iRval1
Local $aResult2 = _SQLite_GetTable2d($hDskDb, "SELECT * FROM Concentrates order by Concentrate asc;", $aArray2, $iRows1, $iColumns1)

Global $aArray = $aArray1
Global $cArray = $aArray2
EndFunc

#Region ### START Koda GUI section ### Form=C:\Users\windows\Documents\Autoit\Add New Concentrate.kxf
			$AddNewConcentrate = GUICreate("ADD NEW CONCENTRATE GUI", 608, 312, ((@DesktopWidth/2)-304), 0)
			GUISetBkColor(0x000000)
			GUISetFont(24, 400, 2, "skrunch")

			$Label1 = GUICtrlCreateLabel("ADD NEW CONCENTRATE GUI", 95, 0, 417, 37)
			_JoltStyle()

			$Label2 = GUICtrlCreateLabel("CONCENTRATE NAME", 16, 50, 184, 24)
			GUICtrlSetFont(-1, 14, 400, 2, "skrunch")
			_JoltStyle()

			$Label3 = GUICtrlCreateLabel("SELECT VENDOR", 15, 99, 149, 24)
			GUICtrlSetFont(-1, 14, 400, 2, "skrunch")
			_JoltStyle()

			$Label4 = GUICtrlCreateLabel("SELECT MINIMUM VOLUME", 19, 149, 244, 24)
			GUICtrlSetFont(-1, 14, 400, 2, "skrunch")
			_JoltStyle()

			$Label5 = GUICtrlCreateLabel("SELECT MAXIMUM VOLUME", 22, 198, 244, 24)
			GUICtrlSetFont(-1, 14, 400, 2, "skrunch")
			_JoltStyle()

			$Name = GUICtrlCreateInput($concentrate, 208, 50, 385, 33)
			GUICtrlSetFont(-1, 18, 400, 2, "skrunch")
			_JoltStyle()

			$Vendor = GUICtrlCreateCombo("", 208, 99, 385, 25, $CBS_DROPDOWNLIST)
			GUICtrlSetData(-1, "TPA|1on1|CAP|FA|CONC", "TPA")
			GUICtrlSetFont(-1, 18, 400, 2, "skrunch")
			_JoltStyle()

			$MIN = GUICtrlCreateCombo("SELECT MIN", 297, 149, 297, 41, $CBS_DROPDOWNLIST)
			GUICtrlSetData(-1, "0.05|0.1|0.15|0.2|0.25|0.3|0.35|0.4|0.5|0.6|0.7|0.8|0.9|1.0|1.1|1.2|1.3|1.4|1.5|1.6")
			GUICtrlSetFont(-1, 18, 400, 2, "skrunch")
			_JoltStyle()

			$MAX = GUICtrlCreateCombo("SELECT MAX", 296, 198, 297, 41, $CBS_DROPDOWNLIST)
			GUICtrlSetData(-1, "0.05|0.1|0.15|0.2|0.25|0.3|0.35|0.4|0.5|0.6|0.7|0.8|0.9|1.0|1.1|1.2|1.3|1.4|1.5|1.6")
			GUICtrlSetFont(-1, 18, 400, 2, "skrunch")
			_JoltStyle()

			$Submit = GUICtrlCreateButton("SUBMIT NEW CONCENTRATE", 15, 248, 577, 57)
			_JoltStyle()

			GUISetState(@SW_SHOW)

			While 1
				$nMsg = GUIGetMsg()
				Switch $nMsg
					Case $GUI_EVENT_CLOSE
						Exit

					Case $Submit
						;Read data from GUI
						$preconcName = GUICtrlRead($Name)
						$concName = StringRegExpReplace($preconcName, " ", "_")

						;Check if concentrate exists
						Local $iIndex = _ArraySearch($cArray, $concName, 0, 0, 0, 0)
							If @error <> 0 Then
								;continue
								$concVendor = GUICtrlRead($Vendor)
								$Volume_mL = 0
								$Volume_oz = 0
								$concMin = GUICtrlRead($MIN)
								$concMax = GUICtrlRead($MAX)
							Else
								MsgBox(4096, "Already Exists", $concName & " has already been added.  Press OK to exit")
								_SQLite_Close()
								_SQLite_Shutdown()
								Exit
							EndIf

						;INSERT DATA INTO DATABASE
						$InsertSQL = "INSERT INTO Concentrates (Concentrate,Vendor,Volume_mL,Volume_oz,Min,Max) VALUES (" & _SQLite_FastEscape($concName) &","&_SQLite_FastEscape($concVendor) &","& _SQLite_FastEscape($Volume_mL) &","& _SQLite_FastEscape($Volume_oz) &","& _SQLite_FastEscape($concMin) &","& _SQLite_FastEscape($concMax) &")"
							If Not _SQLite_Exec(-1, $InsertSQL) = $SQLITE_OK Then
								MsgBox(4096, "SQLite Error", _SQLite_ErrMsg())
								Exit
							Else
								_SQLite_Exec(-1, $InsertSQL)
							EndIf
						IniWrite(@ScriptDir & "\RecipesINI.ini", "RECIPEINFO", "concentrate", "")
						_SQLite_Close()
						_SQLite_Shutdown()
						Exit
				EndSwitch
			WEnd
#EndRegion ### END Koda GUI section ###