#include <IE.au3>

While 1
	While 1
		Do
			Sleep(4000)
		Until WinExists("Touch Order")
		WinWait("Touch Order")
		Local $oIE = _IEAttach("Touch Order")
		Local $oTable = _IETableGetCollection($oIE, 0)
		Local $isData = _IETableWriteToArray($oTable)
		;MsgBox(4096, "TABLE SIZE", "LENGTH = " & UBound($isData, 2))
			If UBound($isData, 2) >= 2 Then
				SoundPlay(@ScriptDir&"\gong.mp3")
				Do
					Sleep(5000)
					Local $oTable = _IETableGetCollection($oIE, 0)
					Local $isData = _IETableWriteToArray($oTable)
					;MsgBox(4096, "TABLE SIZE", "2nd LENGTH = " & UBound($isData, 2))
				Until UBound($isData, 2) = 1
				ExitLoop
			Else
				ExitLoop
			EndIf
	WEnd
WEnd