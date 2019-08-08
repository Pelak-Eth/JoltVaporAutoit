#include <IE.au3>

While 1
	While 1
		Do
			Sleep(4000)
		Until WinExists("Touch Order")
		WinWait("Touch Order")
		Local $oIE = _IEAttach("Touch Order")
		Local $oTable = _IETableGetCollection($oIE, 0)
			If @error = 0 Then
				SoundPlay(@ScriptDir&"\gong.mp3")
				Do
					Sleep(5000)
					Local $oTable = _IETableGetCollection($oIE, 0)
				Until @error = 7
				ExitLoop
			Elseif @error = 7 Then
				ExitLoop
			EndIf
	WEnd
WEnd