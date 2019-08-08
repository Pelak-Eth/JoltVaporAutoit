#include <IE.au3>
#include <MsgBoxConstants.au3>
#include <WinAPI.au3>
#include <Misc.au3>

;##########################
;Allow only one instance ##
;##########################
If _Singleton("Login to POS", 1) = 0 Then
	MsgBox(4096, "ALREADY RUNNING!", "CLOSING DUPLICATE INSTANCE", 3)
    Exit
EndIf

;Function for running other Autoit scripts from here
Func _RunAU3($sFilePath, $sWorkingDir = @ScriptDir, $iShowFlag = @SW_SHOW, $iOptFlag = 0)
    Return Run('"' & @AutoItExe & '" /AutoIt3ExecuteScript "' & $sFilePath & '"', $sWorkingDir, $iShowFlag, $iOptFlag)
EndFunc

ShellExecute("iexplore.exe", "http://192.168.1.76/")
;WinWait("Web Client for EDVS/EDVR", 10)

Do
sleep(1000)
Until WinExists("Web Client for EDVS/EDVR")

MsgBox(4096, "CAMERAS", "IS LOGIN PAGE RESPONSIVE?")

WinWaitActive("Web Client for EDVS/EDVR")
_WinAPI_SetFocus(WinGetHandle("Web Client for EDVS/EDVR"))
ControlSend("Web Client for EDVS/EDVR", "", "", "wyomingjolt")
sleep(500)
ControlSend("Web Client for EDVS/EDVR", "", "", "{TAB}")
sleep(500)
ControlSend("Web Client for EDVS/EDVR", "", "", "42069")
sleep(500)
ControlSend("Web Client for EDVS/EDVR", "", "", "{TAB}")
sleep(500)
ControlSend("Web Client for EDVS/EDVR", "", "", "{TAB}")
sleep(500)
ControlSend("Web Client for EDVS/EDVR", "", "", "{SPACE}")
sleep(500)
ControlSend("Web Client for EDVS/EDVR", "", "", "{ENTER}")

sleep(3000)

WinMove("Web Client for EDVS/EDVR", "", -134, 0, 1240, 1080)

ShellExecute("iexplore.exe", "http://joltpos.com:1337/user/login")
;ShellExecute("iexplore.exe", "http://joltpos.com:1337/order/touch")


WinWait("Krypt POS", "", 240)
$oIE = _IEAttach("Krypt POS 3.0")

Local $oForm = _IEFormGetCollection($oIE, 0)
Local $oUser = _IEFormElementGetCollection($oForm, 0)
_IEFormElementSetValue($oUser, "jolt2@joltvapor.com")

Local $oPassword = _IEFormElementGetCollection($oForm, 1)
_IEFormElementSetValue($oPassword, "joltvapor")

local $submitButton = _IEGetObjByName($oIE, "submit")
_IEAction($submitButton, "click")

;_IELoadWait($oIE)

sleep(2000)

_IENavigate($oIE, "http://joltpos.com:1337/order/touch", 0)

sleep(2000)

MsgBox(4096, "Move Touch POS Window", "Click OK to Move POS Window")

sleep(3000)

_RunAU3("Move POS.au3")

_WinAPI_SetFocus(WinGetHandle("Touch Order"))
WinWaitActive("Touch Order")

_RunAU3("Move POS.au3")

;~ WinMove("Touch Order", "", 915, 0, 905, 938)