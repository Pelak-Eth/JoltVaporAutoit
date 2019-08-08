#include <IE.au3>
#include <MsgBoxConstants.au3>
#include <WinAPI.au3>

;Simple script to login to web based camera monitoring

ShellExecute("iexplore.exe", "http://10.1.10.76/")

MsgBox(4096, "CAMERAS", "IS LOGIN PAGE RESPONSIVE")
Do
sleep(1000)
Until WinExists("Web Client for EDVS/EDVR")

WinWaitActive("Web Client for EDVS/EDVR")
_WinAPI_SetFocus(WinGetHandle("Web Client for EDVS/EDVR"))
ControlSend("Web Client for EDVS/EDVR", "", "", "aUserHere")
sleep(500)
ControlSend("Web Client for EDVS/EDVR", "", "", "{TAB}")
sleep(500)
ControlSend("Web Client for EDVS/EDVR", "", "", "aPasswordHere")
sleep(500)
ControlSend("Web Client for EDVS/EDVR", "", "", "{TAB}")
sleep(500)
ControlSend("Web Client for EDVS/EDVR", "", "", "{TAB}")
sleep(500)
ControlSend("Web Client for EDVS/EDVR", "", "", "{SPACE}")
sleep(500)
ControlSend("Web Client for EDVS/EDVR", "", "", "{ENTER}")

sleep(4000)

WinMove("Web Client for EDVS/EDVR", "", -134, 0, 1240, 1080)
