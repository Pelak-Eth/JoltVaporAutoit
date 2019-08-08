#NoTrayIcon

#Include <Misc.au3>
#Include <Restart.au3>

;_Singleton('MyProgram')

If MsgBox(36, 'Restarting...', 'Press OK to restart this script.') = 6 Then
    _ScriptRestart()
EndIf