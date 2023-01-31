set wshshell = wscript.CreateObject("wscript.shell")
wshshell.run "Notepad"
wscript.sleep 2000
wshshell.AppActivate "Notepad"
WshShell.SendKeys "HA"
WScript.Sleep 500
WshShell.SendKeys "AH"
WScript.Sleep 500
WshShell.SendKeys "H "
WScript.Sleep 500
WshShell.SendKeys "HA"
WScript.Sleep 500
WshShell.SendKeys "L"
WScript.Sleep 500
WshShell.SendKeys "OL!"
WScript.Sleep 500