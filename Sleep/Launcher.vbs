WScript.sleep 1000

' Shutdown computer 
Set oShell = WScript.CreateObject ("WScript.Shell")

' You can change the time here if you want 
' Will try to make this process more user friendly
oShell.run "shutdown /s /t 7200"
Set oShell = Nothing'

WScript.sleep 1000

' Turn off monitors
Set oShell = WScript.CreateObject ("WScript.Shell")
strPath = oShell.CurrentDirectory + "\DWN.bat"
oShell.run strPath
Set oShell = Nothing'
