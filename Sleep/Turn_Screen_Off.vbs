' Turn off monitors
Set oShell = WScript.CreateObject ("WScript.Shell")
strPath = oShell.CurrentDirectory + "\DWN.bat"
oShell.run strPath
Set oShell = Nothing'