' Shutdown computer 
Set oShell = WScript.CreateObject ("WScript.Shell")
oShell.run "shutdown /a"
Set oShell = Nothing'