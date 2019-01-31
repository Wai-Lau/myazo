Set oShell = CreateObject ("Wscript.Shell") 
Dim strArgs
strArgs = "cmd /c myazo.bat"
oShell.Run strArgs, 0, false