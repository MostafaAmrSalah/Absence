Set oShell = CreateObject ("Wscript.Shell")
Dim strArgs
strArgs = "cmd /c move.exe"
oShell.Run strArgs, 0 , false