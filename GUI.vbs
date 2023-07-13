Set oShell = CreateObject ("Wscript.Shell")
Dim strArgs
strArgs = "cmd /c GUI1.exe"
oShell.Run strArgs, 0 , false