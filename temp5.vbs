Set oShell = CreateObject ("Wscript.Shell")
Dim strArgs
strArgs = "cmd /c temp5.exe"
oShell.Run strArgs, 0 , false