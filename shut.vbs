Dim objShell
Set objShell = WScript.CreateObject("WScript.Shell")
objShell.Run "shutdown -s -f -t 10", 0, True