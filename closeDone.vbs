' Wait for 500 milliseconds
WScript.Sleep(500)

' Create Shell instance
Set objShell = CreateObject("WScript.Shell")

' Close the "Done" message box
objShell.SendKeys "{ENTER}"
