Option Explicit

Dim objShell, intCount, intDelay, strInput, strOutput, i

Set objShell = CreateObject("WScript.Shell")

intCount = 0
intDelay = 3

For i = 0 To 9999
    strInput = PadNumber(i)
    strOutput = strInput & "{ENTER}"
    
    objShell.SendKeys strOutput
    
    intCount = intCount + 1
    
    If intCount Mod 4 = 0 Then
        WScript.Sleep 30000 ' 30 seconds delay
    Else
        WScript.Sleep intDelay * 1000 ' 3 seconds delay
    End If
Next

Function PadNumber(intNumber)
    Dim strNumber
    strNumber = CStr(intNumber)
    Do While Len(strNumber) < 4
        strNumber = "0" & strNumber
    Loop
    PadNumber = strNumber
End Function