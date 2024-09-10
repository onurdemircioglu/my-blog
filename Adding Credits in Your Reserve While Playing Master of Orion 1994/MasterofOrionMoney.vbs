Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.AppActivate "Mozilla Firefox"  ' Activate Firefox window
WScript.Sleep 1000  ' Delay for 1 second (1000 milliseconds)

For j = 1 To 20  ' Repeat x times
    WshShell.SendKeys "%m%o%o%l%a"  ' Send "moola" with Alt key
    WScript.Sleep 1100  ' Delay for 1 second between each sequence
Next
