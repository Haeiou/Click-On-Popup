Set oShell = CreateObject("WScript.Shell")
While  true
if oShell.AppActivate("Test")     then
oShell.SendKeys "%{ENTER}"
WScript.Sleep 10
End If 
wend
 