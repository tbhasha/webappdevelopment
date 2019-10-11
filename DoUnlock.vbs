set wsc = CreateObject("WScript.Shell")
Do While True
WScript.Sleep(2*60*1000)
wsc.SendKeys ("{CAPSLOCK}")
Loop