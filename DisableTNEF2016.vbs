Dim WSHShell
Set WSHShell = WScript.CreateObject("WScript.Shell")

WSHShell.RegWrite "HKCU\Software\Microsoft\Office\16.0\Outlook\Preferences\DisableTNEF", 1, "REG_DWORD"

WSHShell.Popup "DisableTNEF Çê›íËÇµÇ‹ÇµÇΩÅB"

