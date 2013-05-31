
Set Sh = CreateObject("WScript.Shell")
key =  "HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\lanmanserver\parameters\AutoShareWks"
sh.regwrite key,0,"REG_DWORD"
key =  "HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\lanmanserver\parameters\AutoShareServer"
sh.regwrite key,0,"REG_DWORD"
MsgBox ("Relancez windows.")
