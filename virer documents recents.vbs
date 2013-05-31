



Set Sh = CreateObject("WScript.Shell")
key = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\"
sh.regwrite key,0
key = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoRecentDocsHistory"
sh.regwrite key,1,"REG_DWORD"

MsgBox "Maintenant redémarrez votre ordinateur."



