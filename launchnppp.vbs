
param=""
for i=0 to wscript.Arguments.length-1
	if i<>0 then
		param=param+" "
	end if
	param=param+wscript.arguments(i)
next

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run """C:\Program Files\Notepad++\notepad++.exe"" """+param+"""", 1, false
