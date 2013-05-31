
Set oFSO = CreateObject("Scripting.FileSystemObject")
curpath = oFSO.getfolder(".\").path
scope = "User"
scoped = false
for i=0 to wscript.Arguments.length-1
	if wscript.arguments(i)="-scope" then
		scoped = true
	elseif scoped then
		scope = wscript.arguments(i)
	else
		curpath = wscript.arguments(i)
	end if
next

Set Sh = CreateObject("WScript.Shell")
set env = sh.environment("User")
wscript.echo env("PATH")


