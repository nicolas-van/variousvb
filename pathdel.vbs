' usage:
' pathdel.vbs [nom du dossier] [-scope User|System]

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
set env = sh.environment(scope)
pathdocs = split(env("PATH"),";")
machin = array()
for i=0 to ubound(pathdocs)
	if pathdocs(i) <> curpath then
		redim preserve machin(ubound(machin)+1)
		machin(ubound(machin))=pathdocs(i)
	end if
next
env("PATH") = join(machin,";")

