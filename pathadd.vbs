' usage:
' pathadd.vbs [nom du dossier] [-scope User|System]

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
pathdocs = split(env("PATH"),";")
shouldadd = true
for i=0 to ubound(pathdocs)
	if pathdocs(i)=curpath then
		shouldadd=false
	end if
next

if shouldadd then
	redim preserve pathdocs(ubound(pathdocs)+1)
	pathdocs(ubound(pathdocs))=curpath
	env("PATH") = join(pathdocs,";")
end if
