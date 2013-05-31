

changeAllFiles "[\._]"," ",".\"

sub changeAllFiles(reg,repl,dir)
	Files=getFiles(dir)

	For i=0 To UBound(Files)
	  File = Split(Files(i),".")
	  If UBound(File)=0 Then
	    ext=""
	  Else
	    ext="."+File(UBound(File))
	    ReDim preserve File(UBound(File)-1)
	  End If
	  
	  File=Join(File,".")
	  Set regEx = New RegExp
	  regEx.Pattern = reg
	  regex.global = True
	  result=regex.replace(File,repl)
	  If result<>file And Files(i)<>WScript.ScriptName Then
	    Set fso = CreateObject("Scripting.FileSystemObject")
	    Set f2 = fso.GetFile(Files(i))
	    f2.Move (result+ext)
	  End If
	Next
end sub

Function getFiles(Folder)
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set f = fso.GetFolder(Folder)
  Set fc = f.Files
  arr=Array()
  
  For Each fich in fc
    ReDim preserve arr(UBound(arr)+1)
    arr(UBound(arr))=fich.name
  Next
  getFiles=arr
End Function




changeAllFiles "[ \._]*\[[^\[\]]*\][ \._]*","",".\"

sub changeAllFiles(reg,repl,dir)
	Files=getFiles(dir)

	For i=0 To UBound(Files)
	  File = Split(Files(i),".")
	  If UBound(File)=0 Then
	    ext=""
	  Else
	    ext="."+File(UBound(File))
	    ReDim preserve File(UBound(File)-1)
	  End If
	  
	  File=Join(File,".")
	  Set regEx = New RegExp
	  regEx.Pattern = reg
	  regex.global = True
	  result=regex.replace(File,repl)
	  If result<>file And Files(i)<>WScript.ScriptName Then
	    Set fso = CreateObject("Scripting.FileSystemObject")
	    Set f2 = fso.GetFile(Files(i))
	    f2.Move (result+ext)
	  End If
	Next
end sub

Function getFiles(Folder)
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set f = fso.GetFolder(Folder)
  Set fc = f.Files
  arr=Array()
  
  For Each fich in fc
    ReDim preserve arr(UBound(arr)+1)
    arr(UBound(arr))=fich.name
  Next
  getFiles=arr
End Function


