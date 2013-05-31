
hellomaster

mode=MsgBox("Voulez vous encoder tous les fichiers wav?"+vbCrLf+ _
            "(Repondre non décodera tous les fichiers mp3)" _
            ,vbYesNoCancel)

If mode=vbCancel Then Wscript.quit

paramEncode=""
ext="mp3"

If mode=vbYes Then
  paramEncode=InputBox("Veuillez spécifier les paramètres d'encodage:"+vbCrLf+ _
                       "(Ceci est facultatif)")
  ext="wav"
End If

arr=getFiles(".\")

reals=Array()
names=Array()
For i=0 To UBound(arr)
  sep=Split(arr(i),".")
  If sep(UBound(sep))=ext Then
    ReDim preserve sep(UBound(sep)-1)
    ReDim preserve names(UBound(names)+1)
    names(UBound(names))=Join(sep)
  End If
Next

For i=0 To UBound(names)
  Set WshShell = WScript.CreateObject("WScript.Shell")
  If mode=vbYes Then
    returned = WshShell.Run("lame "+paramEncode+" "+names(i)+".wav "+ _
                names(i)+".mp3", 1, true)
  Else
    returned = WshShell.Run("lame --decode "+names(i)+".mp3 "+ _
                names(i)+".wav", 1, true)
  End If
Next


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

Sub hellomaster()
  MsgBox "Bonjour maître."+vbCrLf+ _
         "Je vais essayer de ne pas formater votre disque dur."
End Sub
