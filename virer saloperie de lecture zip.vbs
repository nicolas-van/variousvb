
hellomaster

ret=MsgBox("Voulez vous désactiver la lecture des zips dans l'explorateur?"+ _
  vbCrLf+"(Repondre non la réactivera)",vbYesNoCancel)

If ret=vbCancel Then
  Wscript.quit
End If

If ret=vbYes Then
  Set WshShell = WScript.CreateObject("WScript.Shell")
  returned = WshShell.Run("regsvr32 /u /s zipfldr.dll")
  MsgBox ("Relancez windows puis n'oubliez surtout pas d'assigner les"+ _
    " fichiers zip à un programme de lecture d'archive.")
Else
  Set WshShell = WScript.CreateObject("WScript.Shell")
  returned = WshShell.Run("regsvr32 /s zipfldr.dll")
  MsgBox ("Relancez windows.")
End If

Sub hellomaster()
  MsgBox "Bonjour maître."+vbCrLf+ _
         "Je vais essayer de ne pas formater votre disque dur."
End Sub

