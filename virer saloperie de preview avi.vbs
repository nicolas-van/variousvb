
ret=MsgBox("Voulez vous désactiver la preview avi dans l'explorateur?"+ _
  vbCrLf+"(Repondre non la réactivera)",vbYesNoCancel)

If ret=vbCancel Then
  Wscript.quit
End If

If ret=vbYes Then
  Set WshShell = WScript.CreateObject("WScript.Shell")
  returned = WshShell.Run("regsvr32 /u /s shmedia.dll")
  MsgBox ("Relancez windows.")
Else
  Set WshShell = WScript.CreateObject("WScript.Shell")
  returned = WshShell.Run("regsvr32 /s shmedia.dll")
  MsgBox ("Relancez windows.")
End If

