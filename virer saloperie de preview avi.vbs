
ret=MsgBox("Voulez vous d�sactiver la preview avi dans l'explorateur?"+ _
  vbCrLf+"(Repondre non la r�activera)",vbYesNoCancel)

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

