
result = MsgBox("Voulez vous acc�l�rer le partage r�seau?"+vbcrlf _
              +"(r�pondre non remettra les param�tres d'origine)" _
							,vbYesnocancel)

If result = vbyes Then
	desactivate
ElseIf result = vbno Then
	activate
End If
	
Sub Activate()	
	Set Sh = CreateObject("WScript.Shell")
	key =  "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Explorer\RemoteComputer\NameSpace\{D6277990-4C6A-11CF-8D87-00AA0060F5BF}\"
	sh.regwrite key,"T�ches planifi�es"
End Sub

Sub Desactivate()
	Set Sh = CreateObject("WScript.Shell")
	key =  "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Explorer\RemoteComputer\NameSpace\{D6277990-4C6A-11CF-8D87-00AA0060F5BF}\"
	sh.RegDelete key
End Sub

