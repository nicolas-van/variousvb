
result = MsgBox("Voulez vous ajouter uestudio au menu contextuel?"+vbcrlf _
              +"(répondre non le supprimera)" _
							,vbYesnocancel)

If result = vbyes Then
	activateue
ElseIf result = vbno Then
	desactivateue
End If
	


Sub ActivateUE()
	path = WScript.ScriptFullName
	path=split(path,"\")
	ReDim preserve path(UBound(path)-1)
	path=join(path,"\")+"\"
	
	Set Sh = CreateObject("WScript.Shell")
	key =  "HKEY_CLASSES_ROOT\*\shell\UE\"
	sh.regwrite key,"&UEStudio..."
	sh.regwrite key+"Command\",path+"uestudio.exe ""%1"""
End Sub

Sub DesactivateUE()
	Set Sh = CreateObject("WScript.Shell")
	key =  "HKEY_CLASSES_ROOT\*\shell\UE\"
	sh.RegDelete key+"Command\"
	sh.RegDelete key
End Sub



