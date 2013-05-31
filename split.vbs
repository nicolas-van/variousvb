
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFile("tmp.txt")
set text = f.openastextstream(1)
l = text.ReadLine
tab = split(l,"&")
l = join(tab,vbCrlf)

fso.CreateTextFile "result.txt"
set f = fso.getfile("result.txt")
set text = f.openastextstream(2)
text.write(l)


