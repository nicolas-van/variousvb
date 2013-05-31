
icon="D:\videos\mplayerc.exe,0"

list = ".drc .dsm .dsv .dsa .dss .vob .ifo .d2v .flv .fli .flc .flic .ivf .mkv .mpg .mpeg .mpe .m1v .m2v .mpv2 .mp2v .dat .ts .tp .tpr .pva .pss .mp4 .m4v .m4p .m4b .3gp .3gpp .3g2 .3gp2 .ogm .divx .vp6 .asx .m3u .pls .wvx .wax .wmx .mpcpl .mov .qt .amr .ratdvd .rm .ram .rmvb .rpm .rt .rp .smi .smil .roq .swf .smk .bik .avi .wmv .wmp .wm .asf"
files = split(list," ")
for i=0 to ubound(files)
	Set Sh = CreateObject("WScript.Shell")
	firstkey="HKEY_CLASSES_ROOT\"+files(i)+"\"
	secondkey="HKEY_CLASSES_ROOT\"+sh.regread(firstkey)+"\DefaultIcon\"
	sh.regwrite secondkey,icon,"REG_SZ"
	
next

msgbox "Vous pouvez redémarrer l'ordinateur."
