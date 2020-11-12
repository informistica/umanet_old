<%

function ReplaceCar(sInput)
	dim sAns
	sAns=  Replace(sInput,"'",Chr(96)) ' sostituisco l'apice ' con quello storto per non disturbare la sintassi 
	sAns=  Replace(sAns,Chr(39),Chr(96)) 
	sAns=  Replace(sAns,chr(58),Chr(44)) ' sostituisco : con , per non disturbare la creazione del file
	sAns = Replace(sAns,"à","a"&Chr(96))	
	sAns = Replace(sAns,"è","e"&Chr(96))
	sAns = Replace(sAns,"é","e"&Chr(96))
	sAns = Replace(sAns,"ì","i"&Chr(96))
' sAns = Replace(sAns,chr(237),"i"&Chr(96))
	sAns = Replace(sAns,"ò","o"&Chr(96))
	sAns = Replace(sAns,"ù","u"&Chr(96))
	
'  sAns = Replace(sAns,chr(243),"o"&Chr(96))
'	sAns = Replace(sAns,chr(151),"u"&Chr(96))
'  sAns = Replace(sAns,"&#65533;","u"&Chr(96))
'  sAns = Replace(sAns,chr(250),"u"&Chr(96)) 
	sAns = Replace(sAns, Chr(34), Chr(96))' sostituisco gli apici " con l'apice storto
'sAns=  Replace(sAns,"'",) ' sostituisco l'apice ' con quello storto per non disturbare la sintassi sql
	sAns=  Replace(sAns,":",",") ' sostituisco : con , per non disturbare la creazione del file
	sAns=  Replace(sAns,"&","e") 
	sAns=  Replace(sAns,"/","-") 
	sAns=  Replace(sAns,"\","-") 
	sAns=  Replace(sAns,"?",".") 
	sAns=  Replace(sAns,"*","x") 
	sAns=  Replace(sAns,"<","_")
	sAns=  Replace(sAns,">","_") 
	ReplaceCar = sAns
end function
%>


