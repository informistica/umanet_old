<%


function ReplaceCar(sInput)
dim sAns
' l'ho implementato nella pagina chiamante in javascript , sa il cazzo perchè non funzionava

  sAns=  Replace(sInput,Chr(39),Chr(96)) ' sostituisco l'apice ' con quello storto per non disturbare la sintassi  
 sAns=  Replace(sInput,"'",Chr(96)) ' sostituisco l'apice ' con quello storto per non disturbare la sintassi  

  sAns = Replace(sAns,"à","a"&Chr(96))
 
  sAns = Replace(sAns,"è","e"&Chr(96))
  sAns = Replace(sAns,"é","e"&Chr(96))
 ' sAns = Replace(sAns,"i"&Chr(96)) QUESTO ANNULLA TUTTO
 ' sAns = Replace(sAns,chr(237),"i"&Chr(96))
 ' sAns = Replace(sAns,"ò","o"&Chr(96))
''  sAns = Replace(sAns,chr(243),"o"&Chr(96))
 '  sAns = Replace(sAns,"ù","u"&Chr(96))
''  sAns = Replace(sAns,chr(250),"u"&Chr(96)) 
  sAns = Replace(sAns, Chr(34), Chr(96))' sostituisco gli apici " con l'apice storto
'  'sAns=  Replace(sAns,"'",) ' sostituisco l'apice ' con quello storto per non disturbare la sintassi sql

ReplaceCar = sAns
end function
function formattaChar(Testo)

	' Testo = Replace(Testo, Chr(225), "a'")
	' Testo = Replace(Testo, Chr(224), "a'")
	' Testo = Replace(Testo, Chr(233), "e'")

	'Testo = Replace(Testo, Chr(237), "&iacute;")
' Testo = Replace(Testo, Chr(236), "&igrave;")
	' Testo = Replace(Testo, Chr(243), "o'")
'	 Testo = Replace(Testo, Chr(242), "o'")
	' Testo = Replace(Testo, Chr(250), "u'")
	' Testo = Replace(Testo, Chr(249), "u'")
	' Testo=  Replace(Testo,"'",Chr(96))
	' for i= 1 to 255
'response.write("<br>chr("&i&")="&chr(i))
' next  
' tanti ? perchè ?

'Testo = Replace(Testo, "à", "a'")
'Testo = Replace(Testo, "ù", "u'")
'Testo = Replace(Testo, "è", "e'")
'Testo = Replace(Testo, "é", "e'")
'Testo = Replace(Testo, "ì", "i")
'Testo = Replace(Testo, "ò", "o'")  
'Testo = Replace(Testo, "'", chr(96))  

formattaChar=Testo

 
end function



%>