<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "https://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="https://www.w3.org/1999/xhtml">
<head>
<meta https-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Documento senza titolo</title>
</head>

<body>
<%
'  sInput="andiamo è é ora à soreta"
'  sAns=  Replace(sInput,"'",Chr(96)) ' sostituisco l'apice ' con quello storto per non disturbare la sintassi sql
' 
' 
'  
' 
'  sAns=  Replace(sAns,"è","&eacute;") ' sostituisco : con , per non disturbare la creazione del file
'  sAns=  Replace(sAns,"é","&egrave;")
'  response.write(sAns) 
'  response.write("<br>"& chr(234)) 
'  response.write("<br>"& chr(233)) 
'  response.write("<br>"& chr(95)) 
'  response.write("<br>&eacute;") 
'  
  'for i=1 to 255
'       response.write("<br>"& i &"=" & chr(i)) 
'  next
si=0
no=0
n=1000
randomize()
for i=1 to n 
				rand=rnd()
				 
				     if (cint(left(rand*100,2)) mod 2)= 0 then ' se il numero casuale è pari (testa o croce)
					 	si=si+1
						'response.write ("<br>HAI OTTENUTO 1 BONUS!   " &cint(left(rand*100,2)))
					 else
					 '  response.write ("<br>POTEVI OTTENERE 1 BONUS! MA NON SEI STATO FORTUNATO!")	
					   no=no+1				
					end if
							 
next
response.write ("<br>HAI OTTENUTO 1 BONUS!   " & si/n*100 &"%")
response.write ("<br>POTEVI OTTENERE 1 BONUS! MA NON SEI STATO FORTUNATO!" & no/n*100 &"%")	
%>
</body>
</html>
