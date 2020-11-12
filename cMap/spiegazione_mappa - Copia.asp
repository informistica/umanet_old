<%@ Language=VBScript %>

<html>
<head>

   <style>
.loader {
display: block;
position: fixed;
left: 0px;
top: 0px;
width: 100%;
height: 100%;
z-index: 9999;
background: #fafafa url(../image/page-loader.gif) no-repeat center center;
text-align: center;
color: #999;
}
</style>

</head>
<body>
<div class="loader"></div>

<%
condivisione = Request.QueryString("condivisione")
super = Request.QueryString("super")

if (Session("CodiceAllievo")="" or Session("Id_Classe")="") and condivisione <> 1 then
Response.Redirect "../../home.asp"
end if

if 	Session("DB") <> Request.QueryString("DB") and (Session("CodiceAllievo") = "ospite" or Session("CodiceAllievo") = "") then
Response.Redirect "../../home.asp"
end if



if condivisione = 1 then
	Session("Id_Classe") = Request.QueryString("idclasse")
end if

'Recupero i vari elementi dell'indirizzo web

Dim Dominio, Pagina, Qstring

Dominio = Request.ServerVariables("SERVER_NAME")

Pagina = Request.ServerVariables("PATH_INFO")

Qstring = Request.ServerVariables("QUERY_STRING")


Dim Url

'Metto insieme dominio e percorso della pagina

Url = "https://" & Dominio & Pagina

'Verifico se esiste una querystring...

'Se esiste la aggiungo.

If Len(Qstring) > 0 Then

Url = Url & "?" & Qstring

End If

'Stampo tutto a video
Session("urlmappa") = ""
Session("urlmappa") = Url


%>
<%
function ReplaceCar(sInput)
dim sAns
  
  sAns = sInput
  'sAns1 = sInput
  
 sAns = Replace(sInput,chr(236),"i"&Chr(96))
 sAns = Replace(sAns,chr(237),"i"&Chr(96))
 sAns = Replace(sAns,chr(242),"o"&Chr(96))
 sAns = Replace(sAns,chr(243),"o"&Chr(96))
 sAns = Replace(sAns,chr(249),"u"&Chr(96))
 sAns = Replace(sAns,chr(250),"u"&Chr(96)) 
 sAns = Replace(sAns,chr(133),"a"&Chr(96))
 sAns = Replace(sAns,chr(138),"e'")
 sAns = Replace(sAns,"é","e'")
  sAns = Replace(sAns,chr(130),"e'")
 sAns = Replace(sAns, Chr(34), "'") 'sostituisco gli apici " con l'apice singolo
 sAns=  Replace(sAns,"'",Chr(96))  'sostituisco l'apice ' con quello storto per non disturbare la sintassi sql
 sAns=  Replace(sAns,chr(58),Chr(44))  'sostituisco : con , per non disturbare la creazione del file
 sAns=  Replace(sAns,"&","e") 
 sAns=  Replace(sAns,"/","-") 
 sAns=  Replace(sAns,"\","-") 
 sAns=  Replace(sAns,"?",".") 
 sAns=  Replace(sAns,"*","x") 
 sAns=  Replace(sAns,"<","_")
 sAns=  Replace(sAns,">","_") 
   sAns = Replace(sAns,"è","e'" )
   sAns=  Replace(sAns,"'",Chr(96))
   sAns=  Replace(sAns,"«",Chr(96))
   sAns=  Replace(sAns,"»",Chr(96))
   sAns=  Replace(sAns,"à","a'")
   sAns=  Replace(sAns,"ò","o'")
   sAns=  Replace(sAns,"ù","u'")
   sAns = Replace(sAns,"’","'")
   sAns = Replace(sAns,"“","'")
   sAns = Replace(sAns,"”","'")
   sAns = Replace(sAns, Chr(96), "'")
   sAns = Replace(sAns, "È", "E'")
   sAns = Replace(sAns, "ì", "i'")
   sAns = Replace(sAns, "–", "-")
   'sAns = Replace(sAns,VBCrlf,"<br>")
   
 
ReplaceCar = sAns

end function
%>

<% 

  Response.Buffer = true
  On Error Resume Next  
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->

 <%
	
	daQuaderno = Request.QueryString("daQuaderno")
	
	if daQuaderno <> 1 then daQuaderno = 0 end if
	
	if Session("CodiceAllievo") = "" and condivisione = 1 then
	CodiceAllievo = Request.QueryString("cod")
	QuerySQL = "SELECT * FROM Allievi WHERE CodiceAllievo = '"&CodiceAllievo&"'"
	set rsTab = ConnessioneDB.Execute(QuerySQL)
	Session("Id_Classe") = rsTab("Id_Classe")
	Session("ID_Materia") = Request.QueryString("Materia")
	
	else
	CodiceAllievo = Session("CodiceAllievo")
	end if
	
	
	
    QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
	Set rsTabellaCI = ConnessioneDB.Execute(QuerySQL) 
	CIAbilitato=rsTabellaCI("CIAbilitato") 
	rsTabellaCI.close
	' codice per permettere la visualizzazione solo delle proprie domande 
	QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	Privato=rsTabella("Privato") 
	Nodi=rsTabella("Nodi") 
	
	if daQuaderno = 1 then
	Nodi = 0
	end if
	
	rsTabella.close
	
  
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  Stato=Request.QueryString("Stato") 
  Stato0=Request.QueryString("Stato0")
  Codice_Test=Request.QueryString("CodiceTest")
  'File_Mappa="mappa_"&Codice_Test&".json" 
  File_Mappa=CodiceAllievo&".json"   
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  Nome=Request.QueryString("Nome")
  Cognome=Request.QueryString("Cognome")
  Cartella=Request.QueryString("Cartella")
  Codice_Test=Request.QueryString("CodiceTest") 
  
  if super = 1 then
  
  Cartelle = split(Cartella,",")
  Moduli = split(Modulo,",")
  Stato=1
  Stato0=1
  
  end if
  
  Dim objFSO, objTextFile
  
Dim sRead, sReadLine, sReadAll
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.CharSet = "UTF-8"
Set objFSO2 = CreateObject("Scripting.FileSystemObject")
url2=Server.MapPath(homesito & "/script/cMap/data/mappeutenti")&"/"& File_Mappa  
url2=Replace(url2,"\","/")
'response.write(url2)
Set objCreatedFile = objFSO2.CreateTextFile(url2, True)


riga="{""_comment"": ""Created with OWL2VOWL (version 0.3.3-SNAPSHOT), http://vowl.visualdataweb.org [Additional Information added by WebVOWL Exporter Version: 1.0.6]"","
objCreatedFile.WriteLine(riga)
riga="""header"": {"
objCreatedFile.WriteLine(riga)
riga="""languages"" : [ ""Italiano"" ],"
objCreatedFile.WriteLine(riga)
riga="""title"" : {"
objCreatedFile.WriteLine(riga)
riga="""undefined"" : ""Mappa concettuale"""
objCreatedFile.WriteLine(riga)
riga="},"
objCreatedFile.WriteLine(riga)
riga="""description"" : {"
objCreatedFile.WriteLine(riga)

if super <> 1 then
riga="""undefined"" : ""Capitolo '"&ReplaceCar(Capitolo)&"'"""
else

i=0
do while not i = UBound(Moduli)+1

	a = Split(Moduli(i),"_")
	IDModulo = a(0) & "_" & a(1)

	QuerySQL = "SELECT Titolo FROM Moduli WHERE ID_Mod = '"&IDModulo&"'"
	set rsMod = ConnessioneDB.Execute(QuerySQL)
	
	If InStr(Capitoli,rsMod("Titolo")) = 0 Then
	Capitoli = Capitoli&"'"&rsMod("Titolo")&"', "
	end if
i=i+1
loop

'response.write Capitoli

riga="""undefined"" : ""Capitolo/i "&ReplaceCar(Left(Capitoli,Len(Capitoli)-2))&""""

end if

objCreatedFile.WriteLine(riga)
riga="}"
objCreatedFile.WriteLine(riga)
riga="},"
objCreatedFile.WriteLine(riga)
riga="""metrics"": {"
objCreatedFile.WriteLine(riga)
' DA SISTEMARE****** count nodi e connessioni
riga="""classCount"": 22, "
objCreatedFile.WriteLine(riga)
riga="""objectPropertyCount"":22"
objCreatedFile.WriteLine(riga)
riga="},"
objCreatedFile.WriteLine(riga)
		 
costQuerySQL1="SELECT Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Allievi.Nome, Nodi.CodiceNodo, Moduli.ID_Mod,Nodi.Chi,Nodi.Cosa,Nodi.Dove,Nodi.Quando,Nodi.Come,Nodi.Perche,Nodi.Quindi,Nodi.Data,Paragrafi.Posizione, Nodi.Cartella, Nodi.NLink" &_
" FROM Paragrafi INNER JOIN (Moduli INNER JOIN (Allievi INNER JOIN Nodi ON Allievi.CodiceAllievo=Nodi.Id_Stud) ON Moduli.ID_Mod=Nodi.Id_Mod) ON Paragrafi.ID_Paragrafo=Nodi.Id_Arg" &_
" GROUP BY Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome,Allievi.Nome, Nodi.CodiceNodo, Moduli.ID_Mod,Nodi.Chi,Nodi.Cosa,Nodi.Dove,Nodi.Quando,Nodi.Come,Nodi.Perche,Nodi.Quindi,Nodi.Data,Nodi.In_Quiz,Paragrafi.Posizione,Nodi.Cartella,Nodi.NLink,Nodi.Segnalata "

costCountQuerySQL1="SELECT  count(*) " &_
" FROM Paragrafi INNER JOIN (Moduli INNER JOIN (Allievi INNER JOIN Nodi ON Allievi.CodiceAllievo=Nodi.Id_Stud) ON Moduli.ID_Mod=Nodi.Id_Mod) ON Paragrafi.ID_Paragrafo=Nodi.Id_Arg" &_
" GROUP BY Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome,Allievi.Nome, Nodi.CodiceNodo, Moduli.ID_Mod,Nodi.Chi,Nodi.Cosa,Nodi.Dove,Nodi.Quando,Nodi.Come,Nodi.Perche,Nodi.Quindi,Nodi.Data,Nodi.In_Quiz,Paragrafi.Posizione,Nodi.Cartella,Nodi.NLink,Nodi.Segnalata "


costQuerySQL2="SELECT Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome,Allievi.Nome, Nodi.CodiceNodo , Nodi.NLink, Moduli.ID_Mod,Nodi.Chi,Nodi.Cosa,Nodi.Dove,Nodi.Quando,Nodi.Come,Nodi.Perche,Nodi.Quindi,Nodi.Data,Paragrafi.Posizione,Nodi.Id_Stud,Nodi.Cartella" &_
" FROM Paragrafi INNER JOIN (Moduli INNER JOIN (Allievi INNER JOIN Nodi ON Allievi.CodiceAllievo=Nodi.Id_Stud) ON Moduli.ID_Mod=Nodi.Id_Mod) ON Paragrafi.ID_Paragrafo=Nodi.Id_Arg" &_
" GROUP BY Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome,Allievi.Nome, Nodi.CodiceNodo, Moduli.ID_Mod,Nodi.Chi,Nodi.Cosa,Nodi.Dove,Nodi.Quando,Nodi.Come,Nodi.Perche,Nodi.Quindi,Nodi.Data,Nodi.In_Quiz,Paragrafi.Posizione,Nodi.Id_Stud,Nodi.Cartella,Nodi.NLink,Nodi.Segnalata "

costCountQuerySQL2="SELECT count(*)" &_
" FROM Paragrafi INNER JOIN (Moduli INNER JOIN (Allievi INNER JOIN Nodi ON Allievi.CodiceAllievo=Nodi.Id_Stud) ON Moduli.ID_Mod=Nodi.Id_Mod) ON Paragrafi.ID_Paragrafo=Nodi.Id_Arg" &_
" GROUP BY Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome,Allievi.Nome, Nodi.CodiceNodo, Moduli.ID_Mod,Nodi.Chi,Nodi.Cosa,Nodi.Dove,Nodi.Quando,Nodi.Come,Nodi.Perche,Nodi.Quindi,Nodi.Data,Nodi.In_Quiz,Paragrafi.Posizione,Nodi.Id_Stud,Nodi.Cartella,Nodi.NLink,Nodi.Segnalata "


'if (cint(Stato)=0) or (cint(Stato0)=0) then  
if cint(Stato)=0 then
 'Definzione codice SQl della query per ricercare i nodi del paragrafo   
   if (Session("Admin")=True) or (Privato=0)  or (Nodi=1) then  'se vero visualizzo tutte i nodi del paragfrafo altrimenti solo quelle dello       studente loggato  
		QuerySQL=costQuerySQL1 &_
		" HAVING Paragrafi.ID_Paragrafo='" & Codice_Test & "' and Nodi.Chi<>'?' and Nodi.Segnalata=0   " &_   
		" ORDER BY Paragrafi.Posizione,Nodi.CodiceNodo;"
	else
		if condivisione = 1 then
	    QuerySQL=costQuerySQL2 &_
		" HAVING Paragrafi.ID_Paragrafo='" & Codice_Test & "' and Nodi.Chi<>'?' and Nodi.Segnalata=0  and Nodi.Id_Stud='"& CodiceAllievo &_   
		"' ORDER BY Paragrafi.Posizione,Nodi.CodiceNodo;"
		else
		QuerySQL=costQuerySQL2 &_
		" HAVING Paragrafi.ID_Paragrafo='" & Codice_Test & "' and Nodi.Chi<>'?' and Nodi.Segnalata=0  and Nodi.Id_Stud='"& Session("CodiceAllievo") &_   
		"' ORDER BY Paragrafi.Posizione,Nodi.CodiceNodo;"
		end if
	end if
else 

	if super <> 1 then
		 
			if (Session("Admin")=True) or (Privato=0)  or (Nodi=1) then  'se vero visualizzo tutte i nodi del paragfrafo altrimenti solo quelle dello       studente loggato
				QuerySQL= costQuerySQL1 &_
				" HAVING Moduli.ID_Mod='" & Modulo & "' and Nodi.Chi<>'?' and Nodi.Segnalata=0  " &_ 
				" ORDER BY Paragrafi.Posizione,Nodi.CodiceNodo;"
			else
				if condivisione = 1 then
				QuerySQL=costQuerySQL2 &_
				" HAVING Moduli.ID_Mod='" & Modulo & "' and Nodi.Chi<>'?' and Nodi.Segnalata=0   and Nodi.Id_Stud='"& CodiceAllievo &_   
				"' ORDER BY Paragrafi.Posizione,Nodi.CodiceNodo;"
				else
				QuerySQL=costQuerySQL2 &_
				" HAVING Moduli.ID_Mod='" & Modulo & "' and Nodi.Chi<>'?' and Nodi.Segnalata=0   and Nodi.Id_Stud='"& Session("CodiceAllievo") &_   
				"' ORDER BY Paragrafi.Posizione,Nodi.CodiceNodo;"
				end if
			end if
	else
			
			Dim Modulo1
			
			'response.write Modulo &"<br>"&Len(Modulo)&"<br>"
			
			i=0
			do while not i = UBound(Moduli)+1
				a = Split(Moduli(i),"_")
				'response.write UBound(a)
				
				if i = 0 then
					
					if UBound(a) = 2 then
						Modulo1 = "Paragrafi.ID_Paragrafo = '" & Moduli(i) & "'"
						'Modulo1 = Moduli(i)&"'"
					else
						Modulo1 = "Moduli.ID_Mod = '" & Moduli(i) & "'"
					end if
					
				else
					
					if UBound(a) = 2 then 
						Modulo1 = Modulo1&" OR Paragrafi.ID_Paragrafo='" & Moduli(i)  &"'"
					else
						Modulo1 = Modulo1&" OR Moduli.ID_Mod='" & Moduli(i)  &"'"
					end if
				end if
				
				i=i+1
				
			loop
			
			if (Session("Admin")=True) or (Privato=0)  or (Nodi=1) then  'se vero visualizzo tutte i nodi del paragfrafo altrimenti solo quelle dello       studente loggato
				QuerySQL= costQuerySQL1 &_
				" HAVING ("& Modulo1 & ") and Nodi.Chi<>'?' and Nodi.Segnalata=0  " &_ 
				" ORDER BY Paragrafi.Posizione,Nodi.CodiceNodo;"
			else
				QuerySQL=costQuerySQL2 &_
				" HAVING (" & Modulo1 & ") and Nodi.Chi<>'?' and Nodi.Segnalata=0 " &_   
				"' ORDER BY Paragrafi.Posizione,Nodi.CodiceNodo;"
			end if
	
	end if
	
end if  
  
  'Response.write QuerySQL
  
Set rsTabella = ConnessioneDB.Execute(QuerySQL)
cartella=rsTabella("Cartella")

inselect="" ' uno stringone che contiene i nodi della mappa attuale per i quali va cercata l'esistenza di link nella tabella link     
 If rsTabella.BOF=True And rsTabella.EOF=True Then 
		' mappa vuota
		response.write("mappa vuota")
  Else  'scrivo l'elenco dei nodi 
		riga="""class"": [ "
		objCreatedFile.WriteLine(riga)
		Do until rsTabella.EOF
			
			nlink = rsTabella("NLink")
			if IsNumeric(nlink) = False then
				nlink = 0
			end if
			
			riga="{"
			objCreatedFile.WriteLine(riga)
			riga="""id"":"""&rsTabella("CodiceNodo")&""","
			objCreatedFile.WriteLine(riga)
			inselect=inselect&rsTabella("CodiceNodo")

			if nlink = 0 then
				riga="""type"": ""owl:Nothing"""
			else if nlink >=8 or (nlink > 3 and nlink < 8) then
				riga="""type"": ""owl:equivalentClass"""
			else if nlink = 1 then
				riga="""type"": ""rdfs:Datatype"""
			else
				riga="""type"": ""owl:Class"""
			end if
			end if
			end if
			
			objCreatedFile.WriteLine(riga)
			
			rsTabella.MoveNext
			if rsTabella.eof then
			riga="}"
			else
			riga="},"
			inselect=inselect&","
			end if
			objCreatedFile.WriteLine(riga) 
		Loop 
		riga="],"
		objCreatedFile.WriteLine(riga)
				
		rsTabella.MoveFirst
		maxIdNodi=0
		inselect="("&inselect&")"
		'objCreatedFile.WriteLine(inselect)
    
		riga="""classAttribute"": ["
		objCreatedFile.WriteLine(riga)
		
		Do until rsTabella.EOF
			nlink = rsTabella("NLink")
		    if rsTabella("CodiceNodo")> maxIdNodi then
			   maxIdNodi=rsTabella("CodiceNodo")  ' mi serve per iniziare a numerare gli id degli archi che collegano i nodi
			end if
			riga="{"
		    objCreatedFile.WriteLine(riga)
			riga="""label"": {"
			objCreatedFile.WriteLine(riga)
			riga="""IRI-based"":"""&ReplaceCar(rsTabella("Chi"))&""""
			objCreatedFile.WriteLine(riga)
			riga="},"
			objCreatedFile.WriteLine(riga)
			riga="""comment"": {"
			objCreatedFile.WriteLine(riga)
			'devo prelevare la spiegazione nel file di testo
				ID=rsTabella("CodiceNodo")
				
				if super = 1 then
				QuerySQL = "SELECT ID_Mod,Cartella FROM Nodi WHERE CodiceNodo = "&ID
				set rsNodo = ConnessioneDB.Execute(QuerySQL)
				
				Modulo = rsNodo("ID_Mod")
				Cartella = rsNodo("Cartella")
				end if
				
				url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Nodi/"&Modulo&"_"&rsTabella(0)&"_"&ID&".txt"				
				url=Replace(url,"\","/")
				Set objTextFile = objFSO.OpenTextFile(url, ForReading)
				on error resume next
				 If Err.Number <> 0 Then
				    sReadAll="File della spiegazione mancante" & "<br>" & url
				 else
				    sReadAll = objTextFile.ReadAll
				    sReadAll=Replace(sReadAll, VBCrLf, "")
				    Err.Number = 0
				End If
				objTextFile.Close
				sReadAll=sReadAll&" (ID="&ID&")"&" (Autore="&rsTabella("Cognome")& " "& left(rsTabella("Nome"),1)&"."&")"
			riga="""undefined"": """&ReplaceCar(sReadAll) &""""
			objCreatedFile.WriteLine(riga)
			riga=" },"
			objCreatedFile.WriteLine(riga)
			
			if nlink >= 8 then
			riga="""attributes"": [ ""external"" ],"
			objCreatedFile.WriteLine(riga)	
			end if
			
			riga="""id"":"""&rsTabella("CodiceNodo")&""""
			objCreatedFile.WriteLine(riga)			
			rsTabella.MoveNext
			if rsTabella.eof  then
			riga="}"
			else
			riga="},"
			end if
			objCreatedFile.WriteLine(riga) 
		Loop 
		riga="],"
						
		
		QuerySql="Select count(*)"&_
		" FROM LINK_STUD WHERE Id_n1 in "&inselect&" or Id_n2 in "&inselect&";"
		'response.write(QuerySql)
		Set rsTabella = ConnessioneDB.Execute(QuerySQL)		
		
		numlink = rsTabella(0)
		
		if numlink = 0 then
		riga = Left(riga, Len(riga)-1)
		end if
		
		objCreatedFile.WriteLine(riga)
		
		QuerySql="Select ID_Link, Id_n1, L1, Id_n2, L2, Id_Stud,Testo2,Cognome,Nome"&_
		" FROM LINK_STUD WHERE Id_n1 in "&inselect&" or Id_n2 in "&inselect&";"
		'response.write(QuerySql)
		Set rsTabella = ConnessioneDB.Execute(QuerySQL)		
		
		if numlink <> 0 then
		
		stringalink = rsTabella("ID_Link")&","
		'Session("stringalink") = stringalink
		nProperty=maxIdNodi+1
		stringaproperty = nProperty & ","
		stringastud = rsTabella("Id_Stud") & ","
		'nProperty=cInt(rsTabella("Link.ID_Link"))
		riga="""property"": ["
		objCreatedFile.WriteLine(riga)
		
		Do until rsTabella.EOF	
			riga="{"
			objCreatedFile.WriteLine(riga)		
			riga="""id"":"""&nProperty&""","
			objCreatedFile.WriteLine(riga)
			
			if rsTabella("Id_Stud") <> Session("CodiceAllievo") then
				riga="""type"": ""owl:datatypeProperty"""
			else
				riga="""type"": ""rdf:Property"""
			end if
			objCreatedFile.WriteLine(riga)			
			rsTabella.MoveNext
			nProperty=nProperty+1
			stringalink = stringalink & rsTabella("ID_Link")&","
			stringaproperty = stringaproperty & nProperty & ","
			stringastud = stringastud & rsTabella("Id_Stud") & ","
			'nProperty=cInt(rsTabella("Link.ID_Link"))
			if rsTabella.eof then
			riga="}"
			else
			riga="},"
			end if
			objCreatedFile.WriteLine(riga) 
		Loop 
		riga="],"
		objCreatedFile.WriteLine(riga)	
		Session("stringalink") = stringalink
		Session("stringaproperty") = stringaproperty
		Session("stringastud") = stringastud
		
		nProperty=maxIdNodi+1 ' riporto l'indice al valore iniziale
		'nProperty=cInt(rsTabella("Link.ID_Link"))
	    rsTabella.MoveFirst
		riga="""propertyAttribute"": ["
		objCreatedFile.WriteLine(riga)
		
		Do until rsTabella.EOF		
			riga="{"
			objCreatedFile.WriteLine(riga)
			riga="""range"":"""&rsTabella("Id_n2")&""","
			objCreatedFile.WriteLine(riga)
			riga="""label"": {"
			objCreatedFile.WriteLine(riga)
			riga="""IRI-based"":"""&ReplaceCar(rsTabella("Testo2"))&""""
			objCreatedFile.WriteLine(riga)
			riga="},"
			objCreatedFile.WriteLine(riga)
			riga="""domain"":"""&rsTabella("Id_n1")&""","
			objCreatedFile.WriteLine(riga)
			
			
			riga="""comment"": {"
			objCreatedFile.WriteLine(riga)
			riga="""undefined"": """&" (Autore="&rsTabella("Cognome")& " "& left(rsTabella("Nome"),1)&"."&")" &""""
			objCreatedFile.WriteLine(riga)
			riga="},"
			objCreatedFile.WriteLine(riga)
			
					
			riga="""attributes"": ["
			objCreatedFile.WriteLine(riga)
			riga="""object"""
			objCreatedFile.WriteLine(riga)
			riga="],"
			objCreatedFile.WriteLine(riga)
			
			riga="""id"":"""&nProperty&""""
			objCreatedFile.WriteLine(riga)		
			rsTabella.MoveNext
			nProperty=nProperty+1
			'nProperty=cInt(rsTabella("Link.ID_Link"))
			if rsTabella.eof then
			riga="}"
			else
			riga="},"
			end if
			objCreatedFile.WriteLine(riga) 
		Loop 
		riga="]"
		objCreatedFile.WriteLine(riga)	
		
	end if

	end if
	
	riga="}"
	objCreatedFile.WriteLine(riga)
    objCreatedFile.Close
    rsTabella.Close : Set rsTabella = Nothing  ' libera le risorse chiudendo gli oggetti aperti 
    ConnessioneDB.Close : Set ConnessioneDB = Nothing 
	'response.Redirect "index.html#file="&File_Mappa ' da implementare passando il nome della mappa come parametro ad webvowl.app.js
	
	Dim fromURL
	fromURL = Request.ServerVariables("HTTP_REFERER")
	'if fromURL = 
	'response.redirect fromURL
	fromURLs = split(fromURL, "?")
	
	url1 = "https://www.umanetexpo.net/expo2015Server/UECDL/script/cMap/index.asp"
	url2 = "https://www.umanet.net/expo2015Server/UECDL/script/cMap/index.asp"
	url3 = "https://www.elexpo.net/expo2015Server/UECDL/script/cMap/index.asp"
	
	if fromURLs(0) = url1 or fromURLs(0) = url2 or fromURLs(0) = url3 then
	collegamento = 1
	else
	collegamento = 0
	end if
	
	collegamento = 0
	
	'response.Redirect "index.asp?cod="&CodiceAllievo&"&collegamento="&collegamento
	if condivisione <> 1 then 
		urlr = "index.asp?cod="&CodiceAllievo&"&collegamento="&collegamento
	else 
		urlr = "index.asp?cod="&CodiceAllievo&"&collegamento="&collegamento&"&condivisione=1"
	end if
 %>
 
 
	<script>
	localStorage.setItem("dabutton",1); //serve per evitare di capire che sia un refresh
	window.location.href = "<%=urlr%>";
 </script>
 
   </body>
   
   </html>
    

