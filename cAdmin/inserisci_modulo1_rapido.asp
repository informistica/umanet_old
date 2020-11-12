<!-- calcola_risultato_MODBC3.asp -->
<html>
<head>	
	<link rel="stylesheet" type="text/css" href="../../stile.css">
	<style>
	<!--
	 li.MsoNormal
		{mso-style-parent:"";
		margin-bottom:.0001pt;
		font-size:12.0pt;
		font-family:"Times New Roman";
		margin-left:0cm; margin-right:0cm; margin-top:0cm}
	-->
	</style>
   
  
   
</head>
<body>
    <div id="container">
	<div class="risultati_test" >
	<font color=#FF0000 size="4">

<%@ Language=VBScript %>
<!-- #include file = "../extra/test_server.asp" -->
     
<% Response.Buffer=True 
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,RecSet,k
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome
   Dim cartelle(4)
   cartelle(0)="_Domande"
   cartelle(1)="_Frasi"
   cartelle(2)="_Nodi"
   cartelle(3)="_Spiegazioni"
   cartelle(4)="_Esercizi"
   
   'Apertura della connessione al database  
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
       %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
     <%  
 
Num=Request.QueryString("Num")

Classe= Request.QueryString("classe")
'Classe= Session("Cartella")
if Classe="" then
  Classe=Session("classe")
end if
Titolo=Request.form("txtTitolo")
ID_Mod=Request.form("txtID_Mod") 
'ID_Mod=Classe&Request.QueryString("ID_Mod") 

Id_Classe=Request.QueryString("Id_Classe")
  

function ReplaceCar(sInput)
dim sAns
  sAns = Replace(sInput,chr(224),"a"&Chr(96))
  sAns = Replace(sAns,chr(225),"a"&Chr(96))
  sAns = Replace(sAns,chr(232),"e"&Chr(96))
  sAns = Replace(sAns,chr(233),"e"&Chr(96))
  sAns = Replace(sAns,chr(236),"i"&Chr(96))
  sAns = Replace(sAns,chr(237),"i"&Chr(96))
  sAns = Replace(sAns,chr(242),"o"&Chr(96))
  sAns = Replace(sAns,chr(243),"o"&Chr(96))
  sAns = Replace(sAns,chr(249),"u"&Chr(96))
  sAns = Replace(sAns,chr(250),"u"&Chr(96)) 
  sAns = Replace(sAns, Chr(34), "'")' sostituisco gli apici " con l'apice singolo
  sAns=  Replace(sAns,"'",Chr(96)) ' sostituisco l'apice ' con quello storto per non disturbare la sintassi sql
  sAns=  Replace(sAns,chr(58),Chr(44)) ' sostituisco : con , per non disturbare la creazione del file
  sAns=  Replace(sAns,"&","e") 
  sAns=  Replace(sAns,"/","-") 
  sAns=  Replace(sAns,"\","-") 
  sAns=  Replace(sAns,"?",".") 
  sAns=  Replace(sAns,"*","x") 
  sAns=  Replace(sAns,"<","_")
  sAns=  Replace(sAns,">","_") 
  
  ReplaceCar = sAns
end function


  
	   
	Titolo = Replace(Titolo, Chr(34), "'")' sostituisco gli apici " con l'apice singolo
	Titolo=  Replace(Titolo,"'",Chr(96)) ' sostituisco l'apice ' con quello storto per non disturbare la sintassi sql	    
    Titolo=  ReplaceCar(Titolo)
  
  QuerySQL="SELECT Cartella FROM Classi where ID_Classe='"&Id_Classe&"';"
	  set rsTabella1=ConnessioneDB.Execute(QuerySQL)  
	  cartella=rsTabella1(0)


  QuerySQL="SELECT max(posizione) FROM MODULI_NOT_UMANET where Cartella='"&cartella&"';"
	  set rsTabella1=ConnessioneDB.Execute(QuerySQL)  
	  if isnull(rsTabella1(0)) then
	    maxPos=0
	  else
	      maxPos=rsTabella1(0)
	  end if
	  posizione=maxPos+1
 
 
  
  QuerySQL="  INSERT INTO Moduli (ID_Mod,Titolo,Cartella,Posizione,URL_OL,Visibile)  SELECT '" & ID_Mod & "','" & rtrim(Titolo) & "', '" & Classe & "', " & posizione &",'"&urlCopertina&"',1;"  
   response.write(QuerySQL&"<br>") 
   ConnessioneDB.Execute QuerySQL 
 
    Set fso = CreateObject("Scripting.FileSystemObject") 
    for i=0 to 4
		url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&cartella &"/"&Id_Mod&cartelle(i) 
		url=Replace(url,"\","/")
		 response.Write("<br>"&url)
		if fso.FolderExists (url) then
			 response.Write( "La cartella " & url & " esiste gi�.<br>")
		else
			fso.CreateFolder (url) 
			fso.CreateFolder (url&"/Img")
			response.Write( "La cartella " & url&"/Img" & " � stata creata.<br>") 
		end if
    next 
	' creo la cartella per il modulo dentro la cartella risorse del corso  
	    url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&cartella &"/Risorse/Mod_"&right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_")) 
		url=Replace(url,"\","/")
		
	    if fso.FolderExists (url) then
			 response.Write( "La cartella " & url & " esiste gi�.<br>")
		else
		    response.Write( "Creazione della cartella :" & url&"/img" & "....") 
			fso.CreateFolder (url) 
			fso.CreateFolder (url&"/img")
	     	response.Write( "eseguita<br>") 
		 
		end if
	
 'response.write(url)





strText = Request.Form("MyTextArea")
arrLines = Split(strText, vbCrLf)
k=1
'response.write("Numero elementi-1"&ubound(arrLines))
For Each strLine in arrLines
  s=split(strLine," ")
  titolo=""
   For i=1 to ubound(s)
    titolo=titolo&s(i)&" "
  next
   QuerySQL="  INSERT INTO Paragrafi (ID_Paragrafo, Titolo,Posizione,URL_L,URL_O)  SELECT '" &ID_Mod&"_"&k  & "','" & rtrim(titolo) & "','" & k & "','','"& s(0) &"';"
				   ConnessioneDB.Execute QuerySQL 
				   response.Write(QuerySQL&"<br>")
		   QuerySQL="  INSERT INTO Classi_Moduli_Paragrafi (ID_Classe, Id_Modulo,Id_Paragrafo)  SELECT '" & Id_Classe  & "','" & ID_Mod & "', '" & ID_Mod&"_"&k & "';"
				  ConnessioneDB.Execute QuerySQL 
		response.Write(QuerySQL&"<br>")
 ' response.write("<br>url "&k&"="&s(0))
 ' response.write("<br>par "&k&"="&s(1))
  k=k+1
next 

 


	 

 
	On Error Resume Next
	If Err.Number = 0 Then
		Response.Write "Inserimento del modulo avvenuta ! "
	Else
		Response.Write Err.Description 
		Err.Number = 0
	End If

   


   %>
		  
          
<div id=piede_pagina>
				<p><p>
				
				<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
<h3><a href="admin.asp?Id_Classe=<%=Id_Classe%>&divid=<%=divid%>"> Torna alla pagina Amministratore... </a></h3> 
	
  
			</div>
 <!-- se il login � corretto richima la pagina per inserire le domande del test -->

	
	</body>
	</html>
	