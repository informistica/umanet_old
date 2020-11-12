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
   
   <script language="javascript"> 
    function mostra() {
		document["loading"].src = "Upload/loading3.gif";	 
		}
	</script> 
   
</head>
<body>
    <div id="container">
	<div class="risultati_test" >
	<font color=#FF0000 size="4">

<%@ Language=VBScript %>
<!-- #include file = "../extra/test_server.asp" -->
  <!-- #include file = "../service/replacecar.asp" -->
     
<% Response.Buffer=True 
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,RecSet,k
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome
   Dim cartelle(5)
   cartelle(0)="_Domande"
   cartelle(1)="_Frasi"
   cartelle(2)="_Nodi"
   cartelle(3)="_Spiegazioni"
   cartelle(4)="_Metafore"
   
   'Apertura della connessione al database  
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
       %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
     <%  
ID_Mod=Request.QueryString("ID_Mod")  
Num=Request.QueryString("Num")
Titolo=Request.QueryString("Titolo")
Classe= Request.QueryString("Classe")
Id_Classe=Request.QueryString("Id_Classe")
divid=Request.QueryString("divid")
inserito=Request.QueryString("inserito")


  if inserito="" then ' se è la prima chiamata mostro modulo senza risorsa
	   
	   Titolo = Replace(Titolo, Chr(34), "'")' sostituisco gli apici " con l'apice singolo
	   Titolo=  Replace(Titolo,"'",Chr(96)) ' sostituisco l'apice ' con quello storto per non disturbare la sintassi sql
	   Titolo=  Replace(Titolo,Chr(39),Chr(96)) ' come sopra ma indico ' con il suo codice
 
    Titolo=  ReplaceCar(Titolo)
  
   QuerySQL="  INSERT INTO Moduli (ID_Mod,Titolo,Cartella,Posizione)  SELECT '" & ID_Mod & "','" & rtrim(Titolo) & "', '" & Classe & "', " & cint(right(Id_Mod,1)) &";"   
   ConnessioneDB.Execute QuerySQL 
 
    Set fso = CreateObject("Scripting.FileSystemObject") 
    for i=0 to 3 
		url=Server.MapPath(homesito)& "/Db"&Session("DB")& "/Materie/"&Session("ID_Materia")&"/"&Classe&"/"&Id_Mod&cartelle(i) 
		url=Replace(url,"\","/")
		if fso.FolderExists (url) then
			 response.Write( "La cartella " & url & " esiste già.<br>")
		else
			fso.CreateFolder (url) 
			fso.CreateFolder (url&"/Img")
			response.Write( "La cartella " & url&"/Img" & " è stata creata.<br>") 
		end if
    next 
	' creo la cartella per il modulo dentro la cartella risorse del corso  
	    url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&Classe&"/Risorse/Mod_"&cint(right(Id_Mod,1))
		url=Replace(url,"\","/")
		
	    if fso.FolderExists (url) then
			 response.Write( "La cartella " & url & " esiste già.<br>")
		else
		    response.Write( "Creazione della cartella :" & url&"/img" & "....<br>") 
			fso.CreateFolder (url) 
			fso.CreateFolder (url&"/img")
			response.Write( "La cartella " & url & " è stata creata.<br>") 
			response.Write( "La cartella " & url&"/img" & " è stata creata.<br>") 
		end if
		
		' creo file delle risorse comune a TUTTE LE CLASSI
		urlRis=Server.MapPath(homesito)& "/U-ECDL/Risorse/"&Titolo
		urlRisorsa=urlRis
	    urlRisorsa=Replace(urlRis,"\","/")
        if fso.FolderExists (urlRisorsa) then
			 response.Write( "La cartella " & urlRisorsa & " esiste già.<br>")
		else
		    fso.CreateFolder (urlRisorsa) 
			fso.CreateFolder (urlRisorsa&"/img")
			response.Write( "La cartella " & urlRisorsa & " è stata creata.<br>") 
			response.Write( "La cartella " & urlRisorsa&"/img" & " è stata creata.<br>")
	    end if

	
 'response.write(url)
 cont=0
 for k=1 to Num
	   Domanda = Request.Form("txtDomanda"&k)   
	   Domanda = Replace(Domanda, Chr(34), "'")' sostituisco gli apici " con l'apice singolo
	   Domanda=  Replace(Domanda,"'",Chr(96)) ' sostituisco l'apice ' con quello storto per non disturbare la sintassi sql
  	   Domanda=  Replace(Domanda,chr(58),Chr(44)) ' sostituisco : con , per non disturbare la creazione del file
  
	if Domanda<>"" then ' controllo per le righe vuote 
			
		urlRis=Server.MapPath(homesito)& "/U-ECDL/Risorse/"&Titolo&"/"
		ulrRisorsa=urlRis
	    ulrRisorsa=Replace(ulrRis,"\","/")
		ulrRisorsa1=Titolo&"_"&k&".html"
		ulrRisorsa=urlRis&ulrRisorsa1
		ulrRisorsa=Replace(ulrRisorsa,"\","/")
		
			
			
			'Esecuzione della query per  
 

   QuerySQL="  INSERT INTO Paragrafi (ID_Paragrafo, Titolo,Posizione)  SELECT '" & Request.Form("txtId"&k)  & "','" & ReplaceCar(rtrim(Request.Form("txtDomanda"&k))) & "', '" & k & "';"
		   ConnessioneDB.Execute QuerySQL 
		  ' response.Write(QuerySQL&"<br>")
   QuerySQL="  INSERT INTO Classi_Moduli_Paragrafi (ID_Classe, Id_Modulo,Id_Paragrafo)  SELECT '" & Id_Classe  & "','" & ID_Mod & "', '" & Request.Form("txtId"&k) & "';"
		   ConnessioneDB.Execute QuerySQL 
	'response.Write(QuerySQL&"<br>")
  
  
  ' qua genero i file html delle risorse con le immagini cre andranno salvate in img 
	
	
	Set objCreatedFile = fso.CreateTextFile(ulrRisorsa, True)
	
	objCreatedFile.WriteLine("<html>")
	objCreatedFile.WriteLine("<head>")
	objCreatedFile.WriteLine("<link href='../../../stile.css' rel='stylesheet' type='text/css' />")
	objCreatedFile.WriteLine("<title>")
	objCreatedFile.WriteLine(k&") "& Request.Form("txtDomanda"&k))
	objCreatedFile.WriteLine("</title>")
	objCreatedFile.WriteLine("<body>")
	objCreatedFile.WriteLine("	<div id='container' class='contenuti_login'>")
	objCreatedFile.WriteLine("		<div id='bloc_destra_cont'>")
	objCreatedFile.WriteLine("			<div class='contenuti'><div>  ")
	objCreatedFile.WriteLine("				<font size='+3'>"&Titolo &"</font>")
	objCreatedFile.WriteLine("			</div><br><br>")
	objCreatedFile.WriteLine("	<p align=center>")
	' per tutte e 9 le possibili pagine cedo quelle da inserire
	for j=1 to 13
	   if Request.Form("txtPg"&k&"_"&j)<>"" then
			objCreatedFile.WriteLine("	<img class='imground' src='img/"&Request.Form("txtPg"&k&"_"&j)&".jpg'>")
			objCreatedFile.WriteLine("	<br>")
	   end if
	next
	objCreatedFile.WriteLine("		</div>")
	objCreatedFile.WriteLine("	  </p>")
	objCreatedFile.WriteLine("  </div>")
	objCreatedFile.WriteLine("</body>")
	objCreatedFile.WriteLine("</html>")
  
  
  
  
  
    end if 
  
next 

'	On Error Resume Next
	If Err.Number = 0 Then
		Response.Write "Inserimento del modulo avvenuto! "
	Else
		Response.Write Err.Description 
		Err.Number = 0
	End If
    Session("Classe")=Classe
	Session("Id_Mod")=Id_Mod

   %>
	</font>   
	<FORM name="frmDocument" METHOD="Post" ENCTYPE="multipart/form-data" ACTION="Upload/confirm_update.asp?AggRisMod=1&Classe=<%=Classe%>&Id_Mod=<%=ID_Mod%>&Id_Classe=<%=Id_Classe%>&divid=<%=divid%>">
<p align="center"><font face="Verdana" size="2">
Classe : <input type="text" value="<%=Classe%>"><br>
Modulo : <input type="text" name="txtId_Mod "value="<%=ID_Mod%>"><br>
Risorsa : <input type="text" name="txtRis "value="Nessuna"><br><br>
Aggiungi una Risorsa : <INPUT TYPE="file" name="flname" ><BR><br>
 <input type="Submit" name="btnUpload" value="Upload" onClick="mostra()"> 
  <br><img src="Upload/nulla.jpg" width="35" height="35" name="loading">
            
</font>
</FORM>
<%else
	On Error Resume Next
	If Err.Number = 0 Then
		Response.Write "Inserimento della risorsa di modulo avvenuta ! "
	Else
		Response.Write Err.Description 
		Err.Number = 0
	End If

   Dim esecuzione
   set esecuzione = New TestServer ' oggetto di classe per testare dove gira il sito
   QuerySQL=" Select URL,URL_OL from Moduli where Id_Mod='"&ID_Mod&"';"
   set rsTabella=ConnessioneDB.Execute (QuerySQL) 
   if esecuzione.locale=1 then
			url=rsTabella("URL")			 
	  else 		
	        url=rsTabella("URL_OL")
	 end if  
   set esecuzione=nothing
   set rsTabella=nothing



  %>
  </font>
   <FORM name="frmDocument" METHOD="Post" ENCTYPE="multipart/form-data" ACTION="Upload/confirm_update.asp?AggRisMod=1&Classe=<%=Classe%>&Id_Mod=<%=ID_Mod%>&Id_Classe=<%=Id_Classe%>&divid=<%=divid%>">
<p align="center"><font face="Verdana" size="2">
Classe : <input type="text" value="<%=Classe%>"><br>
Modulo : <input type="text" name="txtId_Mod "value="<%=ID_Mod%>"><br>
Risorsa : <input type="text" name="txtRis "value="<%=url%>"><br><br>
Sostituisci la Risorsa : <INPUT TYPE="file" name="flname" ><BR><br>
<input type="Submit" name="btnUpload" value="Upload" onClick="mostra()"> 
 <br><img src="Upload/nulla.jpg" width="35" height="35" name="loading">
 
<% end if%>
		  
          
<div id=piede_pagina>
				<p><p>
				
				<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
<h3><a href="../cClasse/home_app.asp?id_classe=<%=Id_Classe%>&divid=<%=divid%>"> Torna alla pagina Apprendimento... </a></h3> 
	
  
			</div>
 <!-- se il login è corretto richima la pagina per inserire le domande del test -->

	
	</body>
	</html>
	