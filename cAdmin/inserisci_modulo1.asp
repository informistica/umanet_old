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
   <!-- #include file = "../service/replacecar.asp" -->
     <%  
 
Num=Request.QueryString("Num")
Titolo=Request.QueryString("Titolo")
Classe= Request.QueryString("classe")
cartella= Request.QueryString("cartella")
 umanet= Request("umanet")
'Classe= Session("Cartella")
if Classe="" then
  Classe=Session("classe")
end if
ID_Mod=Request.QueryString("ID_Mod") 
sottoparagrafi=Request.QueryString("sottoparagrafi")

posizione=Request("posizione") 
if posizione="" then 
posizione=1
end if
'ID_Mod=Classe&Request.QueryString("ID_Mod") 

Id_Classe=Request.QueryString("Id_Classe")
 
inserito=Request.QueryString("inserito")
urlCopertina=right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_"))&"_0.html"
if request("txtURLMod")<>"" then
urlCopertina=request("txtURLMod")
end if

 


  if inserito="" then ' se � la prima chiamata mostro modulo senza risorsa
	   
	   Titolo = Replace(Titolo, Chr(34), "'")' sostituisco gli apici " con l'apice singolo
	   Titolo=  Replace(Titolo,"'",Chr(96)) ' sostituisco l'apice ' con quello storto per non disturbare la sintassi sql
	    
   Titolo=  ReplaceCar(Titolo)
  
 '  QuerySQL="  INSERT INTO Moduli (ID_Mod,Titolo,Cartella,Posizione,URL_OL)  SELECT '" & ID_Mod & "','" & rtrim(Titolo) & "', '" & Classe & "', " & cint(right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_"))) &",'"&urlCopertina&"';" ' dave errore nel cint
  if umanet<>"" then
     QuerySQL="SELECT max(posizione) FROM MODULI_UMANET1 where Cartella='"&cartella&"';"
  else
   QuerySQL="SELECT max(posizione) FROM MODULI_NOT_UMANET where Cartella='"&cartella&"';"
end if
	  set rsTabella1=ConnessioneDB.Execute(QuerySQL)  
	  if isnull(rsTabella1(0)) then
	    maxPos=0
	  else
	      maxPos=rsTabella1(0)
	  end if
	  posizione=maxPos+1
  
  
  QuerySQL="  INSERT INTO Moduli (ID_Mod,Titolo,Cartella,Posizione,URL_OL, Visibile)  SELECT '" & ID_Mod & "','" & rtrim(Titolo) & "', '" & cartella & "', " & posizione &",'"&urlCopertina&"',1;"  
   response.write(QuerySQL) 
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

		' creo la cartella per il modulo dentro la cartella verifiche del corso  
	    url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&cartella &"/Verifiche/Mod_"&right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_")) 
		url=Replace(url,"\","/")
		
	    if fso.FolderExists (url) then
			 response.Write( "La cartella " & url & " esiste gi�.<br>")
		else
		    response.Write( "Creazione della cartella :" & url) 
			fso.CreateFolder (url) 
			 
	     	response.Write( "eseguita<br>") 
		 
		end if
	
 'response.write(url)
 cont=0
 for k=1 to Num
	   Domanda = Request.Form("txtDomanda"&k)   

	  
   Domanda = ReplaceCar(Domanda)
	if Domanda<>"" then ' controllo per le righe vuote 
			
			'Esecuzione della query per  
	
	if Request.Form("txtUrl"&k)<>"" then
	       ulrRisorsa1= Request.Form("txtUrl"&k) 
	else
					
			urlRis=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&cartella &"/Risorse/Mod_"&right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_"))&"/"
			ulrRisorsa1=right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_"))&"_"&k&".html"
			ulrRisorsa=urlRis&ulrRisorsa1
			ulrRisorsa=Replace(ulrRisorsa,"\","/")
	end if ' if Request.Form("txtUrl"&k)<>"" then	
			
		   QuerySQL="  INSERT INTO Paragrafi (ID_Paragrafo, Titolo,Posizione,URL_L,URL_O)  SELECT '" & Request.Form("txtId"&k)  & "','" & rtrim(Domanda) & "','" & k & "','"& ulrRisorsa1 &"','"& ulrRisorsa1 &"';"
				   ConnessioneDB.Execute QuerySQL 
				   response.Write(QuerySQL&"<br>")
		   QuerySQL="  INSERT INTO Classi_Moduli_Paragrafi (ID_Classe, Id_Modulo,Id_Paragrafo)  SELECT '" & Id_Classe  & "','" & ID_Mod & "', '" & Request.Form("txtId"&k) & "';"
				   ConnessioneDB.Execute QuerySQL 
			response.Write(QuerySQL&"<br>")
		 ' qui vedo se devo inserire i sottoparagrafi
		 
			txtSottoparagrafi = Request.Form("txtSottoparagrafi"&k)
			if  Request.Form("txtSottoparagrafi"&k) <>"" then
			strText = Request.Form("txtSottoparagrafi"&k) 
			arrLines = Split(strText, vbCrLf)
			j=1
			For Each strLine in arrLines	      
			 Sottoparagrafo=strLine 
			 ID_Sottoparagrafo=Request.Form("txtId"&k)&"_"&j
			  QuerySQL="  INSERT INTO Sottoparagrafi (ID_Sottoparagrafo, Titolo,Posizione) SELECT '" & ID_Sottoparagrafo  & "','" & ReplaceCar(Sottoparagrafo) & "'," & j & ";"
				   response.Write(QuerySQL&"<br>" )
				   ConnessioneDB.Execute QuerySQL 
				    
			  QuerySQL="  INSERT INTO ParagrafiSottoparagrafi (Id_Paragrafo,Id_Sottoparagrafo) SELECT '" & Request.Form("txtId"&k)  & "','" & ID_Sottoparagrafo & "';"
				   response.Write(QuerySQL&"<br>" )
				   ConnessioneDB.Execute QuerySQL    
				    
			  j=j+1
			Next
			
			txtUrlSottoparagrafi = Request.Form("txtUrlSottoparagrafi"&k)
			if  Request.Form("txtUrlSottoparagrafi"&k) <>"" then
				strText = Request.Form("txtUrlSottoparagrafi"&k) 
				arrLines = Split(strText, vbCrLf)
				j=1
				For Each strLine in arrLines	      
				 UrlSottoparagrafo=strLine 
				 ID_Sottoparagrafo=Request.Form("txtId"&k)&"_"&j
				 
				 QuerySQL="UPDATE Sottoparagrafi SET Url ='"&  UrlSottoparagrafo&"' where ID_Sottoparagrafo='"&ID_Sottoparagrafo&"'"
				 response.Write(QuerySQL&"<br>" )
					ConnessioneDB.Execute(QuerySQL)
					
				  j=j+1
				Next
			end if 
			
			
			end if
			
			
			' qua genero i file html delle risorse con le immagini cre andranno salvate in img 
			
if Request.Form("txtUrl"&k)="" then			
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
	
   end if ' if Domanda<>"" then 
 end if
next 


' pagina di copertina
    urlRis=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&cartella&"/Risorse/Mod_"&right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_"))&"/"
    ulrRisorsa1=right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_"))&"_0.html"
	ulrRisorsa=urlRis&ulrRisorsa1
	ulrRisorsa=Replace(ulrRisorsa,"\","/")
	
Set objCreatedFile = fso.CreateTextFile(ulrRisorsa, True)
	
	objCreatedFile.WriteLine("<html>")
	objCreatedFile.WriteLine("<head>")
	objCreatedFile.WriteLine("<link href='../../../../../stile.css' rel='stylesheet' type='text/css' />")
	objCreatedFile.WriteLine("<title>")
	objCreatedFile.WriteLine("0) Indice Argomenti")
	objCreatedFile.WriteLine("</title>")
	objCreatedFile.WriteLine("<body>")
	objCreatedFile.WriteLine("	<div id='container' class='contenuti_login'>")
	objCreatedFile.WriteLine("		<div id='bloc_destra_cont'>")
	objCreatedFile.WriteLine("			<div class='contenuti'><div>  ")
	objCreatedFile.WriteLine("				<font size='+3'>Indice argomenti</font>")
	objCreatedFile.WriteLine("			</div><br><br>")
	objCreatedFile.WriteLine("	<p align=center>")
	' per tutte e 9 le possibili pagine cedo quelle da inserire
	      if (Request.Form("txtPg0")<>"") then
			objCreatedFile.WriteLine("	<img class='imground' src='img/"&Request.Form("txtPg0")&".jpg'>")
			objCreatedFile.WriteLine("	<br>")
			if (Request.Form("txtPg1")<>"") then
			objCreatedFile.WriteLine("	<img class='imground' src='img/"&Request.Form("txtPg1")&".jpg'>")
			objCreatedFile.WriteLine("	<br>")
			end if 
			if (Request.Form("txtPg2")<>"") then
			objCreatedFile.WriteLine("	<img class='imground' src='img/"&Request.Form("txtPg2")&".jpg'>")
			objCreatedFile.WriteLine("	<br>")
			end	if	
	      else
		     objCreatedFile.WriteLine("	Copertina del capitolo assente")
			objCreatedFile.WriteLine("	<br>")
		  end if
	objCreatedFile.WriteLine("		</div>")
	objCreatedFile.WriteLine("	  </p>")
	objCreatedFile.WriteLine("  </div>")
	objCreatedFile.WriteLine("</body>")
	objCreatedFile.WriteLine("</html>")


	'On Error Resume Next
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
<input type="Submit" name="btnUpload" value="Upload" onClick="mostra();"> 
 <br><img src="Upload/nulla.jpg" width="35" height="35" name="loading">
 
<% end if%>
		  
          
<div id=piede_pagina>
				<p><p>
				
				<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
<h3><a href="admin.asp?Id_Classe=<%=Id_Classe%>&divid=<%=divid%>"> Torna alla pagina Amministratore... </a></h3> 
	
  
			</div>
 <!-- se il login � corretto richima la pagina per inserire le domande del test -->

	
	</body>
	</html>
	