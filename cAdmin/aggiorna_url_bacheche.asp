 
<html>
<head>	
	<link rel="stylesheet" type="text/css" href="../stile.css">
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
   <% Response.Buffer=True 
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,RecSet
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome
 
   'Apertura della connessione al database  
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
	 Set ConnessioneDB1 = Server.CreateObject("ADODB.Connection") ' per il forum
       %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
   <!-- #include file = "stringa_connessione_forum.inc" -->
 
<%  


QuerySQL="SELECT Cognome,Nome,CodiceAllievo" &_
" FROM Allievi  " &_
" WHERE Id_Classe ='" & Session("Id_Classe") & "'" &_
" ORDER BY Allievi.Cognome Asc; "
Set rsTabella = ConnessioneDB.Execute(QuerySQL) 

' on error resume next
  classe=Request.QueryString("classe")
  id_classe=Request.QueryString("id_classe")
 
  
  txtDomande = Request.Form("txtAzione")		  
  strText = txtDomande
  arrLines = Split(strText, vbCrLf)
    
	k=1
	For Each strLine in arrLines    
		  urlInc=strLine
		  messaggio= "<br><br> <a target=blank href=" & urlInc & "> Apri a TUTTO schermo</a><BR><br>"
		  messaggio=messaggio & "<iframe src="& urlInc &" width= 640  height= 480 ></iframe><BR><BR>" 	 
	  
	   if not rsTabella.eof then   
			  sName =  trim(rsTabella("Cognome")) & " " & left(rsTabella("Nome"),1)&"."  
			  scodAllievo=rsTabella("CodiceAllievo") 
			  sIdClasse=Session("Id_Classe")
			  sTopic = Request.form("txtPostit") 
			  sUrlimg=""
			  sUrlfile=""
			  sBacheca=rsTabella("CodiceAllievo")
			  sComments = messaggio
			  privato=0			 
				 
				  QuerySQL=" INSERT INTO FORUM_MESSAGES (AUTHORNAME,CODICEALLIEVO,ID_CLASSE,TOPIC,URLIMG,URLFILE,BACHECA,COMMENTS,Privato)  SELECT '" & sName & "','" & scodAllievo & "', '" & sIdClasse & "','" & sTopic & "','" & sUrlimg & "', '" & sUrlfile & "','" & sBacheca & "','" & sComments & "', " & privato & ";"
				
				response.write QuerySQL & "<br>"
				ConnessioneDB1.Execute QuerySQL 
				
				QuerySQL = "UPDATE FORUM_MESSAGES SET THREADPARENT = ID WHERE THREADPARENT = 0"
	 			ConnessioneDB1.Execute QuerySQL 
				
				 rsTabella.movenext()
		end if
		'response.write strLine & "<br>"
       k=k+1
	  
	Next
	 
	 
	 
	 
' inserisco l'esercitazione 
'se l'esercitazione si riferisce alla convalida di un quiz mi vado a prendere, a partire dal codice del quiz, il titola del modulo da mettere nella campo descrizione, altrimenti prendo il valore del campo txtVerifica
 
 
ConnessioneDB1.close()
 response.Redirect "../admin.asp?Id_Classe="&id_classe
'end if 
  %>
  
	</font>   
	 
		      <h4><a href="../admin/studente_domande.asp?cla=<%=cla%>">Aggiornamento avvenuto ... continua ...</a></h4>
	<p>&nbsp;</p>
	<div id=piede_pagina>
			<p><p>
			
		 
			 
			  
			</div>
 <!-- se il login è corretto richima la pagina per inserire le domande del test -->

	
	</body>
	</html>
	