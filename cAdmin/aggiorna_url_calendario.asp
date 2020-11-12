 
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
	 
       %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
   
 
<%  

urlCalendar=rtrim(Replace(Request.form("txtCalendar"), Chr(34), ""))
 
QuerySQL = "UPDATE Classi SET Url_calendar = '" &urlCalendar& "' WHERE ID_Classe = '"&  Session("Id_Classe") &"';"
'response.write(QuerySQL)
Set rsTabella = ConnessioneDB.Execute(QuerySQL) 

 
	 
	 
	 
	 
' inserisco l'esercitazione 
'se l'esercitazione si riferisce alla convalida di un quiz mi vado a prendere, a partire dal codice del quiz, il titola del modulo da mettere nella campo descrizione, altrimenti prendo il valore del campo txtVerifica
 
 
ConnessioneDB.close()
 response.Redirect "admin.asp?Id_Classe="&Session("Id_Classe")
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
	