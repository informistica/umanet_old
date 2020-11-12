 
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
   <% Response.Buffer=True 
   Dim ConnessioneDB,  rsTabella, QuerySQL,StringaConnessione,URL,RecSet
   Dim CodiceTest, CodiceAllievo,IdR
   
   CodiceAllievo=Request.QueryString("cod")
   IdR=Request.QueryString("IdR")
   id_classe=request.querystring("id_classe")
   StringaConnessione= Request.Cookies("Dati")("StrConn")
   'Apertura della connessione al database  
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
       %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%  
                            'Lettura dei dati memorizzati nei cookie. 
   'CodiceTest = Request.Cookies("Dati")("CodiceTest")
   
   
   




     QuerySQL ="DELETE Risultati.* FROM Risultati WHERE ID_R =" &IdR&";"

	 ConnessioneDB.Execute(QuerySQL)
	 
	 QuerySQL ="DELETE Risultati1.* FROM Risultati1 WHERE ID_R =" &IdR&";"

	 ConnessioneDB.Execute(QuerySQL)


'CREAZIONE FILE DI TESTO PER INSERIRE LA SPIEGAZIONE DELLA DOMANDA

 

On Error Resume Next
If Err.Number = 0 Then

Response.Write "Cancellazione avvenuta! "
Else
Response.Write Err.Description 
Err.Number = 0
End If





   %>
	</font>   
	 
		
      <h4><a href="../cClasse/studente_domande.asp?cla=<%=d%>&id_classe=<%=id_classe%>">Continua ...</a></h4>
	<p>&nbsp;</p>
	<div id=piede_pagina>
			<p><p>
			
	 
						</div>
 <!-- se il login è corretto richima la pagina per inserire le domande del test -->

	
	</body>
	</html>
	