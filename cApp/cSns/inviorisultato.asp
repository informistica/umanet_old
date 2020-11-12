<%@ Language=VBScript %>

<% Call Response.AddHeader("Access-Control-Allow-Origin", "*") %>
<% Session.CodePage = 65001 %>

<!doctype html>
<html>
<head>
	<title>Invio Risultato</title>
	<meta charset="utf-8">
</head>
<body>

<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
       <!-- #include file = "../../var_globali.inc" --> 
 	<!-- #include file = "../include/stringa_connessione.inc" -->

<%

tempo = Request.QueryString("tempo")
sessione = Request.QueryString("sessione")
allievo = Request.QueryString("CodiceAllievo")

QuerySQL = "INSERT INTO Risultati_SNS (CodiceAllievo, Data, Sessione, Risultato) VALUES ('"&allievo&"', '"&now()&"', '"&sessione&"', '"&tempo&"');" 
ConnessioneDB.Execute(QuerySQL)

'response.write("<br>"&Session.CodePage&"<br>")
'response.write("ok")
'Response.Redirect "../gestionerimorchiapp.asp"

%>
</body>
</html>
