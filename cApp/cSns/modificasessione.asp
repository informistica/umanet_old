<%@ Language=VBScript %>

<% Call Response.AddHeader("Access-Control-Allow-Origin", "*") %>
<% Session.CodePage = 65001 %>

<!doctype html>
<html>
<head>
	<title>Modifica Sessione</title>
	<meta charset="utf-8">
</head>
<body>

<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
       <!-- #include file = "../../var_globali.inc" --> 
 	<!-- #include file = "../include/stringa_connessione.inc" -->

<%

nome = Request.Form("titolomodifica")
id = Request.QueryString("id")
privata = Request.Form("tipomodifica")
chiave = Request.Form("chiavemodifica")

QuerySQL = "UPDATE Sessioni_SNS SET Titolo = '"&nome&"', Privata = '"&privata&"', Chiave = '"&chiave&"' WHERE Id_Sessione='"&id&"';"
'response.write(QuerySQL)
ConnessioneDB.Execute(QuerySQL)

'response.write("<br>"&Session.CodePage&"<br>")
'response.write("ok")

Response.Redirect "../gestionesns.asp"

%>
</body>
</html>
