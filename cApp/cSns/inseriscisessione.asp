<%@ Language=VBScript %>

<% Call Response.AddHeader("Access-Control-Allow-Origin", "*") %>
<% Session.CodePage = 65001 %>

<!doctype html>
<html>
<head>
	<title>Inserisci Nuova Sessione</title>
	<meta charset="utf-8">
</head>
<body>

<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
       <!-- #include file = "../../var_globali.inc" --> 
 	<!-- #include file = "../include/stringa_connessione.inc" -->

<%

nome = Request.Form("nomesessione")
privata = Request.Form("tiposessione")
chiave = Request.Form("passwordsessione")

QuerySQL = "INSERT INTO Sessioni_SNS (Titolo, Data, Aperta, Privata, Chiave) VALUES ('"&nome&"', '"&Now()&"', '1', '"&privata&"', '"&chiave&"');"
'response.write(QuerySQL)
ConnessioneDB.Execute(QuerySQL)

'response.write("<br>"&Session.CodePage&"<br>")
'response.write("ok")

Response.Redirect "../gestionesns.asp"

%>
</body>
</html>
