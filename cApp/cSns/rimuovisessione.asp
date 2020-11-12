<%@ Language=VBScript %>

<% Call Response.AddHeader("Access-Control-Allow-Origin", "*") %>
<% Session.CodePage = 65001 %>

<!doctype html>
<html>
<head>
	<title>Elimina Sessione</title>
	<meta charset="utf-8">
</head>
<body>

<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
       <!-- #include file = "../../var_globali.inc" --> 
 	<!-- #include file = "../include/stringa_connessione.inc" -->

<%

id = Request.QueryString("id")

QuerySQLRis = "DELETE FORM Risultati_SNS WHERE Sessione = '"&id&"';"
ConnessioneDB.Execute(QuerySQLRis)

QuerySQL = "DELETE FROM Sessioni_SNS WHERE Id_Sessione = '"&id&"';"
'response.write(QuerySQL)
ConnessioneDB.Execute(QuerySQL)

'response.write("<br>"&Session.CodePage&"<br>")
'response.write("ok")

Response.Redirect "../gestionesns.asp"

%>
</body>
</html>
