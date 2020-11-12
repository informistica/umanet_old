<%@ Language=VBScript %>

<% Call Response.AddHeader("Access-Control-Allow-Origin", "*") %>
<% Session.CodePage = 65001 %>

<!doctype html>
<html>
<head>
	<title>Proponi Frase</title>
	<meta charset="utf-8">
</head>
<body>

<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
       <!-- #include file = "../../var_globali.inc" --> 
 	<!-- #include file = "../include/stringa_connessione.inc" -->

<%

categoria = Request.QueryString("categoria")
frase = Request.QueryString("frase")
allievo = Request.QueryString("allievo")

frase = Replace(frase,"'",Chr(96))
response.write(frase)

QuerySQL = "INSERT INTO RimorchiApp (Testo, Categoria, Autore, Approvata) VALUES ('"&Replace(frase, Chr(34), "")&"', '"&categoria&"', '"&allievo&"', '0');"
ConnessioneDB.Execute(QuerySQL)

response.write("<br>"&Session.CodePage&"<br>")
response.write("ok")
'Response.Redirect "../gestionerimorchiapp.asp"

%>
</body>
</html>
