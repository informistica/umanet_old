<%@ Language=VBScript %>

<% Call Response.AddHeader("Access-Control-Allow-Origin", "*") %>
<% Session.CodePage = 65001 %>

<!doctype html>
<html>
<head>
	<title>Modifica Categoria</title>
	<meta charset="utf-8">
</head>
<body>

<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
		%>
       <!-- #include file = "../var_globali.inc" -->
 	  	<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
    	<!-- #include file = "../stringhe_connessione/stringa_connessione_social.inc" -->

<%


id = Request.QueryString("id")


QuerySQL = "DELETE CAT_CAT  WHERE Id_Categoria='"&id&"';"
response.write(QuerySQL)
ConnessioneDB.Execute(QuerySQL)

'response.write("<br>"&Session.CodePage&"<br>")
'response.write("ok")

Response.Redirect request.serverVariables("HTTP_REFERER")

%>
</body>
</html>
