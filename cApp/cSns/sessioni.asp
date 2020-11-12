<%@ Language=VBScript %>

<% Call Response.AddHeader("Access-Control-Allow-Origin", "*") %>
<% Session.CodePage = 65001 %>

<!doctype html>
<html>
<head>
	<title>Sessioni Aperte</title>
	<meta charset="utf-8">
</head>
<body>

<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
       <!-- #include file = "../../var_globali.inc" --> 
 	<!-- #include file = "../include/stringa_connessione.inc" -->

<%

QuerySQLC = "SELECT count(*) FROM Sessioni_SNS WHERE Aperta = '1';"
set rsCount = ConnessioneDB.Execute(QuerySQLC)

if rsCount(0) = 0 then
	response.write("<span class='testo'>Non ci sono sessioni classificate aperte</span>")
else %>
	<select id="soflow" style="width:75%; height: 40px" onchange="privata()">
	<option value="-1">Seleziona una Sessione</option>
	
	<% QuerySQL = "SELECT * FROM Sessioni_SNS WHERE Aperta = '1'"
	set rsTabella = ConnessioneDB.Execute(QuerySQL)
	
	
	
	do while not rsTabella.EOF %>
		<option value="<%=rsTabella("Id_Sessione")%>,<%=rsTabella("Titolo")%>,<%=rsTabella("Privata")%>,<%=rsTabella("Chiave")%>,<%=left(rsTabella("Data"),10)%>"><%=rsTabella("Titolo")%></option>
		<% rsTabella.movenext
	loop %>
	
	</select>
<% end if %>

</body>
</html>