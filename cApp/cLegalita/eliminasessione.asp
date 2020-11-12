<%@ Language=VBScript %>

<% Call Response.AddHeader("Access-Control-Allow-Origin", "*") %>
<% Session.CodePage = 65001 %>

	<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    %> 
	<!-- #include file = "../../var_globali.inc" --> 
	<!-- #include file = "../include/stringa_connessione.inc" -->

<%

chiamata=request.querystring("da")  ' vale 1 se provengo da gestione sessioni non loggato 
id_contatto=request.querystring("id_contatto")
id = request.querystring("id")

QuerySQL = "DELETE FROM Leg_Sessioni WHERE id = '"&id&"'"
ConnessioneDB.Execute(QuerySQL)

if chiamata<>"" then
Response.Redirect "../sessionilegalita2.asp?ritorno=1&id_contatto="&id_contatto
else
Response.Redirect "../sessionilegalita.asp"
end if
%>