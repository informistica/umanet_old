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
id_app = request.querystring("id_app")


QuerySQL = "DELETE FROM Leg_Sessioni WHERE id = '"&id&"'"
ConnessioneDB.Execute(QuerySQL)

if chiamata<>"" then
Response.Redirect "../sessioniall2.asp?ritorno=1&id_contatto="&id_contatto&"&id_app="&id_app
else
Response.Redirect "../sessioniall.asp"
end if
%>