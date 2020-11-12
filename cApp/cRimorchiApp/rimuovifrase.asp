<%@ Language=VBScript %>

<% Session.CodePage = 65001 %>
<% if session("DB") <> "1" or session("admin") <> true then
	Response.Redirect "../../../home.asp"
	end if
%>	



<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../../var_globali.inc" --> 
 		<!-- #include file = "../../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../../service/controllo_sessione.asp" -->

<%

id = Request.QueryString("id")

QuerySQL = "DELETE FROM RimorchiApp WHERE Id_Frase = '"&id&"';"
ConnessioneDB.Execute(QuerySQL)

Session.CodePage = 1252

if Request.QueryString("prov") = "approva" then
	Response.Redirect "../approvarimorchiapp.asp"
else
	Response.Redirect "../gestionerimorchiapp.asp"
end if

%>