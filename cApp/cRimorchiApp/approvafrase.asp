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
categoria = Request.Form("categoriamodifica")
frase = Request.Form("frasemodifica")

frase = Replace(frase,"'",Chr(96))

QuerySQL = "UPDATE RimorchiApp SET Testo = '"&Replace(frase, Chr(34), "")&"', Categoria = '"&categoria&"', Approvata = '1' WHERE Id_Frase = '"&id&"';"
ConnessioneDB.Execute(QuerySQL)

Session.CodePage = 1252

Response.Redirect "../approvarimorchiapp.asp"

%>