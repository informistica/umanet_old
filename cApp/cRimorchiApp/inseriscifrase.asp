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

categoria = Request.Form("categoria")
frase = Request.Form("frase")

frase = Replace(frase,"'",Chr(96))

QuerySQL = "INSERT INTO RimorchiApp (Testo, Categoria, Autore, Approvata) VALUES ('"&Replace(frase, Chr(34), "")&"', '"&categoria&"', 'informistica', '1');"
ConnessioneDB.Execute(QuerySQL)

Session.CodePage = 1252

Response.Redirect "../gestionerimorchiapp.asp"

%>
