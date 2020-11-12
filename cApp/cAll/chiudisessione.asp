<%@ Language=VBScript %>

<% Call Response.AddHeader("Access-Control-Allow-Origin", "*") %>
<% Session.CodePage = 65001 %>

<!doctype html>
<html>
<head>
	<title>Chiudi Sessione</title>
	<meta charset="utf-8">
</head>
<body>

	<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    %> 
	<!-- #include file = "../../var_globali.inc" --> 
	<!-- #include file = "../include/stringa_connessione.inc" -->

<%

partita=request.querystring("partita")
mail=request.querystring("mail")
id_contatto=request.querystring("id_contatto")
id_app=request.querystring("id_app")
chiamata=request.querystring("da")  ' vale 1 se provengo da gestione sessioni non loggato 
' modifico sum in avg per risolvere il problema degli invii multipli dei "furbi"
QuerySQL = "SELECT AVG(risultato) as 'Media' FROM Leg_Risultati where partita="&partita&" group by partita, squadra"
set rsRisultati = ConnessioneDB.Execute(QuerySQL)

valore = "R,"
do while not rsRisultati.EOF
valore = valore & rsRisultati("Media") & ","
'response.write valore
rsRisultati.movenext
Loop

lunghezza = Len(valore)

valore = Mid(valore,1,lunghezza-1)


QuerySQL = "UPDATE Leg_Sessioni SET valore = '"&valore&"' WHERE id = '"&partita&"'"
ConnessioneDB.Execute(QuerySQL)

QuerySQL = "DELETE FROM Leg_Risultati WHERE partita="&partita
ConnessioneDB.Execute(QuerySQL)
if chiamata<>"" then
Response.Redirect "../sessioniall2.asp?ritorno=1&id_contatto="&id_contatto&"&id_app="&id_app
else
Response.Redirect "../sessioniall.asp"
end if
%>

</body>
</html>