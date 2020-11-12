<%@ Language=VBScript %>

<% Call Response.AddHeader("Access-Control-Allow-Origin", "*") %>
<% Session.CodePage = 65001 %>

<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    %> 
<!-- #include file = "../../var_globali.inc" --> 
<!-- #include file = "../include/stringa_connessione.inc" -->

<%


partita = request.querystring("partita")
risultato = request.querystring("risultato")
squadra = request.querystring("squadra")

QuerySQL="select count (*) from Leg_Risultati where squadra="&squadra&" and partita="&CInt(partita)
set numris=ConnessioneDB.Execute(QuerySQL)
if numris(0) < 20 then
	QuerySQL = "INSERT INTO Leg_Risultati (squadra, risultato,partita) VALUES ("&CInt(squadra)&", "&CInt(risultato)&", "&CInt(partita)&")"
	ConnessioneDB.Execute(QuerySQL)
	response.write("eseguito")
else
	response.write("raggiunto il massimo dei risultati ammessi")
end if

%>
