<%@ Language=VBScript %>

<% Call Response.AddHeader("Access-Control-Allow-Origin", "*") %>
<% Session.CodePage = 65001 %>

<!doctype html>
<html>
<head>
	<title>Classifica SNS</title>
	<meta charset="utf-8">
</head>
<body>

<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
       <!-- #include file = "../../var_globali.inc" --> 
 	<!-- #include file = "../include/stringa_connessione.inc" -->

<%

sessione = Request.QueryString("sessione")
codice = Request.QueryString("CodiceAllievo")

response.write("<table cellspacing='13'>")
response.write("<tr><td><b>P.</b></td><td><b>Nome</b></td><td><b><center>Best</center></b></td><td><center><b>Last</b></center></td><td><center><b>Tot</b></center></td></tr>")

QuerySQL = "SELECT CodiceAllievo, MAX(Risultato) as Risultato FROM Risultati_SNS WHERE Sessione = '"&sessione&"' GROUP BY CodiceAllievo order by Risultato desc, CodiceAllievo asc" 
 
set rsTabella = ConnessioneDB.Execute(QuerySQL)
i = 1
do while not rsTabella.EOF
	
	response.write("<tr>")
	
	QuerySQLNome = "SELECT Cognome, Nome FROM Allievi WHERE CodiceAllievo = '"&rsTabella("CodiceAllievo")&"';"
	set rsNome = ConnessioneDB.Execute(QuerySQLNome)
	
	nome = rsNome("Cognome")&" "&left(rsNome("Nome"),1)&"."

	
	QuerySQLUltimo = "SELECT Risultato FROM Risultati_SNS WHERE Sessione = '"&sessione&"' and CodiceAllievo = '"&rsTabella("CodiceAllievo")&"' AND Data = (SELECT MAX(Data) FROM Risultati_SNS WHERE CodiceAllievo = '"&rsTabella("CodiceAllievo")&"' and Sessione = '"&sessione&"');"
	set rsUltimo = ConnessioneDB.Execute(QuerySQLUltimo)
	
	
	QuerySQLTot = "SELECT SUM(Risultato) as Totale FROM Risultati_SNS WHERE Sessione = '"&sessione&"' and CodiceAllievo = '"&rsTabella("CodiceAllievo")&"' GROUP BY CodiceAllievo " 
    set rsTabellaTot = ConnessioneDB.Execute(QuerySQLTot)
	 
	'QuerySQLTenta = "SELECT count(Risultato) as Tentativi FROM Risultati_SNS WHERE Sessione = '"&sessione&"' and CodiceAllievo = '"&rsTabella("CodiceAllievo")&"' GROUP BY CodiceAllievo " 
  '  set rsTabellaTenta = ConnessioneDB.Execute(QuerySQLTenta)
	 
	
	migliore = formattatempo(rsTabella("Risultato"))
	ultimo = formattatempo(rsUltimo("Risultato"))
	totale=formattatempo(rsTabellaTot("Totale")) 
	'tentativi=rsTabellaTenta("Tentativi")
	

	'response.write("<td>"&i&".</td><td>"&nome&"</td><td><center>"&migliore&" <span style='font-size:11px'>("&left(datamigliore, 5)&")</span></center></td><td><center>"&ultimo&" <span style='font-size:11px'>("&left(dataultimo, 5)&")</center></span></td>")
	if codice = rsTabella("CodiceAllievo") then
		response.write("<td><b>"&i&".</b></td><td><b>"&nome&"</b></td><td><center><b>"&migliore&"</b></center></td><td><center><b>"&ultimo&"</b></center></td><td><center><b>"&totale&"</b></center></td>")
	else
		response.write("<td>"&i&".</td><td>"&nome&"</td><td><center>"&migliore&"</center></td><td><center>"&ultimo&"</center></td><td><center><b>"&totale&"</b></center></td>")
	end if
	
	response.write("</tr>")
	
	i = i+1
rsTabella.movenext
loop


response.write("</table>")

'response.write("<br>"&Session.CodePage&"<br>")
'response.write("ok")
'Response.Redirect "../gestionerimorchiapp.asp"

%>
</body>
</html>

<% Function formattatempo(tempoiniz)

	min = 0
	sec = 0
	ore = 0
	
	tempo = tempoiniz
	
	if tempo >= 60 then
		ore = fix(tempoiniz/3600)
		tempoiniz = tempoiniz - 3600*ore
		min = fix(tempoiniz/60)
		tempoiniz = tempoiniz - 60*min
		sec = tempoiniz Mod 60
	else
		sec = tempo
	end if
	
	stampa = ""
	
	if tempo>=3600 then
		stampa = stampa & ore & "<span style='font-size:11px'>h </span> "
	end if
	
	if min < 10 and tempo >= 3600 then
		stampa = stampa & "0"
	end if
	
	if tempo >= 60 then
		stampa = stampa & min & "<span style='font-size:11px'>m </span>"
	end if
	
	if sec < 10 and tempo >=60 then
		stampa = stampa & "0"
	end if
	
	stampa = stampa &  sec & "<span style='font-size:11px'>s</span>"

	formattatempo = stampa
	
End Function
 %>