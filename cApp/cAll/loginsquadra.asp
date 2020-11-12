<%@ Language=VBScript %>

<% Call Response.AddHeader("Access-Control-Allow-Origin", "*") %>
<% Session.CodePage = 65001 %>

	<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    %> 
	<!-- #include file = "../../var_globali.inc" --> 
	<!-- #include file = "../include/stringa_connessione.inc" -->

<%

on error resume next
partita= request.querystring("partita")
codice = request.querystring("codice")
login = 0
' se login varrà 0 non ho trovato nè partita nè squadra
' se varrà 1 ho trovato la partita ma non la squadra
' se varrà 2 ok ho trovato entrambe inizia a giocare
squadra = 0
nSess=0	
ndomande=0	

QuerySQL = "SELECT * FROM Leg_Sessioni where id="&partita
set nSessioni = ConnessioneDB.Execute(QuerySQL)
If Not IsNull(nSessioni("id_test")) then ' se il test riguarda un determinato paragrafo cnv
	id_test=nSessioni("id_test")
else
	id_test=0
end if

if not nSessioni.EOF then
    valore=nSessioni("valore")								
	s=Split(valore, ",")
	if  s(0)="P" then
	  nSess=nSess+1
	  login=login+1
	  For i = 1 To UBound(s)
	    if rtrim(s(i)) = codice then
		   squadra = i
		   login = login + 1
	    end if
	  Next
   end if
   ndomande = nSessioni("ndomande")
end if



'response.write valore

'response.write("<br>"&Session.CodePage&"<br>")
'response.write("ok")
'Response.Redirect "../sessionilegalita.asp"

	response.write "{""statologin"": """&login&""", ""id_test"": """&trim(id_test)&""", ""squadra"": """&squadra&""", ""ndomande"": """&ndomande&"""}"

%>