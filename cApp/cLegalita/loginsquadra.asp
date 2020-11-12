<%@ Language=VBScript %>

<% Call Response.AddHeader("Access-Control-Allow-Origin", "*") %>
<% Session.CodePage = 65001 %>

	<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    %> 
	<!-- #include file = "../../var_globali.inc" --> 
	<!-- #include file = "../include/stringa_connessione.inc" -->

<%


partita= request.querystring("partita")
codice = request.querystring("codice")
login = 0
' se login varrà 0 non ho trovato nè partita nè squadra
' se varrà 1 ho trovato la partita ma non la squadra
' se varrà 2 ok ho trovato entrambe inizia a giocare
squadra = 0

nSess=0		

QuerySQL = "SELECT * FROM Leg_Sessioni where id="&partita
set nSessioni = ConnessioneDB.Execute(QuerySQL)
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
end if

 ndomande = nSessioni("ndomande")

'response.write valore

'response.write("<br>"&Session.CodePage&"<br>")
'response.write("ok")
'Response.Redirect "../sessionilegalita.asp"

	response.write "{""statologin"": """&login&""", ""squadra"": """&squadra&""", ""ndomande"": """&ndomande&"""}"

%>