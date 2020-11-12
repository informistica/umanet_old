<%@ Language=VBScript %>

<% Call Response.AddHeader("AccessControlAllowOrigin", "*") %>
<% Session.CodePage = 65001 %>


<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")%> 
		
	<!-- #include file = "../../var_globali.inc" --> 
 	<!-- #include file = "../include/stringa_connessione.inc" -->
	
	<% 'os = Request.ServerVariables("HTTP_USER_AGENT") 
	
	
	Dim cat(4)
	
	' if InStr(os, "iPad") > 0 or InStr(os, "iPhone") > 0 then
	
		' id = Request.QueryString("id")

		' cat(0) = "Frasi Impossibili"
		' cat(1) = "Frasi Pessime"
		' cat(2) = "Frasi Scontate"
		' cat(3) = "Frasi Carine"

		' QuerySQLN = "SELECT count(*) FROM RimorchiApp WHERE Categoria = '"&cat(id)&"' and Approvata = '1';"
		' 'response.write(QuerySQLN)

		' rsNumero = ConnessioneDB.Execute(QuerySQLN)
		' n = rsNumero(0)

		' 'response.write("<br>"&n)

		' 'response.write("<br>"&num)

		' QuerySQL = "SELECT * FROM RimorchiApp WHERE Categoria = '"&cat(id)&"' and Approvata = '1';"
		' 'response.write("<br>"&QuerySQL)

		' set rsTabella = ConnessioneDB.Execute(QuerySQL)
		
		' Randomize()
		' numero = cInt(Rnd()*n)-1
		
		' rsTabella.Move numero
		
		' response.write(rsTabella("Testo"))
		
	' else

	
	%>
<%

id = Request.QueryString("id")

cat(0) = "Frasi Impossibili"
cat(1) = "Frasi Pessime"
cat(2) = "Frasi Scontate"
cat(3) = "Frasi Carine"

QuerySQLN = "SELECT count(*) FROM RimorchiApp WHERE Categoria = '"&cat(id)&"' and Approvata = '1';"
'response.write(QuerySQLN)

rsNumero = ConnessioneDB.Execute(QuerySQLN)
n = rsNumero(0)

'response.write("<br>"&n)


'response.write("<br>"&num)

QuerySQL = "SELECT * FROM RimorchiApp WHERE Categoria = '"&cat(id)&"' and Approvata = '1';"
'response.write("<br>"&QuerySQL)

set rsTabella = ConnessioneDB.Execute(QuerySQL)

i=0
response.write("{""totale"": """&n&""", ")
do while not rsTabella.EOF
	QuerySQLAut = "SELECT Cognome, Nome FROM Allievi WHERE CodiceAllievo = '"&rsTabella("Autore")&"';"
	set rsAutore = ConnessioneDB.Execute(QuerySQLAut)
	autore = rsAutore("Cognome")&" "&left(rsAutore("Nome"), 1)&"."
	
	response.write(""""&i&""": """&Replace(Replace(rsTabella("Testo"), Chr(96), "'"), Chr(34), "")&""", ""autore"&i&""": """&autore&"")
	
	if i<(n-1) then
		response.write(""", ")
	else	
		response.write("""")
	end if	
	
	i=i+1
	rsTabella.movenext
loop	
response.write("}")	
	
'response.write("<input id='fraseprec' type='hidden' value='"&num&"'/>")
'response.write(rsTabella("Testo"))

'Response.Redirect "rimorchiapp_gestione.asp"

%>

<% 


'end if

%>