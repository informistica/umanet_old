<%@ Language=VBScript %>

<% Call Response.AddHeader("Access-Control-Allow-Origin", "*") %>
<% Session.CodePage = 65001 %>

<!doctype html>
<html>
<head>
	<title>Inserisci Sessione</title>
	<meta charset="utf-8">
</head>
<body>

<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
       <!-- #include file = "../../var_globali.inc" --> 
 	<!-- #include file = "../include/stringa_connessione.inc" -->

<%


id_contatto=request.form("txtContatto")
nome = request.form("txtnome")
nsquadre = request.form("nsquadre")
ndomande = request.form("ndomande")
valore = "P,"


chiamata=request.querystring("da")  ' vale 1 se provengo da gestione sessioni non loggato 

'response.write nsquadre
QuerySQL = "INSERT INTO Leg_Sessioni (nome, valore,data,id_contatto,ndomande) VALUES ('1','1','1',1,1)"
ConnessioneDB.Execute(QuerySQL)
response.write(QuerySQL&"<br>")

'devo ottenere l'id della partita come idmax + 1
QuerySQL = "SELECT max(id) FROM Leg_Sessioni"
set rsSessioni = ConnessioneDB.Execute(QuerySQL)
partita=rsSessioni(0)
response.write(partita)


Randomize()

	dim numeri(100)
	dim used(100)
	for a=0 to 99
		used(a)="false"
	next 

Dim i
i=0
While i < CInt(nsquadre)
QuerySQL = "INSERT INTO Leg_Risultati (squadra, risultato,partita) VALUES ("&(i+1)&", 10,"&partita&")"
response.write(QuerySQL&"<br>")

'******provo a toglierlo perchè la presenza del 10 falsa la media
'ConnessioneDB.Execute(QuerySQL)

response.write(i&"<br>"&nsquadre)

	
'	numero = numero+(100*(i+1))

	do 
	    numero = CInt(Rnd()*100)
		if numero=0 then 
		  numero=1
		end if  
	loop until (used(numero)="false") 
	used(numero)="true"
	valore = valore & numero
	if i < (nsquadre-1) then
		valore = valore & ","
	end if
	i = i + 1
WEnd

response.write nome & "<br>" & valore & "<br>" & now()
QuerySQL ="UPDATE Leg_Sessioni SET nome = '" & nome & "', valore = '" &valore & "', data = '" & now() & "', id_contatto = " & id_contatto & ", ndomande = " & ndomande & "    WHERE id =" &partita &";"
response.write(QuerySQL&"<br>")							
ConnessioneDB.Execute(QuerySQL)

'response.write valore

'response.write("<br>"&Session.CodePage&"<br>")
'response.write("ok")
if chiamata<>"" then
 ' Response.Redirect "../sessionilegalita2.asp?ritorno=1&id_contatto="&id_contatto
else
 ' Response.Redirect "../sessionilegalita.asp"
end if
%>
</body>
</html>
