<%@Language=VBScript%>
<!doctype html>
<html>
<head>
<title>Leggi Notifica</title>
</head>
<body>

<% if session("loggato") <> true then %>

<script>alert("Devi essere loggato per accedere a questa sezione"); window.location.href="../../../../"</script>

<%else%>

<% Set ConnessioneDB = Server.CreateObject("ADODB.Connection") %>

<!-- #include file = "../var_globali.inc" --> 
<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" --> 
<!-- #include file = "../service/controllo_sessione.asp" --> 

<%  

IdNotifica = Request.QueryString("IdNotifica")

 parametri = Request.ServerVariables("QUERY_STRING")
 lunghezza = len(parametri)
 
 'response.write("parametri iniziale: "&parametri&"<br>")
 
 if len(IdNotifica) = 4 then
	tofind = left(parametri, 15+38)
	'parametri = right(parametri, lunghezza-53) 
	parametri = Replace(parametri, tofind, "")
 else
	tofind = left(parametri, 14+38)
	'parametri = right(parametri, lunghezza-52)
	parametri = Replace(parametri, tofind, "")
 end if

	'response.write(parametri&"<br>")
 
 parametriv = split(parametri, "%3E")

 'response.write(parametriv(0))
 
 'in questo modo ottengo con parametriv(0) la pagina a cui dovrò andare
 
QuerySQL = "UPDATE AVVISI SET Visto = 1 WHERE ID_Avviso = '" &IdNotifica& "';"
ConnessioneDB.Execute(QuerySQL)

	parametriv2 = split(parametriv(0), "?")
	
	response.write("<br>"&parametriv2(0))
	
	if (parametriv2(0) = "ShowMessage.asp") or (instr(parametriv2(0),"ShowMessage.asp")<>0) then
		parametriv(0) = "../cSocial/"&parametriv(0)&"&byNotifiche=1"
		response.write("1")
	else
		if parametriv2(0) = "social/ShowMessage.asp"  or (instr(parametriv2(0),"social/ShowMessage.asp")<>0) then
			parametriv(0) = "../cS"&right(parametriv(0), len(parametriv(0))-1)
			response.write("2")
		else
			if (parametriv2(0) = "2inserisci_valutazione_frase.asp") or (instr(parametriv2(0),"2inserisci_valutazione_frase.asp")<>0) then
			   response.write("3")
				parametriv(0) = "../cFrasi/"&parametriv(0)
			else
				if (parametriv2(0) = "inserisci_valutazione.asp") or (instr(parametriv2(0),"inserisci_valutazione.asp")<>0)  then
					parametriv(0) = "../cDomande/"&parametriv(0)
					response.write("4")
				end if
			end if	
		end if
		
	end if
	 prova=replace(parametriv(0),"f=","")
	response.write("<br><br><br>"& prova)
		
 Response.Redirect prova


%>

<%end if%>

</body>
</html>