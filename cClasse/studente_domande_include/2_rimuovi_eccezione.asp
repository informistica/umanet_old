<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Elimina Eccezione</title> 
</head>

<body>
   
   <!-- #include file = "../../service/controllo_sessione.asp" -->
   
   
<% fromURL = Request.ServerVariables("HTTP_REFERER") %>

	
	
	<% if Session("admin") = true then 
		
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    

		%>
		
        <!-- #include file = "../../var_globali.inc" --> 
 		<!-- #include file = "../../stringhe_connessione/stringa_connessione.inc" -->
  		

		<%
		frase = Request.QueryString("frase")
		nodo = Request.QueryString("nodo")
		domanda = Request.QueryString("domanda")
		ID = Request.QueryString("ID")

		Id_Stud = Request.QueryString("cod")

		if frase = 1 then
			tabella = "Eccezioni_Frasi"
			nomesing = "frase"
		end if 
		 
		if nodo = 1 then
			tabella = "Eccezioni_Nodi"
			nomesing = "nodo"
		end if
		
		 
		if domanda = 1 then
			tabella = "Eccezioni_Domande"
			nomesing = "domanda"
		end if
		
		if frase="" and nodo="" and domanda="" then
			Response.Redirect fromURL
		end if

		QuerySQL = "DELETE FROM "&tabella&" WHERE Id_Stud='"&Id_Stud&"' and Id_Pre"&nomesing&"='"&ID&"';"
		ConnessioneDB.Execute(QuerySQL)
		
		response.write("<script>alert('Eliminazione Eccezione "&nomesing&" completata correttamente'); location.href='"&fromURL&"&daRimuovi=1'</script>")
		
		'Response.Redirect fromURL
	
	else
		
		response.write("<script>alert('Devi essere amministratore per vedere questa pagina'); location.href='../../../../../'</script>")
	
end if%>
	
</body>
</html>