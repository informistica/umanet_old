<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Resetta Eccezioni</title> 
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
		frasi = Request.QueryString("frasi")
		nodi = Request.QueryString("nodi")
		domande = Request.QueryString("domande")

		Id_Stud = Request.QueryString("cod")

		if frasi = 1 then
			tabella = "Eccezioni_Frasi"
			nome = "Frasi"
		end if 
		 
		if nodi = 1 then
			tabella = "Eccezioni_Nodi"
			nome = "Nodi"
		end if
		
		 
		if domande = 1 then
			tabella = "Eccezioni_Domande"
			nome = "Domande"
		end if
		
		if frasi="" and nodi="" and domande="" then
			Response.Redirect fromURL
		end if

		QuerySQL = "DELETE FROM "&tabella&" WHERE Id_Stud='"&Id_Stud&"';"
		ConnessioneDB.Execute(QuerySQL)
		
		response.write("<script>alert('Eliminazione Eccezioni "&nome&" completata correttamente'); location.href='"&fromURL&"'</script>")
		
		'Response.Redirect fromURL
	
	else
		
		response.write("<script>alert('Devi essere amministratore per vedere questa pagina'); location.href='../../../../../'</script>")
	
end if%>
	
</body>
</html>