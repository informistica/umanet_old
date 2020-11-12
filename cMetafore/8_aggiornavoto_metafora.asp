<%@ Language=VBScript %>

<%
  Response.Buffer = true
  'On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if Session("Admin") = false then %>
	 Sessione scaduta, rieffettuare il login.
  <% else %>
     
   <%
		
		id = Request.QueryString("id")
		voto = Request.QueryString("voto")
		  
   %>
	
		<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
		
 						<%
						
						   QuerySQL="UPDATE FORUM_MESSAGES SET Punti = "&cInt(voto)&" WHERE ID = "&cInt(id)&";"						 
						   response.write(QuerySql&"<br>")
						   ConnessioneDB.Execute QuerySQL 	
						
						%>			
modificato	 
				<%
				'Response.AddHeader "REFRESH","2;URL=inserisci_collegamento.asp?Tipo=0&Stato="&Stato&"&Cartella="&Cartella&"&CodiceTest="&CodiceTest&"&Capitolo="&Capitolo&"&Paragrafo="&Paragrafo&"&Modulo="&Modulo

				'Response.Redirect Session("urlmappa")
				
				%>
				
	<% end if %>			
				