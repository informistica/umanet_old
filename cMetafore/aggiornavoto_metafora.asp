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
        codiceTest=Request.QueryString("codiceTest")
        cartella=Request.QueryString("cartella")
		  
   %>
	
		<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
		
 						<%
						


 Select Case codiceTest
  Case cartella&"_U_2_3" 'Topolino 
   QuerySQL="UPDATE M_Topolino SET Voto = "&cInt(voto)&" WHERE CodiceMetafora = "&cInt(id)&";"	
						
  Case cartella&"_U_2_5" 'Navigazione
  QuerySQL="UPDATE M_Navigazione SET Voto = "&cInt(voto)&" WHERE CodiceMetafora = "&cInt(id)&";"	
' 
  Case cartella&"_U_2_8" 'ClientServer
 End Select


						  					 
						  ' response.write(QuerySql&"<br>")
						   ConnessioneDB.Execute QuerySQL 	
						
						%>			
Voto assegnato	
				<%
				'Response.AddHeader "REFRESH","2;URL=inserisci_collegamento.asp?Tipo=0&Stato="&Stato&"&Cartella="&Cartella&"&CodiceTest="&CodiceTest&"&Capitolo="&Capitolo&"&Paragrafo="&Paragrafo&"&Modulo="&Modulo

				'Response.Redirect Session("urlmappa")
				
				%>
				
	<% end if %>			
				