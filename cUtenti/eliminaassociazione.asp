<%@ Language=VBScript %>
<%
  Response.Buffer = true
  'On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 Sessione scaduta, rieffettuare il login.
  <% else %>
   <%

		    CodiceAllievo = Request("all1")
			UtenteAssociato = Request("all2")
   %>
	
		<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
 
 						<%
						
						   QuerySQL = "DELETE FROM AssociazioniAllievi WHERE CodiceAllievo = '"&CodiceAllievo&"' AND UtenteAssociato = '"&UtenteAssociato&"';"
						   ConnessioneDB.Execute(QuerySQL)
						   
						   'response.write(QuerySQL)
						   
						   response.write("eliminato")
						   
						   
							
						%>	

<% end if %>