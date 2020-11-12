<%@ Language=VBScript %>
<%
  Response.Buffer = true
  'On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 Sessione scaduta, rieffettuare il login.
  <% else %>
   <%

		    CodiceAllievo = Session("CodiceAllievo")
			UtenteAssociato = Request("user")
			PasswordAssociato = Request("pass")
   %>
	
		<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
 
 						<%
						
						   QuerySQL = "SELECT count(*) FROM Allievi WHERE CodiceAllievo = '"&UtenteAssociato&"' AND PasswordSHA256 = '"&PasswordAssociato&"';"
						   Set rsTabella = ConnessioneDB.Execute(QuerySQL)
						   
						   
						   if rsTabella(0) = 1 then
								
								QuerySQL = "SELECT count(*) FROM AssociazioniAllievi WHERE (CodiceAllievo = '"&CodiceAllievo&"' AND UtenteAssociato = '"&UtenteAssociato&"') OR (CodiceAllievo = '"&UtenteAssociato&"' AND UtenteAssociato = '"&CodiceAllievo&"')"
								set rsTabella1 = ConnessioneDB.Execute(QuerySQL)
								
								if rsTabella1(0) = 0 AND UtenteAssociato <> CodiceAllievo then
								
									QuerySQL = "INSERT INTO AssociazioniAllievi (CodiceAllievo, UtenteAssociato) VALUES ('"&CodiceAllievo&"', '"&UtenteAssociato&"')"
									ConnessioneDB.Execute(QuerySQL)
									
									response.write "associato"
								
								else
								
									response.write "presente"
								end if	
						   
						   else
						   
								response.write "errorecred"
						   
						   end if
							
						%>	

<% end if %>