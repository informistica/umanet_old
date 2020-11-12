<%@ Language=VBScript %>

<%
  Response.Buffer = true
  'On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
   %>

	
		<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
		
 						<%
						
						 
						idclasse=  Request("idclasse")
                        url=  Request("url")

						
		on error resume next				 
		QuerySQL="UPDATE Classi Set Url_feedback = '" & url & "' where ID_Classe='"&idclasse&"';"
		ConnessioneDB.Execute(QuerySQL)

                 If Err.Number = 0 Then
					Response.Write "Inserimento avvenuto!"
				Else
					Response.Write Err.Description
					Err.Number = 0
				End If

'response.write(QuerySQL)						   				 
						  
						   
							
						   
		 			
						%>		
				
					 	
				