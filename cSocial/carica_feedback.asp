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
						
						 
						idxSel = Request("id")
						
						 
							querySQL="SELECT  *  FROM Feedback " &_
							" WHERE id_poli="& idxSel &_
						" ORDER BY Segno, Posizione;"
						   				 
						  ' response.write(querySql&"<br>")
						  set rsTabella =  ConnessioneDB.Execute(querySQL) 	
						  do while not rsTabella.EOF
							if rsTabella("Segno")="-" Then
						  		colore="red"
							else
						 		 colore="green"
							end if
							response.write("<option style='color:"&colore&"' value='"& rsTabella("Descrizione")& "'> (" &rsTabella("Segno")&") "&rsTabella("Descrizione")&"</option>")
							rsTabella.movenext
							loop
						   
							
						   
		 			
						%>		
				
					 	
				