<%@ Language=VBScript %>
 
	
		<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
 
 						<%
						
					 
							'QuerySQL = "SELECT * FROM Nodi WHERE CodiceNodo = '"&clng(Id_n1)&"'"
							
							QuerySQL = "SELECT * FROM Nodi"
							Set rsTabella = ConnessioneDB.Execute(QuerySQL)
							do while not rsTabella.eof
								id_nodo=rsTabella("CodiceNodo")
								QuerySQL = "SELECT count(*) FROM Link WHERE Id_n1 = "&id_nodo&" or Id_n2="&id_nodo&";"
								Set rsTabellaNLink = ConnessioneDB.Execute(QuerySQL)
								response.write(QuerySQL&"<br>")
								nlink=rsTabellaNLink(0)
								if nlink > 0 then
									QuerySQL = "UPDATE Nodi SET NLink = "&(nlink)&" WHERE CodiceNodo = '"&id_nodo&"';"
									ConnessioneDB.Execute(QuerySQL)
									response.write(QuerySQL&"<br>")
								else 
									QuerySQL = "UPDATE Nodi SET NLink = 0 WHERE CodiceNodo = '"&id_nodo&"';"
									ConnessioneDB.Execute(QuerySQL)
									response.write(QuerySQL&"<br>")
								end if
							
							
							rsTabella.movenext
							loop
							
							 
					  %>