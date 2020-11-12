<%@ Language=VBScript %>
<%
  Response.Buffer = true
  'On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 Sessione scaduta, rieffettuare il login.
  <% else %>
   <%

		  Id_n1=Request.QueryString("Id_n1")   'id del nodo di partenza del link (href che punto all'ancora nel documento)
		  Id_n2=Request.QueryString("Id_n2")  ' 'id del nodo di arrivo del link   (ancora puntata dall'href)
		  L1=Request.QueryString("L1") ' livello del primo nodo da cui parte il link (chi, cosa, dove, ecc...)
		  L2=Request.QueryString("L2")' livello del secondo nodo a cui arriva il link (chi, cosa, dove, ecc...)
		  T2=Request.QueryString("T2") ' testo nel livello di arrivo da visualizzare sull'arco che collega i nodi
		  T2 = Replace(T2, "'",Chr(96))
		 
		  T2 = Replace(T2,chr(133),"a"&Chr(96))
		  T2 = Replace(T2,chr(236),"i"&Chr(96))
		  T2 = Replace(T2,chr(237),"i"&Chr(96))
		  T2 = Replace(T2,chr(242),"o"&Chr(96))
		  T2 = Replace(T2,chr(243),"o"&Chr(96))
		  T2 = Replace(T2,chr(151),"u"&Chr(96))
		  T2 = Replace(T2,chr(250),"u"&Chr(96))
		 T2 = Replace(T2,chr(138),"e"&Chr(96))
		 T2 = Replace(T2,chr(130),"e"&Chr(96))		  
   %>
	
		<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
 
 						<%
						
						   QuerySQL="INSERT INTO Link (Id_n1, L1, Id_n2,L2,Id_Stud,Testo2) VALUES (" & clng(Id_n1) & "," &L1 & ", " & clng(Id_n2) & "," & L2 & ",'" & Session("CodiceAllievo")& "','" &T2& "');"						 
						   'response.write(QuerySql&"<br>")
						   ConnessioneDB.Execute QuerySQL 	
						   
						   	QuerySQL = "UPDATE Nodi SET NLink = NLink + 1 WHERE CodiceNodo = '"&clng(Id_n1)&"';"
							ConnessioneDB.Execute(QuerySQL)
							QuerySQL = "UPDATE Nodi SET NLink = NLink + 1 WHERE CodiceNodo = '"&clng(Id_n2)&"';"
							ConnessioneDB.Execute(QuerySQL)
						   
						   ' punti nodo di partenza
							
							QuerySQL = "SELECT * FROM Nodi WHERE CodiceNodo = '"&clng(Id_n1)&"'"
							Set rsTabella = ConnessioneDB.Execute(QuerySQL)
							proprietario = rsTabella("Id_Stud")
							voto = rsTabella("Voto")
							
							if proprietario = Session("CodiceAllievo") and voto < 3 then
								QuerySQL = "UPDATE Nodi SET Voto = "&(voto+1)&" WHERE CodiceNodo = '"&clng(Id_n1)&"';"
								ConnessioneDB.Execute(QuerySQL)
								partenza = 1
							else
								partenza = 0
							end if
							
							' punti nodo di arrivo
							
							if partenza = 0 then
							
							QuerySQL = "SELECT * FROM Nodi WHERE CodiceNodo = '"&clng(Id_n2)&"'"
							Set rsTabella = ConnessioneDB.Execute(QuerySQL)
							proprietario = rsTabella("Id_Stud")
							voto = rsTabella("Voto")
							
							if proprietario = Session("CodiceAllievo") and voto < 3 then
								QuerySQL = "UPDATE Nodi SET Voto = "&(voto+1)&" WHERE CodiceNodo = '"&clng(Id_n2)&"';"
								ConnessioneDB.Execute(QuerySQL)
							end if
							
							end if
							
						%>	
inserito
				<%
				'Response.AddHeader "REFRESH","2;URL=inserisci_collegamento.asp?Tipo=0&Stato="&Stato&"&Cartella="&Cartella&"&CodiceTest="&CodiceTest&"&Capitolo="&Capitolo&"&Paragrafo="&Paragrafo&"&Modulo="&Modulo

				'Response.Redirect Session("urlmappa")
				
				%> 
<% end if %>