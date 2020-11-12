<%@ Language=VBScript %>

<%
  Response.Buffer = true
  'On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") or Session("Admin") = false then %>
	 Sessione scaduta, rieffettuare il login.
  <% else %>
     

	
		<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
		
 						<%
						
						byUmanet = Request("byUmanet")
						capitolo = Request("modulo")
						paragrafo = Request("paragrafo")
						
						idxSel = Session("cartella")&"_"&capitolo&"_"&paragrafo
						Paragrafo = '3Ct$6 _ 3Ct$6_14 _Introduzione a javascript'
						if byUmanet="" then
						querySQL="SELECT  Titolo, PosMod ,PosPar,ID_Mod,ID_Paragrafo,TitPar FROM MODULI_NOT_UMANET " &_
							" WHERE Id_Classe='"&Session("Id_Classe") &"'" &_
						" ORDER BY PosMod,PosPar;"
						 else
						 querySQL="SELECT  Titolo, PosMod ,PosPar,ID_Mod,ID_Paragrafo,TitPar FROM MODULI_UMANET1" &_
							" WHERE Id_Classe='"&Session("Id_Classe") &"'" &_
						" ORDER BY PosMod,PosPar;"
						 end if
						cont=1
						 response.write(querySql&"<br>")
						Set rsTabellaPos = ConnessioneDB.Execute(QuerySQL) 
							do while not rsTabellaPos.eof
							   if strcomp(idxSel,rsTabellaPos("ID_Paragrafo"))=0 then
								segnalibro=cont 
								'response.write("<br>" & Session("ID_ParSel") & "="& rsTabellaPos("ID_Paragrafo") &"Cont="&cont)
							end if
							cont=cont+1 'veniva fatoo cont = cont+0 per cui non veniva incrementato il numero del totale del box -> rimaneva sempre 1
							rsTabellaPos.movenext
						loop
						
						Session("segnalibro") = segnalibro
						
						
						querySQL = "SELECT * FROM ParagrafiSottoparagrafi WHERE Id_Paragrafo = '"&idxSel&"'"	 
						   response.write("<br>"&querySql&"<br>")
						  set rsTabella =  ConnessioneDB.Execute(querySQL) 	
						   
							do while not rsTabella.EOF
							
							QuerySQL = "SELECT * FROM Sottoparagrafi WHERE ID_Sottoparagrafo = '"&rsTabella("Id_Sottoparagrafo")&"'"
							set rsSP = ConnessioneDB.Execute(QuerySQL)
							response.write("<option value='"&rsSP("Posizione")&"'>"&rsSP("Titolo")&"</option>")
							rsTabella.movenext
							loop
						   
						
						%>		
				
	<% end if %>			
				