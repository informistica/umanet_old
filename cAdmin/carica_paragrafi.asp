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
						idxSel = Request("modulo")
						
						if byUmanet="" then
							querySQL="SELECT  Titolo,PosMod ,PosPar,ID_Mod,ID_Paragrafo,TitPar  FROM MODULI_NOT_UMANET " &_
							" WHERE Id_Classe='"&Session("Id_Classe") &"' and ID_Mod='"& idxSel &"' and Visibile=1" &_
						" ORDER BY PosPar;"
						   else
						  querySQL="SELECT  Titolo,PosMod ,PosPar,ID_Mod,ID_Paragrafo,TitPar  FROM MODULI_UMANET1" &_
							" WHERE Id_Classe='"&Session("Id_Classe") &"' and ID_Mod='"& idxSel &"' and Visibile=1" &_
						" ORDER BY PosPar;"
						   end if					 
						   'response.write(querySql&"<br>")
						  set rsTabella =  ConnessioneDB.Execute(querySQL) 	
						   pospar=1
							do while not rsTabella.EOF
							'response.write("<option value='"&rsTabella("PosPar")&"'>"&rsTabella("TitPar")&"</option>")
							response.write("<option value='"&rsTabella("TitPar")&"'>"&pospar&"-"&rsTabella("TitPar")&"</option>")
							rsTabella.movenext
							pospar=pospar+1
							loop
						   
						
						%>		
				
	<% end if %>			
				