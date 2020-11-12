<%@ Language=VBScript %>

<%
Response.charset="utf-8" 'codifica caratteri speciali funzionante!! 
Call Response.AddHeader("Access-Control-Allow-Origin", "*") 
id_test=request.querystring("id_test")
id_app=request.querystring("id_app")
'paragrafo = Request.QueryString("paragrafo")

%>


<%
  
 %>
<% Response.Buffer=True %>
 

<%  
  'On Error Resume Next  
    
		 
 ' per generare un ordinamento casuale delle domande in base ad uno dei seguenti campi
 
 
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection") 
		 
 
	 
		%> 
        <!-- #include file = "../../var_globali.inc" --> 
		<!-- #include file = "../include/stringa_connessione.inc" --> 
 	     
	 
                 
<%  

classe="Expo"	
TestAbilitato=1
 
 
 
 
											    splittato=Split(id_test, "_")  'es Expo_12
												if (UBound(splittato)=1) then ' non contiene il numero di paragrfo quindi è l'id del modulo, devo trovare il numero minimo di domande per paragrafo
												' seleziono tutti i paragrafi del capitolo per i quali esistono delle domande
														QuerySQL="Select * from MODULI_PARAGRAFI_CLASSE where ID_Mod='"&id_test&"'"
														set rsParagrafi=ConnessioneDB.execute(QuerySQL)
														 minNumDom=20												
														do while not rsParagrafi.eof
															q1="Select nDomande from DomandeQuizN where ID_Paragrafo='"&rsParagrafi("ID_paragrafo")&"'"
															set rsNumDom=ConnessioneDB.execute(q1)
															if not rsNumDom.eof then
																 if rsNumDom(0)<minNumDom then
																   minNumDom = rsNumDom(0)
																 end if
			
															end if
															
															
															rsParagrafi.movenext
														loop
														response.write(minNumDom)
												else												
															' conto il numero di domande disponibili per il quiz estratto
													if id_test<>"" then 'se anzichè tutti i paragrafi della vista ne prendo solo uno a scelta
													 extraSql=" and (Id_Arg='"&id_test&"')"
													else
													 extraSql=""
													end	 if
													Select Case id_app
														  Case 1
															 QuerySQL = "SELECT count(*)  FROM Leg_Domande  where 1=1 " &extraSql
														  Case 2
															 QuerySQL = "SELECT count(*)  FROM Cnv_Domande where 1=1 " &extraSql
														 
													End Select
													

													set rsDomande = ConnessioneDB.Execute(QuerySQL)
													ndomande=rsDomande(0)
													response.write(ndomande)
													'response.write(QuerySQL)
 
												
												end if
 

 
 
	
	

ConnessioneDB.Close : Set ConnessioneDB = Nothing 
		  
         
                     
                      
%>
  
   



                

