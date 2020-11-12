<!-- calcola_risultato_MODBC3.asp -->
 

<%@ Language=VBScript %>
   <% Response.Buffer=True 
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,RecSet
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome
   
   
 
   'Apertura della connessione al database  
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
	 
       %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
   <!-- #include file = "../var_globali.inc" -->
<%  
                     'Lettura dei dati memorizzati nei cookie. 
   
 
 				 stile=request.QueryString("stile")

	  			 QuerySQL ="UPDATE Allievi SET Stile = '" & stile & "' WHERE CodiceAllievo ='" &Session("CodiceAllievo")&"';"
				 ' response.write("Aggiornato stile")
				 Set rsTabella = ConnessioneDB.Execute(QuerySQL)	 
  				 ConnessioneDB.Execute QuerySQL 
				 session("stile")=stile
				 if Request.ServerVariables("HTTP_REFERER") <>"" then 
							response.Redirect request.serverVariables("HTTP_REFERER") 
	        	 end if 
 
  %>
  
	 
 <!-- se il login è corretto richima la pagina per inserire le domande del test -->

	 