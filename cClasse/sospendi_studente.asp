<%@ Language=VBScript %>
 

   <% Response.Buffer=True 
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,RecSet
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome
  
   CodiceAllievo=Request.QueryString("CodiceAllievo")
   
   'Apertura della connessione al database  
    
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")   
	 
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
     
        	  

<%  
                            'Lettura dei dati memorizzati nei cookie. 
   'CodiceTest = Request.Cookies("Dati")("CodiceTest")
   
   
   
   
  
 
if  (Session("Admin")=true) then  
QuerySQL = "UPDATE Allievi SET Attivo =0 WHERE CodiceAllievo='"&CodiceAllievo&"';"

ConnessioneDB.execute(QuerySQL)
'response.write(QuerySQL)
Response.Redirect request.serverVariables("HTTP_REFERER")

     
 
On Error Resume Next
If Err.Number = 0 Then

Response.Write "Sospensione avvenuta! "
Else
Response.Write Err.Description 
Err.Number = 0
End If
end if





   %>
	 