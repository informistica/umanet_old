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
 
 

	
	QuerySQL = "SELECT Titolo,Descrizione  FROM Paragrafi where (Id_Paragrafo='"&id_test&"')"
	set rsTitolo = ConnessioneDB.Execute(QuerySQL)
	titolo=rsTitolo(0)
	descrizione=rsTitolo(1)
	 
	'response.write(QuerySQL)
		response.write(" { "  &_
 """titolo"": """ & replace(titolo,"VbCrLf","")& """," &_
 """descrizione"": """ & descrizione & """}")
 

ConnessioneDB.Close : Set ConnessioneDB = Nothing 
		  
         
                     
                      
%>
  
   



                

