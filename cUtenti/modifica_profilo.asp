<%@ Language=VBScript %>
<html>
<head>	
	<link rel="stylesheet" type="text/css" href="../../stile.css">
	 
   
</head>


   <% Response.Buffer=True 
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,RecSet
  ' mipiace = Replace(Request.Form("mipiace"), "'", "''")
 '  nonmipiace=Request.Form("nonmipiace")
  ' descriviti=Request.Form("descriviti")
   
   
   cognome=ltrim(Request("txtCognome"))
   nome=ltrim(Request("txtNome"))
   mipiace = Replace(Request("txtmipiace"), "'", Chr(96))
   nonmipiace = Replace(Request("txtnonmipiace"), "'", Chr(96))
   descriviti = Replace(Request("S1"), "'", Chr(96))
    cod=Request.QueryString("CodiceAllievo")
	gruppo=Request("txtTag")
	inquiz=Request("txtInquiz")
  ' username = Request("txtusername") 
  ' password = Replace(Request.Form("password"), "'", "''")
   
	Set ConnessioneDB = Server.CreateObject("ADODB.Connection") 
			 
			
			%>   
			   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
              
                <!-- #include file="../upload_resize/resizecheck.asp" -->
				<!-- #include file = "../var_globali.inc" -->
   				<!-- #include file = "../service/controllo_sessione.asp" -->
			<%  
		
	
       %>   
   
  
<%  
                           
    QuerySQL ="UPDATE Allievi  SET  Cognome ='" &cognome& "', Nome ='" &nome& "',Mipiace ='" &mipiace& "', Nonmipiace ='" &nonmipiace& "', Descriviti ='" &Descriviti& "', Tags='"&gruppo&"', In_Quiz="&inquiz&"  WHERE CodiceAllievo= '"&cod& "';"	
	response.write(QuerySQL)
	ConnessioneDB.Execute(QuerySQL)
	 
	
	

   %>
	</font>   
	 
		
  			<p><p>
			
<%		<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
 if Request.ServerVariables("HTTP_REFERER") <>"" then 
							response.Redirect request.serverVariables("HTTP_REFERER") 
		 end if 
	 
    %>
	</body>
	</html>
	